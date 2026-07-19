"""
Phase X.38 — Global Render Queue & Excel COM Worker.

Serializes all Excel COM operations through a single background thread,
preventing concurrency failures caused by Excel's Single-Threaded Apartment (STA).

Architecture:
    HTTP Requests
         |
         v
    Global Render Queue  (thread-safe, FIFO)
         |
         v
    Single Worker Thread  (owns Excel.Application)
         |
         v
    Excel COM  -->  PDF Export  -->  Return Response

The worker keeps Excel.Application alive between requests to avoid
re-launch overhead.  Every COM operation runs on this thread only.
"""

import os
import queue
import threading
import time
import uuid
import traceback
from pathlib import Path
from typing import Any, Callable, Optional


# ── Module-level reference to the queue's Excel instance ────────────────
# Set by the worker thread before executing a job.  xlsx_to_pdf() and
# other COM functions check this and use it instead of creating their own
# Excel.Application.  This is the ONLY mechanism for sharing Excel across
# modules without modifying their interfaces.
_queue_excel = None
_queue_excel_lock = threading.Lock()


def get_queue_excel():
    """Return the queue's persistent Excel.Application, or None."""
    with _queue_excel_lock:
        return _queue_excel


def _set_queue_excel(excel):
    with _queue_excel_lock:
        global _queue_excel
        _queue_excel = excel


def _clear_queue_excel():
    with _queue_excel_lock:
        global _queue_excel
        _queue_excel = None


# ── Job definition ──────────────────────────────────────────────────────

class _RenderJob:
    """A unit of work for the render queue."""

    __slots__ = (
        "id", "fn", "args", "kwargs",
        "result", "error", "event",
        "queued_at", "started_at", "completed_at",
    )

    def __init__(self, fn: Callable, args: tuple, kwargs: dict):
        self.id = uuid.uuid4().hex[:8]
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.result: Any = None
        self.error: Optional[str] = None
        self.event = threading.Event()
        self.queued_at = time.time()
        self.started_at: Optional[float] = None
        self.completed_at: Optional[float] = None


# ── Queue metrics ───────────────────────────────────────────────────────

class _QueueMetrics:
    __slots__ = (
        "total_submitted", "total_completed", "total_failed",
        "total_timeouts", "total_wait_ms", "total_render_ms",
        "active_request_id", "active_start_time",
        "excel_running", "worker_alive",
    )

    def __init__(self):
        self.total_submitted = 0
        self.total_completed = 0
        self.total_failed = 0
        self.total_timeouts = 0
        self.total_wait_ms = 0.0
        self.total_render_ms = 0.0
        self.active_request_id: Optional[str] = None
        self.active_start_time: Optional[float] = None
        self.excel_running = False
        self.worker_alive = False


# ── Render Queue ────────────────────────────────────────────────────────

class RenderQueue:
    """Global render queue that serializes all Excel COM operations.

    Usage:
        queue = RenderQueue()
        queue.start()

        # In endpoint handler:
        result = queue.submit(my_com_function, arg1, arg2)

        # On shutdown:
        queue.shutdown()
    """

    MAX_RESTARTS = 3
    RESTART_DELAY_S = 2.0

    def __init__(self, max_wait: int = 60):
        self._queue: queue.Queue = queue.Queue()
        self._max_wait = max_wait
        self._excel = None
        self._thread: Optional[threading.Thread] = None
        self._shutdown_event = threading.Event()
        self._metrics = _QueueMetrics()
        self._lock = threading.Lock()
        self._started = False
        self._restart_count = 0

    # ── Lifecycle ──────────────────────────────────────────────────────

    def start(self):
        """Start the worker thread and create the persistent Excel instance."""
        if self._started:
            return
        self._started = True
        self._restart_count = 0
        self._shutdown_event.clear()
        self._thread = threading.Thread(target=self._worker_loop, daemon=True)
        self._thread.start()
        print("[queue] Render queue started")

    def shutdown(self, timeout: float = 5.0):
        """Gracefully shut down the worker and quit Excel."""
        if not self._started:
            return
        print("[queue] Shutting down render queue...")
        self._shutdown_event.set()
        # Drain the queue so no jobs get stuck
        while not self._queue.empty():
            try:
                self._queue.get_nowait()
            except queue.Empty:
                break
        # Push sentinel to wake the worker
        self._queue.put(None)
        if self._thread:
            self._thread.join(timeout=timeout)
        self._cleanup_excel()
        self._started = False
        self._metrics.worker_alive = False
        print("[queue] Render queue shutdown complete")

    @property
    def is_running(self) -> bool:
        return self._started and self._thread is not None and self._thread.is_alive()

    def restart_worker(self):
        """Restart the worker thread after a crash."""
        with self._lock:
            self._restart_count += 1
            if self._restart_count > self.MAX_RESTARTS:
                print(f"[queue] Max restarts ({self.MAX_RESTARTS}) reached — not restarting")
                return False
        print(f"[queue] Restarting worker (attempt {self._restart_count}/{self.MAX_RESTARTS})...")
        self._cleanup_excel()
        self._thread = threading.Thread(target=self._worker_loop, daemon=True)
        self._thread.start()
        return True

    # ── Submit ─────────────────────────────────────────────────────────

    def submit(self, fn: Callable, *args, **kwargs) -> Any:
        """Submit a function to execute on the worker thread.

        Blocks until complete.  All COM operations within *fn* are safe
        because they execute on the single worker thread.

        Returns:
            The return value of *fn*.

        Raises:
            TimeoutError: If the job takes longer than max_wait seconds.
            RuntimeError: If the job raised an exception.
        """
        job = _RenderJob(fn=fn, args=args, kwargs=kwargs)

        with self._lock:
            self._metrics.total_submitted += 1

        self._queue.put(job)
        print(f"[queue] Job {job.id} queued | size={self._queue.qsize()}")

        if not job.event.wait(timeout=self._max_wait):
            with self._lock:
                self._metrics.total_timeouts += 1
            print(f"[queue] Job {job.id} TIMEOUT after {self._max_wait}s")
            raise TimeoutError(f"Render queue timeout after {self._max_wait}s")

        if job.error:
            raise RuntimeError(job.error)

        return job.result

    # ── Worker ─────────────────────────────────────────────────────────

    def _worker_loop(self):
        """Main worker loop — owns Excel COM for its entire lifetime.

        On fatal errors the worker self-restarts (up to MAX_RESTARTS times).
        """
        try:
            import pythoncom
            import win32com.client
        except ImportError:
            print("[worker] CRITICAL: pywin32 not installed — queue disabled")
            return

        pythoncom.CoInitialize()
        self._metrics.worker_alive = True
        self._metrics.excel_running = False
        print("[worker] Worker started, initializing Excel...")

        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.DisplayAlerts = False
            excel.Visible = False
            self._excel = excel
            self._metrics.excel_running = True
            # Publish Excel reference so xlsx_to_pdf() can use it
            _set_queue_excel(excel)
            print("[worker] Excel initialized and published to queue")

            while not self._shutdown_event.is_set():
                try:
                    job = self._queue.get(timeout=1.0)
                except queue.Empty:
                    continue

                if job is None:  # Sentinel — shutdown
                    print("[worker] Received shutdown sentinel")
                    break

                try:
                    self._process_job(job, excel)
                except Exception as e:
                    # Per-job exception is already caught inside _process_job,
                    # but be defensive — if a job raises unexpectedly, log and
                    # mark it failed rather than crashing the entire worker.
                    print(f"[worker] Unhandled job exception: {e}")
                    traceback.print_exc()
                    if not job.event.is_set():
                        job.error = str(e)
                        job.event.set()

        except Exception as e:
            print(f"[worker] Fatal worker error: {e}")
            traceback.print_exc()
            # Self-restart (new thread, fresh Excel)
            print(f"[worker] Attempting restart in {self.RESTART_DELAY_S}s...")
            time.sleep(self.RESTART_DELAY_S)
            if not self._shutdown_event.is_set():
                self.restart_worker()
        finally:
            _clear_queue_excel()
            self._cleanup_excel()
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            self._metrics.worker_alive = False
            print("[worker] Worker stopped")

    def _process_job(self, job: _RenderJob, excel):
        """Execute a single job on the worker thread."""
        job.started_at = time.time()

        with self._lock:
            self._metrics.active_request_id = job.id
            self._metrics.active_start_time = job.started_at

        fn_name = getattr(job.fn, "__name__", str(job.fn))
        print(f"[worker] Processing {fn_name} | job={job.id}")

        try:
            job.result = job.fn(*job.args, **job.kwargs)
            job.completed_at = time.time()
            render_ms = (job.completed_at - job.started_at) * 1000
            wait_ms = (job.started_at - job.queued_at) * 1000

            with self._lock:
                self._metrics.total_completed += 1
                self._metrics.total_render_ms += render_ms
                self._metrics.total_wait_ms += wait_ms

            print(f"[worker] Completed {fn_name} | job={job.id} | "
                  f"wait={wait_ms:.0f}ms | render={render_ms:.0f}ms")

        except Exception as e:
            job.error = str(e)
            job.completed_at = time.time()
            with self._lock:
                self._metrics.total_failed += 1
            print(f"[worker] FAILED {fn_name} | job={job.id}: {e}")
            traceback.print_exc()
        finally:
            job.event.set()
            with self._lock:
                self._metrics.active_request_id = None
                self._metrics.active_start_time = None

    # ── Excel Cleanup ──────────────────────────────────────────────────

    def _cleanup_excel(self):
        """Quit Excel and release COM resources."""
        if self._excel is not None:
            try:
                self._excel.Quit()
                print("[worker] Excel quit")
            except Exception:
                pass
            self._excel = None
            self._metrics.excel_running = False

    # ── Metrics ────────────────────────────────────────────────────────

    def get_metrics(self) -> dict:
        """Return current queue metrics for the /health endpoint."""
        with self._lock:
            avg_wait = 0.0
            avg_render = 0.0
            if self._metrics.total_completed > 0:
                avg_wait = self._metrics.total_wait_ms / self._metrics.total_completed
                avg_render = self._metrics.total_render_ms / self._metrics.total_completed

            active_ms = None
            if self._metrics.active_start_time is not None:
                active_ms = round((time.time() - self._metrics.active_start_time) * 1000)

            return {
                "excelRunning": self._metrics.excel_running,
                "workerAlive": self._metrics.worker_alive,
                "queueLength": self._queue.qsize(),
                "activeRequest": self._metrics.active_request_id,
                "activeRequestRunningMs": active_ms,
                "totalSubmitted": self._metrics.total_submitted,
                "totalCompleted": self._metrics.total_completed,
                "totalFailed": self._metrics.total_failed,
                "totalTimeouts": self._metrics.total_timeouts,
                "avgWaitTimeMs": round(avg_wait, 1),
                "avgRenderTimeMs": round(avg_render, 1),
                "restartCount": self._restart_count,
            }


# ── Global singleton ────────────────────────────────────────────────────

_queue_instance: Optional[RenderQueue] = None
_queue_lock = threading.Lock()


def get_queue() -> RenderQueue:
    """Return the global RenderQueue singleton."""
    global _queue_instance
    if _queue_instance is None:
        with _queue_lock:
            if _queue_instance is None:
                _queue_instance = RenderQueue()
    return _queue_instance
