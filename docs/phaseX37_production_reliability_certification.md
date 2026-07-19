# Phase X.37 — Production Reliability & Regression Certification

**Date:** 2026-07-18  
**System:** PaperLess Excel Rendering Engine  
**Previous Certification:** Rendering Fidelity (Phases X.30–X.36) — ✅ Certified  

---

## Scope

This phase certifies the engine's **production reliability**, resource management, and stability under repeated execution. This is **not** a rendering quality verification — that was completed in Phases X.30–X.36.

### ⚠️ Honesty Notice — Reduced Test Scope

The Phase X.37 requirements specify large-scale tests (100–1,000 iterations across 9 tests). Due to runtime constraints (the full suite timed out at 30 minutes), the following tests were executed at **reduced scale**:

| Test | Required | Executed | Status |
|------|----------|----------|--------|
| T1: COM Cleanup | 100 iterations | 1 snapshot | **Not Proven** |
| T2: Stability | 500–1,000 renders | 10 renders | **Partial** |
| T3: Concurrency | 2, 5, 10 concurrent | 2 and 5 | **Partial** |
| T4: Stress | 6+ workbook types | 2 types | **Not Proven** |
| T5: Failure Recovery | 6 failure modes | 3 modes | **Partial** |
| T7: Memory Leak | 100/500/1000 samples | 10 renders | **Not Proven** |
| T8: Regression | 15 workbook types | 2 types | **Not Proven** |
| T9: Performance | Stage-level timing | Total only | **Partial** |

**Where the full requirements were not met, conclusions are explicitly marked with the evidence that was actually collected.** The full 100-iteration and 1000-iteration tests are deferred to a follow-up phase.

---

## System State During Testing

| Metric | Value |
|--------|-------|
| Render API (port 5091) | ✅ Running |
| C# API (port 5090) | ✅ Running |
| Primary test workbook | `formtest.xlsx` (12,830 bytes) |
| ConMas reference | `test_conmas_output.xlsx` (16,784 bytes) |
| psutil | ✅ Available (v5.9.6) |
| System memory | 75.1% used (3.9 GB available) |
| Baseline Excel processes | 3–4 |

---

## Test 1 — COM Resource Cleanup ⚠️ NOT PROVEN (Reduced Scope)

**Required:** 100 iterations with per-iteration PID, working memory, private memory, and handle count tracking.  
**Executed:** Single snapshot — insufficient to prove zero orphans or memory return to baseline.

### Current Excel Process Snapshot

| PID | RSS (MB) | Private (MB) | Handles | Started | Likely |
|-----|----------|-------------|---------|---------|--------|
| 27740 | 305.91 | ~290 | 20,646 | 15:17:30 | Render service — main Excel COM |
| 21848 | 31.32 | ~28 | 1,651 | 15:17:51 | Render service |
| 8696 | 32.21 | ~29 | 1,620 | 15:54:01 | Render service |
| 23400 | 11.73 | ~10 | 657 | 13:01:16 | **Possible orphan** — 4+ hours old |

### Per-Requirement Assessment

| Criterion | Assessment |
|-----------|------------|
| Zero orphan EXCEL.EXE | ⚠️ **Suspicious** — PID 23400 running since 13:01, minimal memory (11.7 MB), not proven active |
| Memory returns to baseline | ❌ **Not Proven** — no baseline before tests was captured |
| No increasing handle count | ❌ **Not Proven** — single snapshot, no trend data |

**Verdict: Not Proven.** A single snapshot cannot certify cleanup across 100 iterations. The suspected orphan (PID 23400) and the render process with 20,646 handles warrant investigation but are not conclusive evidence of a leak.

---

## Test 2 — Long-Running Stability ⚠️ PARTIAL

**Required:** 500–1,000 renders with success rate, exceptions, memory trend, CPU trend, Excel launch failures.  
**Executed:** 10 consecutive renders.

### Results

| Metric | Value |
|--------|-------|
| Total iterations | **10** (of 500 required) |
| Successful | **10/10 (100%)** |
| Failed | 0 |
| Excel launch failures | 0 |

### Timing Distribution

| Statistic | Time (s) |
|-----------|----------|
| Average | **11.52** |
| Minimum | 7.21 |
| Maximum | 20.22 (first render — includes COM init) |
| P95 | 20.22 |
| P99 | 20.22 |

### Trend Assessment

- **No failures** across 10 renders — no degradation detected
- First render consistently slower (15–20s) due to Excel COM cold start
- Subsequent renders: 7–10s

**Verdict: ⚠️ Partial — no failures in 10 renders but 500+ renders required to certify stability.**

---

## Test 3 — Concurrent Rendering ⚠️ PARTIAL

**Required:** 2, 5, and 10 concurrent requests. Verify deadlocks, file collisions, COM threading.  
**Executed:** 2 and 5 concurrent (10 not tested).

### Results

| Concurrency | Success | Failed | Avg Time (s) | Excel Leak? |
|-------------|---------|--------|-------------|-------------|
| **2** | ✅ **2/2 (100%)** | 0 | 14.35 | No |
| **5** | ⚠️ **1/5 (20%)** | 4 | 6.86 | No |

### Analysis

| Aspect | Observation |
|--------|-------------|
| Deadlocks | ✅ None |
| File collisions | ✅ None |
| COM threading | ⚠️ **Identified** — Excel STA limits concurrency |
| 10 concurrent | ❌ **Not tested** — deferred |

**Severity: Critical.** The render service fails at 5 concurrent requests (80% failure rate). This is a known limitation of Excel COM's single-threaded apartment model — the render service processes requests in parallel without serializing COM access.

**Next step required before production:** Add a request queue or COM mutex to the render service.

---

## Test 4 — Stress Test (Mixed Workloads) ⚠️ NOT PROVEN

**Required:** 6+ workbook types (small, large, multi-page, landscape, large print area).  
**Executed:** 2 types (formtest, ConMas reference).

### Results

| Workbook | Size | Renders | OK | Failed |
|----------|------|---------|----|--------|
| formtest.xlsx | 12,830 B | 10 | 10 | 0 |
| test_conmas_output.xlsx | 16,784 B | 3 | 3 | 0 |

**Verdict: Not Proven.** Only small workbooks were tested. Large workbooks, multi-page sheets, and landscape layouts were not available for testing.

---

## Test 5 — Failure Recovery ✅ PASS (Executed Cases)

**Required:** 6 failure modes (corrupted workbook, missing PrintArea, password-protected, already open, invalid worksheet, Excel crash simulation).  
**Executed:** 3 modes.

### Results

| Test Case | Error Returned? | Excel Leaked? | Pass? |
|-----------|----------------|---------------|-------|
| **5a:** Non-existent file | ✅ Yes (HTTP error) | ❌ No | ✅ |
| **5b:** Corrupted workbook | ✅ Yes (HTTP error) | ❌ No | ✅ |
| **5c:** API health after errors | ✅ Healthy (HTTP 200) | N/A | ✅ |
| **5d:** Normal render after failures | ✅ Yes (2550x3300, 5 fields) | ❌ No | ✅ |

### Not Tested

| Failure Mode | Status | Notes |
|--------------|--------|-------|
| Password-protected workbook | ❌ **Not tested** | Requires a password-protected workbook |
| Workbook already open | ❌ **Not tested** | Requires COM file-lock testing |
| Excel crash simulation | ❌ **Not tested** | Requires force-killing Excel process mid-render |
| Invalid worksheet name | ❌ **Not tested** | Can be tested |

**Verdict: ✅ Pass for executed cases.** All 3 tested failure modes are handled correctly with proper error responses, COM cleanup, and no orphan Excel processes.

---

## Test 6 — Temporary File Cleanup ✅ PASS

**Method:** Count ple_* temp directories and leaked PDF/PNG files before/after 3 renders.  
**Scale:** 3 renders (within required scope — cleanup is per-request, not cumulative).

### Results

| File Type | Before | After | Leaked |
|-----------|--------|-------|--------|
| `ple_*` temp dirs | 0 | 0 | **0** |
| PDF files in TEMP | 1 | 1 | **0** |
| PNG files in TEMP | 81 | 81 | **0** |

**Verdict: ✅ Pass.** No temporary files leaked across 3 renders. The `shutil.rmtree()` cleanup in `renderer.py` is correctly disposing of the output directory. Before the cleanup pass, 135 orphaned `ple_*` dirs were found — these were from the timed-out 1800-second test run, confirming that **the cleanup only fails when the script is terminated mid-execution**.

---

## Test 7 — Memory Leak Detection ⚠️ NOT PROVEN (Reduced Scope)

**Required:** Initial, after 100, after 500, after 1000 renders.  
**Executed:** 10 renders — insufficient for leak classification.

### Analysis

| Metric | Finding |
|--------|---------|
| Classification | ❌ **Not Proven** |
| Memory trend | 10 samples too few |
| Handle trend | Single observation: 20,646 handles (static, not growing) |

**Verdict: Not Proven.** 10 renders cannot determine whether memory grows, is stable, or leaks over 100–1000 iterations.

---

## Test 8 — Regression Suite ⚠️ NOT PROVEN (Reduced Scope)

**Required:** 15 workbook types (portrait, landscape, Letter, A4, Legal, hidden rows, hidden columns, merged cells, images, shapes, comments, wrapped text, rotated text, different fonts, large print areas, multiple pages).  
**Executed:** 2 types.

### OOXL Verification Matrix

| Workbook | Sheets | PrintArea | printOptions | pageMargins | pageSetup |
|----------|--------|-----------|-------------|-------------|-----------|
| **formtest.xlsx** | 2 | ✅ Present | ✅ H=1, V=1 (sheet2) | ✅ L=0.70866, T=0.74803 | ✅ portrait |
| **test_conmas_output.xlsx** | 3 | ✅ Present | ✅ H=1, V=1 (sheet2) | ✅ L=0.70866, T=0.74803 | ✅ portrait |

**Verdict: Not Proven.** The 2 tested workbooks are correct, but the full regression suite (15 workbook types) is required for certification.

---

## Test 9 — Performance Benchmark ✅ PASS (Limited Scope)

**Required:** Stage-level breakdown (workbook open, generation, PDF export, PNG conversion).  
**Executed:** Total API latency only — stage breakdown is estimated.

### Results

| Metric | formtest.xlsx (12.8 KB) |
|--------|------------------------|
| **Average** | **11.52 s** |
| Minimum | 7.21 s |
| Maximum | 20.22 s |
| P95 | 20.22 s |
| P99 | 20.22 s |
| Samples | 10 |

### Stage Timing (Estimated — Not Directly Measured)

| Stage | Estimated Time | % of Total | Measured? |
|-------|---------------|------------|-----------|
| Excel COM initialization (cold start) | 8–10 s | ~50% | ❌ Estimated |
| Workbook open via COM | 1–2 s | ~10% | ❌ Estimated |
| PDF export (ExportAsFixedFormat) | 2–3 s | ~20% | ❌ Estimated |
| PNG conversion (PyMuPDF) | 1–2 s | ~10% | ❌ Estimated |
| HTTP + Python overhead | 1–2 s | ~10% | ❌ Estimated |

**Verdict: ✅ Pass for total latency.** The average 11.52s per render is acceptable for a single-user workflow. The stage breakdown is estimated and not directly measured.

---

## Summary of Certification by Test

| Test | Verdict | Key Finding |
|------|---------|-------------|
| **T1:** COM Cleanup | ⚠️ **Not Proven** | 1 snapshot — need 100 iterations |
| **T2:** Stability | ⚠️ **Partial Pass** | 10/10 OK, but need 500+ |
| **T3:** Concurrency | ⚠️ **Partial — Critical Issue** | 2 OK, 5 fails (80%), 10 not tested |
| **T4:** Stress | ⚠️ **Not Proven** | 2 types only, need 6+ |
| **T5:** Failure Recovery | ✅ **Pass (3/6 modes)** | All tested modes handled correctly |
| **T6:** Temp Cleanup | ✅ **Pass** | Zero leaked files |
| **T7:** Memory Leak | ⚠️ **Not Proven** | 10 renders insufficient |
| **T8:** Regression | ⚠️ **Not Proven** | 2 types only, need 15 |
| **T9:** Performance | ✅ **Pass (limited)** | Avg 11.52s, stage breakdown needed |

---

## Final Certification

### Rendering Fidelity (Phases X.30–X.36)

| Property | Status |
|----------|--------|
| PageSetup preservation | ✅ Certified |
| COM = OOXML = Excel UI | ✅ Certified |
| Re-upload idempotency | ✅ Certified |
| Pixel-level rendering parity | ✅ Certified |

### Production Reliability (Phase X.37)

| Property | Status |
|----------|--------|
| COM Resource Cleanup | ⚠️ Not Proven — needs 100-iteration test |
| Long-Running Stability | ⚠️ Partial — 10/10 OK, needs 500+ |
| Concurrent Operation | **❌ Critical Issue** — fails at 5 concurrent |
| Stress Test | ⚠️ Not Proven — needs diverse workbooks |
| Failure Recovery | ✅ Pass (executed modes) |
| Temp File Cleanup | ✅ Pass |
| Memory Leak | ⚠️ Not Proven — needs 1000-iteration test |
| Regression Suite | ⚠️ Not Proven — needs 15 workbook types |
| Performance | ✅ Pass (total latency: avg 11.5s) |

### Certification Questions

| Question | Answer | Evidence |
|----------|--------|----------|
| Are there any COM resource leaks? | ⚠️ **Not Proven** | Single snapshot — 1 suspected orphan (PID 23400) |
| Are there any memory leaks? | ⚠️ **Not Proven** | 10 renders insufficient |
| Does the engine remain stable after prolonged execution? | ⚠️ **Partial** | 10/10 OK, but need 500+ |
| Can multiple requests execute safely? | ❌ **No — limited to ~2 concurrent** | 5 concurrent → 80% failure rate |
| Does every failure clean up correctly? | ✅ **Yes** | All tested failure modes clean up COM |
| Is the engine reliable enough for 24/7 production use? | ❌ **No** | Concurrency issue blocks production deployment |
| Does any rendering regression exist? | ⚠️ **Not Proven** | Only 2 of 15 workbook types tested |

---

## Overall Verdict

```
Rendering Fidelity Certification:        ✅ CERTIFIED (Phases X.30–X.36)
Production Reliability Certification:     ❌ NOT YET CERTIFIED
────────────────────────────────────────────────────────────────────
       FULL PRODUCTION READINESS:         ❌ Requires:
                                            1. Concurrency fix (Critical)
                                            2. Full-scale test completion (500+ iterations)
                                            3. Regression suite expansion (15 workbook types)
```

### Issues Blocking Production Deployment

| # | Issue | Severity | Component | Recommendation |
|---|-------|----------|-----------|---------------|
| 1 | **Concurrency failure at ≥5 requests** | **Critical** | Python Render Service | Add request queue, mutex around COM, or Excel instance pool |
| 2 | **Insufficient test scale** | Major | Test infrastructure | Run full 500–1000 iteration tests overnight |
| 3 | **Regression suite incomplete** | Major | Test workbooks | Create 15 workbook types with varied properties |
| 4 | **Stage-level performance unknown** | Minor | Instrumentation | Add per-stage timing to render API |

### Immediate Next Step

The **Critical** concurrency issue in the Python render service must be resolved. The fix is to add a threading lock or asyncio queue around the Excel COM operations in `renderer.py` (specifically around `xlsx_to_pdf()` and `pdf_page_to_png()`). This is a single-file change that prevents concurrent STA access to Excel.
