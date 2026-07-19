"""Convert XLSX to PDF using Microsoft Excel COM ExportAsFixedFormat.

Matches the original ConMas pipeline: Workbook.ExportAsFixedFormat(xlTypePDF)

Phase X.38: When the global RenderQueue is active, xlsx_to_pdf() uses the
queue's persistent Excel.Application instead of creating its own.  This
ensures all COM operations are serialized through a single worker thread.
"""

import os
import sys
import tempfile
from pathlib import Path


def xlsx_to_pdf(xlsx_path: str, output_dir: str | None = None) -> str:
    """Generate PDF from XLSX via Excel COM ExportAsFixedFormat.

    This matches the original ConMas coordinate capture pipeline exactly.
    Excel COM ensures the workbook's declared page size is respected
    (unlike LibreOffice which may use the system printer default).

    Phase X.38: If the global RenderQueue has a persistent Excel instance,
    this function uses it (no new Excel created).  Otherwise falls back
    to creating its own Excel.Application (original behavior).

    Args:
        xlsx_path: Path to the source .xlsx file.
        output_dir: Output directory for the PDF.  If omitted, a temp dir
            is created (caller is responsible for cleanup).

    Returns:
        Path to the generated PDF.

    Raises:
        RuntimeError: If Excel is not installed or the export fails.
        FileNotFoundError: If the PDF was not created.
    """
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        raise RuntimeError(
            "win32com (pywin32) not installed.  "
            "Run: pip install pywin32"
        )

    if output_dir is None:
        output_dir = tempfile.mkdtemp(prefix="ple_pdf_")
    os.makedirs(output_dir, exist_ok=True)

    stem = Path(xlsx_path).stem
    pdf_path = os.path.join(output_dir, f"{stem}.pdf")

    # ── Phase X.38: Use queue's persistent Excel if available ──────
    from render_service.render_queue import get_queue_excel
    queue_excel = get_queue_excel()

    if queue_excel is not None:
        # Use the queue's Excel — no CoInitialize/Quit needed
        # The queue worker owns the Excel lifecycle.
        try:
            wb = queue_excel.Workbooks.Open(os.path.abspath(xlsx_path))
            try:
                wb.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
            finally:
                wb.Close(False)
        except Exception as e:
            raise RuntimeError(f"Excel COM PDF export failed: {e}") from e

        if not os.path.isfile(pdf_path):
            pdfs = [f for f in os.listdir(output_dir) if f.endswith(".pdf")]
            if pdfs:
                pdf_path = os.path.join(output_dir, pdfs[0])
            else:
                raise FileNotFoundError(
                    f"PDF was not created by Excel COM in {output_dir}"
                )
        return pdf_path

    # ── Original behavior: create own Excel ────────────────────────
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = False
        try:
            wb = excel.Workbooks.Open(os.path.abspath(xlsx_path))
            # xlTypePDF = 0, xlQualityStandard = 0
            wb.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
            wb.Close(False)
        except Exception as e:
            raise RuntimeError(f"Excel COM PDF export failed: {e}") from e
        finally:
            excel.Quit()
    finally:
        pythoncom.CoUninitialize()

    if not os.path.isfile(pdf_path):
        # Try to find any PDF created in the output dir
        pdfs = [f for f in os.listdir(output_dir) if f.endswith(".pdf")]
        if pdfs:
            pdf_path = os.path.join(output_dir, pdfs[0])
        else:
            raise FileNotFoundError(
                f"PDF was not created by Excel COM in {output_dir}"
            )

    return pdf_path
