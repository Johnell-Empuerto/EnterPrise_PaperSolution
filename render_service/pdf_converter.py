"""Convert XLSX to PDF using Microsoft Excel COM ExportAsFixedFormat.

Matches the original ConMas pipeline: Workbook.ExportAsFixedFormat(xlTypePDF)
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
