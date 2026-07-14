"""Render PDF first page to high-resolution PNG using PyMuPDF."""

import os
import fitz  # PyMuPDF


def pdf_page_to_png(
    pdf_path: str,
    output_path: str,
    dpi: int = 300,
    page_index: int = 0,
) -> str:
    doc = fitz.open(pdf_path)

    if page_index >= len(doc):
        raise IndexError(
            f"PDF has {len(doc)} pages, requested page {page_index}"
        )

    page = doc[page_index]
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    pix.save(output_path)
    doc.close()

    return output_path


def get_page_dimensions(
    pdf_path: str,
    dpi: int = 300,
    page_index: int = 0,
) -> tuple[int, int]:
    doc = fitz.open(pdf_path)
    if page_index >= len(doc):
        doc.close()
        raise IndexError(
            f"PDF has {len(doc)} pages, requested page {page_index}"
        )

    page = doc[page_index]
    zoom = dpi / 72.0
    w_pt = page.rect.width
    h_pt = page.rect.height
    w_px = int(round(w_pt * zoom))
    h_px = int(round(h_pt * zoom))
    doc.close()

    return w_px, h_px
