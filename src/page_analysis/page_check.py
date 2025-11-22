"""
page_check.py

Module for analyzing DOCX documents page by page.
Converts DOCX to PDF and identifies pages with potential issues.
"""

import os
from docx2pdf import convert
import fitz  # PyMuPDF


def count_issues_by_page(docx_file: str) -> list[int]:
    """
    Counts pages in a DOCX document that may contain font issues.
    The method converts the DOCX to a temporary PDF and checks
    for pages containing any text (as a proxy for potential issues).

    Args:
        docx_file (str): Path to the DOCX file to check.

    Returns:
        list[int]: Sorted list of page numbers with potential issues.
    """
    temp_pdf = "temp.pdf"
    
    # Convert DOCX to PDF
    convert(docx_file, temp_pdf)
    
    doc_pdf = fitz.open(temp_pdf)
    pages_with_issues = set()

    # Simple heuristic: any page containing text is considered
    # as potentially having issues (can be improved with more checks)
    for page_number, page in enumerate(doc_pdf, start=1):
        if page.get_text().strip():
            pages_with_issues.add(page_number)

    doc_pdf.close()
    os.remove(temp_pdf)

    return sorted(list(pages_with_issues))
