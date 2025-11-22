""" 
docx_utils/docx_operations.py 

Module for checking DOCX documents for font compliance. 
Highlights runs with incorrect fonts in red and returns the total number of discrepancies. 
"""

from typing import List, Tuple, Dict
from docx.document import Document as DocumentObject
from docx.text.run import Run
from docx import Document
from docx_utils.font_check import check_paragraph_font, check_table_font
from config.config import ReportItem
from docx_utils.alignment_check import check_alignment

def analyze_docx(docx_path: str) -> Tuple[List[ReportItem], DocumentObject]:
    """
    Checks a DOCX file for font compliance with TARGET_FONT.
    Highlights runs with incorrect fonts in red in memory only.
    
    Args:
        docx_path (str): Path to the original DOCX file.
    
    Returns:
        Tuple[List[ReportItem], DocumentObject]: A tuple containing a list of runs with font discrepancies and the loaded Document object with highlights.
    """
    docx: DocumentObject = Document(docx_path)
    report: List[ReportItem] = []

    # Check all paragraphs
    for paragraph in docx.paragraphs:
        check_paragraph_font(paragraph, report)

    # Check all tables
    for table in docx.tables:
        check_table_font(table, report)

    check_alignment(docx, report)

    return report, docx  # Return doc object for optional saving


def save_docx(docx: DocumentObject, output_path: str):
    """
    Saves a DOCX Document object to the specified file.

    Args:
        docx (DocumentObject): The Document object to save.
        output_path (str): Path to save the DOCX file.
    """
    docx.save(output_path)
