"""
docx_operations.py

Module for checking DOCX documents for font compliance.
Highlights runs with incorrect fonts in red and returns the total number of discrepancies.
"""

from docx import Document
from docx_utils.font_check import check_paragraph_font, check_table_font


def check_docx_file(input_file: str, output_file: str):
    """
    Checks a DOCX file for font compliance with TARGET_FONT.
    Highlights runs with incorrect fonts in red and saves a copy to output_file.

    Args:
        input_file (str): Path to the original DOCX file.
        output_file (str): Path to save the modified DOCX file.

    Returns:
        tuple: (total_discrepancies: int, output_file: str)
    """
    doc = Document(input_file)
    report = []

    # Check all paragraphs
    for paragraph in doc.paragraphs:
        check_paragraph_font(paragraph, report)

    # Check all tables
    for table in doc.tables:
        check_table_font(table, report)

    # Save the modified document
    doc.save(output_file)

    return len(report), output_file
