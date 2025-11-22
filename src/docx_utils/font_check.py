"""
font_check.py

Module for checking fonts in a DOCX document.
Highlights runs with incorrect fonts in red and counts discrepancies.
"""

from typing import List, Optional
from docx.shared import RGBColor
from config.config import TARGET_FONT
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.text.run import Run

# -----------------------------
# Highlight a run if its font is different from TARGET_FONT
# -----------------------------
def highlight_run_if_wrong_font(run: Run, report: List[Run]) -> None:
    """
    Checks the font of a run and highlights it in red if it doesn't match TARGET_FONT.
    
    Args:
        run: docx.text.run.Run object.
        report: list to store runs with font discrepancies.
    """
    font: Optional[str] = run.font.name
    if font is None:
        # None indicates the run uses the default style, considered OK
        return
    if font != TARGET_FONT:
        report.append(run)  # Store the Run object itself
        run.font.color.rgb = RGBColor(255, 0, 0)  # Highlight in red


# -----------------------------
# Check all runs in a paragraph
# -----------------------------
def check_paragraph_font(paragraph: Paragraph, report: List[Run]) -> None:
    """
    Checks each run in a paragraph for font discrepancies.
    
    Args:
        paragraph: docx.text.paragraph.Paragraph object.
        report: list to store counts of discrepancies.
    """
    for run in paragraph.runs:
        highlight_run_if_wrong_font(run, report)


# -----------------------------
# Check all paragraphs in a table
# -----------------------------
def check_table_font(table: Table, report: List[Run]) -> None:
    """
    Checks all cells and paragraphs in a table for font discrepancies.
    
    Args:
        table: docx.table.Table object.
        report: list to store counts of discrepancies.
    """
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                check_paragraph_font(paragraph, report)
