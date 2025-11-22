"""
docx_utils/font_fix.py

Module for fixing fonts and font sizes in a DOCX document.
Automatically sets all runs to the TARGET_FONT and a standard font size (e.g., 12pt).
"""

from docx.shared import Pt
from config.config import TARGET_FONT, MIN_FONT_SIZE
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.text.run import Run


def fix_run_style(run: Run) -> None:
    """
    Fix a single run to have the correct font family and size.
    Also resets any previous red highlighting.
    """
    run.font.name = TARGET_FONT
    run.font.size = Pt(MIN_FONT_SIZE)  # Use minimum font size as standard
    run.font.color.rgb = None  # Remove red highlighting if present
    run.bold = False
    run.italic = False
    run.underline = False


def fix_paragraph_font(paragraph: Paragraph) -> None:
    """
    Fix all runs in a paragraph.
    """
    for run in paragraph.runs:
        fix_run_style(run)


def fix_table_font(table: Table) -> None:
    """
    Fix all paragraphs inside all cells of a table.
    """
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                fix_paragraph_font(paragraph)
