"""
docx_utils/font_check.py

Module for checking fonts and font sizes in a DOCX document.
Highlights runs with incorrect fonts/sizes in red and records discrepancies
as tuples (run, reason) in the report list.
"""

from typing import List, Optional
from docx.shared import RGBColor, Pt
from config.config import TARGET_FONT, ReportItem, MIN_FONT_SIZE, MAX_FONT_SIZE
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.text.run import Run


def highlight_run(run: Run, paragraph: Paragraph, report: List[ReportItem], reason: str) -> None:
    """
    Highlight the given run in red and append a reasoned dictionary to report.

    Args:
        run: docx.text.run.Run object to highlight.
        paragraph: the Paragraph object containing this run (for context)
        report: list that collects dicts with run, paragraph_text, and reason.
        reason: short textual description why this run is considered incorrect.
    """
    run.font.color.rgb = RGBColor(255, 0, 0)
    report.append({
        "run": run,
        "paragraph_text": paragraph.text,
        "reason": reason
    })


def check_run_style(run: Run, paragraph: Paragraph, report: List[ReportItem]) -> None:
    """
    Check a single run for font family and font size rules and highlight it
    if any rule is violated. Appends (run, reason) to report for each violation.

    Rules:
      - Font size must be between 12 and 14 pt (inclusive)
      - Font family must equal TARGET_FONT when explicitly set
    """
    font = run.font

    # Font family check
    font_name: Optional[str] = font.name
    if font_name is not None and font_name != TARGET_FONT:
        highlight_run(run, paragraph, report,
                      f"Wrong font family: {font_name} (expected {TARGET_FONT})")

    # Font size check
    size: Optional[Pt] = font.size
    if size is None:
        return  # No explicit size (inherits style) — treat as OK

    size_pt = size.pt
    if not (MIN_FONT_SIZE <= size_pt <= MAX_FONT_SIZE):
        highlight_run(run, paragraph, report,
                      f"Text should be {MIN_FONT_SIZE}-{MAX_FONT_SIZE} pt (found {size_pt} pt)")


def check_paragraph_font(paragraph: Paragraph, report: List[ReportItem]) -> None:
    """
    Check all runs in a paragraph for font-family and size issues.
    """
    for run in paragraph.runs:
        check_run_style(run, paragraph, report)


def check_table_font(table: Table, report: List[ReportItem]) -> None:
    """
    Check all paragraphs inside all cells of a table.
    """
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                check_paragraph_font(paragraph, report)
