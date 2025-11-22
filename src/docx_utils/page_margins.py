"""
docx_utils/page_margins.py

Module for checking page margins in a DOCX document.
Adds entries to the report if margins do not match the required values.
"""

from typing import List
from docx.document import Document as DocumentObject
from config.config import ReportItem
from docx.shared import Cm
from config.config import TOP_MARGIN_CM, BOTTOM_MARGIN_CM, LEFT_MARGIN_CM, RIGHT_MARGIN_CM

def check_page_margins(docx: DocumentObject, report: List[ReportItem]) -> None:
    """
    Check that all sections of the document have the required page margins.

    Args:
        docx: Document object.
        report: List to append any margin inconsistencies.
    """
    for i, section in enumerate(docx.sections):
        if round(section.top_margin.cm, 2) != TOP_MARGIN_CM:
            report.append({
                "run": None,
                "paragraph_text": f"Section {i+1}",
                "reason": f"Top margin should be {TOP_MARGIN_CM} cm (found {section.top_margin.cm:.2f} cm)"
            })
        if round(section.bottom_margin.cm, 2) != BOTTOM_MARGIN_CM:
            report.append({
                "run": None,
                "paragraph_text": f"Section {i+1}",
                "reason": f"Bottom margin should be {BOTTOM_MARGIN_CM} cm (found {section.bottom_margin.cm:.2f} cm)"
            })
        if round(section.left_margin.cm, 2) != LEFT_MARGIN_CM:
            report.append({
                "run": None,
                "paragraph_text": f"Section {i+1}",
                "reason": f"Left margin should be {LEFT_MARGIN_CM} cm (found {section.left_margin.cm:.2f} cm)"
            })
        if round(section.right_margin.cm, 2) != RIGHT_MARGIN_CM:
            report.append({
                "run": None,
                "paragraph_text": f"Section {i+1}",
                "reason": f"Right margin should be {RIGHT_MARGIN_CM} cm (found {section.right_margin.cm:.2f} cm)"
            })
