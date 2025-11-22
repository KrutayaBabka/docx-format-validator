"""
docx_utils/page_margins_fix.py

Module for fixing page margins in a DOCX document.
Automatically sets all sections to the required margins.
"""

from docx.document import Document as DocumentObject
from config.config import TOP_MARGIN_CM, BOTTOM_MARGIN_CM, LEFT_MARGIN_CM, RIGHT_MARGIN_CM
from docx.shared import Cm

def fix_page_margins(docx: DocumentObject) -> None:
    """
    Fix all sections of the document to have the required page margins.
    
    Args:
        docx: Document object to modify.
    """
    for section in docx.sections:
        section.top_margin = Cm(TOP_MARGIN_CM)
        section.bottom_margin = Cm(BOTTOM_MARGIN_CM)
        section.left_margin = Cm(LEFT_MARGIN_CM)
        section.right_margin = Cm(RIGHT_MARGIN_CM)
