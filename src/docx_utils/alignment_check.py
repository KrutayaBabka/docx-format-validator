""" 
docx_utils/alignment_check.py

Module for checking paragraph alignment in a DOCX document.
Highlights runs with incorrect alignment in red and records discrepancies
as tuples (run, reason) in the report list.
"""

from typing import List
from docx.shared import RGBColor
from docx.text.paragraph import Paragraph
from docx.document import Document as DocumentObject
from config.config import ReportItem
import re
from docx.enum.text import WD_ALIGN_PARAGRAPH


def highlight_alignment(paragraph: Paragraph, report: List[ReportItem], reason: str) -> None:
    """
    Highlight the entire paragraph in red and append a reasoned dictionary to report.

    Args:
        paragraph: the Paragraph object to highlight.
        report: list that collects dicts with run, paragraph_text, and reason.
        reason: short textual description why this paragraph is considered incorrect.
    """
    for run in paragraph.runs:
        run.font.color.rgb = RGBColor(255, 0, 0)
    report.append({
        "run": paragraph.runs[0] if paragraph.runs else None,
        "paragraph_text": paragraph.text,
        "reason": reason
    })


def is_title_page(paragraphs: List[Paragraph], index: int) -> bool:
    """
    Determine if a paragraph belongs to the title page based on 'Moscow <year> g.' pattern.

    Args:
        paragraphs: list of all paragraphs in the document.
        index: index of the current paragraph.
    Returns:
        True if the paragraph belongs to the title page.
    """
    text = paragraphs[index].text.strip()
    return bool(re.search(r"Москва\s+\d{4}\s+г\.", text))


def is_image_caption(paragraph: Paragraph) -> bool:
    """
    Check if the paragraph is a caption under an image (starts with 'Рис.').

    Args:
        paragraph: the Paragraph object to check.
    Returns:
        True if paragraph is an image caption.
    """
    text = paragraph.text.strip()
    return bool(re.match(r"^Рис\.\s*\d*", text))


def is_table_caption(paragraph: Paragraph) -> bool:
    """
    Check if the paragraph is a caption above a table (starts with 'Табл.').

    Args:
        paragraph: the Paragraph object to check.
    Returns:
        True if paragraph is a table caption.
    """
    text = paragraph.text.strip()
    return bool(re.match(r"^Табл\.\s*\d*", text))


def check_alignment(docx: DocumentObject, report: List[ReportItem]) -> None:
    """
    Check all paragraphs in a document for correct alignment.

    Rules:
      - Skip the title page (up to 'Moscow <year> g.')
      - Image captions must be center-aligned
      - Table captions must be right-aligned
      - All other text must be justified

    Args:
        docx: loaded Document object.
        report: list to store alignment issues.
    """
    paragraphs = docx.paragraphs
    skip_until_index = -1

    # Detect the title page
    for i, paragraph in enumerate(paragraphs):
        if is_title_page(paragraphs, i):
            skip_until_index = i
            break

    for i, paragraph in enumerate(paragraphs):
        if i <= skip_until_index:
            continue  # skip title page

        text = paragraph.text.strip()
        if not text:
            continue  # skip empty paragraphs

        # Image captions
        if is_image_caption(paragraph):
            if paragraph.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                highlight_alignment(paragraph, report, "Caption under image should be center aligned")
            continue

        # Table captions
        if is_table_caption(paragraph):
            if paragraph.alignment != WD_ALIGN_PARAGRAPH.RIGHT:
                highlight_alignment(paragraph, report, "Caption above table should be right aligned")
            continue

        # Normal text
        if paragraph.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
            highlight_alignment(paragraph, report, "Normal text should be justified")
