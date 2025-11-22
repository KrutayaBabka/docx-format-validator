"""
main.py

This script allows the user to select a .docx document, checks it for font inconsistencies,
highlights the problematic text in red, and saves a new copy of the document.
Additionally, it analyzes pages for potential issues and reports them.
"""


from docx.document import Document as DocumentObject
from typing import List
from docx.text.run import Run
import tkinter as tk
from tkinter import filedialog
from docx_utils.docx_operations import analyze_docx, save_docx
from pathlib import Path

from config.config import ReportItem


def main():
    # Initialize Tkinter and hide the main window
    root: tk.Tk = tk.Tk()
    root.withdraw()

    # -----------------------------
    # Select the DOCX file
    # -----------------------------
    print("Please select a .docx file...")
    docx_path: str = filedialog.askopenfilename(
        title="Select a Word document",
        filetypes=[("Word Documents", "*.docx")]
    )
    if not docx_path:
        print("No file selected. Exiting...")
        return

    # -----------------------------
    # Select the directory to save the checked file
    # -----------------------------
    print("\nPlease select a folder to save the checked file...")
    save_dir = filedialog.askdirectory(title="Select a folder")
    if not save_dir:
        print("No folder selected. Exiting...")
        return

    # Construct the new filename: original_name + "_checked"
    base_name: str = Path(docx_path).stem
    checked_doc_path: Path = Path(save_dir) / f"{base_name}_checked.docx"

    # -----------------------------
    # Check the document and save a new copy
    # -----------------------------
    # Analyze document
    report: List[ReportItem]
    docx: DocumentObject
    report, docx = analyze_docx(docx_path)

    for issue in report:
        run = issue.get('run')
        run_text = run.text if run else "<No run text>"
        print(f"Issue found: '{run_text}' - {issue['reason']}")
        print(f"Context paragraph: '{issue['paragraph_text']}'\n")

    total_issues: int = len(report)

    # Save a checked copy of the document
    save_docx(docx, checked_doc_path)
    print(f"\n✅ Checked file saved as: {checked_doc_path}")

    # Summary of findings
    print(f"Total issues found: {total_issues}")

    if total_issues > 0:
        print("⚠️ Some formatting issues were found and highlighted in red:")
        print("- Paragraph alignment")
        print("- First-line, left, and right indents")
        print("- Line spacing")
        print("\nPlease review highlighted text and correct formatting as needed.\n")
    else:
        print("🎉 No formatting issues found. The document conforms to the required standards (Times New Roman, 1.5 spacing, correct indents).")


if __name__ == "__main__":
    main()
