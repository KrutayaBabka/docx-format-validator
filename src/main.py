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
from page_analysis.page_check import count_issues_by_page
from pathlib import Path


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
    report: List[Run]
    docx: DocumentObject
    report, docx = analyze_docx(docx_path)

    total_issues: int = len(report)

    # Save a copy
    save_docx(docx, checked_doc_path)
    print(f"\nChecked file saved as: {checked_doc_path}")
    print(f"Total font inconsistencies found: {total_issues}")

    # -----------------------------
    # Analyze pages for issues
    # -----------------------------
    if total_issues > 0:
        print("All problematic text has been highlighted in red.\n")
        pages_with_issues: List[int] = count_issues_by_page(checked_doc_path)
        for page_num in pages_with_issues:
            print(f"Issue detected on page {page_num}")
    else:
        print("The document fully conforms to the Times New Roman font.")


if __name__ == "__main__":
    main()
