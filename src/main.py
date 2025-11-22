"""
main.py

This script allows the user to select a .docx document, checks it for font inconsistencies,
highlights the problematic text in red, and saves a new copy of the document.
Additionally, it analyzes pages for potential issues and reports them.
"""

import os
import tkinter as tk
from tkinter import filedialog
from docx_utils.docx_operations import check_docx_file
from page_analysis.page_check import count_issues_by_page


def main():
    # Initialize Tkinter and hide the main window
    root = tk.Tk()
    root.withdraw()

    # -----------------------------
    # Select the DOCX file
    # -----------------------------
    print("Please select a .docx file...")
    filename = filedialog.askopenfilename(
        title="Select a Word document",
        filetypes=[("Word Documents", "*.docx")]
    )
    if not filename:
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
    base_name = os.path.splitext(os.path.basename(filename))[0]
    new_filename = os.path.join(save_dir, f"{base_name}_checked.docx")

    # -----------------------------
    # Check the document and save a new copy
    # -----------------------------
    total_issues, out_file = check_docx_file(filename, output_file=new_filename)
    print(f"\nChecked file saved as: {out_file}")
    print(f"Total font inconsistencies found: {total_issues}")

    # -----------------------------
    # Analyze pages for issues
    # -----------------------------
    if total_issues > 0:
        print("All problematic text has been highlighted in red.\n")
        pages_with_issues = count_issues_by_page(out_file)
        for page_num in pages_with_issues:
            print(f"Issue detected on page {page_num}")
    else:
        print("The document fully conforms to the Times New Roman font.")


if __name__ == "__main__":
    main()
