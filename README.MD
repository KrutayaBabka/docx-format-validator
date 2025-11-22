# DOCX Format Validator & Auto-Fixer

This project analyzes **DOCX documents** for formatting issues and automatically fixes many of them according to common academic and technical formatting standards (e.g., typical requirements for reports and coursework).

The tool highlights incorrect formatting, generates a detailed report, and optionally applies automatic corrections.

## Features

### ✔ **Checks formatting:**

* Font family and size
* Paragraph alignment
* First-line indent, left/right indentation
* Line spacing
* Page margins
* Captions for images and tables (e.g., `Рис. 1. ...`, `Табл. 2. ...`)
* Formatting in tables

### ✔ **Automatically fixes:**

* Alignment (justify, center for images, right for table captions)
* Indentation (including special rules for captions)
* Line spacing
* Font family and size
* Page margins
* Caption formatting (removes bold/italic/underline)


### ✔ **Generates a detailed report**

The program produces two independent diagnostic outputs:

* **A DOCX copy with highlighted issues**
  Every formatting error is visually marked (red-highlighted runs or paragraphs), allowing you to review problems directly inside the document.

* **A `report.txt` file**
  Contains a structured list of all detected issues, including:

  * the fragment where the problem occurred,
  * the type of violation,
  * a clear explanation of what exactly is wrong.

These two files together provide both a human-readable overview and a precise technical description of all formatting inconsistencies.


## ⚠ Limitations

The program does **not** process:

* **Table of Contents** — Word-generated structure that should not be modified programmatically
* **Images floating outside text flow** (floating images are ignored)
* **Punctuation or grammar issues**
* **Semantic correctness of captions or headings**
* **Complex nested structures in headers/footers**

Additionally:

> **Even with auto-fixing enabled, it's strongly recommended to manually review the final document.**
> Rare edge cases or unusual DOCX structures may cause unexpected behavior.

---

## Project Structure

```
docx-format-validator/
├── src/
│   ├── docx_utils/
│   │   ├── auto_fix/
│   │   │   ├── alignment_fix.py
│   │   │   ├── font_fix.py
│   │   │   ├── page_margins_fix.py
│   │   ├── alignment_check.py
│   │   ├── font_check.py
│   │   ├── page_margins.py
│   │   └── docx_operations.py
│   ├── main.py
├── requirements.txt
├──.gitignore
└── README.md
```

## Installation

1. **Clone the repository:**

```bash
git clone https://github.com/KrutayaBabka/docx-format-validator.git
cd docx-format-validator
```

2. **Create a virtual environment:**

```bash
python -m venv venv
```

3. **Activate the environment:**

```bash
# Windows
venv\Scripts\activate

# Linux / Mac
source venv/bin/activate
```

4. **Install dependencies:**

```bash
pip install -r requirements.txt
```

## Usage

Run the main script:

```bash
python src/main.py
```

The program will ask you to **select a DOCX file**, and then it will also ask you to **select a folder where all output files will be saved**.

The application will then:

1. Analyze the document formatting
2. Generate a detailed report (`report.txt`)
3. Create a version with **highlighted issues** → `*_checked.docx`
4. Optionally apply automatic fixes
5. Save the fully corrected version → `*_fixed.docx`

## Output Files

After processing, the selected output folder will contain:

* **`report.txt`** — a readable list of all issues found
* **`<original_name>_checked.docx`** — document with incorrect fragments highlighted
* **`<original_name>_fixed.docx`** — automatically corrected version (if auto-fix is enabled)

---
