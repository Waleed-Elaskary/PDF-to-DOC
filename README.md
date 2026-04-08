# PDF-to-DOC Toolkit

A collection of four command-line tools for working with PDF and Word documents:

1. **pdf-to-doc** — Batch convert PDF files to editable `.docx` documents
2. **pdf-to-libre** — Batch convert PDF files to `.odg` (LibreOffice Draw) or `.odt` (LibreOffice Writer)
3. **docx-hf-replace** — Replace headers and footers across many `.docx` files using a template
4. **pdf-download** — Scrape a web page for PDF links, create named folders, and download each file

All tools are cross-platform Python CLI applications. The Word COM engine (best editable output for `.docx`) is available on Windows with Microsoft Word installed. LibreOffice conversions require [LibreOffice](https://www.libreoffice.org/download/) installed.

---

## Installation

**Prerequisites:** Python 3.10+ and `pip`.

```bash
# Clone the repository
git clone https://github.com/Waleed-Elaskary/PDF-to-DOC.git
cd PDF-to-DOC

# Create and activate a virtual environment
python -m venv .venv
# Windows (PowerShell)
.venv\Scripts\Activate.ps1
# Windows (cmd)
.venv\Scripts\activate.bat
# macOS / Linux
source .venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

**Optional — for the PDF downloader on JavaScript-rendered pages:**

```bash
pip install playwright
python -m playwright install chromium
```

**Optional — for OCR of scanned PDFs:**

Install [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki) and ensure it is on your PATH or in a standard install directory.

---

## Tool 1: PDF to DOCX Converter

Batch convert PDF files into Microsoft Word `.docx` documents. One `.docx` is written per input PDF, saved in the same folder using the same base filename.

### Usage

```bash
# Convert all PDFs in a folder (non-recursive)
python -m pdf_to_doc /path/to/folder

# Recursive — include subfolders
python -m pdf_to_doc /path/to/folder -r

# Convert specific files
python -m pdf_to_doc file1.pdf file2.pdf

# Use Microsoft Word engine (Windows, best editable output)
python -m pdf_to_doc /path/to/folder --engine word

# Overwrite existing .docx files
python -m pdf_to_doc /path/to/folder --overwrite

# OCR scanned PDFs (requires Tesseract)
python -m pdf_to_doc /path/to/folder --ocr

# OCR with specific language
python -m pdf_to_doc /path/to/folder --ocr --ocr-language eng+ara

# Verbose logging
python -m pdf_to_doc /path/to/folder -v
```

### Options

| Flag | Description |
|------|-------------|
| `-r`, `--recursive` | Scan subfolders for PDFs |
| `--engine {auto,word,pdf2docx,text}` | Conversion engine. `auto` picks `word` if available, else `pdf2docx` |
| `--overwrite` | Overwrite existing `.docx` instead of creating `name (1).docx` |
| `--ocr` | Run Tesseract OCR on pages with no text layer |
| `--ocr-language LANG` | Tesseract language code (default: `eng`) |
| `--ocr-dpi DPI` | Rendering DPI for OCR (default: `300`) |
| `--tessdata PATH` | Path to tessdata directory |
| `-v`, `--verbose` | Debug logging |

### Engines

| Engine | Platform | Output Quality | Notes |
|--------|----------|---------------|-------|
| `word` | Windows | Best — fully editable text, tables, images, styles | Requires Microsoft Word installed + `pywin32` |
| `pdf2docx` | Any | Good — preserves layout, tables, images in floating frames | Pure Python |
| `text` | Any | Basic — plain text with page breaks | Fastest, no formatting |

### Behavior

- For each `name.pdf`, writes `name.docx` in the same directory.
- If `name.docx` already exists and `--overwrite` is not set, creates `name (1).docx`, `name (2).docx`, etc.
- Source PDFs are never modified.
- If a single file fails, the tool logs the error and continues with the rest.
- Exit codes: `0` = all succeeded, `1` = no PDFs found, `2` = some files failed.

---

## Tool 2: PDF to LibreOffice Formats (ODG / ODT)

Batch convert PDF files to LibreOffice Draw (`.odg`) or LibreOffice Writer (`.odt`) using headless LibreOffice. One output file is written per input PDF, saved in the same folder using the same base filename.

### Prerequisites

Install [LibreOffice](https://www.libreoffice.org/download/). The tool auto-detects common install locations, or you can specify the path explicitly.

### Usage

```bash
# Convert all PDFs in a folder to .odg (LibreOffice Draw, default)
python -m pdf_to_doc.lo_cli /path/to/folder

# Convert to .odt (LibreOffice Writer)
python -m pdf_to_doc.lo_cli /path/to/folder -f odt

# Recursive — include subfolders
python -m pdf_to_doc.lo_cli /path/to/folder -r

# Convert specific files
python -m pdf_to_doc.lo_cli file1.pdf file2.pdf

# Overwrite existing output files
python -m pdf_to_doc.lo_cli /path/to/folder --overwrite

# Specify LibreOffice path explicitly
python -m pdf_to_doc.lo_cli /path/to/folder --soffice "/usr/bin/soffice"

# Custom timeout per file (default: 300 seconds)
python -m pdf_to_doc.lo_cli /path/to/folder --timeout 600

# Verbose logging
python -m pdf_to_doc.lo_cli /path/to/folder -v
```

### Options

| Flag | Description |
|------|-------------|
| `-f`, `--format {odg,odt}` | Output format: `odg` (Draw, default) or `odt` (Writer) |
| `-r`, `--recursive` | Scan subfolders for PDFs |
| `--overwrite` | Overwrite existing output files instead of creating `name (1).odg` |
| `--soffice PATH` | Explicit path to the `soffice` executable |
| `--timeout SECONDS` | Per-file conversion timeout (default: `300`) |
| `-v`, `--verbose` | Debug logging |

### Formats

| Format | Extension | Application | Best For |
|--------|-----------|-------------|----------|
| `odg` | `.odg` | LibreOffice Draw | Drawings, diagrams, PDF pages as editable graphics |
| `odt` | `.odt` | LibreOffice Writer | Text documents, reports, editable text content |

### Behavior

- For each `name.pdf`, writes `name.odg` (or `name.odt`) in the same directory.
- If the output file already exists and `--overwrite` is not set, creates `name (1).odg`, `name (2).odg`, etc.
- Source PDFs are never modified.
- If a single file fails, the tool logs the error and continues with the rest.
- LibreOffice auto-detection checks: `--soffice` flag, `SOFFICE_PATH` environment variable, system PATH, and common install directories.
- Exit codes: `0` = all succeeded, `1` = no PDFs found or LibreOffice not found, `2` = some files failed.

---

## Tool 3: Header & Footer Replacer

Replace the header and footer in every `.docx` file inside a folder (recursively by default) with the header and footer from a template `.docx` file.

### Usage

```bash
# Replace headers/footers in all .docx files under a folder (recursive)
python -m pdf_to_doc.hf_cli /path/to/template.docx /path/to/target/folder

# Non-recursive (top folder only)
python -m pdf_to_doc.hf_cli template.docx folder --no-recursive

# Skip creating .bak backup copies
python -m pdf_to_doc.hf_cli template.docx folder --no-backup

# Force pure-Python engine (no Word COM)
python -m pdf_to_doc.hf_cli template.docx folder --engine python

# Verbose logging
python -m pdf_to_doc.hf_cli template.docx folder -v
```

### Options

| Flag | Description |
|------|-------------|
| `--no-recursive` | Only process the top-level folder |
| `--engine {auto,word,python}` | `auto` picks `word` if available |
| `--no-backup` | Do not create `.docx.bak` files before modifying |
| `-v`, `--verbose` | Debug logging |

### Engines

| Engine | Platform | Capability |
|--------|----------|------------|
| `word` | Windows | Copies images, tables, page numbers, fields, and all formatting |
| `python` | Any | Copies paragraphs and tables; does **not** copy images |

### Behavior

- The template's **first section** header/footer is applied to **every section** of each target document.
- All three variants are overwritten: primary, first-page, and even-page headers/footers — ensuring **all pages** display the template content.
- "Link to Previous" is broken on every section so each receives its own copy.
- By default, a `.docx.bak` backup is created for each file before modification.
- The template file is skipped if it resides inside the target folder.
- Word temp/lock files (`~$*.docx`) are automatically skipped.

---

## Tool 4: PDF Link Downloader

Scrape a web page for PDF hyperlinks, create a named subfolder per link, and download each PDF.

### Usage

```bash
# Download all PDFs linked on a page (uses browser rendering by default)
python -m pdf_to_doc.dl_cli "https://example.com/downloads" /path/to/output

# Overwrite existing PDFs (log history is preserved)
python -m pdf_to_doc.dl_cli "https://example.com/downloads" /path/to/output --overwrite

# Use plain HTTP (faster, for simple static HTML pages without JavaScript)
python -m pdf_to_doc.dl_cli "https://example.com/downloads" /path/to/output --no-browser

# Verbose logging
python -m pdf_to_doc.dl_cli "https://example.com/downloads" /path/to/output -v
```

### Options

| Flag | Description |
|------|-------------|
| `--overwrite` | Re-download and overwrite existing PDFs. The download log is appended to (not replaced) so history is preserved |
| `--no-browser` | Skip browser rendering; use plain HTTP (only for static HTML pages) |
| `-v`, `--verbose` | Debug logging |

### Behavior

Given a page containing:
```html
<a href="/files/report.pdf">Annual Report</a>
<a href="/files/budget.pdf">Budget Summary</a>
```

The tool creates:
```
output-folder/
├── Annual Report/
│   ├── Annual Report.pdf
│   └── download_log.json
└── Budget Summary/
    ├── Budget Summary.pdf
    └── download_log.json
```

- **Folder name** = sanitized link text from the `<a>` tag.
- **PDF filename** = same as the folder name (with `.pdf` extension).
- By default, uses a headless Chromium browser (Playwright) to handle JavaScript-rendered pages (Wix, React, Angular, etc.).
- Use `--no-browser` for faster downloads on simple static HTML pages.
- Skips already-downloaded files by default; use `--overwrite` to re-download.
- Deduplicates links found on the page.
- Invisible Unicode characters (zero-width spaces, etc.) are stripped from link text.

### Download Log

Each subfolder contains a `download_log.json` with metadata for every download. On re-downloads with `--overwrite`, new entries are **appended** (not replaced), preserving full history:

```json
[
  {
    "source_page": "https://example.com/downloads",
    "download_url": "https://example.com/files/report.pdf",
    "original_filename": "report.pdf",
    "saved_filename": "Annual Report.pdf",
    "saved_path": "/path/to/output/Annual Report/Annual Report.pdf",
    "folder_name": "Annual Report",
    "file_size_bytes": 245760,
    "file_size_mb": 0.234,
    "content_type": "application/pdf",
    "server": "nginx",
    "download_started_utc": "2026-04-07T14:32:01.123456+00:00",
    "download_completed_utc": "2026-04-07T14:32:03.456789+00:00",
    "download_duration_seconds": 2.33
  },
  {
    "_note": "--- Re-downloaded (entry #2) ---",
    "source_page": "https://example.com/downloads",
    "download_url": "https://example.com/files/report.pdf",
    "saved_filename": "Annual Report.pdf",
    "download_completed_utc": "2026-04-08T09:15:22.789012+00:00",
    "..."
  }
]
```

**Logged fields:** source page URL, direct download URL, original filename (from URL), saved filename, full path, folder name, file size (bytes and MB), content type, server, download start/end timestamps (UTC), and download duration.

---

## Chaining Tools

Download PDFs from a web page, convert them to multiple formats, and apply a custom header/footer:

```bash
# 1. Download PDFs
python -m pdf_to_doc.dl_cli "https://example.com/specs" /path/to/output -v

# 2a. Convert all downloaded PDFs to .docx (recursive)
python -m pdf_to_doc /path/to/output -r --engine word --overwrite -v

# 2b. Convert all downloaded PDFs to .odg (LibreOffice Draw)
python -m pdf_to_doc.lo_cli /path/to/output -r -f odg --overwrite -v

# 2c. Convert all downloaded PDFs to .odt (LibreOffice Writer)
python -m pdf_to_doc.lo_cli /path/to/output -r -f odt --overwrite -v

# 3. Apply custom header/footer to .docx files from a template
python -m pdf_to_doc.hf_cli /path/to/template.docx /path/to/output -v
```

---

## Project Structure

```
PDF-to-DOC/
├── README.md
├── requirements.txt
├── pyproject.toml
├── .gitignore
├── pdf_to_doc/
│   ├── __init__.py
│   ├── __main__.py
│   ├── cli.py              # CLI for PDF-to-DOCX converter
│   ├── converter.py         # Core PDF-to-DOCX conversion logic
│   ├── hf_cli.py            # CLI for header/footer replacer
│   ├── hf_replace.py        # Core header/footer replacement logic
│   ├── lo_cli.py            # CLI for LibreOffice converter
│   ├── lo_converter.py      # Core LibreOffice conversion logic
│   ├── dl_cli.py            # CLI for PDF link downloader
│   └── pdf_downloader.py    # Core page scraping and download logic
├── tests/
│   └── test_converter.py
└── scripts/
    └── smoke_test.py
```

## Running Tests

```bash
pip install pytest
pytest -v
```

Or run the smoke test:

```bash
python scripts/smoke_test.py
```

---

## License

MIT
