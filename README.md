# PDF-to-DOC Toolkit

A collection of six command-line tools for working with PDF and Word documents:

1. **pdf-to-doc** — Batch convert PDF files to editable `.docx` documents
2. **pdf-to-libre** — Batch convert PDF files to `.odg` (LibreOffice Draw) or `.odt` (LibreOffice Writer)
3. **odt-to-docx** — Batch convert `.odt` files to `.docx` via LibreOffice, with automatic removal of white background rectangles from PDF conversion artifacts
4. **odt-hf-apply** — Apply header/footer from a template `.odt` to matching `.odt` files with auto-incrementing filenames
5. **docx-hf-replace** — Replace headers and footers across many `.docx` files using a template
6. **pdf-download** — Scrape a web page for PDF links, create named folders, and download each file

All tools are cross-platform Python CLI applications. The Word COM engine (best editable `.docx` output) is available on Windows with Microsoft Word installed. LibreOffice conversions require [LibreOffice](https://www.libreoffice.org/download/) installed.

---

## Installation

**Prerequisites:** Python 3.10+ and `pip`.

```bash
# Clone the repository
git clone <repository-url>
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

**Optional — for LibreOffice conversions (.odg / .odt):**

Install [LibreOffice](https://www.libreoffice.org/download/). The tool auto-detects common install locations, or you can specify the path with `--soffice`.

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

# OCR with specific language(s)
python -m pdf_to_doc /path/to/folder --ocr --ocr-language eng+fra

# Increase OCR resolution for better accuracy
python -m pdf_to_doc /path/to/folder --ocr --ocr-dpi 400

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
| `--ocr-language LANG` | Tesseract language code(s), e.g. `eng`, `eng+fra`, `eng+ara` (default: `eng`) |
| `--ocr-dpi DPI` | Rendering DPI for OCR (default: `300`). Higher values improve accuracy for small text |
| `--tessdata PATH` | Path to tessdata directory (auto-detected if omitted) |
| `-v`, `--verbose` | Debug logging |

### Engines

| Engine | Platform | Output Quality | Notes |
|--------|----------|---------------|-------|
| `word` | Windows | Best — fully editable text, tables, images, styles | Requires Microsoft Word installed + `pywin32` |
| `pdf2docx` | Any | Good — preserves layout, tables, images in floating frames | Pure Python, no external software needed |
| `text` | Any | Basic — plain text with page breaks | Fastest, no formatting preserved |

### Behavior

- For each `name.pdf`, writes `name.docx` in the same directory.
- If `name.docx` already exists and `--overwrite` is not set, creates `name (1).docx`, `name (2).docx`, etc.
- Source PDFs are never modified.
- If a single file fails, the tool logs the error and continues with the rest.
- For scanned/image-based PDFs with no text layer, enable `--ocr` to extract text via Tesseract. The tool first creates a searchable PDF, then runs the conversion engine on it.
- Exit codes: `0` = all succeeded, `1` = no PDFs found, `2` = some files failed.

---

## Tool 2: PDF to LibreOffice Formats (ODG / ODT)

Batch convert PDF files to LibreOffice Draw (`.odg`) or LibreOffice Writer (`.odt`) using headless LibreOffice. One output file is written per input PDF, saved in the same folder using the same base filename. Optionally prepend a prefix to all output filenames.

### Prerequisites

Install [LibreOffice](https://www.libreoffice.org/download/). The tool auto-detects common install locations on Windows, macOS, and Linux. You can also set the `SOFFICE_PATH` environment variable or pass `--soffice` explicitly.

### Usage

```bash
# Convert all PDFs in a folder to .odg (LibreOffice Draw, default)
python -m pdf_to_doc.lo_cli /path/to/folder

# Convert to .odt (LibreOffice Writer)
python -m pdf_to_doc.lo_cli /path/to/folder -f odt

# Recursive — include subfolders
python -m pdf_to_doc.lo_cli /path/to/folder -r

# Convert specific files
python -m pdf_to_doc.lo_cli file1.pdf file2.pdf -f odt

# Add a prefix to output filenames
python -m pdf_to_doc.lo_cli /path/to/folder --prefix "PROJ-"

# Overwrite existing output files
python -m pdf_to_doc.lo_cli /path/to/folder --overwrite

# Combine: recursive + prefix + format + overwrite
python -m pdf_to_doc.lo_cli /path/to/folder -r -f odt --prefix "PROJ-" --overwrite

# Specify LibreOffice path explicitly
python -m pdf_to_doc.lo_cli /path/to/folder --soffice "/path/to/soffice"

# Custom timeout per file (default: 300 seconds)
python -m pdf_to_doc.lo_cli /path/to/folder --timeout 600

# Verbose logging
python -m pdf_to_doc.lo_cli /path/to/folder -v
```

### Options

| Flag | Description |
|------|-------------|
| `-f`, `--format {odg,odt}` | Output format: `odg` for LibreOffice Draw (default), `odt` for LibreOffice Writer |
| `-r`, `--recursive` | Scan subfolders for PDFs |
| `--overwrite` | Overwrite existing output files instead of creating `name (1).odg` |
| `--prefix TEXT` | Prefix to prepend to output filenames (e.g. `--prefix "PROJ-"` turns `report.pdf` into `PROJ-report.odg`) |
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
- With `--prefix "PROJ-"`, output becomes `PROJ-name.odg`.
- If the output file already exists and `--overwrite` is not set, creates `name (1).odg`, `name (2).odg`, etc.
- Source PDFs are never modified.
- If a single file fails, the tool logs the error and continues with the rest.
- Uses an isolated user profile for headless LibreOffice, so conversions work even when LibreOffice is already open in the GUI.
- For ODT conversion, the tool forces Writer to import the PDF (not Draw) and exports with the correct Writer filter.
- For ODG conversion, Draw's native PDF import is used.
- LibreOffice auto-detection order: `--soffice` flag > `SOFFICE_PATH` environment variable > system PATH > common install directories.
- Exit codes: `0` = all succeeded, `1` = no PDFs found or LibreOffice not found, `2` = some files failed.

### Troubleshooting

| Issue | Solution |
|-------|----------|
| LibreOffice not found | Install LibreOffice, or pass `--soffice /path/to/soffice` |
| Silent conversion failure | Close any open LibreOffice windows and retry |
| Timeout on large files | Increase with `--timeout 600` (or higher) |
| Permission errors | Ensure write access to the PDF's parent directory |

---

## Tool 3: ODT to DOCX Converter

Batch convert `.odt` (LibreOffice Writer) files to `.docx` using headless LibreOffice. Automatically detects and removes white background rectangles that are artifacts from PDF-to-ODT conversions — these appear as opaque white boxes in Word that hide content underneath. The removal scans **all pages**, including headers and footers.

### Prerequisites

Install [LibreOffice](https://www.libreoffice.org/download/).

### Usage

```bash
# Convert all .odt files in a folder
python -m pdf_to_doc.odt_cli /path/to/folder

# Recursive — include subfolders
python -m pdf_to_doc.odt_cli /path/to/folder -r

# Convert specific files
python -m pdf_to_doc.odt_cli file1.odt file2.odt

# Add a prefix to output filenames
python -m pdf_to_doc.odt_cli /path/to/folder --prefix "PROJ-"

# Overwrite existing .docx files
python -m pdf_to_doc.odt_cli /path/to/folder --overwrite

# Keep white background rectangles (do not strip them)
python -m pdf_to_doc.odt_cli /path/to/folder --keep-bg

# Combine options
python -m pdf_to_doc.odt_cli /path/to/folder -r --prefix "PROJ-" --overwrite -v

# Specify LibreOffice path explicitly
python -m pdf_to_doc.odt_cli /path/to/folder --soffice "/path/to/soffice"

# Custom timeout per file
python -m pdf_to_doc.odt_cli /path/to/folder --timeout 600

# Verbose logging
python -m pdf_to_doc.odt_cli /path/to/folder -v
```

### Options

| Flag | Description |
|------|-------------|
| `-r`, `--recursive` | Scan subfolders for `.odt` files |
| `--overwrite` | Overwrite existing `.docx` instead of creating `name (1).docx` |
| `--prefix TEXT` | Prefix to prepend to output filenames (e.g. `--prefix "PROJ-"` turns `report.odt` into `PROJ-report.docx`) |
| `--keep-bg` | Do **not** strip white background rectangles (by default they are removed) |
| `--soffice PATH` | Explicit path to the `soffice` executable |
| `--timeout SECONDS` | Per-file conversion timeout (default: `300`) |
| `-v`, `--verbose` | Debug logging |

### White Background Removal

When PDFs are converted to `.odt` via LibreOffice, each page gets a white rectangle shape as a background layer. When the `.odt` is then converted to `.docx`, these rectangles appear as opaque white boxes in Microsoft Word that hide content behind them.

This tool automatically:
- Scans the **entire document** — body, headers, and footers across all pages and sections.
- Detects both modern DrawingML shapes (`<wps:wsp>`) and legacy VML shapes (`<v:rect>`) with white or near-white fills.
- Removes the shape and its container element without affecting other content.
- Use `--keep-bg` to disable this behavior if the rectangles are intentional.

### Behavior

- For each `name.odt`, writes `name.docx` in the same directory.
- With `--prefix "PROJ-"`, output becomes `PROJ-name.docx`.
- If the output file already exists and `--overwrite` is not set, creates `name (1).docx`, `name (2).docx`, etc.
- Source `.odt` files are never modified.
- If a single file fails, the tool logs the error and continues with the rest.
- Uses an isolated user profile for headless LibreOffice (works even with LibreOffice GUI open).
- Exit codes: `0` = all succeeded, `1` = no `.odt` files found or LibreOffice not found, `2` = some files failed.

### Troubleshooting

| Issue | Solution |
|-------|----------|
| LibreOffice not found | Install LibreOffice, or pass `--soffice /path/to/soffice` |
| Silent conversion failure | Close any open LibreOffice windows and retry |
| White boxes still visible | Try opening the `.docx` in Word — some viewers render shapes differently. If still present, file a bug with the problematic `.odt` |
| Timeout on large files | Increase with `--timeout 600` |

---

## Tool 4: ODT Header & Footer Applicator

Apply header and footer from a template `.odt` file to all `.odt` files in a folder that match a filename pattern. The output file has the trailing number in the filename incremented (e.g. `-001.odt` becomes `-002.odt`). This is a pure-Python tool — no LibreOffice needed at runtime.

### Usage

```bash
# Apply header/footer to all files matching default pattern (*-NNN.odt)
python -m pdf_to_doc.odt_hf_cli /path/to/template.odt /path/to/folder

# Recursive
python -m pdf_to_doc.odt_hf_cli /path/to/template.odt /path/to/folder -r

# Custom filename pattern (regex with one capture group for the number)
python -m pdf_to_doc.odt_hf_cli /path/to/template.odt /path/to/folder -p "^EBB-.+-(\d+)\.odt$"

# Overwrite existing output files
python -m pdf_to_doc.odt_hf_cli /path/to/template.odt /path/to/folder --overwrite

# Custom increment (e.g. +2 instead of +1)
python -m pdf_to_doc.odt_hf_cli /path/to/template.odt /path/to/folder --increment 2

# Verbose logging
python -m pdf_to_doc.odt_hf_cli /path/to/template.odt /path/to/folder -v
```

### Options

| Flag | Description |
|------|-------------|
| `-p`, `--pattern REGEX` | Regex to match filenames. Must have one capture group for the trailing number. Default: `^.+-(\d+)\.odt$` |
| `-r`, `--recursive` | Scan subfolders |
| `--overwrite` | Overwrite existing output files |
| `--increment N` | Number to add to trailing digit (default: `1`) |
| `-v`, `--verbose` | Debug logging |

### Pattern Examples

| Pattern | Matches | Output |
|---------|---------|--------|
| `^.+-(\d+)\.odt$` (default) | `EBB-Report-001.odt` | `EBB-Report-002.odt` |
| `^EBB-.+-(\d+)\.odt$` | Only files starting with `EBB-` | `EBB-Name-002.odt` |
| `^.*_v(\d+)\.odt$` | `Document_v3.odt` | `Document_v4.odt` |
| `^Report-(\d+)\.odt$` | `Report-05.odt` | `Report-06.odt` |

### Behavior

- Copies all header/footer variants (primary, left-page, first-page) from the template's first master-page into every master-page of each target.
- Copies images referenced by the header/footer from the template into the output file.
- Copies automatic styles used by the header/footer content.
- The trailing number is zero-padded to the same width as the original (e.g. `001` → `002`, not `2`).
- The template file itself is skipped if it resides inside the target folder.
- Source `.odt` files are never modified — output is always a new file.

---

## Tool 5: Header & Footer Replacer

Replace the header and footer in every `.docx` file inside a folder (recursively by default) with the header and footer from a template `.docx` file.

### Usage

```bash
# Replace headers/footers in all .docx files under a folder (recursive)
python -m pdf_to_doc.hf_cli /path/to/template.docx /path/to/target/folder

# Non-recursive (top folder only)
python -m pdf_to_doc.hf_cli /path/to/template.docx /path/to/folder --no-recursive

# Skip creating .bak backup copies
python -m pdf_to_doc.hf_cli /path/to/template.docx /path/to/folder --no-backup

# Force pure-Python engine (no Word COM)
python -m pdf_to_doc.hf_cli /path/to/template.docx /path/to/folder --engine python

# Verbose logging
python -m pdf_to_doc.hf_cli /path/to/template.docx /path/to/folder -v
```

### Options

| Flag | Description |
|------|-------------|
| `--no-recursive` | Only process the top-level folder (default is recursive) |
| `--engine {auto,word,python}` | `auto` picks `word` if available, else `python` |
| `--no-backup` | Do not create `.docx.bak` files before modifying |
| `-v`, `--verbose` | Debug logging |

### Engines

| Engine | Platform | Capability |
|--------|----------|------------|
| `word` | Windows | Copies images, tables, page numbers, fields, and all formatting |
| `python` | Any | Copies paragraphs and tables; does **not** copy images |

### Behavior

- The template's **first section** header/footer is applied to **every section** of each target document.
- All three header/footer variants are overwritten: primary, first-page, and even-page — ensuring **all pages** display the template content regardless of per-section settings in the target.
- "Link to Previous" is broken on every section so each receives its own independent copy.
- By default, a `.docx.bak` backup is created for each file before modification.
- The template file itself is skipped if it resides inside the target folder.
- Word temp/lock files (`~$*.docx`) are automatically skipped.

---

## Tool 6: PDF Link Downloader

Scrape a web page for PDF hyperlinks, create a named subfolder per link, and download each PDF. Supports both static HTML pages and JavaScript-rendered pages (single-page apps, dynamic content).

### Prerequisites (optional)

For JavaScript-rendered pages, install Playwright:

```bash
pip install playwright
python -m playwright install chromium
```

### Usage

```bash
# Download all PDFs linked on a page (uses browser rendering by default)
python -m pdf_to_doc.dl_cli "https://example.com/downloads" /path/to/output

# Overwrite existing PDFs (download log history is preserved)
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
| `--no-browser` | Skip browser rendering; use plain HTTP only. Faster, but only works for static HTML pages without JavaScript |
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
- By default, uses a headless Chromium browser (Playwright) to handle JavaScript-rendered pages (single-page apps, dynamically loaded content, etc.).
- Use `--no-browser` for faster downloads on simple static HTML pages.
- Skips already-downloaded files by default; use `--overwrite` to re-download.
- Deduplicates links found on the page.
- Invisible Unicode characters (zero-width spaces, etc.) are automatically stripped from link text to prevent filesystem encoding errors.

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
    "download_started_utc": "2025-01-15T14:32:01.123456+00:00",
    "download_completed_utc": "2025-01-15T14:32:03.456789+00:00",
    "download_duration_seconds": 2.33
  },
  {
    "_note": "--- Re-downloaded (entry #2) ---",
    "source_page": "https://example.com/downloads",
    "download_url": "https://example.com/files/report.pdf",
    "saved_filename": "Annual Report.pdf",
    "download_completed_utc": "2025-01-16T09:15:22.789012+00:00",
    "..."
  }
]
```

**Logged fields:** source page URL, direct download URL, original filename (from URL), saved filename, full local path, folder name, file size (bytes and MB), content type, server, download start/end timestamps (UTC), and download duration.

---

## Chaining Tools

Download PDFs from a web page, convert them to multiple formats, and apply a custom header/footer:

```bash
# 1. Download PDFs from a page
python -m pdf_to_doc.dl_cli "https://example.com/specs" /path/to/output -v

# 2a. Convert all downloaded PDFs to .docx (recursive)
python -m pdf_to_doc /path/to/output -r --engine word --overwrite -v

# 2b. Convert all downloaded PDFs to .odg (LibreOffice Draw) with prefix
python -m pdf_to_doc.lo_cli /path/to/output -r -f odg --prefix "PROJ-" --overwrite -v

# 2c. Convert all downloaded PDFs to .odt (LibreOffice Writer)
python -m pdf_to_doc.lo_cli /path/to/output -r -f odt --overwrite -v

# 2d. Convert .odt files to .docx (strips white background artifacts)
python -m pdf_to_doc.odt_cli /path/to/output -r --overwrite -v

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
│   ├── lo_cli.py            # CLI for LibreOffice converter (ODG/ODT)
│   ├── lo_converter.py      # Core LibreOffice conversion logic
│   ├── odt_cli.py           # CLI for ODT-to-DOCX converter
│   ├── odt_to_docx.py       # Core ODT-to-DOCX + white rect removal logic
│   ├── odt_hf_cli.py        # CLI for ODT header/footer applicator
│   ├── odt_hf.py            # Core ODT header/footer replacement logic
│   ├── hf_cli.py            # CLI for header/footer replacer
│   ├── hf_replace.py        # Core header/footer replacement logic
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
