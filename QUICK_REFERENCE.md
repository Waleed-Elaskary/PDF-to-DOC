# PDF-to-DOC Toolkit — Quick Reference

## Tool 1: pdf-to-doc (PDF to DOCX)

```bash
python -m pdf_to_doc <inputs> [options]
```

| Parameter | Description |
|-----------|-------------|
| `inputs` | One folder OR one-or-more PDF file paths |
| `-r`, `--recursive` | Scan subfolders for PDFs |
| `--engine {auto,word,pdf2docx,text}` | Conversion engine (default: `auto`) |
| `--overwrite` | Overwrite existing `.docx` files |
| `--ocr` | OCR scanned PDFs via Tesseract |
| `--ocr-language LANG` | Tesseract language(s) (default: `eng`) |
| `--ocr-dpi DPI` | OCR rendering DPI (default: `300`) |
| `--tessdata PATH` | Path to tessdata directory |
| `-v`, `--verbose` | Debug logging |

**Engines:** `word` (Windows, best), `pdf2docx` (cross-platform), `text` (plain text only)

**Examples:**
```bash
python -m pdf_to_doc /path/to/folder -r --engine word --overwrite -v
python -m pdf_to_doc file1.pdf file2.pdf --ocr --ocr-language eng+fra
```

---

## Tool 2: pdf-to-libre (PDF to ODG / ODT)

```bash
python -m pdf_to_doc.lo_cli <inputs> [options]
```

| Parameter | Description |
|-----------|-------------|
| `inputs` | One folder OR one-or-more PDF file paths |
| `-f`, `--format {odg,odt}` | Output format (default: `odg`) |
| `-r`, `--recursive` | Scan subfolders for PDFs |
| `--overwrite` | Overwrite existing output files |
| `--prefix TEXT` | Prefix for output filenames (e.g. `--prefix "PROJ-"`) |
| `--soffice PATH` | Path to `soffice` executable |
| `--timeout SECONDS` | Per-file timeout (default: `300`) |
| `-v`, `--verbose` | Debug logging |

**Requires:** LibreOffice installed

**Examples:**
```bash
python -m pdf_to_doc.lo_cli /path/to/folder -r -f odt --prefix "PROJ-" --overwrite -v
python -m pdf_to_doc.lo_cli file1.pdf -f odg --timeout 600
```

---

## Tool 3: odt-to-docx (ODT to DOCX)

```bash
python -m pdf_to_doc.odt_cli <inputs> [options]
```

| Parameter | Description |
|-----------|-------------|
| `inputs` | One folder OR one-or-more `.odt` file paths |
| `-r`, `--recursive` | Scan subfolders for `.odt` files |
| `--overwrite` | Overwrite existing `.docx` files |
| `--prefix TEXT` | Prefix for output filenames (e.g. `--prefix "PROJ-"`) |
| `--keep-bg` | Do NOT remove white background rectangles |
| `--soffice PATH` | Path to `soffice` executable |
| `--timeout SECONDS` | Per-file timeout (default: `300`) |
| `-v`, `--verbose` | Debug logging |

**Requires:** LibreOffice installed

**Note:** By default, removes white background rectangles (PDF-to-ODT artifacts) from body, headers, and footers across all pages.

**Examples:**
```bash
python -m pdf_to_doc.odt_cli /path/to/folder -r --prefix "PROJ-" --overwrite -v
python -m pdf_to_doc.odt_cli file.odt --keep-bg --timeout 600
```

---

## Tool 4: odt-hf-apply (ODT Header & Footer Applicator)

```bash
python -m pdf_to_doc.odt_hf_cli <template.odt> <folder> [options]
```

| Parameter | Description |
|-----------|-------------|
| `template` | Template `.odt` file (source of header/footer) |
| `folder` | Folder containing target `.odt` files |
| `-p`, `--pattern REGEX` | Filename pattern with one capture group for trailing number (default: `^.+-(\d+)\.odt$`) |
| `-r`, `--recursive` | Scan subfolders |
| `--overwrite` | Overwrite existing output files |
| `--increment N` | Number to add to trailing digit (default: `1`) |
| `-v`, `--verbose` | Debug logging |

**Note:** Pure Python — no LibreOffice needed. Source files are never modified; output is a new file with incremented number.

**Examples:**
```bash
python -m pdf_to_doc.odt_hf_cli /path/to/template.odt /path/to/folder -r -v
python -m pdf_to_doc.odt_hf_cli template.odt folder -p "^EBB-.+-(\d+)\.odt$" --overwrite
python -m pdf_to_doc.odt_hf_cli template.odt folder --increment 2
```

**Filename behavior:**
| Input | Output |
|-------|--------|
| `EBB-Report-001.odt` | `EBB-Report-002.odt` |
| `Doc-05.odt` | `Doc-06.odt` |
| `File-999.odt` | `File-1000.odt` |

---

## Tool 5: docx-hf-replace (Header & Footer Replacer)

```bash
python -m pdf_to_doc.hf_cli <template.docx> <folder> [options]
```

| Parameter | Description |
|-----------|-------------|
| `template` | Template `.docx` file (source of header/footer) |
| `folder` | Folder containing target `.docx` files |
| `--no-recursive` | Only process top-level folder (default: recursive) |
| `--engine {auto,word,python}` | Replacement engine (default: `auto`) |
| `--no-backup` | Skip creating `.docx.bak` backups |
| `-v`, `--verbose` | Debug logging |

**Engines:** `word` (Windows, handles images/fields), `python` (cross-platform, text/tables only)

**Note:** Overwrites all header/footer variants (primary, first-page, even-page) on every section of every target document.

**Examples:**
```bash
python -m pdf_to_doc.hf_cli /path/to/template.docx /path/to/folder -v
python -m pdf_to_doc.hf_cli template.docx folder --no-recursive --no-backup
python -m pdf_to_doc.hf_cli template.docx folder --engine python
```

---

## Tool 6: pdf-download (PDF Link Downloader)

```bash
python -m pdf_to_doc.dl_cli <url> <folder> [options]
```

| Parameter | Description |
|-----------|-------------|
| `url` | URL of web page containing PDF links |
| `folder` | Root folder for downloaded files |
| `--overwrite` | Re-download existing PDFs (log is appended, not replaced) |
| `--no-browser` | Use plain HTTP instead of browser (static pages only) |
| `-v`, `--verbose` | Debug logging |

**Requires (optional):** Playwright + Chromium for JavaScript-rendered pages

**Note:** Creates one subfolder per link (named after link text). Each folder gets a `download_log.json` with metadata. PDF filename matches folder name.

**Examples:**
```bash
python -m pdf_to_doc.dl_cli "https://example.com/downloads" /path/to/output -v
python -m pdf_to_doc.dl_cli "https://example.com/downloads" /path/to/output --overwrite
python -m pdf_to_doc.dl_cli "https://example.com/downloads" /path/to/output --no-browser
```

---

## Tool 7: odt-remove (ODT Object Remover)

```bash
python -m pdf_to_doc.odt_remove_cli <template.odt> <folder> [options]
```

| Parameter | Description |
|-----------|-------------|
| `template` | Remove-template `.odt` file (objects to remove) |
| `folder` | Folder containing target `.odt` files |
| `-p`, `--pattern REGEX` | Filename pattern (default: `^.+\.odt$` — all `.odt` files) |
| `-r`, `--recursive` | Scan subfolders |
| `--overwrite` | Overwrite existing output files |
| `--suffix NUM` | Number appended before extension (default: `001`) |
| `--keep-page-bg` | Do NOT remove page-size white background rectangles |
| `-v`, `--verbose` | Debug logging |

**Matching (fuzzy):** Text frames by case-insensitive keyword overlap; images by hash or approximate position+size; page-size white backgrounds removed by default; thin-line polygons by height.

**Note:** Pure Python — no LibreOffice needed. Source files are never modified; output appends suffix (e.g. `report.odt` → `report-001.odt`). Removes from all pages including headers/footers.

**Examples:**
```bash
python -m pdf_to_doc.odt_remove_cli template.odt /path/to/folder -r -v
python -m pdf_to_doc.odt_remove_cli template.odt folder -p "^EBB-.+\.odt$" --overwrite
python -m pdf_to_doc.odt_remove_cli template.odt folder --suffix 002
```

---

## Tool 8: odt-to-pdf (ODT to PDF)

```bash
python -m pdf_to_doc.odt_pdf_cli <inputs> [options]
```

| Parameter | Description |
|-----------|-------------|
| `inputs` | One folder OR one-or-more `.odt` file paths |
| `-p`, `--pattern REGEX` | Filename pattern to filter (e.g. `^EBB-.+\.odt$`) |
| `-r`, `--recursive` | Scan subfolders for `.odt` files |
| `--overwrite` | Overwrite existing `.pdf` files |
| `--prefix TEXT` | Prefix for output filenames (e.g. `--prefix "PROJ-"`) |
| `--soffice PATH` | Path to `soffice` executable |
| `--timeout SECONDS` | Per-file timeout (default: `300`) |
| `-v`, `--verbose` | Debug logging |

**Requires:** LibreOffice installed

**Note:** Uses `writer_pdf_Export` filter. Source files are never modified. Lock files (`~$*.odt`) are skipped.

**Examples:**
```bash
python -m pdf_to_doc.odt_pdf_cli /path/to/folder -r --overwrite -v
python -m pdf_to_doc.odt_pdf_cli /path/to/folder -p "^EBB-.+\.odt$" --prefix "PROJ-"
python -m pdf_to_doc.odt_pdf_cli file1.odt file2.odt --timeout 600
```

---

## Common Workflows

### Download + Convert to DOCX
```bash
python -m pdf_to_doc.dl_cli "https://example.com/specs" /path/to/output -v
python -m pdf_to_doc /path/to/output -r --engine word --overwrite -v
```

### Download + Convert to ODT + Convert to DOCX (clean)
```bash
python -m pdf_to_doc.dl_cli "https://example.com/specs" /path/to/output -v
python -m pdf_to_doc.lo_cli /path/to/output -r -f odt --prefix "PROJ-" --overwrite -v
python -m pdf_to_doc.odt_cli /path/to/output -r --prefix "PROJ-" --overwrite -v
```

### Convert + Apply Header/Footer
```bash
python -m pdf_to_doc /path/to/pdfs -r --engine word --overwrite -v
python -m pdf_to_doc.hf_cli /path/to/template.docx /path/to/pdfs -v
```

### Full Pipeline
```bash
# 1. Download
python -m pdf_to_doc.dl_cli "https://example.com/specs" /path/to/output -v
# 2. Convert to DOCX
python -m pdf_to_doc /path/to/output -r --engine word --overwrite -v
# 3. Convert to ODG
python -m pdf_to_doc.lo_cli /path/to/output -r -f odg --overwrite -v
# 4. Convert to ODT then clean DOCX
python -m pdf_to_doc.lo_cli /path/to/output -r -f odt --overwrite -v
python -m pdf_to_doc.odt_cli /path/to/output -r --overwrite -v
# 5. Apply header/footer
python -m pdf_to_doc.hf_cli /path/to/template.docx /path/to/output -v
# 6. Export .odt to .pdf
python -m pdf_to_doc.odt_pdf_cli /path/to/output -r --overwrite -v
```

---

## Exit Codes (all tools)

| Code | Meaning |
|------|---------|
| `0` | All files processed successfully |
| `1` | No input files found, or missing prerequisite |
| `2` | Some files failed (others may have succeeded) |

## Environment Variables

| Variable | Description |
|----------|-------------|
| `SOFFICE_PATH` | Path to LibreOffice `soffice` binary |
| `TESSDATA_PREFIX` | Path to Tesseract `tessdata` directory |

## Dependencies

| Package | Required By | Purpose |
|---------|-------------|---------|
| `PyMuPDF` | Tool 1 | PDF text extraction, OCR |
| `python-docx` | Tools 1, 3, 5 | Read/write `.docx` files |
| `pdf2docx` | Tool 1 | Layout-aware PDF-to-DOCX |
| `lxml` | Tools 3, 4, 7 | XML manipulation and shape detection |
| `requests` | Tool 6 | HTTP downloads |
| `beautifulsoup4` | Tool 6 | HTML parsing |
| `playwright` | Tool 6 (optional) | JavaScript-rendered page support |
| `pywin32` | Tools 1, 5 (Windows) | Microsoft Word COM automation |
| LibreOffice | Tools 2, 3, 8 | Headless document conversion |
| Tesseract OCR | Tool 1 (optional) | OCR for scanned PDFs |
