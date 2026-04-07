"""Scrape a web page for PDF links, create a folder per link, download each PDF.

Given a URL to a page that contains hyperlinks to PDF files, this module:
  1. Fetches the page HTML (via Playwright browser for JS-rendered sites,
     or plain requests for simple HTML pages).
  2. Finds every <a> whose href points to a .pdf file.
  3. For each link, creates a subfolder named after the link text (sanitised)
     inside a root output folder.
  4. Downloads the PDF into that subfolder.
"""
from __future__ import annotations

import json
import logging
import re
import unicodedata
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urljoin, urlparse, unquote

import requests
from bs4 import BeautifulSoup

logger = logging.getLogger(__name__)

_CHUNK_SIZE = 1024 * 256  # 256 KB


def _sanitize_name(name: str, *, max_len: int = 200) -> str:
    """Turn arbitrary link text into a safe folder / file name."""
    name = unicodedata.normalize("NFKC", name)
    # Strip zero-width and invisible Unicode characters.
    name = re.sub(r'[\u200b\u200c\u200d\u200e\u200f\ufeff\u00ad]', "", name)
    name = name.strip()
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name)
    name = re.sub(r'[\s_]+', " ", name).strip()
    if len(name) > max_len:
        name = name[:max_len].rstrip()
    return name or "unnamed"


def _pdf_filename_from_url(url: str) -> str:
    """Extract the .pdf filename from a URL, or fall back to 'document.pdf'."""
    path = unquote(urlparse(url).path)
    name = Path(path).name if path else ""
    if name.lower().endswith(".pdf"):
        return name
    return "document.pdf"


def _fetch_html_simple(page_url: str) -> str:
    """Fetch page HTML with plain requests (no JS execution)."""
    logger.debug("Fetching with requests: %s", page_url)
    resp = requests.get(page_url, timeout=60)
    resp.raise_for_status()
    return resp.text


def _fetch_html_browser(page_url: str) -> str:
    """Fetch page HTML using Playwright (renders JavaScript)."""
    logger.info("Launching browser to render JS-heavy page...")
    from playwright.sync_api import sync_playwright

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=True)
        try:
            page = browser.new_page()
            page.goto(page_url, wait_until="networkidle", timeout=60000)
            # Extra wait for late-loading Wix / SPA content.
            page.wait_for_timeout(3000)
            html = page.content()
        finally:
            browser.close()

    return html


def discover_pdf_links(
    page_url: str,
    *,
    use_browser: bool = True,
) -> list[dict]:
    """Fetch *page_url* and return a list of ``{text, href}`` for PDF links.

    Args:
        page_url: URL of the page to scrape.
        use_browser: If True (default), render JS with Playwright.
                     Set False for simple static HTML pages.
    """
    logger.info("Fetching page: %s (browser=%s)", page_url, use_browser)
    if use_browser:
        html = _fetch_html_browser(page_url)
    else:
        html = _fetch_html_simple(page_url)

    soup = BeautifulSoup(html, "html.parser")
    results: list[dict] = []
    seen_hrefs: set[str] = set()

    for a_tag in soup.find_all("a", href=True):
        href = a_tag["href"].strip()
        abs_url = urljoin(page_url, href)

        parsed = urlparse(abs_url)
        if not parsed.path.lower().endswith(".pdf"):
            continue

        if abs_url in seen_hrefs:
            continue
        seen_hrefs.add(abs_url)

        text = a_tag.get_text(separator=" ", strip=True)
        if not text:
            text = _pdf_filename_from_url(abs_url)
        text = _sanitize_name(text)

        results.append({"text": text, "href": abs_url})

    logger.info("Found %d PDF link(s) on the page.", len(results))
    for link in results:
        logger.debug("  %s -> %s", link["text"], link["href"])
    return results


def download_pdf(url: str, dest: Path, *, timeout: int = 120) -> dict:
    """Download a PDF from *url* into *dest*.

    Returns a metadata dict with download details (used for the log file).
    """
    dest.parent.mkdir(parents=True, exist_ok=True)
    logger.info("Downloading %s", url)

    start_time = datetime.now(timezone.utc)

    with requests.get(url, stream=True, timeout=timeout) as r:
        r.raise_for_status()
        content_type = r.headers.get("Content-Type", "unknown")
        server = r.headers.get("Server", "unknown")
        with open(dest, "wb") as f:
            for chunk in r.iter_content(_CHUNK_SIZE):
                f.write(chunk)

    end_time = datetime.now(timezone.utc)
    file_size = dest.stat().st_size
    size_mb = file_size / (1024 * 1024)
    logger.info("Saved %s (%.2f MB)", dest.name, size_mb)

    return {
        "download_url": url,
        "original_filename": _pdf_filename_from_url(url),
        "saved_filename": dest.name,
        "saved_path": str(dest),
        "folder_name": dest.parent.name,
        "file_size_bytes": file_size,
        "file_size_mb": round(size_mb, 3),
        "content_type": content_type,
        "server": server,
        "download_started_utc": start_time.isoformat(),
        "download_completed_utc": end_time.isoformat(),
        "download_duration_seconds": round(
            (end_time - start_time).total_seconds(), 2
        ),
    }


def _write_log_file(
    target_dir: Path, metadata: dict, page_url: str, *, overwrite: bool = False,
) -> None:
    """Write or append a download entry to download_log.json.

    When overwriting an existing file, the new entry is appended to the log
    (as a JSON array) with a visual separator, preserving the download history.
    On first download the file is created with a single-entry array.
    """
    log_path = target_dir / "download_log.json"
    log_entry = {
        "source_page": page_url,
        **metadata,
    }

    entries: list[dict] = []
    if log_path.exists():
        try:
            existing = json.loads(log_path.read_text(encoding="utf-8"))
            if isinstance(existing, list):
                entries = existing
            else:
                # Migrate old single-object format into a list.
                entries = [existing]
        except (json.JSONDecodeError, OSError):
            entries = []

    # Add a separator marker so it's easy to spot re-downloads in the log.
    if entries:
        log_entry["_note"] = (
            f"--- Re-downloaded (entry #{len(entries) + 1}) ---"
        )

    entries.append(log_entry)

    log_path.write_text(
        json.dumps(entries, indent=2, ensure_ascii=False), encoding="utf-8",
    )
    logger.debug("Log written: %s (%d entries)", log_path, len(entries))


def scrape_and_download(
    page_url: str,
    root_folder: Path,
    *,
    use_browser: bool = True,
    overwrite: bool = False,
) -> tuple[int, int, int]:
    """End-to-end: discover PDF links, create folders, download files.

    For each link with text ``"Some Report"``, creates::

        root_folder/
        └── Some Report/
            └── Some Report.pdf

    Args:
        overwrite: If True, re-download and overwrite existing PDFs. The log
                   file is appended to (not replaced) so download history is
                   preserved.

    Returns ``(found, downloaded, failed)``.
    """
    root_folder = Path(root_folder)
    root_folder.mkdir(parents=True, exist_ok=True)

    links = discover_pdf_links(page_url, use_browser=use_browser)
    if not links:
        logger.warning("No PDF links found on %s", page_url)
        return (0, 0, 0)

    downloaded = 0
    failed = 0
    for link in links:
        folder_name = link["text"]
        pdf_url = link["href"]

        target_dir = root_folder / folder_name
        target_file = target_dir / (folder_name + ".pdf")

        if not overwrite and target_file.exists():
            logger.info("Skipping (exists): %s", target_file)
            continue

        try:
            metadata = download_pdf(pdf_url, target_file)
            _write_log_file(target_dir, metadata, page_url, overwrite=overwrite)
            downloaded += 1
        except Exception as exc:
            failed += 1
            logger.error("Failed to download %s: %s", pdf_url, exc)

    return (len(links), downloaded, failed)
