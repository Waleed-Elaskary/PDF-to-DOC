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

import logging
import re
import unicodedata
from pathlib import Path
from urllib.parse import urljoin, urlparse, unquote

import requests
from bs4 import BeautifulSoup

logger = logging.getLogger(__name__)

_CHUNK_SIZE = 1024 * 256  # 256 KB


def _sanitize_name(name: str, *, max_len: int = 200) -> str:
    """Turn arbitrary link text into a safe folder / file name."""
    name = unicodedata.normalize("NFKC", name)
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


def download_pdf(url: str, dest: Path, *, timeout: int = 120) -> Path:
    """Download a PDF from *url* into *dest* (a file path). Returns *dest*."""
    dest.parent.mkdir(parents=True, exist_ok=True)
    logger.info("Downloading %s", url)
    with requests.get(url, stream=True, timeout=timeout) as r:
        r.raise_for_status()
        with open(dest, "wb") as f:
            for chunk in r.iter_content(_CHUNK_SIZE):
                f.write(chunk)
    size_mb = dest.stat().st_size / (1024 * 1024)
    logger.info("Saved %s (%.2f MB)", dest.name, size_mb)
    return dest


def scrape_and_download(
    page_url: str,
    root_folder: Path,
    *,
    use_browser: bool = True,
    skip_existing: bool = True,
) -> tuple[int, int, int]:
    """End-to-end: discover PDF links, create folders, download files.

    For each link with text ``"Some Report"``, creates::

        root_folder/
        └── Some Report/
            └── some-report.pdf   (original filename from the URL)

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
        pdf_filename = _pdf_filename_from_url(pdf_url)

        target_dir = root_folder / folder_name
        target_file = target_dir / (folder_name + ".pdf")

        if skip_existing and target_file.exists():
            logger.info("Skipping (exists): %s", target_file)
            continue

        try:
            download_pdf(pdf_url, target_file)
            downloaded += 1
        except Exception as exc:
            failed += 1
            logger.error("Failed to download %s: %s", pdf_url, exc)

    return (len(links), downloaded, failed)
