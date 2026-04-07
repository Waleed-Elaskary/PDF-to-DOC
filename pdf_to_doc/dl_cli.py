"""CLI: scrape a web page for PDF links, create folders, download each PDF."""
from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from .pdf_downloader import scrape_and_download

logger = logging.getLogger("pdf_to_doc.dl")


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="pdf-download",
        description=(
            "Given a URL to a web page that contains hyperlinks to PDF files, "
            "create a subfolder per link (named after the link text) under a "
            "root folder and download each PDF into its subfolder."
        ),
    )
    p.add_argument("url", help="URL of the page containing PDF links.")
    p.add_argument(
        "folder", type=Path,
        help="Root folder where subfolders and PDFs will be created.",
    )
    p.add_argument(
        "--overwrite", action="store_true",
        help=(
            "Re-download and overwrite existing PDFs. The download log is "
            "appended to (not replaced) so history is preserved."
        ),
    )
    p.add_argument(
        "--no-browser", dest="use_browser", action="store_false",
        help=(
            "Use plain HTTP instead of a browser. Only works for simple HTML "
            "pages (no JavaScript rendering). Default uses a browser."
        ),
    )
    p.add_argument("-v", "--verbose", action="store_true")
    p.set_defaults(use_browser=True)
    return p


def main(argv: list[str] | None = None) -> int:
    args = _build_parser().parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )

    try:
        found, downloaded, failed = scrape_and_download(
            page_url=args.url,
            root_folder=args.folder,
            use_browser=args.use_browser,
            overwrite=args.overwrite,
        )
    except Exception as exc:
        logger.error("%s", exc)
        return 2

    logger.info(
        "Links found: %d | Downloaded: %d | Failed: %d",
        found, downloaded, failed,
    )
    return 0 if failed == 0 else 2


if __name__ == "__main__":
    sys.exit(main())
