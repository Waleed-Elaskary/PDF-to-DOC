"""CLI: convert PDFs to .odg (Draw) or .odt (Writer) via LibreOffice."""
from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from .lo_converter import (
    FORMAT_ODG,
    FORMAT_ODT,
    VALID_FORMATS,
    collect_pdfs,
    convert_pdf,
    find_soffice,
)

logger = logging.getLogger("pdf_to_doc.lo")


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="pdf-to-libre",
        description=(
            "Batch convert PDF files to LibreOffice formats using headless "
            "LibreOffice. Supports .odg (Draw) and .odt (Writer). Accepts "
            "either a folder (non-recursive by default) or an explicit list "
            "of PDF files. Each output is written next to its source PDF."
        ),
    )
    p.add_argument(
        "inputs", nargs="+", type=Path,
        help="One folder OR one-or-more PDF file paths.",
    )
    p.add_argument(
        "-f", "--format",
        choices=VALID_FORMATS,
        default=FORMAT_ODG,
        help=(
            "Output format: 'odg' for LibreOffice Draw (default), "
            "'odt' for LibreOffice Writer."
        ),
    )
    p.add_argument(
        "-r", "--recursive", action="store_true",
        help="When an input is a folder, also scan its subfolders for PDFs.",
    )
    p.add_argument(
        "--overwrite", action="store_true",
        help="Overwrite existing output files instead of creating 'name (1).odg'.",
    )
    p.add_argument(
        "--prefix", default="",
        help=(
            "Prefix to prepend to output filenames. "
            "E.g. --prefix 'EBB-' turns 'report.pdf' into 'EBB-report.odg'."
        ),
    )
    p.add_argument(
        "--soffice", default=None,
        help=(
            "Explicit path to the soffice executable. If omitted, the tool "
            "checks SOFFICE_PATH, PATH, and common install locations."
        ),
    )
    p.add_argument(
        "--timeout", type=int, default=300,
        help="Per-file conversion timeout in seconds (default: 300).",
    )
    p.add_argument(
        "-v", "--verbose", action="store_true",
        help="Enable debug logging.",
    )
    return p


def _configure_logging(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )


def main(argv: list[str] | None = None) -> int:
    args = _build_parser().parse_args(argv)
    _configure_logging(args.verbose)

    try:
        soffice = find_soffice(args.soffice)
    except FileNotFoundError as exc:
        logger.error("%s", exc)
        return 1
    logger.info("Using LibreOffice: %s", soffice)

    pdfs = collect_pdfs(args.inputs, recursive=args.recursive)
    if not pdfs:
        logger.error("No PDF files to process.")
        return 1

    logger.info(
        "Found %d PDF file(s) to convert to .%s",
        len(pdfs), args.format,
    )

    succeeded = 0
    failed = 0
    for pdf in pdfs:
        try:
            convert_pdf(
                pdf,
                args.format,
                soffice=soffice,
                overwrite=args.overwrite,
                prefix=args.prefix,
                timeout=args.timeout,
            )
            succeeded += 1
        except Exception as exc:
            failed += 1
            logger.error("Failed to convert %s: %s", pdf, exc)
            continue

    logger.info("Done. Succeeded: %d, Failed: %d", succeeded, failed)
    return 0 if failed == 0 else 2


if __name__ == "__main__":
    sys.exit(main())
