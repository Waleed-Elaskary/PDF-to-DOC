"""CLI: convert .odt files to .docx via LibreOffice, strip white backgrounds."""
from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from .odt_to_docx import collect_odt, convert_odt_to_docx
from .lo_converter import find_soffice

logger = logging.getLogger("pdf_to_doc.odt")


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="odt-to-docx",
        description=(
            "Batch convert .odt (LibreOffice Writer) files to .docx using "
            "headless LibreOffice. Automatically removes white background "
            "rectangles that originate from PDF-to-ODT conversions. Accepts "
            "either a folder or an explicit list of .odt files."
        ),
    )
    p.add_argument(
        "inputs", nargs="+", type=Path,
        help="One folder OR one-or-more .odt file paths.",
    )
    p.add_argument(
        "-r", "--recursive", action="store_true",
        help="When an input is a folder, also scan its subfolders for .odt files.",
    )
    p.add_argument(
        "--overwrite", action="store_true",
        help="Overwrite existing .docx files instead of creating 'name (1).docx'.",
    )
    p.add_argument(
        "--prefix", default="",
        help=(
            "Prefix to prepend to output filenames. "
            "E.g. --prefix 'PROJ-' turns 'report.odt' into 'PROJ-report.docx'."
        ),
    )
    p.add_argument(
        "--keep-bg", dest="strip_bg", action="store_false",
        help=(
            "Do NOT strip white background rectangles from the output .docx. "
            "By default they are removed (they are artifacts from PDF-to-ODT "
            "conversion)."
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
    p.set_defaults(strip_bg=True)
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

    odts = collect_odt(args.inputs, recursive=args.recursive)
    if not odts:
        logger.error("No .odt files to process.")
        return 1

    logger.info("Found %d .odt file(s) to convert.", len(odts))

    succeeded = 0
    failed = 0
    for odt in odts:
        try:
            convert_odt_to_docx(
                odt,
                soffice=soffice,
                overwrite=args.overwrite,
                prefix=args.prefix,
                strip_bg=args.strip_bg,
                timeout=args.timeout,
            )
            succeeded += 1
        except Exception as exc:
            failed += 1
            logger.error("Failed to convert %s: %s", odt, exc)
            continue

    logger.info("Done. Succeeded: %d, Failed: %d", succeeded, failed)
    return 0 if failed == 0 else 2


if __name__ == "__main__":
    sys.exit(main())
