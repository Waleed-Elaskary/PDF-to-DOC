"""CLI: convert .odt files to .pdf via LibreOffice."""
from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from .lo_converter import find_soffice
from .odt_to_pdf import collect_odts, convert_odt_to_pdf

logger = logging.getLogger("pdf_to_doc.odt_pdf")


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="odt-to-pdf",
        description=(
            "Batch convert .odt (LibreOffice Writer) files to .pdf using "
            "headless LibreOffice. Accepts either a folder (non-recursive by "
            "default) or an explicit list of .odt files. Each output is "
            "written next to its source file with the same base filename."
        ),
    )
    p.add_argument(
        "inputs", nargs="+", type=Path,
        help="One folder OR one-or-more .odt file paths.",
    )
    p.add_argument(
        "-p", "--pattern",
        default=None,
        help=(
            "Regex pattern to filter target filenames. "
            "E.g. '^EBB-.+\\.odt$' to only convert files starting with EBB-."
        ),
    )
    p.add_argument(
        "-r", "--recursive", action="store_true",
        help="When an input is a folder, also scan its subfolders for .odt files.",
    )
    p.add_argument(
        "--overwrite", action="store_true",
        help="Overwrite existing .pdf files instead of creating 'name (1).pdf'.",
    )
    p.add_argument(
        "--prefix", default="",
        help=(
            "Prefix to prepend to output filenames. "
            "E.g. --prefix 'PROJ-' turns 'report.odt' into 'PROJ-report.pdf'."
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

    odts = collect_odts(
        args.inputs,
        recursive=args.recursive,
        pattern=args.pattern,
    )
    if not odts:
        logger.error("No .odt files to process.")
        return 1

    logger.info("Found %d .odt file(s) to convert to .pdf", len(odts))

    succeeded = 0
    failed = 0
    for odt in odts:
        try:
            convert_odt_to_pdf(
                odt,
                soffice=soffice,
                overwrite=args.overwrite,
                prefix=args.prefix,
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
