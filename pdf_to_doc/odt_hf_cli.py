"""CLI: apply header/footer from a template .odt to matching .odt files."""
from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from .odt_hf import process_folder, _DEFAULT_PATTERN

logger = logging.getLogger("pdf_to_doc.odt_hf")


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="odt-hf-apply",
        description=(
            "Apply header and footer from a template .odt file to all .odt "
            "files in a folder that match a filename pattern. The output file "
            "has the trailing number in the filename incremented "
            "(e.g. -001.odt becomes -002.odt)."
        ),
    )
    p.add_argument(
        "template", type=Path,
        help="Template .odt file whose header/footer will be applied.",
    )
    p.add_argument(
        "folder", type=Path,
        help="Folder containing target .odt files.",
    )
    p.add_argument(
        "-p", "--pattern", default=_DEFAULT_PATTERN,
        help=(
            "Regex pattern to match filenames. Must contain exactly one "
            "capture group for the trailing number. "
            f"Default: '{_DEFAULT_PATTERN}' "
            "(matches filenames ending with -<digits>.odt)."
        ),
    )
    p.add_argument(
        "-r", "--recursive", action="store_true",
        help="Scan subfolders for matching .odt files.",
    )
    p.add_argument(
        "--overwrite", action="store_true",
        help="Overwrite existing output files.",
    )
    p.add_argument(
        "--increment", type=int, default=1,
        help="Number to add to the trailing digit (default: 1).",
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
        found, ok, bad = process_folder(
            template_path=args.template,
            folder=args.folder,
            pattern=args.pattern,
            recursive=args.recursive,
            overwrite=args.overwrite,
            increment=args.increment,
        )
    except Exception as exc:
        logger.error("%s", exc)
        return 2

    logger.info("Files found: %d | Succeeded: %d | Failed: %d", found, ok, bad)
    return 0 if bad == 0 else 2


if __name__ == "__main__":
    sys.exit(main())
