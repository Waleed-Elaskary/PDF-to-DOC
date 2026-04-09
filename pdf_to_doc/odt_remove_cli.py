"""CLI: remove objects from .odt files using a remove-template."""
from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from .odt_remove import process_folder, _DEFAULT_PATTERN

logger = logging.getLogger("pdf_to_doc.odt_rm")


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="odt-remove",
        description=(
            "Remove specific objects (images, shapes, lines) from .odt files "
            "using a template .odt as a reference. Every object found in the "
            "template is matched (by image content hash, shape geometry, etc.) "
            "and removed from all matching .odt files in the folder. Output "
            "files are saved with a suffix number appended before the "
            "extension (e.g. report.odt -> report-001.odt)."
        ),
    )
    p.add_argument(
        "template", type=Path,
        help="Remove-template .odt file containing objects to be removed.",
    )
    p.add_argument(
        "folder", type=Path,
        help="Folder containing target .odt files.",
    )
    p.add_argument(
        "-p", "--pattern", default=_DEFAULT_PATTERN,
        help=(
            "Regex pattern to match target filenames. "
            f"Default: '{_DEFAULT_PATTERN}' (all .odt files)."
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
        "--suffix", default="001",
        help=(
            "Number to append before the extension in output filenames. "
            "Default: '001' (e.g. report.odt -> report-001.odt)."
        ),
    )
    p.add_argument(
        "--keep-page-bg", action="store_true",
        help=(
            "Do NOT remove page-size white background rectangles. "
            "By default these PDF-to-ODT artifacts are always removed."
        ),
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
        found, ok, bad, total_obj = process_folder(
            template_path=args.template,
            folder=args.folder,
            pattern=args.pattern,
            recursive=args.recursive,
            overwrite=args.overwrite,
            suffix_num=args.suffix,
            remove_page_bg=not args.keep_page_bg,
        )
    except Exception as exc:
        logger.error("%s", exc)
        return 2

    logger.info(
        "Files found: %d | Succeeded: %d | Failed: %d | Objects removed: %d",
        found, ok, bad, total_obj,
    )
    return 0 if bad == 0 else 2


if __name__ == "__main__":
    sys.exit(main())
