"""CLI: replace headers and footers in .docx files using a template."""
from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from .hf_replace import VALID_ENGINES, ENGINE_AUTO, replace_headers_footers

logger = logging.getLogger("pdf_to_doc.hf")


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="docx-hf-replace",
        description=(
            "Replace header and footer in every .docx file inside a folder "
            "(recursively by default) with the header/footer from a template "
            ".docx file."
        ),
    )
    p.add_argument("template", type=Path, help="Template .docx file.")
    p.add_argument("folder", type=Path, help="Folder containing target .docx files.")
    p.add_argument(
        "--no-recursive", dest="recursive", action="store_false",
        help="Do NOT scan subfolders (default is recursive).",
    )
    p.add_argument(
        "--engine", choices=VALID_ENGINES, default=ENGINE_AUTO,
        help="Conversion engine (default: auto -> Word if available).",
    )
    p.add_argument(
        "--no-backup", dest="backup", action="store_false",
        help="Do NOT create .bak copies before modifying files.",
    )
    p.add_argument("-v", "--verbose", action="store_true")
    p.set_defaults(recursive=True, backup=True)
    return p


def main(argv: list[str] | None = None) -> int:
    args = _build_parser().parse_args(argv)
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )
    try:
        found, ok, bad = replace_headers_footers(
            template_path=args.template,
            target_folder=args.folder,
            recursive=args.recursive,
            engine=args.engine,
            backup=args.backup,
        )
    except Exception as exc:
        logger.error("%s", exc)
        return 2

    logger.info("Files found: %d | Succeeded: %d | Failed: %d", found, ok, bad)
    return 0 if bad == 0 else 2


if __name__ == "__main__":
    sys.exit(main())
