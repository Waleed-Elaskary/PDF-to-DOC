"""Command-line interface for pdf-to-doc."""
from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from .converter import (
    ENGINE_AUTO,
    VALID_ENGINES,
    collect_pdfs,
    convert_pdf_to_docx,
)

logger = logging.getLogger("pdf_to_doc")


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pdf-to-doc",
        description=(
            "Batch convert PDF files to .docx. Accepts either a folder "
            "(non-recursive) or an explicit list of PDF files. Each output "
            "is written next to its source PDF with the same base name."
        ),
    )
    parser.add_argument(
        "inputs", nargs="+", type=Path,
        help="One folder OR one-or-more PDF file paths.",
    )
    parser.add_argument(
        "-r", "--recursive", action="store_true",
        help="When an input is a folder, also scan its subfolders for PDFs.",
    )
    parser.add_argument(
        "--overwrite", action="store_true",
        help="Overwrite existing .docx files instead of creating 'name (1).docx'.",
    )
    parser.add_argument(
        "--engine", choices=VALID_ENGINES, default=ENGINE_AUTO,
        help=(
            "Conversion engine. 'word' = MS Word COM (best editable output, "
            "Windows only). 'pdf2docx' = Python layout-aware. 'text' = plain "
            "text. 'auto' (default) picks 'word' if available else 'pdf2docx'."
        ),
    )
    parser.add_argument(
        "--ocr", action="store_true",
        help="Run OCR (Tesseract) for PDFs with no text layer (scanned).",
    )
    parser.add_argument(
        "--ocr-language", default="eng",
        help="Tesseract language code(s), e.g. 'eng', 'eng+ara' (default: eng).",
    )
    parser.add_argument(
        "--ocr-dpi", type=int, default=300,
        help="Rendering DPI for OCR (default: 300).",
    )
    parser.add_argument(
        "--tessdata", default=None,
        help="Path to tessdata directory (else TESSDATA_PREFIX or a default).",
    )
    parser.add_argument(
        "-v", "--verbose", action="store_true",
        help="Enable debug logging.",
    )
    return parser


def _configure_logging(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )


def main(argv: list[str] | None = None) -> int:
    args = _build_parser().parse_args(argv)
    _configure_logging(args.verbose)

    pdfs = collect_pdfs(args.inputs, recursive=args.recursive)
    if not pdfs:
        logger.error("No PDF files to process.")
        return 1

    logger.info("Found %d PDF file(s) to convert.", len(pdfs))
    succeeded = failed = 0
    for pdf in pdfs:
        try:
            convert_pdf_to_docx(
                pdf,
                overwrite=args.overwrite,
                engine=args.engine,
                ocr=args.ocr,
                ocr_language=args.ocr_language,
                ocr_dpi=args.ocr_dpi,
                tessdata=args.tessdata,
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
