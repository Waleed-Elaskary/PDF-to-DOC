"""Smoke test: generate a couple of sample PDFs and convert them.

Usage:
    python scripts/smoke_test.py
"""
from __future__ import annotations

import logging
import sys
import tempfile
from pathlib import Path

import fitz

# Make the package importable when run from the repo root.
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from pdf_to_doc.converter import collect_pdfs, convert_pdf_to_docx  # noqa: E402


def make_sample_pdf(path: Path, pages: list[str]) -> None:
    doc = fitz.open()
    for text in pages:
        page = doc.new_page()
        page.insert_text((72, 72), text, fontsize=12)
    doc.save(path)
    doc.close()


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        make_sample_pdf(tmp_path / "alpha.pdf", ["Alpha page 1", "Alpha page 2"])
        make_sample_pdf(tmp_path / "beta.pdf", ["Beta only page"])

        pdfs = collect_pdfs([tmp_path])
        assert len(pdfs) == 2, pdfs

        for pdf in pdfs:
            out = convert_pdf_to_docx(pdf)
            assert out.exists(), out
            print(f"OK: {pdf.name} -> {out.name}")

        # Collision check
        again = convert_pdf_to_docx(tmp_path / "alpha.pdf")
        assert again.name == "alpha (1).docx", again
        print(f"OK collision: {again.name}")

    print("Smoke test passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
