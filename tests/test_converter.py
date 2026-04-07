"""Minimal tests for pdf_to_doc.converter."""
from __future__ import annotations

from pathlib import Path

import fitz
import pytest
from docx import Document

from pdf_to_doc.converter import (
    collect_pdfs,
    convert_pdf_to_docx,
    resolve_output_path,
)


def _make_pdf(path: Path, pages: list[str]) -> None:
    doc = fitz.open()
    for text in pages:
        page = doc.new_page()
        page.insert_text((72, 72), text, fontsize=12)
    doc.save(path)
    doc.close()


def test_resolve_output_path_no_conflict(tmp_path: Path):
    pdf = tmp_path / "example.pdf"
    pdf.write_bytes(b"%PDF-1.4\n")
    out = resolve_output_path(pdf, overwrite=False)
    assert out == tmp_path / "example.docx"


def test_resolve_output_path_collisions(tmp_path: Path):
    pdf = tmp_path / "example.pdf"
    pdf.write_bytes(b"%PDF-1.4\n")
    (tmp_path / "example.docx").write_bytes(b"x")
    (tmp_path / "example (1).docx").write_bytes(b"x")

    out = resolve_output_path(pdf, overwrite=False)
    assert out.name == "example (2).docx"


def test_resolve_output_path_overwrite(tmp_path: Path):
    pdf = tmp_path / "example.pdf"
    pdf.write_bytes(b"%PDF-1.4\n")
    (tmp_path / "example.docx").write_bytes(b"x")

    out = resolve_output_path(pdf, overwrite=True)
    assert out.name == "example.docx"


def test_convert_pdf_to_docx_basic(tmp_path: Path):
    pdf = tmp_path / "hello.pdf"
    _make_pdf(pdf, ["Hello World", "Second Page"])

    out = convert_pdf_to_docx(pdf)
    assert out.exists()
    assert out.name == "hello.docx"

    doc = Document(out)
    all_text = "\n".join(p.text for p in doc.paragraphs)
    assert "Hello World" in all_text
    assert "Second Page" in all_text


def test_convert_pdf_to_docx_collision(tmp_path: Path):
    pdf = tmp_path / "hello.pdf"
    _make_pdf(pdf, ["One"])

    first = convert_pdf_to_docx(pdf)
    second = convert_pdf_to_docx(pdf)

    assert first.name == "hello.docx"
    assert second.name == "hello (1).docx"
    assert first.exists() and second.exists()


def test_collect_pdfs_folder_nonrecursive(tmp_path: Path):
    (tmp_path / "a.pdf").write_bytes(b"%PDF-1.4\n")
    (tmp_path / "b.PDF").write_bytes(b"%PDF-1.4\n")
    (tmp_path / "note.txt").write_text("nope")
    sub = tmp_path / "sub"
    sub.mkdir()
    (sub / "c.pdf").write_bytes(b"%PDF-1.4\n")

    found = collect_pdfs([tmp_path])
    names = sorted(p.name for p in found)
    assert names == ["a.pdf", "b.PDF"]


def test_collect_pdfs_folder_recursive(tmp_path: Path):
    (tmp_path / "a.pdf").write_bytes(b"%PDF-1.4\n")
    sub = tmp_path / "sub"
    sub.mkdir()
    (sub / "c.pdf").write_bytes(b"%PDF-1.4\n")
    deep = sub / "deeper"
    deep.mkdir()
    (deep / "d.pdf").write_bytes(b"%PDF-1.4\n")
    found = collect_pdfs([tmp_path], recursive=True)
    names = sorted(p.name for p in found)
    assert names == ["a.pdf", "c.pdf", "d.pdf"]


def test_collect_pdfs_dedup(tmp_path: Path):
    p = tmp_path / "a.pdf"
    p.write_bytes(b"%PDF-1.4\n")
    found = collect_pdfs([p, p, tmp_path])
    assert len(found) == 1


if __name__ == "__main__":
    raise SystemExit(pytest.main([__file__, "-v"]))
