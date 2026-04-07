"""Replace headers and footers in .docx files using a template .docx.

Engines:
  word    MS Word COM automation (Windows). Handles images, tables, fields,
          page numbers -- recommended.
  python  Pure python-docx. Copies paragraphs/tables from template headers
          and footers. Does NOT copy images.

The template's FIRST section's headers/footers are applied to EVERY section
of each target document (primary header/footer, plus first-page and even-page
variants if present in the template).
"""
from __future__ import annotations

import copy
import logging
import shutil
import sys
from pathlib import Path
from typing import Iterable

logger = logging.getLogger(__name__)

ENGINE_AUTO = "auto"
ENGINE_WORD = "word"
ENGINE_PYTHON = "python"
VALID_ENGINES = (ENGINE_AUTO, ENGINE_WORD, ENGINE_PYTHON)

# Word constants
_WD_HEADER_FOOTER_PRIMARY = 1
_WD_HEADER_FOOTER_FIRST_PAGE = 2
_WD_HEADER_FOOTER_EVEN_PAGES = 3
_WD_FORMAT_DOCX = 16
_WD_SEEK_CURRENT = 0


def _word_available() -> bool:
    if sys.platform != "win32":
        return False
    try:
        import win32com.client  # noqa: F401
        import pythoncom  # noqa: F401
    except ImportError:
        return False
    try:
        import win32com.client as w
        app = w.DispatchEx("Word.Application")
        app.Quit()
        return True
    except Exception:
        return False


def _pick_engine(requested: str) -> str:
    if requested != ENGINE_AUTO:
        return requested
    if _word_available():
        return ENGINE_WORD
    return ENGINE_PYTHON


# --- Discovery ---------------------------------------------------------------

def collect_docx(folder: Path, *, recursive: bool) -> list[Path]:
    """Find all .docx files in folder (optionally recursive).

    Skips Word lock/temp files (~$*.docx).
    """
    it = folder.rglob("*") if recursive else folder.iterdir()
    out: list[Path] = []
    for p in sorted(it):
        if not (p.is_file() and p.suffix.lower() == ".docx"):
            continue
        if p.name.startswith("~$"):
            continue
        out.append(p)
    return out


# --- Engine: Word COM --------------------------------------------------------

def _word_replace_all(
    template_path: Path,
    targets: Iterable[Path],
) -> tuple[int, int]:
    """Open each target in Word and copy headers/footers from the template.

    Returns (succeeded, failed).
    """
    import pythoncom
    import win32com.client as w

    pythoncom.CoInitialize()
    app = None
    template_doc = None
    succeeded = 0
    failed = 0
    try:
        app = w.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0

        template_doc = app.Documents.Open(
            FileName=str(template_path.resolve()),
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
            ConfirmConversions=False,
        )
        template_section = template_doc.Sections(1)
        template_page_setup = template_section.PageSetup

        template_diff_first = bool(
            template_page_setup.DifferentFirstPageHeaderFooter
        )
        template_odd_even = bool(
            template_doc.PageSetup.OddAndEvenPagesHeaderFooter
        )

        all_variants = (
            _WD_HEADER_FOOTER_PRIMARY,
            _WD_HEADER_FOOTER_FIRST_PAGE,
            _WD_HEADER_FOOTER_EVEN_PAGES,
        )

        def _src_hf(collection, hf_type: int):
            """Pick the matching template variant; fall back to primary."""
            if hf_type == _WD_HEADER_FOOTER_FIRST_PAGE and not template_diff_first:
                return collection(_WD_HEADER_FOOTER_PRIMARY)
            if hf_type == _WD_HEADER_FOOTER_EVEN_PAGES and not template_odd_even:
                return collection(_WD_HEADER_FOOTER_PRIMARY)
            return collection(hf_type)

        logger.info(
            "Template variants: primary%s%s",
            ", first-page" if template_diff_first else "",
            ", even-pages" if template_odd_even else "",
        )

        for target_path in targets:
            target_doc = None
            try:
                logger.info("Updating %s", target_path.name)
                target_doc = app.Documents.Open(
                    FileName=str(target_path.resolve()),
                    ReadOnly=False,
                    AddToRecentFiles=False,
                    Visible=False,
                    ConfirmConversions=False,
                )

                # Mirror document-level even/odd flag on the target.
                try:
                    target_doc.PageSetup.OddAndEvenPagesHeaderFooter = (
                        template_odd_even
                    )
                except Exception:
                    pass

                for sec in target_doc.Sections:
                    # Mirror first-page flag on this section.
                    try:
                        sec.PageSetup.DifferentFirstPageHeaderFooter = (
                            template_diff_first
                        )
                    except Exception:
                        pass

                    # Break link-to-previous on EVERY variant so this section
                    # keeps its own copy.
                    for hf_type in all_variants:
                        try:
                            sec.Headers(hf_type).LinkToPrevious = False
                            sec.Footers(hf_type).LinkToPrevious = False
                        except Exception:
                            pass

                    # Overwrite EVERY variant slot in the target with template
                    # content. Variants the template doesn't use receive the
                    # template's primary content, so no old data remains
                    # visible on any page.
                    for hf_type in all_variants:
                        _copy_hf_range(
                            _src_hf(template_section.Headers, hf_type).Range,
                            sec.Headers(hf_type).Range,
                        )
                        _copy_hf_range(
                            _src_hf(template_section.Footers, hf_type).Range,
                            sec.Footers(hf_type).Range,
                        )

                target_doc.Save()
                succeeded += 1
            except Exception as exc:
                failed += 1
                logger.error("Failed on %s: %s", target_path, exc)
            finally:
                if target_doc is not None:
                    try:
                        target_doc.Close(SaveChanges=0)
                    except Exception:
                        pass
    finally:
        if template_doc is not None:
            try:
                template_doc.Close(SaveChanges=0)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
    return succeeded, failed


def _copy_hf_range(src_range, dst_range) -> None:
    """Overwrite dst header/footer range with src content.

    Direct assignment of FormattedText replaces the full content -- no manual
    Delete() needed (Delete collapses the range and breaks the assignment).
    """
    dst_range.FormattedText = src_range.FormattedText


# --- Engine: pure python-docx -----------------------------------------------

def _python_replace_all(
    template_path: Path,
    targets: Iterable[Path],
) -> tuple[int, int]:
    from docx import Document

    template_doc = Document(str(template_path))
    template_section = template_doc.sections[0]

    succeeded = 0
    failed = 0
    for target_path in targets:
        try:
            logger.info("Updating %s", target_path.name)
            target_doc = Document(str(target_path))
            for sec in target_doc.sections:
                _py_copy_hf(template_section.header, sec.header)
                _py_copy_hf(template_section.footer, sec.footer)
                # First-page and even variants (if enabled on the template)
                if template_section.different_first_page_header_footer:
                    sec.different_first_page_header_footer = True
                    _py_copy_hf(
                        template_section.first_page_header,
                        sec.first_page_header,
                    )
                    _py_copy_hf(
                        template_section.first_page_footer,
                        sec.first_page_footer,
                    )
            target_doc.save(str(target_path))
            succeeded += 1
        except Exception as exc:
            failed += 1
            logger.error("Failed on %s: %s", target_path, exc)
    return succeeded, failed


def _py_copy_hf(src_hf, dst_hf) -> None:
    """Copy paragraphs/tables from src header/footer to dst header/footer."""
    dst_hf.is_linked_to_previous = False
    # Remove existing paragraphs/tables in destination
    dst_elem = dst_hf._element
    for child in list(dst_elem):
        # Keep sectPr if present, remove p and tbl
        tag = child.tag.split("}", 1)[-1]
        if tag in ("p", "tbl"):
            dst_elem.remove(child)

    # Deep-copy source paragraphs/tables into destination
    src_elem = src_hf._element
    for child in src_elem:
        tag = child.tag.split("}", 1)[-1]
        if tag in ("p", "tbl"):
            dst_elem.append(copy.deepcopy(child))


# --- Orchestration -----------------------------------------------------------

def replace_headers_footers(
    template_path: Path,
    target_folder: Path,
    *,
    recursive: bool = True,
    engine: str = ENGINE_AUTO,
    backup: bool = True,
) -> tuple[int, int, int]:
    """Replace header/footer in every .docx under target_folder with those
    from template_path.

    Returns (found, succeeded, failed).
    """
    template_path = Path(template_path)
    target_folder = Path(target_folder)
    if not template_path.is_file() or template_path.suffix.lower() != ".docx":
        raise ValueError(f"Template must be a .docx file: {template_path}")
    if not target_folder.is_dir():
        raise ValueError(f"Target folder not found: {target_folder}")

    targets = collect_docx(target_folder, recursive=recursive)
    # Don't process the template itself if it happens to live under the folder.
    template_resolved = template_path.resolve()
    targets = [t for t in targets if t.resolve() != template_resolved]

    if not targets:
        logger.warning("No .docx files found under %s", target_folder)
        return (0, 0, 0)

    if backup:
        for t in targets:
            bak = t.with_suffix(t.suffix + ".bak")
            if not bak.exists():
                shutil.copy2(t, bak)

    chosen = _pick_engine(engine)
    logger.info("Engine: %s | %d target file(s)", chosen, len(targets))

    if chosen == ENGINE_WORD:
        ok, bad = _word_replace_all(template_path, targets)
    elif chosen == ENGINE_PYTHON:
        ok, bad = _python_replace_all(template_path, targets)
    else:
        raise ValueError(f"Unknown engine: {chosen}")

    return (len(targets), ok, bad)
