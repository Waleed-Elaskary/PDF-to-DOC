"""Convert .odt files to .docx using headless LibreOffice, then strip white
background rectangles left over from PDF-to-ODT conversion.

The white rectangles are page-sized shapes with a solid white (or near-white)
fill that LibreOffice inserts as a background layer when importing PDFs.  They
appear as opaque white boxes in Word and hide content underneath.  This module
removes them automatically after conversion (opt out with --keep-bg).
"""
from __future__ import annotations

import copy
import logging
import os
import re
import shutil
import subprocess
import tempfile
from lxml import etree
from pathlib import Path
from typing import Iterable

from docx import Document

logger = logging.getLogger(__name__)

# Reuse soffice discovery from the LO converter.
from .lo_converter import find_soffice, _get_lo_profile

# XML namespaces used in .docx
_NS = {
    "w":   "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "mc":  "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "wp":  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "v":   "urn:schemas-microsoft-com:vml",
    "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

# Colours considered "white" (hex, case-insensitive)
_WHITE_COLOURS = {"ffffff", "white", "fefefe", "fdfdfd", "f8f8f8"}

# Minimum area (EMU^2) to consider a shape "page-sized".
# A4 ≈ 7560000 x 10692000 EMU → ~80 trillion EMU^2.  We use a generous lower
# bound so it catches letter, A4, and slightly smaller page rects.
_MIN_AREA_EMU2 = 30_000_000_000_000  # ~30 trillion


def _is_white_colour(fill_elem) -> bool:
    """Return True if the fill element represents a white / near-white colour."""
    # <a:srgbClr val="FFFFFF"/>
    for srgb in fill_elem.iter(f'{{{_NS["a"]}}}srgbClr'):
        val = (srgb.get("val") or "").lower()
        if val in _WHITE_COLOURS:
            return True
    # <a:schemeClr val="bg1"/>  or  <a:schemeClr val="lt1"/>
    for scheme in fill_elem.iter(f'{{{_NS["a"]}}}schemeClr'):
        val = (scheme.get("val") or "").lower()
        if val in ("bg1", "lt1"):
            return True
    return False


def _shape_is_white_rect(elem) -> bool:
    """Check if a wps:wsp or wps:spPr element describes a large white rect."""
    # Look for preset geometry = rectangle
    is_rect = False
    for prst in elem.iter(f'{{{_NS["a"]}}}prstGeom'):
        if prst.get("prst") in ("rect", "roundRect"):
            is_rect = True
            break
    if not is_rect:
        return False

    # Check for solid white fill
    has_white_fill = False
    for sf in elem.iter(f'{{{_NS["a"]}}}solidFill'):
        if _is_white_colour(sf):
            has_white_fill = True
            break
    if not has_white_fill:
        return False

    # Check dimensions — look for extent (cx, cy) in the drawing wrapper.
    # Walk up to the <w:drawing> or <mc:AlternateContent> and check <wp:extent>
    # or <a:ext> for large dimensions.
    for ext in elem.iter(f'{{{_NS["a"]}}}ext'):
        cx = int(ext.get("cx", "0"))
        cy = int(ext.get("cy", "0"))
        if cx * cy >= _MIN_AREA_EMU2:
            return True
    for ext in elem.iter(f'{{{_NS["wp"]}}}extent'):
        cx = int(ext.get("cx", "0"))
        cy = int(ext.get("cy", "0"))
        if cx * cy >= _MIN_AREA_EMU2:
            return True

    # If we can't determine size, still flag it if it's a white rect
    # (small white rects are unusual and likely also artifacts).
    return True


def _vml_is_white_rect(elem) -> bool:
    """Check if a VML v:rect element has white fill."""
    fillcolor = (elem.get("fillcolor") or "").lower().lstrip("#")
    if fillcolor in _WHITE_COLOURS:
        return True
    # Check child <v:fill> element
    for fill in elem.iter(f'{{{_NS["v"]}}}fill'):
        color = (fill.get("color") or "").lower().lstrip("#")
        if color in _WHITE_COLOURS:
            return True
    return False


def _scan_root_for_white_rects(root) -> list[tuple]:
    """Scan an XML root element for white background rectangles.

    Returns a list of (parent, child) tuples to remove.
    """
    to_remove: list[tuple] = []

    # Modern DrawingML shapes
    for drawing in root.iter(f'{{{_NS["w"]}}}drawing'):
        for wsp in drawing.iter(f'{{{_NS["wps"]}}}wsp'):
            if _shape_is_white_rect(wsp):
                target = _find_removable_ancestor(drawing, root)
                if target is not None:
                    to_remove.append(target)
                break

    # Legacy VML shapes
    for pict in root.iter(f'{{{_NS["w"]}}}pict'):
        for rect in pict.iter(f'{{{_NS["v"]}}}rect'):
            if _vml_is_white_rect(rect):
                target = _find_removable_ancestor(pict, root)
                if target is not None:
                    to_remove.append(target)
                break

    # <mc:AlternateContent> wrappers
    for mc in root.iter(f'{{{_NS["mc"]}}}AlternateContent'):
        for wsp in mc.iter(f'{{{_NS["wps"]}}}wsp'):
            if _shape_is_white_rect(wsp):
                target = _find_removable_ancestor(mc, root)
                if target is not None:
                    to_remove.append(target)
                break
        for rect in mc.iter(f'{{{_NS["v"]}}}rect'):
            if _vml_is_white_rect(rect):
                target = _find_removable_ancestor(mc, root)
                if target is not None:
                    to_remove.append(target)
                break

    return to_remove


def strip_white_rectangles(docx_path: Path) -> int:
    """Open a .docx, remove white background rectangles from ALL pages
    (body, headers, and footers), save in place.

    Returns the number of shapes removed.
    """
    doc = Document(str(docx_path))
    removed = 0

    # Collect all XML roots to scan: document body + every header/footer part.
    roots_to_scan = [doc.element.body]

    for section in doc.sections:
        for hf in (
            section.header, section.footer,
            section.first_page_header, section.first_page_footer,
            section.even_page_header, section.even_page_footer,
        ):
            try:
                if hf is not None and hf._element is not None:
                    roots_to_scan.append(hf._element)
            except Exception:
                pass

    to_remove: list[tuple] = []
    for root in roots_to_scan:
        to_remove.extend(_scan_root_for_white_rects(root))

    # Deduplicate (same element could be found via multiple paths).
    seen_ids: set[int] = set()
    for parent, child in to_remove:
        eid = id(child)
        if eid in seen_ids:
            continue
        seen_ids.add(eid)
        try:
            parent.remove(child)
            removed += 1
        except ValueError:
            pass  # Already removed via a parent.

    if removed:
        doc.save(str(docx_path))

    return removed


def _find_removable_ancestor(elem, body):
    """Walk up from elem to find the best ancestor to remove.

    Prefers: <mc:AlternateContent> > <w:r> > <w:p> (if paragraph has no text).
    Returns (parent, child_to_remove) or None.
    """
    # Build ancestor chain.
    parent_map = {c: p for p in body.iter() for c in p}
    current = elem
    while current is not None:
        parent = parent_map.get(current)
        if parent is None:
            break
        tag = current.tag.split("}")[-1] if "}" in current.tag else current.tag
        if tag == "AlternateContent":
            return (parent, current)
        if tag == "r":
            return (parent, current)
        if tag == "drawing":
            return (parent, current)
        current = parent
    # Fallback: remove the element from its direct parent.
    parent = parent_map.get(elem)
    if parent is not None:
        return (parent, elem)
    return None


# --- Conversion via LibreOffice -----------------------------------------------

def resolve_output_path(
    odt_path: Path, *, overwrite: bool, prefix: str = "",
) -> Path:
    new_stem = f"{prefix}{odt_path.stem}"
    base = odt_path.with_name(f"{new_stem}.docx")
    if overwrite or not base.exists():
        return base
    counter = 1
    while True:
        candidate = base.with_name(f"{new_stem} ({counter}).docx")
        if not candidate.exists():
            return candidate
        counter += 1


def _run_libreoffice(
    soffice: str,
    odt_path: Path,
    outdir: Path,
    *,
    timeout: int = 300,
) -> Path:
    profile = _get_lo_profile()
    profile_uri = profile.resolve().as_uri()

    attempts = [
        [soffice, "--headless", "--norestore", "--nolockcheck",
         f"-env:UserInstallation={profile_uri}",
         "--convert-to", "docx",
         "--outdir", str(outdir),
         str(odt_path)],
        [soffice, "--headless", "--norestore", "--nolockcheck",
         f"-env:UserInstallation={profile_uri}",
         "--convert-to", "docx:MS Word 2007 XML",
         "--outdir", str(outdir),
         str(odt_path)],
    ]

    result = None
    for cmd in attempts:
        logger.debug("Running: %s", " ".join(cmd))
        result = subprocess.run(
            cmd, capture_output=True, text=True, timeout=timeout,
        )
        logger.debug("stdout: %s", result.stdout.strip())
        logger.debug("stderr: %s", result.stderr.strip())

        expected = outdir / odt_path.with_suffix(".docx").name
        if expected.exists():
            return expected

        candidates = list(outdir.glob("*.docx"))
        if candidates:
            return candidates[0]

    stdout = result.stdout.strip() if result else ""
    stderr = result.stderr.strip() if result else ""
    raise FileNotFoundError(
        f"LibreOffice did not produce a .docx file.\n"
        f"Input: {odt_path}\n"
        f"stdout: {stdout}\n"
        f"stderr: {stderr}\n"
        f"Tip: close any running LibreOffice windows and retry."
    )


def convert_odt_to_docx(
    odt_path: Path,
    *,
    soffice: str,
    overwrite: bool = False,
    prefix: str = "",
    strip_bg: bool = True,
    timeout: int = 300,
) -> Path:
    """Convert one .odt to .docx via LibreOffice, optionally strip white rects.

    Returns the path of the created .docx.
    """
    odt_path = Path(odt_path)
    if not odt_path.is_file():
        raise FileNotFoundError(f"ODT not found: {odt_path}")

    output_path = resolve_output_path(odt_path, overwrite=overwrite, prefix=prefix)
    logger.info("Converting %s -> %s", odt_path.name, output_path.name)

    parent_dir = odt_path.parent
    try:
        tmp_ctx = tempfile.TemporaryDirectory(dir=parent_dir, prefix=".lo_tmp_")
        tmp_path = Path(tmp_ctx.name)
    except OSError:
        tmp_ctx = tempfile.TemporaryDirectory(prefix="lo_tmp_")
        tmp_path = Path(tmp_ctx.name)

    try:
        produced = _run_libreoffice(soffice, odt_path, tmp_path, timeout=timeout)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(produced), str(output_path))
    finally:
        tmp_ctx.cleanup()

    if strip_bg:
        try:
            removed = strip_white_rectangles(output_path)
            if removed:
                logger.info(
                    "Removed %d white background rectangle(s) from %s",
                    removed, output_path.name,
                )
        except Exception as exc:
            logger.warning(
                "Could not strip background shapes from %s: %s",
                output_path.name, exc,
            )

    logger.info("Wrote %s", output_path.name)
    return output_path


def collect_odt(
    inputs: Iterable[Path],
    *,
    recursive: bool = False,
) -> list[Path]:
    """Resolve CLI inputs into a list of .odt files."""
    resolved: list[Path] = []
    seen: set[Path] = set()
    for item in inputs:
        p = Path(item)
        if p.is_dir():
            iterator = p.rglob("*") if recursive else p.iterdir()
            for child in sorted(iterator):
                if child.is_file() and child.suffix.lower() == ".odt":
                    rp = child.resolve()
                    if rp not in seen:
                        seen.add(rp)
                        resolved.append(child)
        elif p.is_file():
            if p.suffix.lower() != ".odt":
                logger.warning("Skipping non-ODT file: %s", p)
                continue
            rp = p.resolve()
            if rp not in seen:
                seen.add(rp)
                resolved.append(p)
        else:
            logger.warning("Input does not exist: %s", p)
    return resolved
