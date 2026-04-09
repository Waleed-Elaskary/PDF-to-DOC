"""Remove specific objects from .odt files using a "remove template".

Given a template .odt that contains objects (images, shapes, lines, text
blocks), find and remove all matching objects from target .odt files in a
folder.  Matching is *fuzzy* — designed to work across files produced by
independent PDF-to-ODT conversions where exact positions, sizes and text
can vary slightly.

Matching strategies:

  - **Text frames**: case-insensitive keyword overlap between template and
    target text-box content.
  - **Image frames**: approximate position + approximate size (within
    configurable tolerance).
  - **Page-size rectangles**: any polygon/rect whose dimensions are close to
    US-Letter (8.5 × 11 in) is treated as a white background and removed.
  - **Thin-line polygons**: matched by approximate height (< 0.1 in).
  - **Small decorative frames**: empty text-boxes matched by approximate size.

Objects are removed from both the document body (content.xml) AND
headers/footers (styles.xml) across ALL pages.
"""
from __future__ import annotations

import hashlib
import logging
import os
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

logger = logging.getLogger(__name__)

_NS = {
    "office":   "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "style":    "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    "text":     "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "table":    "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "draw":     "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
    "fo":       "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
    "xlink":    "http://www.w3.org/1999/xlink",
    "svg":      "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
    "manifest": "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0",
    "loext":    "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0",
}

_DEFAULT_PATTERN = r"^.+\.odt$"

_DRAW_NS = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
_SVG_NS  = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
_XLINK_NS = "http://www.w3.org/1999/xlink"
_TEXT_NS = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"

# Fuzzy tolerances
_POS_TOLERANCE_IN = 0.5     # ±0.5 inch for position matching
_SIZE_TOLERANCE_IN = 0.5    # ±0.5 inch for size matching
_PAGE_W = 8.5               # US Letter width
_PAGE_H = 11.0              # US Letter height
_PAGE_TOL = 0.15            # tolerance for page-size detection
_THIN_THRESHOLD = 0.1       # polygons thinner than this = decorative line


# ---------------------------------------------------------------------------
# Dimension helpers
# ---------------------------------------------------------------------------

def _to_inches(val: str | None) -> float | None:
    """Parse an ODF dimension string to inches."""
    if val is None:
        return None
    val = val.strip()
    try:
        if val.endswith("in"):
            return float(val[:-2])
        if val.endswith("cm"):
            return float(val[:-2]) / 2.54
        if val.endswith("mm"):
            return float(val[:-2]) / 25.4
        if val.endswith("pt"):
            return float(val[:-2]) / 72.0
        if val.endswith("pc"):
            return float(val[:-2]) / 6.0
        return float(val)  # assume inches
    except ValueError:
        return None


def _is_page_size(w: float | None, h: float | None) -> bool:
    """True if dimensions ≈ US Letter (8.5 × 11 in)."""
    if w is None or h is None:
        return False
    return abs(w - _PAGE_W) < _PAGE_TOL and abs(h - _PAGE_H) < _PAGE_TOL


def _approx_eq(a: float | None, b: float | None, tol: float) -> bool:
    """True if two values are within tolerance."""
    if a is None or b is None:
        return False
    return abs(a - b) <= tol


# ---------------------------------------------------------------------------
# Text helpers
# ---------------------------------------------------------------------------

def _get_text(elem) -> str:
    """Extract all text from an element, normalised."""
    parts = []
    for node in elem.iter():
        if node.text:
            parts.append(node.text.strip())
        if node.tail:
            parts.append(node.tail.strip())
    return " ".join(p for p in parts if p)


def _extract_keywords(text: str) -> set[str]:
    """Extract meaningful keywords (lowercase, 3+ chars) from text."""
    words = re.findall(r'[a-zA-Z0-9]+', text.lower())
    # Skip very short words and common ones
    skip = {"the", "and", "for", "are", "but", "not", "you", "all",
            "can", "had", "her", "was", "one", "our", "out", "com",
            "www", "tel", "fax", "address"}
    return {w for w in words if len(w) >= 3 and w not in skip}


def _text_match(tmpl_keywords: set[str], target_text: str) -> bool:
    """Fuzzy text match: ≥50% of template keywords found in target."""
    if not tmpl_keywords:
        return False
    target_lower = target_text.lower()
    hits = sum(1 for kw in tmpl_keywords if kw in target_lower)
    ratio = hits / len(tmpl_keywords)
    return ratio >= 0.5


# ---------------------------------------------------------------------------
# Template signature collection
# ---------------------------------------------------------------------------

class _FrameSig:
    """Signature for a draw:frame in the template."""
    __slots__ = ("x", "y", "w", "h", "has_image", "has_textbox",
                 "text_keywords", "image_hash")

    def __init__(self, *, x=None, y=None, w=None, h=None,
                 has_image=False, has_textbox=False,
                 text_keywords=None, image_hash=None):
        self.x = x
        self.y = y
        self.w = w
        self.h = h
        self.has_image = has_image
        self.has_textbox = has_textbox
        self.text_keywords = text_keywords or set()
        self.image_hash = image_hash

    def __repr__(self):
        kind = "img" if self.has_image else ("txt" if self.has_textbox else "empty")
        return f"FrameSig({kind}, pos=({self.x:.2f},{self.y:.2f}), {self.w:.2f}x{self.h:.2f})"


class _PolySig:
    """Signature for a draw:polygon in the template."""
    __slots__ = ("w", "h", "points", "is_page_size", "is_thin")

    def __init__(self, *, w=None, h=None, points=None):
        self.w = w
        self.h = h
        self.points = points
        self.is_page_size = _is_page_size(w, h)
        self.is_thin = h is not None and h < _THIN_THRESHOLD


class _TemplateSigs:
    """All signatures collected from the remove template."""
    def __init__(self):
        self.frames: list[_FrameSig] = []
        self.polygons: list[_PolySig] = []
        self.has_page_rect: bool = False
        self.has_thin_lines: bool = False
        # Collect all text keywords for broad text matching
        self.all_text_keywords: set[str] = set()
        self.text_sigs: list[set[str]] = []  # per-frame keywords
        self.image_hashes: set[str] = set()

    def __repr__(self):
        return (f"TemplateSigs(frames={len(self.frames)}, "
                f"polys={len(self.polygons)}, page_rect={self.has_page_rect})")


def _collect_sigs_from_xml(xml_root, zf: zipfile.ZipFile, sigs: _TemplateSigs):
    """Collect frame and polygon signatures from one XML file."""
    draw_frame_tag = f'{{{_DRAW_NS}}}frame'
    draw_image_tag = f'{{{_DRAW_NS}}}image'
    draw_textbox_tag = f'{{{_DRAW_NS}}}text-box'
    xlink_href = f'{{{_XLINK_NS}}}href'
    svg_x = f'{{{_SVG_NS}}}x'
    svg_y = f'{{{_SVG_NS}}}y'
    svg_w = f'{{{_SVG_NS}}}width'
    svg_h = f'{{{_SVG_NS}}}height'
    draw_pts = f'{{{_DRAW_NS}}}points'

    # Frames
    for frame in xml_root.iter(draw_frame_tag):
        x = _to_inches(frame.get(svg_x))
        y = _to_inches(frame.get(svg_y))
        w = _to_inches(frame.get(svg_w))
        h = _to_inches(frame.get(svg_h))
        if x is None: x = 0.0
        if y is None: y = 0.0
        if w is None: w = 0.0
        if h is None: h = 0.0

        fs = _FrameSig(x=x, y=y, w=w, h=h)

        # Check children
        for child in frame:
            local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if local == "text-box":
                fs.has_textbox = True
                text = _get_text(child)
                if text:
                    kw = _extract_keywords(text)
                    fs.text_keywords = kw
                    sigs.all_text_keywords |= kw
                    sigs.text_sigs.append(kw)
                break
            elif local == "image":
                fs.has_image = True
                href = child.get(xlink_href)
                if href and href in zf.namelist():
                    ih = hashlib.sha256(zf.read(href)).hexdigest()
                    fs.image_hash = ih
                    sigs.image_hashes.add(ih)
                break

        sigs.frames.append(fs)
        logger.debug("Template frame: %s", fs)

    # Polygons / polylines
    for tag_name in ("polygon", "polyline"):
        full_tag = f'{{{_DRAW_NS}}}{tag_name}'
        for shape in xml_root.iter(full_tag):
            w = _to_inches(shape.get(svg_w))
            h = _to_inches(shape.get(svg_h))
            pts = shape.get(draw_pts)
            ps = _PolySig(w=w, h=h, points=pts)
            sigs.polygons.append(ps)
            if ps.is_page_size:
                sigs.has_page_rect = True
            if ps.is_thin:
                sigs.has_thin_lines = True
            logger.debug("Template polygon: w=%s h=%s page=%s thin=%s",
                         w, h, ps.is_page_size, ps.is_thin)


def collect_template_sigs(template_path: Path) -> _TemplateSigs:
    """Collect all signatures from the remove template."""
    sigs = _TemplateSigs()
    with zipfile.ZipFile(template_path, "r") as zf:
        for xml_name in ("content.xml", "styles.xml"):
            if xml_name in zf.namelist():
                root = etree.fromstring(zf.read(xml_name))
                _collect_sigs_from_xml(root, zf, sigs)
    return sigs


# ---------------------------------------------------------------------------
# Removal logic — scan target and remove matching objects
# ---------------------------------------------------------------------------

def _remove_from_xml(
    xml_root,
    target_zf: zipfile.ZipFile,
    sigs: _TemplateSigs,
    *,
    remove_page_bg: bool = True,
) -> int:
    """Remove matching objects from an XML root. Returns count removed."""
    draw_frame_tag = f'{{{_DRAW_NS}}}frame'
    draw_image_tag = f'{{{_DRAW_NS}}}image'
    draw_textbox_tag = f'{{{_DRAW_NS}}}text-box'
    xlink_href = f'{{{_XLINK_NS}}}href'
    svg_x = f'{{{_SVG_NS}}}x'
    svg_y = f'{{{_SVG_NS}}}y'
    svg_w = f'{{{_SVG_NS}}}width'
    svg_h = f'{{{_SVG_NS}}}height'
    draw_pts = f'{{{_DRAW_NS}}}points'

    parent_map: dict = {c: p for p in xml_root.iter() for c in p}
    to_remove: list[tuple] = []
    removed_ids: set[int] = set()

    def _mark(elem):
        """Walk up to find top-level draw ancestor and mark for removal."""
        current = elem
        best = None
        while current is not None:
            par = parent_map.get(current)
            if par is None:
                break
            local = current.tag.split("}")[-1] if "}" in current.tag else current.tag
            if local in ("frame", "line", "rect", "circle", "ellipse",
                         "polygon", "polyline", "path", "custom-shape",
                         "connector"):
                best = (par, current)
            current = par
        if best:
            to_remove.append(best)
        elif parent_map.get(elem) is not None:
            to_remove.append((parent_map[elem], elem))

    # --- 1. Match draw:frame elements ---
    for frame in xml_root.iter(draw_frame_tag):
        tx = _to_inches(frame.get(svg_x)) or 0.0
        ty = _to_inches(frame.get(svg_y)) or 0.0
        tw = _to_inches(frame.get(svg_w)) or 0.0
        th = _to_inches(frame.get(svg_h)) or 0.0
        matched = False

        # 1a. Text-box: fuzzy keyword matching
        for child in frame:
            local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if local == "text-box":
                text = _get_text(child)
                if text:
                    # Check against each template text signature
                    for tmpl_kw in sigs.text_sigs:
                        if _text_match(tmpl_kw, text):
                            logger.debug("Text match: '%s'", text[:60])
                            _mark(frame)
                            matched = True
                            break
                elif not text:
                    # Empty text-box: match by approximate position+size
                    for fs in sigs.frames:
                        if (fs.has_textbox and not fs.text_keywords and
                                _approx_eq(tw, fs.w, _SIZE_TOLERANCE_IN) and
                                _approx_eq(th, fs.h, _SIZE_TOLERANCE_IN)):
                            logger.debug("Empty frame match at (%.2f,%.2f)", tx, ty)
                            _mark(frame)
                            matched = True
                            break
                break

            elif local == "image":
                # 1b. Image: hash match first, then position+size match
                href = child.get(xlink_href)
                if href and href in target_zf.namelist():
                    img_hash = hashlib.sha256(target_zf.read(href)).hexdigest()
                    if img_hash in sigs.image_hashes:
                        logger.debug("Image hash match: %s", href)
                        _mark(frame)
                        matched = True
                        break

                # Fallback: approximate position + size match
                for fs in sigs.frames:
                    if fs.has_image:
                        if (_approx_eq(tx, fs.x, _POS_TOLERANCE_IN) and
                                _approx_eq(ty, fs.y, _POS_TOLERANCE_IN) and
                                _approx_eq(tw, fs.w, _SIZE_TOLERANCE_IN) and
                                _approx_eq(th, fs.h, _SIZE_TOLERANCE_IN)):
                            logger.debug("Image pos+size match at (%.2f,%.2f) "
                                         "%.2fx%.2f", tx, ty, tw, th)
                            _mark(frame)
                            matched = True
                            break
                break

        if matched:
            continue

    # --- 2. Match polygons ---
    for tag_name in ("polygon", "polyline"):
        full_tag = f'{{{_DRAW_NS}}}{tag_name}'
        for shape in xml_root.iter(full_tag):
            w = _to_inches(shape.get(svg_w))
            h = _to_inches(shape.get(svg_h))

            # 2a. Page-size white background rectangle
            # Always remove if remove_page_bg is on, OR if template had one
            if _is_page_size(w, h) and (remove_page_bg or sigs.has_page_rect):
                logger.debug("Page-size polygon: %.2fx%.2f", w or 0, h or 0)
                _mark(shape)
                continue

            # 2b. Thin decorative lines
            if sigs.has_thin_lines and h is not None and h < _THIN_THRESHOLD:
                # Check if any template polygon is similarly thin
                for ps in sigs.polygons:
                    if ps.is_thin:
                        logger.debug("Thin line match: h=%.4f", h)
                        _mark(shape)
                        break
                continue

            # 2c. Exact points match (in case same conversion)
            pts = shape.get(draw_pts)
            if pts:
                for ps in sigs.polygons:
                    if ps.points == pts:
                        _mark(shape)
                        break

    # --- 3. Other shapes (rect, custom-shape, etc.) ---
    other_tags = ("line", "rect", "circle", "ellipse", "path",
                  "custom-shape", "connector")
    for tag_name in other_tags:
        full_tag = f'{{{_DRAW_NS}}}{tag_name}'
        for shape in xml_root.iter(full_tag):
            w = _to_inches(shape.get(svg_w))
            h = _to_inches(shape.get(svg_h))
            if _is_page_size(w, h) and (remove_page_bg or sigs.has_page_rect):
                logger.debug("Page-size %s: %.2fx%.2f", tag_name, w or 0, h or 0)
                _mark(shape)

    # --- Deduplicate and remove ---
    removed = 0
    for parent, child in to_remove:
        eid = id(child)
        if eid in removed_ids:
            continue
        removed_ids.add(eid)
        try:
            parent.remove(child)
            removed += 1
        except ValueError:
            pass

    return removed


# ---------------------------------------------------------------------------
# File-level operations
# ---------------------------------------------------------------------------

def _output_filename(original: str, suffix_num: str = "001") -> str:
    """Append a number before the extension.

    "EBB-Report-002.odt" -> "EBB-Report-002-001.odt"
    """
    stem, ext = os.path.splitext(original)
    return f"{stem}-{suffix_num}{ext}"


def collect_matching_odt(
    folder: Path,
    pattern: str,
    *,
    recursive: bool = False,
) -> list[Path]:
    """Find .odt files matching a regex pattern."""
    folder = Path(folder)
    iterator = folder.rglob("*") if recursive else folder.iterdir()
    results: list[Path] = []
    seen: set[Path] = set()
    for p in sorted(iterator):
        if not (p.is_file() and p.suffix.lower() == ".odt"):
            continue
        if p.name.startswith("~"):
            continue
        if not re.match(pattern, p.name, re.IGNORECASE):
            continue
        rp = p.resolve()
        if rp not in seen:
            seen.add(rp)
            results.append(p)
    return results


def remove_objects_from_odt(
    template_path: Path,
    target_path: Path,
    output_path: Path,
    *,
    remove_page_bg: bool = True,
) -> int:
    """Remove template objects from target .odt, save to output_path.

    Returns total number of objects removed.
    """
    template_path = Path(template_path)
    target_path = Path(target_path)
    output_path = Path(output_path)

    # 1. Collect signatures from the template.
    sigs = collect_template_sigs(template_path)
    logger.info("Template: %s", sigs)

    total_sigs = (len(sigs.frames) + len(sigs.polygons))
    if total_sigs == 0:
        logger.warning("No removable objects found in template.")

    # 2. Copy target to temp, modify, save.
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        suffix=".odt", delete=False, dir=output_path.parent,
    ) as tmp:
        tmp_name = tmp.name

    try:
        shutil.copy2(target_path, tmp_name)
        total_removed = 0

        with zipfile.ZipFile(tmp_name, "r") as zf_target:
            target_content = etree.fromstring(zf_target.read("content.xml"))
            target_styles = etree.fromstring(zf_target.read("styles.xml"))

            # Remove from content.xml (document body — all pages).
            r1 = _remove_from_xml(target_content, zf_target, sigs,
                                  remove_page_bg=remove_page_bg)
            logger.debug("Removed %d objects from content.xml", r1)

            # Remove from styles.xml (headers/footers — all pages).
            r2 = _remove_from_xml(target_styles, zf_target, sigs,
                                  remove_page_bg=remove_page_bg)
            logger.debug("Removed %d objects from styles.xml", r2)
            total_removed = r1 + r2

        new_content = etree.tostring(
            target_content, xml_declaration=True, encoding="UTF-8",
        )
        new_styles = etree.tostring(
            target_styles, xml_declaration=True, encoding="UTF-8",
        )

        # 3. Rebuild the ODT properly.
        final_tmp = tmp_name + ".new"
        with zipfile.ZipFile(tmp_name, "r") as zf_in, \
             zipfile.ZipFile(final_tmp, "w") as zf_out:

            # mimetype first, uncompressed.
            if "mimetype" in zf_in.namelist():
                zf_out.writestr(
                    zipfile.ZipInfo("mimetype"),
                    zf_in.read("mimetype"),
                    compress_type=zipfile.ZIP_STORED,
                )

            for item in zf_in.namelist():
                if item == "mimetype":
                    continue
                elif item == "content.xml":
                    zf_out.writestr(
                        zipfile.ZipInfo("content.xml"),
                        new_content,
                        compress_type=zipfile.ZIP_DEFLATED,
                    )
                elif item == "styles.xml":
                    zf_out.writestr(
                        zipfile.ZipInfo("styles.xml"),
                        new_styles,
                        compress_type=zipfile.ZIP_DEFLATED,
                    )
                else:
                    info = zf_in.getinfo(item)
                    zf_out.writestr(info, zf_in.read(item))

        os.replace(final_tmp, tmp_name)
        shutil.move(tmp_name, str(output_path))
        return total_removed

    except Exception:
        for f in (tmp_name, tmp_name + ".new"):
            try:
                os.unlink(f)
            except OSError:
                pass
        raise


def process_folder(
    template_path: Path,
    folder: Path,
    *,
    pattern: str = _DEFAULT_PATTERN,
    recursive: bool = False,
    overwrite: bool = False,
    suffix_num: str = "001",
    remove_page_bg: bool = True,
) -> tuple[int, int, int, int]:
    """Remove template objects from all matching .odt files in a folder.

    Returns (found, succeeded, failed, total_objects_removed).
    """
    template_path = Path(template_path)
    folder = Path(folder)

    if not template_path.is_file():
        raise FileNotFoundError(f"Template not found: {template_path}")
    if not folder.is_dir():
        raise FileNotFoundError(f"Folder not found: {folder}")

    template_resolved = template_path.resolve()
    matches = collect_matching_odt(folder, pattern, recursive=recursive)
    matches = [m for m in matches if m.resolve() != template_resolved]

    if not matches:
        logger.warning("No .odt files matching pattern '%s' in %s", pattern, folder)
        return (0, 0, 0, 0)

    logger.info("Found %d matching .odt file(s).", len(matches))

    succeeded = 0
    failed = 0
    total_removed = 0
    for odt_path in matches:
        try:
            new_name = _output_filename(odt_path.name, suffix_num)
            output_path = odt_path.parent / new_name

            if not overwrite and output_path.exists():
                logger.info("Skipping (exists): %s", output_path.name)
                continue

            logger.info("%s -> %s", odt_path.name, new_name)
            removed = remove_objects_from_odt(
                template_path, odt_path, output_path,
                remove_page_bg=remove_page_bg,
            )
            total_removed += removed
            logger.info("Wrote %s (removed %d object(s))", output_path.name, removed)
            succeeded += 1
        except Exception as exc:
            failed += 1
            logger.error("Failed on %s: %s", odt_path.name, exc)

    return (len(matches), succeeded, failed, total_removed)
