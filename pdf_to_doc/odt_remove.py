"""Remove specific objects from .odt files using a "remove template".

Given a template .odt that contains objects (images, shapes, lines, text
blocks), find and remove all matching objects from target .odt files in a
folder. Matching is done by:

  - **Images**: SHA-256 hash of the actual image data inside the ODT archive.
  - **Shapes/lines**: draw:style-name + geometry attributes (coordinates, size).
  - **Text blocks**: exact text content comparison.

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
from copy import deepcopy
from pathlib import Path
from typing import Iterable

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


# ---------------------------------------------------------------------------
# Fingerprinting: build "signatures" for objects in the template so we can
# find them in targets.
# ---------------------------------------------------------------------------

class _ImageSig:
    """Signature for a draw:image — based on SHA-256 of the image bytes."""
    __slots__ = ("hash",)
    def __init__(self, h: str):
        self.hash = h
    def __eq__(self, other):
        return isinstance(other, _ImageSig) and self.hash == other.hash
    def __hash__(self):
        return hash(("img", self.hash))
    def __repr__(self):
        return f"ImageSig({self.hash[:12]}...)"


class _ShapeSig:
    """Signature for a draw:line / draw:rect / draw:custom-shape etc."""
    __slots__ = ("tag", "attrs",)
    def __init__(self, tag: str, attrs: frozenset):
        self.tag = tag
        self.attrs = attrs
    def __eq__(self, other):
        return isinstance(other, _ShapeSig) and self.tag == other.tag and self.attrs == other.attrs
    def __hash__(self):
        return hash(("shape", self.tag, self.attrs))
    def __repr__(self):
        return f"ShapeSig({self.tag}, {len(self.attrs)} attrs)"


class _TextSig:
    """Signature for a text block — normalised text content."""
    __slots__ = ("text",)
    def __init__(self, text: str):
        self.text = text
    def __eq__(self, other):
        return isinstance(other, _TextSig) and self.text == other.text
    def __hash__(self):
        return hash(("text", self.text))
    def __repr__(self):
        return f"TextSig({self.text[:40]}...)"


def _get_text_content(elem) -> str:
    """Extract all text from an element tree, normalised."""
    parts = []
    for node in elem.iter():
        if node.text:
            parts.append(node.text.strip())
        if node.tail:
            parts.append(node.tail.strip())
    return " ".join(p for p in parts if p)


def _geometry_attrs(elem) -> frozenset:
    """Extract geometry-related attributes from a shape element."""
    # Attributes that define position/size/geometry
    keys = []
    for attr_name in ("svg:x", "svg:y", "svg:x1", "svg:y1", "svg:x2", "svg:y2",
                       "svg:width", "svg:height", "svg:cx", "svg:cy", "svg:r",
                       "draw:style-name", "draw:text-style-name"):
        ns_prefix, local = attr_name.split(":")
        full = f'{{{_NS.get(ns_prefix, "")}}}{local}'
        val = elem.get(full)
        if val:
            keys.append((attr_name, val))
    return frozenset(keys)


def _collect_signatures(
    xml_root,
    zf: zipfile.ZipFile,
) -> set:
    """Collect all object signatures from an XML root (content.xml or styles.xml)."""
    sigs: set = set()
    xlink_href = f'{{{_NS["xlink"]}}}href'

    # Images inside draw:frame > draw:image
    for frame in xml_root.iter(f'{{{_NS["draw"]}}}frame'):
        for img in frame.iter(f'{{{_NS["draw"]}}}image'):
            href = img.get(xlink_href)
            if href and href in zf.namelist():
                h = hashlib.sha256(zf.read(href)).hexdigest()
                sigs.add(_ImageSig(h))
                logger.debug("Template image sig: %s -> %s", href, h[:12])

    # Shapes: draw:line, draw:rect, draw:custom-shape, draw:circle, etc.
    draw_ns = _NS["draw"]
    shape_tags = (
        f'{{{draw_ns}}}line',
        f'{{{draw_ns}}}rect',
        f'{{{draw_ns}}}circle',
        f'{{{draw_ns}}}ellipse',
        f'{{{draw_ns}}}polygon',
        f'{{{draw_ns}}}polyline',
        f'{{{draw_ns}}}path',
        f'{{{draw_ns}}}custom-shape',
        f'{{{draw_ns}}}connector',
    )
    for tag in shape_tags:
        for shape in xml_root.iter(tag):
            local_tag = tag.split("}")[-1]
            attrs = _geometry_attrs(shape)
            if attrs:
                sigs.add(_ShapeSig(local_tag, attrs))
                logger.debug("Template shape sig: %s %s", local_tag, dict(attrs))

    return sigs


def _collect_image_hashes(zf: zipfile.ZipFile, xml_root) -> dict[str, str]:
    """Map image href -> sha256 hash for all images referenced in the XML."""
    hashes: dict[str, str] = {}
    xlink_href = f'{{{_NS["xlink"]}}}href'
    for img in xml_root.iter(f'{{{_NS["draw"]}}}image'):
        href = img.get(xlink_href)
        if href and href in zf.namelist():
            hashes[href] = hashlib.sha256(zf.read(href)).hexdigest()
    return hashes


def _remove_matching_objects(
    xml_root,
    target_zf: zipfile.ZipFile,
    sigs: set,
    tmpl_image_hashes: set[str],
) -> int:
    """Remove elements from xml_root that match any signature. Returns count."""
    removed = 0
    xlink_href = f'{{{_NS["xlink"]}}}href'

    # Build parent map.
    parent_map = {c: p for p in xml_root.iter() for c in p}

    # Build image hash map for target.
    target_img_hashes: dict[str, str] = {}
    for img in xml_root.iter(f'{{{_NS["draw"]}}}image'):
        href = img.get(xlink_href)
        if href and href in target_zf.namelist():
            target_img_hashes[href] = hashlib.sha256(target_zf.read(href)).hexdigest()

    to_remove: list[tuple] = []

    # Match draw:frame containing matching images.
    for frame in xml_root.iter(f'{{{_NS["draw"]}}}frame'):
        for img in frame.iter(f'{{{_NS["draw"]}}}image'):
            href = img.get(xlink_href)
            if href and href in target_img_hashes:
                if target_img_hashes[href] in tmpl_image_hashes:
                    # Find removable ancestor.
                    ancestor = _find_removable(frame, parent_map)
                    if ancestor:
                        to_remove.append(ancestor)
                    break

    # Match shapes.
    draw_ns = _NS["draw"]
    shape_tags = (
        f'{{{draw_ns}}}line', f'{{{draw_ns}}}rect', f'{{{draw_ns}}}circle',
        f'{{{draw_ns}}}ellipse', f'{{{draw_ns}}}polygon',
        f'{{{draw_ns}}}polyline', f'{{{draw_ns}}}path',
        f'{{{draw_ns}}}custom-shape', f'{{{draw_ns}}}connector',
    )
    for tag in shape_tags:
        local_tag = tag.split("}")[-1]
        for shape in xml_root.iter(tag):
            attrs = _geometry_attrs(shape)
            if attrs and _ShapeSig(local_tag, attrs) in sigs:
                ancestor = _find_removable(shape, parent_map)
                if ancestor:
                    to_remove.append(ancestor)

    # Deduplicate and remove.
    seen: set[int] = set()
    for parent, child in to_remove:
        eid = id(child)
        if eid in seen:
            continue
        seen.add(eid)
        try:
            parent.remove(child)
            removed += 1
        except ValueError:
            pass

    return removed


def _find_removable(elem, parent_map) -> tuple | None:
    """Walk up to find the best ancestor to remove (draw:frame > w:p etc.)."""
    current = elem
    best = None
    while current is not None:
        parent = parent_map.get(current)
        if parent is None:
            break
        tag = current.tag.split("}")[-1] if "}" in current.tag else current.tag
        if tag in ("frame", "line", "rect", "circle", "ellipse", "polygon",
                    "polyline", "path", "custom-shape", "connector"):
            best = (parent, current)
        current = parent
    if best:
        return best
    parent = parent_map.get(elem)
    if parent is not None:
        return (parent, elem)
    return None


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
) -> int:
    """Remove template objects from target .odt, save to output_path.

    Returns total number of objects removed.
    """
    template_path = Path(template_path)
    target_path = Path(target_path)
    output_path = Path(output_path)

    # 1. Collect signatures from the template.
    with zipfile.ZipFile(template_path, "r") as zf_tmpl:
        tmpl_content = etree.fromstring(zf_tmpl.read("content.xml"))
        tmpl_styles = etree.fromstring(zf_tmpl.read("styles.xml"))
        sigs = _collect_signatures(tmpl_content, zf_tmpl)
        sigs |= _collect_signatures(tmpl_styles, zf_tmpl)

        # Also collect raw image hashes for content-based matching.
        tmpl_image_hashes: set[str] = set()
        img_map_c = _collect_image_hashes(zf_tmpl, tmpl_content)
        img_map_s = _collect_image_hashes(zf_tmpl, tmpl_styles)
        tmpl_image_hashes = set(img_map_c.values()) | set(img_map_s.values())

    logger.info(
        "Template signatures: %d sigs, %d unique images",
        len(sigs), len(tmpl_image_hashes),
    )

    if not sigs and not tmpl_image_hashes:
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
            all_names = zf_target.namelist()

            # Remove from content.xml (document body — all pages).
            r1 = _remove_matching_objects(
                target_content, zf_target, sigs, tmpl_image_hashes,
            )
            logger.debug("Removed %d objects from content.xml", r1)

            # Remove from styles.xml (headers/footers — all pages).
            r2 = _remove_matching_objects(
                target_styles, zf_target, sigs, tmpl_image_hashes,
            )
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
            removed = remove_objects_from_odt(template_path, odt_path, output_path)
            total_removed += removed
            logger.info("Wrote %s (removed %d object(s))", output_path.name, removed)
            succeeded += 1
        except Exception as exc:
            failed += 1
            logger.error("Failed on %s: %s", odt_path.name, exc)

    return (len(matches), succeeded, failed, total_removed)
