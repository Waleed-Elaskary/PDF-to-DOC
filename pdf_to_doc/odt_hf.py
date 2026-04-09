"""Apply header/footer from a template .odt to matching .odt files in a folder.

Finds .odt files matching a filename pattern (e.g. ending in -001.odt), copies
header and footer content from a template .odt into each match, and saves the
result with the trailing number incremented (001 -> 002).

ODT files are ZIP archives.  Header/footer content lives in ``styles.xml``
inside ``<style:master-page>`` elements.  This module copies those elements
(and any referenced images) from the template into the target.
"""
from __future__ import annotations

import logging
import os
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Iterable
from copy import deepcopy

from lxml import etree

logger = logging.getLogger(__name__)

# ODT XML namespaces
_NS = {
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "style":  "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    "text":   "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "table":  "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "draw":   "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
    "fo":     "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
    "xlink":  "http://www.w3.org/1999/xlink",
    "svg":    "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
    "manifest": "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0",
}

# Default pattern: filename ends with a dash and digits before .odt
# e.g. EBB-SomeName-001.odt, Report-42.odt
_DEFAULT_PATTERN = r"^.+-(\d+)\.odt$"


def _parse_trailing_number(filename: str, pattern: str) -> tuple[str, int, int] | None:
    """Extract the trailing number from a filename using the given regex.

    Returns (prefix, number, num_digits) or None if no match.
    The regex must have exactly one capture group for the number portion.
    """
    m = re.match(pattern, filename, re.IGNORECASE)
    if not m:
        return None
    num_str = m.group(1)
    # Derive the prefix = everything before the captured number
    start, end = m.span(1)
    prefix = filename[:start]
    return (prefix, int(num_str), len(num_str))


def _increment_filename(filename: str, pattern: str, increment: int = 1) -> str:
    """Return a new filename with the trailing number incremented.

    E.g. "EBB-Report-001.odt" -> "EBB-Report-002.odt"
    """
    parsed = _parse_trailing_number(filename, pattern)
    if parsed is None:
        raise ValueError(f"Filename does not match pattern: {filename}")
    prefix, number, num_digits = parsed
    new_number = number + increment
    return f"{prefix}{str(new_number).zfill(num_digits)}.odt"


def collect_matching_odt(
    folder: Path,
    pattern: str,
    *,
    recursive: bool = False,
) -> list[Path]:
    """Find .odt files in folder whose filename matches the regex pattern."""
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


def _read_xml_from_odt(odt_path: Path, xml_name: str) -> etree._Element:
    """Read and parse an XML file from inside an ODT archive."""
    with zipfile.ZipFile(odt_path, "r") as zf:
        data = zf.read(xml_name)
    return etree.fromstring(data)


def _get_master_pages(styles_root: etree._Element) -> list[etree._Element]:
    """Return all <style:master-page> elements from a styles.xml root."""
    return styles_root.findall(
        ".//office:master-styles/style:master-page", _NS
    )


def _get_hf_elements(master_page: etree._Element) -> dict[str, etree._Element | None]:
    """Extract header and footer elements from a master-page."""
    return {
        "header": master_page.find("style:header", _NS),
        "footer": master_page.find("style:footer", _NS),
        "header-left": master_page.find("style:header-left", _NS),
        "footer-left": master_page.find("style:footer-left", _NS),
        "header-first": master_page.find("style:header-first", _NS),
        "footer-first": master_page.find("style:footer-first", _NS),
    }


def _collect_image_refs(elem: etree._Element) -> set[str]:
    """Find all xlink:href image references in an element tree."""
    refs: set[str] = set()
    xlink_href = f'{{{_NS["xlink"]}}}href'
    for node in elem.iter():
        href = node.get(xlink_href)
        if href and (href.startswith("Pictures/") or href.startswith("media/")):
            refs.add(href)
    return refs


def _collect_referenced_styles(elem: etree._Element) -> set[str]:
    """Find style names referenced by elements in the header/footer."""
    style_names: set[str] = set()
    text_style = f'{{{_NS["text"]}}}style-name'
    style_name_attr = f'{{{_NS["style"]}}}name'  # not used for lookup
    # Common attributes that reference style names
    for node in elem.iter():
        for attr in ("text:style-name", "table:style-name",
                      "draw:style-name", "draw:text-style-name"):
            ns, local = attr.split(":")
            full = f'{{{_NS.get(ns, "")}}}{local}'
            val = node.get(full)
            if val:
                style_names.add(val)
    return style_names


def apply_hf_to_odt(
    template_path: Path,
    target_path: Path,
    output_path: Path,
) -> None:
    """Copy header/footer from template .odt into target .odt, save to output.

    Copies:
      - All header/footer variants from the template's first master-page
        into every master-page in the target.
      - Any images referenced by the header/footer.
      - Style definitions used by the header/footer.
    """
    template_path = Path(template_path)
    target_path = Path(target_path)
    output_path = Path(output_path)

    # Read template's styles.xml and extract header/footer
    tmpl_styles = _read_xml_from_odt(template_path, "styles.xml")
    tmpl_master_pages = _get_master_pages(tmpl_styles)
    if not tmpl_master_pages:
        raise ValueError(f"No master pages found in template: {template_path}")

    tmpl_mp = tmpl_master_pages[0]
    tmpl_hf = _get_hf_elements(tmpl_mp)

    # Collect image references from header/footer
    image_refs: set[str] = set()
    for elem in tmpl_hf.values():
        if elem is not None:
            image_refs |= _collect_image_refs(elem)

    # Copy the target ODT to a temp file, modify, then move to output.
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(
        suffix=".odt", delete=False, dir=output_path.parent,
    ) as tmp:
        tmp_name = tmp.name

    try:
        shutil.copy2(target_path, tmp_name)

        # Modify styles.xml inside the target ODT
        with zipfile.ZipFile(tmp_name, "r") as zf_in:
            target_styles_data = zf_in.read("styles.xml")
            all_names = zf_in.namelist()

        target_styles = etree.fromstring(target_styles_data)
        target_master_pages = _get_master_pages(target_styles)

        if not target_master_pages:
            # Create a master-styles section if missing
            auto_styles = target_styles.find(
                "office:automatic-styles", _NS
            )
            master_styles = etree.SubElement(
                target_styles,
                f'{{{_NS["office"]}}}master-styles',
            )
            mp = etree.SubElement(
                master_styles,
                f'{{{_NS["style"]}}}master-page',
            )
            mp.set(f'{{{_NS["style"]}}}name', "Default")
            mp.set(
                f'{{{_NS["style"]}}}page-layout-name', "pm1"
            )
            target_master_pages = [mp]

        # Apply template header/footer to every master-page in the target.
        hf_tags = {
            "header":       f'{{{_NS["style"]}}}header',
            "footer":       f'{{{_NS["style"]}}}footer',
            "header-left":  f'{{{_NS["style"]}}}header-left',
            "footer-left":  f'{{{_NS["style"]}}}footer-left',
            "header-first": f'{{{_NS["style"]}}}header-first',
            "footer-first": f'{{{_NS["style"]}}}footer-first',
        }

        for mp in target_master_pages:
            for key, tag in hf_tags.items():
                # Remove existing
                for old in mp.findall(tag.split("}")[-1], _NS):
                    mp.remove(old)
                existing = mp.findall(
                    f'{{{_NS["style"]}}}{key.replace("-", "-")}',
                )
                # Direct tag search
                for child in list(mp):
                    child_local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
                    if child_local == key:
                        mp.remove(child)

                # Insert from template
                tmpl_elem = tmpl_hf.get(key)
                if tmpl_elem is not None:
                    mp.append(deepcopy(tmpl_elem))

        # Also copy automatic styles used by header/footer from template
        tmpl_auto = tmpl_styles.find("office:automatic-styles", _NS)
        target_auto = target_styles.find("office:automatic-styles", _NS)
        if tmpl_auto is not None and target_auto is not None:
            # Collect style names referenced by template hf
            needed_styles: set[str] = set()
            for elem in tmpl_hf.values():
                if elem is not None:
                    needed_styles |= _collect_referenced_styles(elem)

            # Existing style names in target
            existing_names: set[str] = set()
            for s in target_auto:
                n = s.get(f'{{{_NS["style"]}}}name')
                if n:
                    existing_names.add(n)

            # Copy needed styles that don't exist in target
            for s in tmpl_auto:
                n = s.get(f'{{{_NS["style"]}}}name')
                if n and n in needed_styles and n not in existing_names:
                    target_auto.append(deepcopy(s))

        # Serialize modified styles.xml
        new_styles_data = etree.tostring(
            target_styles, xml_declaration=True, encoding="UTF-8",
        )

        # Rebuild the ODT with modified styles.xml and template images.
        final_tmp = tmp_name + ".new"
        with zipfile.ZipFile(tmp_name, "r") as zf_in, \
             zipfile.ZipFile(final_tmp, "w", zipfile.ZIP_DEFLATED) as zf_out:

            for item in zf_in.namelist():
                if item == "styles.xml":
                    zf_out.writestr("styles.xml", new_styles_data)
                else:
                    zf_out.writestr(item, zf_in.read(item))

            # Copy images from template that are referenced by header/footer.
            if image_refs:
                with zipfile.ZipFile(template_path, "r") as zf_tmpl:
                    for ref in image_refs:
                        if ref in zf_tmpl.namelist() and ref not in all_names:
                            zf_out.writestr(ref, zf_tmpl.read(ref))

        os.replace(final_tmp, tmp_name)
        shutil.move(tmp_name, str(output_path))

    except Exception:
        # Clean up temp files on failure.
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
    increment: int = 1,
) -> tuple[int, int, int]:
    """Apply header/footer to all matching .odt files in a folder.

    For each matching file (e.g. ``EBB-Report-001.odt``), creates a new file
    with the trailing number incremented (``EBB-Report-002.odt``).

    Returns (found, succeeded, failed).
    """
    template_path = Path(template_path)
    folder = Path(folder)

    if not template_path.is_file():
        raise FileNotFoundError(f"Template not found: {template_path}")
    if not folder.is_dir():
        raise FileNotFoundError(f"Folder not found: {folder}")

    # Don't process the template itself.
    template_resolved = template_path.resolve()

    matches = collect_matching_odt(folder, pattern, recursive=recursive)
    matches = [m for m in matches if m.resolve() != template_resolved]

    if not matches:
        logger.warning("No .odt files matching pattern '%s' in %s", pattern, folder)
        return (0, 0, 0)

    logger.info("Found %d matching .odt file(s).", len(matches))

    succeeded = 0
    failed = 0
    for odt_path in matches:
        try:
            new_name = _increment_filename(odt_path.name, pattern, increment)
            output_path = odt_path.parent / new_name

            if not overwrite and output_path.exists():
                logger.info("Skipping (exists): %s", output_path.name)
                continue

            logger.info("%s -> %s", odt_path.name, new_name)
            apply_hf_to_odt(template_path, odt_path, output_path)
            logger.info("Wrote %s", output_path.name)
            succeeded += 1
        except Exception as exc:
            failed += 1
            logger.error("Failed on %s: %s", odt_path.name, exc)

    return (len(matches), succeeded, failed)
