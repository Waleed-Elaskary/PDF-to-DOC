"""Convert .odt files to .pdf using LibreOffice in headless mode.

LibreOffice must be installed. The tool auto-detects common install paths or
accepts an explicit --soffice flag.

Uses an isolated user profile so conversions work even if LibreOffice GUI is
already open.
"""
from __future__ import annotations

import logging
import os
import re
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Iterable

logger = logging.getLogger(__name__)

# Reuse soffice finder from lo_converter
from .lo_converter import find_soffice, _get_lo_profile


def resolve_output_path(
    odt_path: Path,
    *,
    overwrite: bool,
    prefix: str = "",
) -> Path:
    """Return the .pdf output path for a given .odt.

    If prefix is set, prepend it to the stem.
    If the target exists and overwrite is False, append " (1)", " (2)", ...
    """
    new_stem = f"{prefix}{odt_path.stem}"
    base = odt_path.with_name(f"{new_stem}.pdf")
    if overwrite or not base.exists():
        return base

    counter = 1
    while True:
        candidate = base.with_name(f"{new_stem} ({counter}).pdf")
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
    """Run LibreOffice headless to export an .odt to .pdf.

    Returns the path of the produced .pdf file.
    """
    profile = _get_lo_profile()
    profile_uri = profile.resolve().as_uri()

    cmd = [
        soffice,
        "--headless",
        "--norestore",
        "--nolockcheck",
        f"-env:UserInstallation={profile_uri}",
        "--convert-to", "pdf:writer_pdf_Export",
        "--outdir", str(outdir),
        str(odt_path),
    ]

    logger.debug("Running: %s", " ".join(cmd))
    result = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        timeout=timeout,
    )
    logger.debug("stdout: %s", result.stdout.strip())
    logger.debug("stderr: %s", result.stderr.strip())

    # Check if file was produced.
    expected = outdir / odt_path.with_suffix(".pdf").name
    if expected.exists():
        return expected

    # Scan for any .pdf in case of name mangling.
    candidates = list(outdir.glob("*.pdf"))
    if candidates:
        logger.debug("Found output via scan: %s", candidates[0].name)
        return candidates[0]

    raise FileNotFoundError(
        f"LibreOffice did not produce a PDF.\n"
        f"Input: {odt_path}\n"
        f"stdout: {result.stdout.strip()}\n"
        f"stderr: {result.stderr.strip()}\n"
        f"Tip: close any running LibreOffice windows and retry."
    )


def convert_odt_to_pdf(
    odt_path: Path,
    *,
    soffice: str,
    overwrite: bool = False,
    prefix: str = "",
    timeout: int = 300,
) -> Path:
    """Convert a single .odt to .pdf via LibreOffice.

    The output is saved next to the source .odt with collision handling.
    Returns the path of the created file.
    """
    odt_path = Path(odt_path)
    if not odt_path.is_file():
        raise FileNotFoundError(f"ODT not found: {odt_path}")

    output_path = resolve_output_path(
        odt_path, overwrite=overwrite, prefix=prefix,
    )
    logger.info("Converting %s -> %s", odt_path.name, output_path.name)

    # Use a temp dir next to the .odt (avoids cross-drive issues).
    parent_dir = odt_path.parent
    try:
        tmp_ctx = tempfile.TemporaryDirectory(dir=parent_dir, prefix=".lo_tmp_")
        tmp_path = Path(tmp_ctx.name)
    except OSError:
        logger.debug("Could not create temp dir next to ODT, using system temp.")
        tmp_ctx = tempfile.TemporaryDirectory(prefix="lo_tmp_")
        tmp_path = Path(tmp_ctx.name)

    try:
        produced = _run_libreoffice(
            soffice, odt_path, tmp_path, timeout=timeout,
        )
        output_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(produced), str(output_path))
    finally:
        tmp_ctx.cleanup()

    logger.info("Wrote %s", output_path.name)
    return output_path


def collect_odts(
    inputs: Iterable[Path],
    *,
    recursive: bool = False,
    pattern: str | None = None,
) -> list[Path]:
    """Resolve CLI inputs (files and/or folders) into a list of .odt files.

    If pattern is provided, only filenames matching the regex are included.
    """
    resolved: list[Path] = []
    seen: set[Path] = set()
    for item in inputs:
        p = Path(item)
        if p.is_dir():
            iterator = p.rglob("*") if recursive else p.iterdir()
            for child in sorted(iterator):
                if not (child.is_file() and child.suffix.lower() == ".odt"):
                    continue
                if child.name.startswith("~"):
                    continue
                if pattern and not re.match(pattern, child.name, re.IGNORECASE):
                    continue
                rp = child.resolve()
                if rp not in seen:
                    seen.add(rp)
                    resolved.append(child)
        elif p.is_file():
            if p.suffix.lower() != ".odt":
                logger.warning("Skipping non-ODT file: %s", p)
                continue
            if pattern and not re.match(pattern, p.name, re.IGNORECASE):
                logger.warning("Skipping (pattern mismatch): %s", p)
                continue
            rp = p.resolve()
            if rp not in seen:
                seen.add(rp)
                resolved.append(p)
        else:
            logger.warning("Input does not exist: %s", p)
    return resolved
