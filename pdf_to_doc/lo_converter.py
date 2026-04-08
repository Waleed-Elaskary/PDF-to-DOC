"""Convert PDFs to .odg (LibreOffice Draw) and .odt (LibreOffice Writer) using
LibreOffice in headless mode.

LibreOffice must be installed. The tool auto-detects common install paths or
accepts an explicit --soffice flag.

IMPORTANT: If another LibreOffice window (or background process) is running,
headless conversion silently fails and produces nothing. This module works
around that by giving the headless instance its own user profile directory.
"""
from __future__ import annotations

import logging
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Iterable

logger = logging.getLogger(__name__)

FORMAT_ODG = "odg"
FORMAT_ODT = "odt"
VALID_FORMATS = (FORMAT_ODG, FORMAT_ODT)

_SOFFICE_CANDIDATES = [
    # Windows
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    # macOS
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    # Linux
    "/usr/bin/soffice",
    "/usr/bin/libreoffice",
    "/usr/local/bin/soffice",
    "/snap/bin/libreoffice",
]


def find_soffice(explicit: str | None = None) -> str:
    """Return path to the soffice binary.

    Priority: explicit arg > SOFFICE_PATH env > PATH lookup > known locations.
    """
    if explicit:
        p = Path(explicit)
        if p.is_file():
            return str(p)
        raise FileNotFoundError(f"soffice not found at: {explicit}")

    env = os.environ.get("SOFFICE_PATH")
    if env and Path(env).is_file():
        return env

    on_path = shutil.which("soffice") or shutil.which("libreoffice")
    if on_path:
        return on_path

    for candidate in _SOFFICE_CANDIDATES:
        if Path(candidate).is_file():
            return candidate

    raise FileNotFoundError(
        "LibreOffice (soffice) not found. Install it or pass --soffice <path>."
    )


def resolve_output_path(
    pdf_path: Path, fmt: str, *, overwrite: bool, prefix: str = "",
) -> Path:
    """Return the output path for a given PDF and format.

    If prefix is set, prepend it to the stem (e.g. prefix="EBB-" turns
    "report.pdf" into "EBB-report.odg").

    If the target exists and overwrite is False, append " (1)", " (2)", ... to
    the stem until a free name is found.
    """
    new_stem = f"{prefix}{pdf_path.stem}"
    base = pdf_path.with_name(f"{new_stem}.{fmt}")
    if overwrite or not base.exists():
        return base

    counter = 1
    while True:
        candidate = base.with_name(f"{new_stem} ({counter}).{fmt}")
        if not candidate.exists():
            return candidate
        counter += 1


# Persistent profile dir for headless LibreOffice (avoids conflict with any
# running GUI instance).
_LO_PROFILE_DIR: Path | None = None


def _get_lo_profile() -> Path:
    """Return a persistent temp directory for the headless LO user profile."""
    global _LO_PROFILE_DIR
    if _LO_PROFILE_DIR is None or not _LO_PROFILE_DIR.exists():
        _LO_PROFILE_DIR = Path(tempfile.mkdtemp(prefix="lo_profile_"))
    return _LO_PROFILE_DIR


def _run_libreoffice(
    soffice: str,
    pdf_path: Path,
    fmt: str,
    outdir: Path,
    *,
    timeout: int = 300,
) -> Path:
    """Run LibreOffice headless to convert a PDF to the given format.

    Uses an isolated user profile so conversions work even if LibreOffice
    is already open.

    Returns the path of the file LibreOffice produced.
    """
    profile = _get_lo_profile()
    # file:/// URI with forward slashes (required by LibreOffice on all platforms).
    profile_uri = profile.resolve().as_uri()

    # Build the command depending on target format.
    #
    #   ODG: LibreOffice Draw opens PDFs natively.
    #        Export filter: "draw8" (or just "odg").
    #
    #   ODT: We must force Writer to import the PDF (not Draw) using
    #        --infilter="writer_pdf_import", then export with "writer8".
    #
    base_cmd = [
        soffice,
        "--headless",
        "--norestore",
        "--nolockcheck",
        f"-env:UserInstallation={profile_uri}",
    ]

    if fmt == FORMAT_ODT:
        attempts = [
            # Attempt 1: force Writer import + writer8 export
            [*base_cmd,
             "--infilter=writer_pdf_import",
             "--convert-to", "odt:writer8",
             "--outdir", str(outdir),
             str(pdf_path)],
            # Attempt 2: just odt (some LO versions pick Writer automatically)
            [*base_cmd,
             "--infilter=writer_pdf_import",
             "--convert-to", "odt",
             "--outdir", str(outdir),
             str(pdf_path)],
        ]
    else:  # ODG
        attempts = [
            [*base_cmd,
             "--convert-to", "odg",
             "--outdir", str(outdir),
             str(pdf_path)],
            [*base_cmd,
             "--convert-to", "odg:draw8",
             "--outdir", str(outdir),
             str(pdf_path)],
        ]

    result = None
    for cmd in attempts:
        logger.debug("Running: %s", " ".join(cmd))
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=timeout,
        )

        logger.debug("stdout: %s", result.stdout.strip())
        logger.debug("stderr: %s", result.stderr.strip())

        # Check if the file was produced.
        expected = outdir / pdf_path.with_suffix(f".{fmt}").name
        if expected.exists():
            return expected

        candidates = list(outdir.glob(f"*.{fmt}"))
        if candidates:
            logger.debug("Found output via scan: %s", candidates[0].name)
            return candidates[0]

        logger.debug("Attempt did not produce output, trying next...")

    # Nothing worked.
    stdout = result.stdout.strip() if result else ""
    stderr = result.stderr.strip() if result else ""
    raise FileNotFoundError(
        f"LibreOffice did not produce an output file.\n"
        f"Input: {pdf_path}\n"
        f"stdout: {stdout}\n"
        f"stderr: {stderr}\n"
        f"Tip: close any running LibreOffice windows and retry."
    )


def convert_pdf(
    pdf_path: Path,
    fmt: str,
    *,
    soffice: str,
    overwrite: bool = False,
    prefix: str = "",
    timeout: int = 300,
) -> Path:
    """Convert a single PDF to .odg or .odt via LibreOffice.

    The output is saved next to the source PDF with collision handling.
    Returns the path of the created file.
    """
    pdf_path = Path(pdf_path)
    if not pdf_path.is_file():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")
    if fmt not in VALID_FORMATS:
        raise ValueError(f"format must be one of {VALID_FORMATS}, got {fmt!r}")

    output_path = resolve_output_path(
        pdf_path, fmt, overwrite=overwrite, prefix=prefix,
    )
    logger.info(
        "Converting %s -> %s (format=%s)",
        pdf_path.name, output_path.name, fmt,
    )

    # Use a temp dir next to the PDF (avoids cross-drive / permission issues).
    parent_dir = pdf_path.parent
    try:
        tmp_ctx = tempfile.TemporaryDirectory(dir=parent_dir, prefix=".lo_tmp_")
        tmp_path = Path(tmp_ctx.name)
    except OSError:
        logger.debug("Could not create temp dir next to PDF, using system temp.")
        tmp_ctx = tempfile.TemporaryDirectory(prefix="lo_tmp_")
        tmp_path = Path(tmp_ctx.name)

    try:
        produced = _run_libreoffice(
            soffice, pdf_path, fmt, tmp_path, timeout=timeout,
        )
        output_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(produced), str(output_path))
    finally:
        tmp_ctx.cleanup()

    logger.info("Wrote %s", output_path.name)
    return output_path


def collect_pdfs(
    inputs: Iterable[Path],
    *,
    recursive: bool = False,
) -> list[Path]:
    """Resolve CLI inputs (files and/or folders) into a list of PDFs.

    Mirrors the same logic as the PDF-to-DOCX tool.
    """
    resolved: list[Path] = []
    seen: set[Path] = set()
    for item in inputs:
        p = Path(item)
        if p.is_dir():
            iterator = p.rglob("*") if recursive else p.iterdir()
            for child in sorted(iterator):
                if child.is_file() and child.suffix.lower() == ".pdf":
                    rp = child.resolve()
                    if rp not in seen:
                        seen.add(rp)
                        resolved.append(child)
        elif p.is_file():
            if p.suffix.lower() != ".pdf":
                logger.warning("Skipping non-PDF file: %s", p)
                continue
            rp = p.resolve()
            if rp not in seen:
                seen.add(rp)
                resolved.append(p)
        else:
            logger.warning("Input does not exist: %s", p)
    return resolved
