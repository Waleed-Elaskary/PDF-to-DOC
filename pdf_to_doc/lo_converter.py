"""Convert PDFs to .odg (LibreOffice Draw) and .odt (LibreOffice Writer) using
LibreOffice in headless mode.

LibreOffice must be installed. The tool auto-detects common install paths or
accepts an explicit --soffice flag.
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

    # Check PATH
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
    pdf_path: Path, fmt: str, *, overwrite: bool
) -> Path:
    """Return the output path for a given PDF and format.

    If the target exists and overwrite is False, append " (1)", " (2)", ... to
    the stem until a free name is found.
    """
    base = pdf_path.with_suffix(f".{fmt}")
    if overwrite or not base.exists():
        return base

    counter = 1
    while True:
        candidate = base.with_name(f"{base.stem} ({counter}).{fmt}")
        if not candidate.exists():
            return candidate
        counter += 1


def _run_libreoffice(
    soffice: str,
    pdf_path: Path,
    fmt: str,
    outdir: Path,
    *,
    timeout: int = 300,
) -> Path:
    """Run LibreOffice headless to convert a PDF to the given format.

    Returns the path of the file LibreOffice produced.
    """
    # LibreOffice filter hints for PDF import
    if fmt == FORMAT_ODG:
        convert_arg = "odg:draw_pdf_import"
    elif fmt == FORMAT_ODT:
        convert_arg = "odt:writer_pdf_import"
    else:
        convert_arg = fmt

    cmd = [
        soffice,
        "--headless",
        "--norestore",
        "--convert-to", convert_arg,
        "--outdir", str(outdir),
        str(pdf_path),
    ]

    logger.debug("Running: %s", " ".join(cmd))
    result = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        timeout=timeout,
    )

    if result.returncode != 0:
        logger.debug("stdout: %s", result.stdout)
        logger.debug("stderr: %s", result.stderr)
        raise RuntimeError(
            f"LibreOffice exited with code {result.returncode}: "
            f"{result.stderr.strip() or result.stdout.strip()}"
        )

    # LibreOffice writes <stem>.<fmt> into outdir.
    expected = outdir / pdf_path.with_suffix(f".{fmt}").name
    if not expected.exists():
        raise FileNotFoundError(
            f"LibreOffice did not produce expected file: {expected}"
        )
    return expected


def convert_pdf(
    pdf_path: Path,
    fmt: str,
    *,
    soffice: str,
    overwrite: bool = False,
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

    output_path = resolve_output_path(pdf_path, fmt, overwrite=overwrite)
    logger.info(
        "Converting %s -> %s (format=%s)",
        pdf_path.name, output_path.name, fmt,
    )

    # LibreOffice always writes to outdir with the original stem, so use a temp
    # dir then move to the final name (handles collision renames).
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        produced = _run_libreoffice(
            soffice, pdf_path, fmt, tmp_path, timeout=timeout,
        )
        # Move to final destination.
        output_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(produced), str(output_path))

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
