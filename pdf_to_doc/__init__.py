"""Batch PDF to DOCX conversion tool."""
from .converter import convert_pdf_to_docx, resolve_output_path

__version__ = "1.0.0"
__all__ = ["convert_pdf_to_docx", "resolve_output_path"]
