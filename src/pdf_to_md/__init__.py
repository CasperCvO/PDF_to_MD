"""PDF to Markdown converter package.

Provides high-fidelity conversion of PDFs and Office documents to Markdown,
with support for directory trees and batch PDF merging.  Also ships an
add-on vision-based slide extractor (``convert_presentation_vision``) that
uses the Gemini API for presentations with embedded UI screenshots, terminal
dumps, and metric dashboards.
"""

from .main import convert_tree
from .merge_pdfs import merge_pdfs
from .office_converter import SUPPORTED_EXTENSIONS as OFFICE_EXTENSIONS
from .office_converter import convert_office
from .pdf_converter import SUPPORTED_EXTENSIONS as PDF_EXTENSIONS
from .pdf_converter import convert_pdf
from .vision_converter import convert_presentation_vision

__all__ = [
    "convert_tree",
    "convert_pdf",
    "convert_office",
    "convert_presentation_vision",
    "merge_pdfs",
    "PDF_EXTENSIONS",
    "OFFICE_EXTENSIONS",
]

__version__ = "0.3.0"
