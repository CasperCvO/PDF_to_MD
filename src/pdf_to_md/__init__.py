"""PDF to Markdown converter package.

Provides high-fidelity conversion of PDFs and Office documents to Markdown,
with support for directory trees and batch PDF merging.
"""

from .main import convert_tree
from .merge_pdfs import merge_pdfs
from .office_converter import SUPPORTED_EXTENSIONS as OFFICE_EXTENSIONS
from .office_converter import convert_office
from .pdf_converter import SUPPORTED_EXTENSIONS as PDF_EXTENSIONS
from .pdf_converter import convert_pdf

__all__ = [
    "convert_tree",
    "convert_pdf",
    "convert_office",
    "merge_pdfs",
    "PDF_EXTENSIONS",
    "OFFICE_EXTENSIONS",
]

__version__ = "0.1.0"
