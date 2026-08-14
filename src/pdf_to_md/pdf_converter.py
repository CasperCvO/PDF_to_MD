"""PDF-to-Markdown converter module.

Converts a single PDF file to Markdown using pymupdf4llm (with a plain
PyMuPDF fallback).  Extracted images are saved into a per-document folder.
Each output file is prepended with a Wikilink to the original PDF.
"""

from __future__ import annotations

import logging
from pathlib import Path

LOGGER = logging.getLogger(__name__)

SUPPORTED_EXTENSIONS: set[str] = {".pdf"}


def _extract_markdown(pdf_path: Path, image_dir: Path) -> str:
    """Return Markdown text for *pdf_path*, saving images to *image_dir*."""
    try:
        from pymupdf4llm import to_markdown
    except ImportError:
        to_markdown = None

    try:
        import fitz
    except ImportError as exc:
        raise RuntimeError(
            "Missing dependencies. Install pymupdf4llm and pymupdf to enable extraction."
        ) from exc

    with fitz.open(pdf_path) as doc:
        if to_markdown is not None:
            try:
                image_dir.mkdir(parents=True, exist_ok=True)
                return to_markdown(
                    doc,
                    write_images=True,
                    image_path=str(image_dir),
                ).strip()
            except Exception as exc:
                LOGGER.warning(
                    "pymupdf4llm failed for %s (%s). Falling back to PyMuPDF text.",
                    pdf_path,
                    exc,
                )

        pages = [page.get_text("text").rstrip() for page in doc]
        return "\n\n".join(page for page in pages if page).strip()


def convert_pdf(pdf_path: Path, output_path: Path) -> None:
    """Convert a single PDF to Markdown and write the result to *output_path*.

    Images are extracted into a sibling ``<stem>_images/`` folder next to
    *output_path*.
    """
    image_dir = output_path.parent / f"{pdf_path.stem}_images"

    markdown_body = _extract_markdown(pdf_path, image_dir)
    wikilink = f"[[{pdf_path.name}]]"
    content = f"{wikilink}\n\n{markdown_body}\n" if markdown_body else f"{wikilink}\n"

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(content, encoding="utf-8")
    LOGGER.info("Wrote %s", output_path)
