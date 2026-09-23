"""PDF-to-Markdown converter module.

Converts a single PDF file to Markdown using pymupdf4llm (with a plain
PyMuPDF fallback).  Extracted images are saved into a per-document folder.
Each output file is prepended with a Wikilink to the original PDF, and every
page starts with a ``<!-- page: N -->`` marker (including empty pages) so
downstream consumers can locate the source page of any content.
"""

from __future__ import annotations

import logging
import re
from collections.abc import Iterable
from pathlib import Path

LOGGER = logging.getLogger(__name__)

SUPPORTED_EXTENSIONS: set[str] = {".pdf"}


def _patch_pymupdf4llm_image_saver() -> None:
    """Ensure pymupdf4llm creates directories before saving images even if paths are transformed."""
    try:
        from pymupdf4llm.helpers import utils as llm_utils

        if getattr(llm_utils, "_pdf_to_md_patched", False):
            return

        orig_md_path = llm_utils.md_path

        def safe_md_path(folder: str, filename: str):
            md_ref, save_ref = orig_md_path(folder, filename)
            Path(save_ref).parent.mkdir(parents=True, exist_ok=True)
            return md_ref, save_ref

        llm_utils.md_path = safe_md_path
        llm_utils._pdf_to_md_patched = True

        try:
            from pymupdf4llm.helpers import document_layout as doc_layout

            doc_layout.utils.md_path = safe_md_path
        except (ImportError, AttributeError):
            pass
    except (ImportError, AttributeError):
        pass


def _extract_markdown(pdf_path: Path, image_dir: Path, clean_stem: str) -> str:
    """Return Markdown text for *pdf_path*, saving images to *image_dir*."""
    try:
        from pymupdf4llm import to_markdown
    except ImportError:
        to_markdown = None

    try:
        import pymupdf
    except ImportError as exc:
        raise RuntimeError(
            "Missing dependencies. Install pymupdf4llm and pymupdf to enable extraction."
        ) from exc

    pdf_bytes = pdf_path.read_bytes()

    if to_markdown is not None:
        try:
            _patch_pymupdf4llm_image_saver()
            image_dir.mkdir(parents=True, exist_ok=True)
            with pymupdf.open(stream=pdf_bytes, filetype="pdf") as doc:
                chunks = to_markdown(
                    doc,
                    write_images=True,
                    image_path=str(image_dir),
                    filename=clean_stem,
                    page_chunks=True,
                )
            pages = []
            for index, chunk in enumerate(chunks):
                meta = chunk.get("metadata") or {}
                number = meta.get("page_number") or meta.get("page") or index + 1
                pages.append((number, (chunk.get("text") or "").strip()))
            return _join_pages(pages)
        except Exception as exc:
            LOGGER.warning(
                "pymupdf4llm failed for %s (%s). Falling back to PyMuPDF text.",
                pdf_path,
                exc,
            )

    with pymupdf.open(stream=pdf_bytes, filetype="pdf") as doc:
        pages = [
            (index, page.get_text("text").strip())
            for index, page in enumerate(doc, start=1)
        ]

    return _join_pages(pages)


def _join_pages(pages: Iterable[tuple[int, str]]) -> str:
    """Join ``(page_number, text)`` pairs into page-marker-delimited Markdown."""
    blocks = []
    for number, text in pages:
        marker = f"<!-- page: {number} -->"
        blocks.append(f"{marker}\n\n{text}" if text else marker)
    return "\n\n".join(blocks)


def convert_pdf(pdf_path: Path, output_path: Path) -> None:
    """Convert a single PDF to Markdown and write the result to *output_path*.

    Images are extracted into a sibling ``<stem>_images/`` folder next to
    *output_path*.
    """
    clean_stem = re.sub(r"[\s\(\)\[\]]+", "_", pdf_path.stem)
    image_dir = output_path.parent / f"{clean_stem}_images"

    markdown_body = _extract_markdown(pdf_path, image_dir, clean_stem)
    wikilink = f"[[{pdf_path.name}]]"
    content = f"{wikilink}\n\n{markdown_body}\n" if markdown_body else f"{wikilink}\n"

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(content, encoding="utf-8")
    LOGGER.info("Wrote %s", output_path)
