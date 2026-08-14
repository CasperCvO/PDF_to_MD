"""Office-document-to-Markdown converter module.

Converts office files (docx, xlsx, pptx, xls, doc, …) to Markdown using
Microsoft's *markitdown* library.  Each output file is prepended with a
Wikilink to the original document for easy back-referencing in Obsidian.
"""

from __future__ import annotations

import logging
from pathlib import Path

LOGGER = logging.getLogger(__name__)

SUPPORTED_EXTENSIONS: set[str] = {
    ".docx",
    ".doc",
    ".xlsx",
    ".xls",
    ".pptx",
    ".ppt",
}


def _extract_markdown(file_path: Path) -> str:
    """Return Markdown text for *file_path* using markitdown."""
    try:
        from markitdown import MarkItDown
    except ImportError as exc:
        raise RuntimeError(
            "Missing dependency. Install markitdown to enable office-document conversion: "
            "pip install 'markitdown[docx,xlsx,pptx]'"
        ) from exc

    converter = MarkItDown()
    result = converter.convert(str(file_path))
    return (result.text_content or "").strip()


def convert_office(file_path: Path, output_path: Path) -> None:
    """Convert a single office document to Markdown and write to *output_path*."""
    markdown_body = _extract_markdown(file_path)
    wikilink = f"[[{file_path.name}]]"
    content = f"{wikilink}\n\n{markdown_body}\n" if markdown_body else f"{wikilink}\n"

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(content, encoding="utf-8")
    LOGGER.info("Wrote %s", output_path)
