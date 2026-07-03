"""PDF merger.

Merges all PDF files found in a source directory into a single output PDF,
using PyMuPDF (already a dependency via pymupdf4llm/pymupdf-layout).

Usage
-----
    python merge_pdfs.py [source_dir] [-o output.pdf]

If *source_dir* is omitted, ``pdf_to_merge`` (relative to this script) is
used. If *-o/--output* is omitted, ``merged.pdf`` is written to the current
working directory.

Files are merged in alphabetical filename order.
"""

from __future__ import annotations

import argparse
import logging
from pathlib import Path

LOGGER = logging.getLogger(__name__)

DEFAULT_SOURCE_DIR_NAME = "pdf_to_merge"
DEFAULT_OUTPUT_NAME = "merged.pdf"


def merge_pdfs(source_dir: Path, output_path: Path) -> int:
    """Merge every ``.pdf`` file in *source_dir* into *output_path*.

    Files are merged in alphabetical filename order. Returns the number of
    files merged.
    """
    try:
        import fitz
    except ImportError as exc:
        raise RuntimeError(
            "Missing dependency. Install pymupdf to enable PDF merging."
        ) from exc

    pdf_paths = sorted(source_dir.glob("*.pdf"), key=lambda p: p.name.lower())
    if not pdf_paths:
        raise SystemExit(f"No PDF files found in {source_dir}")

    merged = fitz.open()
    try:
        for pdf_path in pdf_paths:
            with fitz.open(pdf_path) as doc:
                merged.insert_pdf(doc)
            LOGGER.info("Added %s", pdf_path.name)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        merged.save(output_path)
    finally:
        merged.close()

    return len(pdf_paths)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Merge all PDFs in a directory into a single PDF.",
    )
    parser.add_argument(
        "source_dir",
        nargs="?",
        default=None,
        help=f"Directory containing PDFs to merge (default: {DEFAULT_SOURCE_DIR_NAME}).",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help=f"Output PDF path (default: {DEFAULT_OUTPUT_NAME} in the current directory).",
    )
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

    source_dir = (
        Path(args.source_dir).resolve()
        if args.source_dir
        else (Path(__file__).parent / DEFAULT_SOURCE_DIR_NAME).resolve()
    )
    if not source_dir.is_dir():
        raise SystemExit(f"Source directory not found: {source_dir}")

    output_path = (
        Path(args.output).resolve() if args.output else (Path.cwd() / DEFAULT_OUTPUT_NAME)
    )

    count = merge_pdfs(source_dir, output_path)
    LOGGER.info("Merged %d PDF(s) into %s", count, output_path)


if __name__ == "__main__":
    main()
