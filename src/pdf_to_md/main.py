"""Document-to-Markdown batch converter.

Recursively scans a source directory for supported files (PDFs and office
documents) and converts each one to Markdown.  The resulting Markdown files
are written into an output folder (default: ``Extracts``) that mirrors the
original directory structure.

Supported formats
-----------------
- **PDF** — via pymupdf4llm (with a plain PyMuPDF fallback).
- **Office** — .docx, .doc, .xlsx, .xls, .pptx, .ppt via markitdown.

Features
--------
- Preserves tables, headers, and reading order where possible.
- Extracts embedded images from PDFs into per-document ``<stem>_images/`` folders.
- Prepends a Wikilink ``[[OriginalFile.ext]]`` at the top of every Markdown
  file for easy back-referencing (e.g. in Obsidian).
- Adds approximate ``<!-- page: N -->`` markers to ``.docx`` output
  (disable with ``--no-docx-page-markers``).
- Automatically skips the output folder to avoid re-processing.

Usage
-----
    python -m pdf_to_md.main [source_dir] [-o output_dir]

If *source_dir* is omitted the current working directory is used.
If *output_dir* is omitted, an ``Extracts`` folder inside *source_dir* is used.
"""

from __future__ import annotations

import argparse
import logging
import os
from pathlib import Path

from .hybrid_vision import (
    DEFAULT_HIGH_THRESHOLD,
    DEFAULT_HYBRID_DPI,
    DEFAULT_LOW_THRESHOLD,
    VALID_VISION_MODES,
    VisionScreenReport,
    screen_document,
    validate_vision_options,
)
from .office_converter import SUPPORTED_EXTENSIONS as OFFICE_EXTENSIONS
from .office_converter import convert_office
from .pdf_converter import SUPPORTED_EXTENSIONS as PDF_EXTENSIONS
from .pdf_converter import convert_pdf
from .vision_converter import DEFAULT_CONCURRENCY, DEFAULT_MAX_RETRIES, DEFAULT_MODEL

LOGGER = logging.getLogger(__name__)

OUTPUT_DIR_NAME = "Extracts"


def build_output_path(source_root: Path, output_root: Path, file_path: Path) -> Path:
    relative_path = file_path.relative_to(source_root)
    output_dir = output_root / relative_path.parent
    output_dir.mkdir(parents=True, exist_ok=True)
    return output_dir / f"{file_path.stem}.md"


def screen_tree(
    source_root: Path | str,
    *,
    low_threshold: float = DEFAULT_LOW_THRESHOLD,
    high_threshold: float = DEFAULT_HIGH_THRESHOLD,
    exclude_dir: Path | None = None,
) -> list[VisionScreenReport]:
    """Screen all supported documents under *source_root* (or a single file)."""
    validate_vision_options("auto", low_threshold, high_threshold)
    resolved = Path(source_root).resolve()
    if resolved.is_file():
        return [
            screen_document(
                resolved, low_threshold=low_threshold, high_threshold=high_threshold
            )
        ]
    if not resolved.is_dir():
        raise FileNotFoundError(f"Source path not found: {resolved}")

    skip_dir = (exclude_dir or (resolved / OUTPUT_DIR_NAME)).resolve()
    supported = PDF_EXTENSIONS | OFFICE_EXTENSIONS
    reports: list[VisionScreenReport] = []

    for root, dirs, files in os.walk(resolved):
        root_path = Path(root).resolve()
        dirs[:] = sorted(
            d for d in dirs if (root_path / d).resolve() != skip_dir
        )
        for filename in sorted(files):
            file_path = root_path / filename
            if file_path.suffix.lower() in supported:
                reports.append(
                    screen_document(
                        file_path,
                        low_threshold=low_threshold,
                        high_threshold=high_threshold,
                    )
                )
    return reports


def convert_tree(
    source_root: Path | str,
    output_root: Path | str | None = None,
    *,
    docx_page_markers: bool = True,
    vision_mode: str = "off",
    api_key: str | None = None,
    vision_model: str | None = None,
    low_threshold: float = DEFAULT_LOW_THRESHOLD,
    high_threshold: float = DEFAULT_HIGH_THRESHOLD,
    dpi: int = DEFAULT_HYBRID_DPI,
    concurrency: int = DEFAULT_CONCURRENCY,
    max_retries: int = DEFAULT_MAX_RETRIES,
) -> dict[str, int]:
    """Recursively convert supported documents in *source_root* to Markdown.

    Parameters
    ----------
    source_root : Path | str
        Root directory to scan.
    output_root : Path | str | None, optional
        Directory where converted Markdown files will be written. If None,
        defaults to ``<source_root>/Extracts``.
    docx_page_markers : bool, optional
        Insert approximate ``<!-- page: N -->`` markers into ``.docx``
        output (default: True).
    vision_mode : str, optional
        One of ``"off"`` (default), ``"auto"`` (screen and route by
        *low_threshold* / *high_threshold*), ``"images"``, ``"pages"``, or
        ``"full"``.

    Returns
    -------
    dict[str, int]
        Conversion counts for "pdf", "office", and "skipped".
    """
    if vision_mode != "off":
        validate_vision_options(
            vision_mode, low_threshold, high_threshold, dpi, concurrency, max_retries
        )

    source_root = Path(source_root).resolve()
    if not source_root.is_dir():
        raise FileNotFoundError(f"Source directory not found: {source_root}")

    if output_root is None:
        resolved_output_root = (source_root / OUTPUT_DIR_NAME).resolve()
    else:
        resolved_output_root = Path(output_root).resolve()

    resolved_output_root.mkdir(parents=True, exist_ok=True)

    counts: dict[str, int] = {"pdf": 0, "office": 0, "skipped": 0}
    vision_kwargs = (
        {}
        if vision_mode == "off"
        else {
            "vision_mode": vision_mode,
            "api_key": api_key,
            "vision_model": vision_model,
            "low_threshold": low_threshold,
            "high_threshold": high_threshold,
            "dpi": dpi,
            "concurrency": concurrency,
            "max_retries": max_retries,
        }
    )

    for root, dirs, files in os.walk(source_root):
        root_path = Path(root).resolve()
        dirs[:] = [
            directory
            for directory in dirs
            if (root_path / directory).resolve() != resolved_output_root
        ]

        for filename in files:
            file_path = root_path / filename
            ext = file_path.suffix.lower()

            if ext in PDF_EXTENSIONS:
                output_path = build_output_path(source_root, resolved_output_root, file_path)
                try:
                    convert_pdf(file_path, output_path, **vision_kwargs)
                    counts["pdf"] += 1
                except Exception as exc:
                    LOGGER.error("Failed to convert PDF %s: %s", file_path, exc)

            elif ext in OFFICE_EXTENSIONS:
                output_path = build_output_path(source_root, resolved_output_root, file_path)
                try:
                    convert_office(
                        file_path,
                        output_path,
                        docx_page_markers=docx_page_markers,
                        **vision_kwargs,
                    )
                    counts["office"] += 1
                except Exception as exc:
                    LOGGER.error("Failed to convert office file %s: %s", file_path, exc)

            else:
                LOGGER.info("Skipped %s (unsupported extension: %s)", file_path, ext)
                counts["skipped"] += 1

    return counts


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Convert PDFs and office documents in a directory tree to Markdown.",
    )
    parser.add_argument(
        "source_dir",
        nargs="?",
        default=".",
        help="Root directory (or single file) to scan (default: current directory).",
    )
    parser.add_argument(
        "-o",
        "--output-dir",
        default=None,
        help=f"Output directory (default: <source_dir>/{OUTPUT_DIR_NAME}).",
    )
    parser.add_argument(
        "--docx-page-markers",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Insert <!-- page: N --> markers into .docx output (default: enabled).",
    )
    parser.add_argument(
        "--vision",
        action="store_true",
        default=False,
        help="Enable automatic hybrid vision screening and extraction (--vision-mode auto).",
    )
    parser.add_argument(
        "--vision-mode",
        choices=list(VALID_VISION_MODES),
        default=None,
        help="Vision strategy: off (default), auto, images, pages, or full.",
    )
    parser.add_argument(
        "--vision-low-threshold",
        type=float,
        default=DEFAULT_LOW_THRESHOLD,
        help=(
            "Max visual-page ratio for targeted image-in-context extraction "
            f"(default: {DEFAULT_LOW_THRESHOLD})."
        ),
    )
    parser.add_argument(
        "--vision-high-threshold",
        type=float,
        default=DEFAULT_HIGH_THRESHOLD,
        help=(
            "Visual-page ratio above which the entire document is vision-extracted "
            f"(default: {DEFAULT_HIGH_THRESHOLD})."
        ),
    )
    parser.add_argument(
        "--api-key",
        default=None,
        help="Gemini API key (default: GEMINI_API_KEY env var or .env file).",
    )
    parser.add_argument(
        "--vision-model",
        default=DEFAULT_MODEL,
        help=f"Gemini model id for vision extraction (default: {DEFAULT_MODEL}).",
    )
    parser.add_argument(
        "--vision-dpi",
        type=int,
        default=DEFAULT_HYBRID_DPI,
        help=f"Render DPI for page-level vision extraction (default: {DEFAULT_HYBRID_DPI}).",
    )
    parser.add_argument(
        "--vision-concurrency",
        type=int,
        default=DEFAULT_CONCURRENCY,
        help=f"Max concurrent Gemini API requests (default: {DEFAULT_CONCURRENCY}).",
    )
    parser.add_argument(
        "--vision-max-retries",
        type=int,
        default=DEFAULT_MAX_RETRIES,
        help=f"Max retries per vision request (default: {DEFAULT_MAX_RETRIES}).",
    )
    parser.add_argument(
        "--screen-only",
        action="store_true",
        default=False,
        help="Screen document(s) for vision candidacy and print strategy without converting.",
    )
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    for noisy_logger in ("httpx", "httpcore"):
        logging.getLogger(noisy_logger).setLevel(logging.WARNING)
    logging.getLogger("google_genai").setLevel(logging.ERROR)

    effective_vision_mode = args.vision_mode or ("auto" if args.vision else "off")
    source_path = Path(args.source_dir).resolve()
    output_root = Path(args.output_dir).resolve() if args.output_dir else None

    if args.screen_only:
        if not source_path.exists():
            raise SystemExit(f"Source path not found: {source_path}")
        reports = screen_tree(
            source_path,
            low_threshold=args.vision_low_threshold,
            high_threshold=args.vision_high_threshold,
            exclude_dir=output_root,
        )
        for rep in reports:
            pages_str = (
                f" (units: {rep.candidate_unit_numbers})" if rep.candidate_units else ""
            )
            LOGGER.info(
                "%s: %d/%d visual (%.1f%%) -> %s%s",
                rep.file_path.name,
                rep.candidate_units,
                rep.total_units,
                rep.vision_ratio * 100.0,
                rep.recommended_strategy,
                pages_str,
            )
        return

    if source_path.is_file():
        target_dir = output_root or (source_path.parent / OUTPUT_DIR_NAME)
        target_dir.mkdir(parents=True, exist_ok=True)
        out_md = target_dir / f"{source_path.stem}.md"
        ext = source_path.suffix.lower()
        vision_kwargs = (
            {}
            if effective_vision_mode == "off"
            else {
                "vision_mode": effective_vision_mode,
                "api_key": args.api_key,
                "vision_model": args.vision_model,
                "low_threshold": args.vision_low_threshold,
                "high_threshold": args.vision_high_threshold,
                "dpi": args.vision_dpi,
                "concurrency": args.vision_concurrency,
                "max_retries": args.vision_max_retries,
            }
        )
        if ext in PDF_EXTENSIONS:
            convert_pdf(source_path, out_md, **vision_kwargs)
        elif ext in OFFICE_EXTENSIONS:
            convert_office(
                source_path,
                out_md,
                docx_page_markers=args.docx_page_markers,
                **vision_kwargs,
            )
        else:
            raise SystemExit(f"Unsupported file extension: {ext}")
        return

    if not source_path.is_dir():
        raise SystemExit(f"Source directory not found: {source_path}")

    if effective_vision_mode == "off":
        counts = convert_tree(
            source_path,
            output_root=output_root,
            docx_page_markers=args.docx_page_markers,
        )
    else:
        counts = convert_tree(
            source_path,
            output_root=output_root,
            docx_page_markers=args.docx_page_markers,
            vision_mode=effective_vision_mode,
            api_key=args.api_key,
            vision_model=args.vision_model,
            low_threshold=args.vision_low_threshold,
            high_threshold=args.vision_high_threshold,
            dpi=args.vision_dpi,
            concurrency=args.vision_concurrency,
            max_retries=args.vision_max_retries,
        )
    total = counts["pdf"] + counts["office"]
    target_output_dir = output_root if output_root else (source_path / OUTPUT_DIR_NAME)
    LOGGER.info(
        "Converted %d file(s) (%d PDF, %d office, %d skipped). Output is in %s",
        total,
        counts["pdf"],
        counts["office"],
        counts["skipped"],
        target_output_dir,
    )


if __name__ == "__main__":
    main()
