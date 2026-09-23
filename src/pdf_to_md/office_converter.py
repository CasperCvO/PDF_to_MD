"""Office-document-to-Markdown converter module.

Converts office files (docx, xlsx, pptx, xls, doc, …) to Markdown using
Microsoft's *markitdown* library — except ``.xlsx`` workbooks, which are
converted with openpyxl so every sheet gets a ``<!-- sheet: "Name"
range: A1:D4 -->`` locator and a row/column-addressed table for provenance.
Each output file is prepended with a Wikilink to the original document for
easy back-referencing in Obsidian.
"""

from __future__ import annotations

import datetime
import logging
from pathlib import Path
from typing import Any

LOGGER = logging.getLogger(__name__)

SUPPORTED_EXTENSIONS: set[str] = {
    ".docx",
    ".doc",
    ".xlsx",
    ".xls",
    ".pptx",
    ".ppt",
}


def _format_cell(value: Any) -> str:
    """Render a single cell value as Markdown-safe text."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, datetime.datetime):
        if value.time() == datetime.time(0, 0):
            return value.date().isoformat()
        return value.isoformat(sep=" ")
    if isinstance(value, (datetime.date, datetime.time)):
        return value.isoformat()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, str):
        text = value.strip()
        text = (
            text.replace("\r\n", "<br>").replace("\r", "<br>").replace("\n", "<br>")
        )
        return text.replace("|", "\\|")
    return str(value)


def _xlsx_to_markdown(file_path: Path) -> str:
    """Return Markdown for an ``.xlsx`` workbook with sheet/cell locators."""
    try:
        from openpyxl import load_workbook
        from openpyxl.utils import get_column_letter
    except ImportError as exc:
        raise RuntimeError(
            "Missing dependency. Install openpyxl to enable .xlsx conversion."
        ) from exc

    formulas_wb = load_workbook(file_path, data_only=False)
    values_wb = load_workbook(file_path, data_only=True)

    sheets = []
    for ws in formulas_wb.worksheets:
        values_ws = values_wb[ws.title]
        non_empty: set[tuple[int, int]] = set()
        for row_cells in ws.iter_rows():
            for cell in row_cells:
                value = cell.value
                if value is not None and (
                    not isinstance(value, str) or value.strip()
                ):
                    non_empty.add((cell.row, cell.column))

        heading = f"## Sheet: {ws.title}"
        if not non_empty:
            sheets.append(f"{heading}\n\n_(empty sheet)_")
            continue

        rows = sorted({row for row, _ in non_empty})
        cols = sorted({col for _, col in non_empty})
        min_row, max_row = rows[0], rows[-1]
        min_col, max_col = cols[0], cols[-1]
        top_left = f"{get_column_letter(min_col)}{min_row}"
        bottom_right = f"{get_column_letter(max_col)}{max_row}"
        escaped_title = ws.title.replace('"', '\\"')
        locator = f'<!-- sheet: "{escaped_title}" range: {top_left}:{bottom_right} -->'

        header = "| Row | " + " | ".join(get_column_letter(col) for col in cols) + " |"
        separator = "| --- |" + " --- |" * len(cols)
        lines = [header, separator]
        for row in rows:
            cells = []
            for col in cols:
                if (row, col) not in non_empty:
                    cells.append("")
                    continue
                value = values_ws.cell(row=row, column=col).value
                if value is None:
                    formula = ws.cell(row=row, column=col).value
                    value = getattr(formula, "text", None) or str(formula)
                cells.append(_format_cell(value))
            lines.append(f"| {row} | " + " | ".join(cells) + " |")

        sheets.append(f"{heading}\n\n{locator}\n\n" + "\n".join(lines))

    return "\n\n".join(sheets)


def _extract_markdown(file_path: Path) -> str:
    """Return Markdown text for *file_path*."""
    if file_path.suffix.lower() == ".xlsx":
        return _xlsx_to_markdown(file_path)

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
