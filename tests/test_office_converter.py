"""Tests for XLSX sheet/cell locators in office_converter."""

from datetime import datetime
from pathlib import Path

from openpyxl import Workbook

from pdf_to_md.office_converter import convert_office

EXPECTED_DATA_BLOCK = '''## Sheet: Data

<!-- sheet: "Data" range: A1:D4 -->

| Row | A | B | D |
| --- | --- | --- | --- |
| 1 | Name | Amount |  |
| 2 | Alice | 100 |  |
| 4 | Bob\\|Smith | =SUM(B2:B2) | x |'''


def _write_test_workbook(path: Path) -> None:
    wb = Workbook()

    ws = wb.active
    ws.title = "Data"
    ws["A1"] = "Name"
    ws["B1"] = "Amount"
    ws["A2"] = "Alice"
    ws["B2"] = 100
    ws["A4"] = "Bob|Smith"
    ws["B4"] = "=SUM(B2:B2)"
    ws["D4"] = "x"

    wb.create_sheet("Notes")

    offset = wb.create_sheet("Offset")
    offset["C5"] = 2.5
    offset["D6"] = datetime(2024, 1, 31)

    misc = wb.create_sheet("Misc")
    misc["A1"] = "line one\nline two"
    misc["B1"] = True

    wb.save(path)


def test_xlsx_markdown(tmp_path):
    xlsx_path = tmp_path / "book.xlsx"
    _write_test_workbook(xlsx_path)
    output_path = tmp_path / "out" / "book.md"

    convert_office(xlsx_path, output_path)
    text = output_path.read_text(encoding="utf-8")

    assert text.splitlines()[0] == "[[book.xlsx]]"
    assert EXPECTED_DATA_BLOCK in text
    assert '<!-- sheet: "Offset" range: C5:D6 -->' in text
    assert "| 5 | 2.5 |  |" in text
    assert "| 6 |  | 2024-01-31 |" in text
    assert "## Sheet: Notes\n\n_(empty sheet)_" in text
    assert "| 3 |" not in text
    assert "line one<br>line two" in text
    assert "| TRUE |" in text
