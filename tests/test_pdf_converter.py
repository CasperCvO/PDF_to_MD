"""Tests for PDF page-locator markers in pdf_converter."""

from pathlib import Path

import pymupdf

from pdf_to_md.pdf_converter import convert_pdf


def _write_test_pdf(path: Path) -> None:
    doc = pymupdf.open()
    try:
        page = doc.new_page()
        page.insert_text((72, 72), "Alpha page one")
        doc.new_page()
        page = doc.new_page()
        page.insert_text((72, 72), "Gamma page three")
        doc.save(path)
    finally:
        doc.close()


def _assert_page_markers(text: str, name: str) -> None:
    assert text.splitlines()[0] == f"[[{name}]]"

    m1 = text.index("<!-- page: 1 -->")
    m2 = text.index("<!-- page: 2 -->")
    m3 = text.index("<!-- page: 3 -->")
    assert m1 < m2 < m3

    alpha = text.index("Alpha page one")
    gamma = text.index("Gamma page three")
    assert m1 < alpha < m2
    assert gamma > m3


def test_convert_pdf_adds_page_markers(tmp_path):
    pdf_path = tmp_path / "sample.pdf"
    _write_test_pdf(pdf_path)
    output_path = tmp_path / "out" / "sample.md"

    convert_pdf(pdf_path, output_path)

    _assert_page_markers(output_path.read_text(encoding="utf-8"), pdf_path.name)


def test_convert_pdf_fallback_adds_page_markers(tmp_path, monkeypatch):
    import pymupdf4llm

    def _failing_to_markdown(*args, **kwargs):
        raise RuntimeError("forced pymupdf4llm failure")

    monkeypatch.setattr(pymupdf4llm, "to_markdown", _failing_to_markdown)

    pdf_path = tmp_path / "sample.pdf"
    _write_test_pdf(pdf_path)
    output_path = tmp_path / "out" / "sample.md"

    convert_pdf(pdf_path, output_path)

    _assert_page_markers(output_path.read_text(encoding="utf-8"), pdf_path.name)
