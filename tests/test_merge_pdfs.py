"""Tests for PDF merging in merge_pdfs."""

from __future__ import annotations

import sys
from pathlib import Path

import pymupdf
import pytest

from pdf_to_md.merge_pdfs import main as merge_main
from pdf_to_md.merge_pdfs import merge_pdfs


def _make_pdf(path: Path, text: str) -> None:
    doc = pymupdf.open()
    try:
        page = doc.new_page()
        page.insert_text((72, 72), text)
        doc.save(path)
    finally:
        doc.close()


def test_merge_pdfs_alphabetical_order(tmp_path):
    src = tmp_path / "pdfs"
    src.mkdir()
    _make_pdf(src / "b_second.pdf", "Second PDF")
    _make_pdf(src / "A_first.pdf", "First PDF")
    _make_pdf(src / "c_third.pdf", "Third PDF")

    out_pdf = tmp_path / "out" / "merged.pdf"
    count = merge_pdfs(src, out_pdf)

    assert count == 3
    assert out_pdf.is_file()
    with pymupdf.open(out_pdf) as merged:
        assert merged.page_count == 3
        texts = [merged[i].get_text("text").strip() for i in range(3)]
    assert texts == ["First PDF", "Second PDF", "Third PDF"]


def test_merge_pdfs_missing_or_empty_raises_file_not_found(tmp_path):
    missing_dir = tmp_path / "does_not_exist"
    with pytest.raises(FileNotFoundError):
        merge_pdfs(missing_dir, tmp_path / "merged.pdf")

    empty_dir = tmp_path / "empty"
    empty_dir.mkdir()
    with pytest.raises(FileNotFoundError):
        merge_pdfs(empty_dir, tmp_path / "merged.pdf")


def test_merge_pdfs_cli_exits_when_empty(tmp_path, monkeypatch):
    empty_dir = tmp_path / "empty"
    empty_dir.mkdir()
    monkeypatch.setattr(
        sys,
        "argv",
        ["merge-pdfs", str(empty_dir), "-o", str(tmp_path / "out.pdf")],
    )
    with pytest.raises(SystemExit):
        merge_main()
