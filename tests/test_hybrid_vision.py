"""Tests for document vision screening and hybrid local + Gemini vision extraction."""

from __future__ import annotations

import io
import json
import sys
import zipfile
from pathlib import Path
from types import SimpleNamespace

import pymupdf
import pytest
from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.util import Inches

import pdf_to_md.hybrid_vision as hv
from pdf_to_md import convert_office, convert_pdf, convert_tree, screen_document, screen_tree
from pdf_to_md.hybrid_vision import choose_strategy_from_ratio


def _make_detailed_png(width: int = 240, height: int = 120, seed_color: tuple[int, int, int] = (20, 80, 180)) -> bytes:
    """Create a multi-color diagram-like PNG (passes visual detail screening)."""
    img = Image.new("RGB", (width, height), (248, 250, 252))
    draw = ImageDraw.Draw(img)
    draw.rectangle((12, 12, width // 2 - 8, height - 12), fill=seed_color, outline=(10, 20, 40))
    draw.rectangle((width // 2 + 8, 20, width - 12, height - 20), fill=(40, 150, 90), outline=(10, 20, 40))
    draw.line((width // 2 - 8, height // 2, width // 2 + 8, height // 2), fill=(0, 0, 0), width=3)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def _make_solid_png(width: int = 240, height: int = 80) -> bytes:
    """Create a single-color spacer image (rejected by visual detail screening)."""
    img = Image.new("RGB", (width, height), (220, 220, 220))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


class _FakeModels:
    def __init__(self, handler):
        self._handler = handler
        self.calls: list[dict] = []

    async def generate_content(self, *, model, contents, config):
        parts = contents[0]["parts"]
        prompt_text = parts[0]["text"]
        inline_data = parts[1]["inline_data"]
        record = {
            "model": model,
            "config": config,
            "prompt": prompt_text,
            "mime_type": inline_data["mime_type"],
            "data_len": len(inline_data["data"]),
        }
        self.calls.append(record)
        payload = self._handler(prompt_text)
        return SimpleNamespace(text=json.dumps(payload))


class _FakeAsyncClient:
    def __init__(self, handler):
        self.models = _FakeModels(handler)
        self.closed = False

    async def aclose(self):
        self.closed = True


def test_threshold_routing_boundaries():
    assert choose_strategy_from_ratio(10, 0) == "none"
    assert choose_strategy_from_ratio(10, 1) == "targeted"
    assert choose_strategy_from_ratio(10, 3) == "targeted"
    assert choose_strategy_from_ratio(10, 4) == "hybrid_pages"
    assert choose_strategy_from_ratio(10, 7) == "hybrid_pages"
    assert choose_strategy_from_ratio(10, 8) == "full_document"
    assert choose_strategy_from_ratio(10, 10) == "full_document"

    assert (
        choose_strategy_from_ratio(10, 3, low_threshold=0.20, high_threshold=0.60)
        == "hybrid_pages"
    )
    assert (
        choose_strategy_from_ratio(10, 7, low_threshold=0.20, high_threshold=0.60)
        == "full_document"
    )

    with pytest.raises(ValueError, match="low_threshold"):
        screen_document(Path(__file__), low_threshold=0.8, high_threshold=0.3)


def test_screening_filters_repeating_logos_and_solid_spacers(tmp_path):
    pdf_path = tmp_path / "company_spec.pdf"
    logo_png = _make_detailed_png(120, 50, seed_color=(180, 30, 30))
    diagram_png = _make_detailed_png(300, 150, seed_color=(25, 95, 190))
    solid_png = _make_solid_png(300, 90)

    doc = pymupdf.open()
    try:
        logo_xref = None
        for p_idx in range(5):
            page = doc.new_page(width=595, height=842)
            page.insert_text(
                (56, 110),
                f"Section {p_idx + 1}: Native architecture prose describing module responsibilities.",
                fontsize=11,
            )
            # Repeating header logo across all 5 pages (reusing xref via stream first time)
            if logo_xref is None:
                logo_xref = page.insert_image(pymupdf.Rect(450, 20, 550, 62), stream=logo_png)
            else:
                page.insert_image(pymupdf.Rect(450, 20, 550, 62), xref=logo_xref)

            if p_idx == 1:
                # Page 2: solid spacer bar (should be ignored)
                page.insert_image(pymupdf.Rect(56, 160, 400, 240), stream=solid_png)
            if p_idx == 3:
                # Page 4: real substantive diagram
                page.insert_image(pymupdf.Rect(56, 180, 420, 340), stream=diagram_png)
        doc.save(pdf_path)
    finally:
        doc.close()

    report = screen_document(pdf_path)
    assert report.total_units == 5
    assert report.candidate_units == 1
    assert report.candidate_unit_numbers == [4]
    assert report.vision_ratio == pytest.approx(0.20)
    assert report.recommended_strategy == "targeted"


def test_pdf_targeted_vision_passes_context_and_keeps_single_valid_md(tmp_path, monkeypatch):
    pdf_path = tmp_path / "pega_like.pdf"
    swagger_png = _make_detailed_png(320, 110)
    standalone_png = _make_detailed_png(340, 160, seed_color=(90, 40, 160))

    doc = pymupdf.open()
    try:
        for p_idx in range(5):
            page = doc.new_page(width=595, height=842)
            if p_idx != 2:
                page.insert_text(
                    (56, 80),
                    f"Pure native prose on page {p_idx + 1} that needs zero vision calls.",
                    fontsize=11,
                )
                continue

            # Page 3 (1/5 = 20% <= 30% -> targeted image mode)
            page.insert_text((56, 55), "CrossAppComm", fontsize=16, fontname="hebo")
            page.insert_text((56, 85), "Architecture flow diagram below:", fontsize=11)
            page.insert_image(pymupdf.Rect(56, 100, 380, 230), stream=standalone_png)

            # 2-column stub table containing an embedded API screenshot inside Purpose cell
            edges = [56.0, 150.0, 540.0]
            page.draw_rect(pymupdf.Rect(56, 280, 150, 460), color=None, fill=(0.94, 0.94, 0.95))
            page.draw_rect(pymupdf.Rect(150, 280, 540, 460), color=None, fill=(1.0, 1.0, 1.0))
            for x in edges:
                page.draw_line((x, 280), (x, 460), color=(0.8, 0.8, 0.8), width=1)
            page.draw_line((56, 280), (540, 280), color=(0.8, 0.8, 0.8), width=1)
            page.draw_line((56, 460), (540, 460), color=(0.8, 0.8, 0.8), width=1)

            page.insert_text((62, 302), "Purpose", fontsize=11)
            page.insert_text(
                (158, 302),
                "Exposes out-of-the-box Pega REST case endpoints.",
                fontsize=11,
            )
            page.insert_image(pymupdf.Rect(158, 320, 460, 415), stream=swagger_png)
            page.insert_text((158, 440), "Pega API library", fontsize=11)
        doc.save(pdf_path)
    finally:
        doc.close()

    def _handler(prompt: str) -> dict:
        if "Exposes out-of-the-box Pega REST" in prompt:
            return {
                "is_decorative": False,
                "markdown": (
                    "### Swagger Endpoints\n"
                    "| Method | Path |\n"
                    "| --- | --- |\n"
                    "| GET | /cases |\n"
                    "| POST | /cases |"
                ),
            }
        return {
            "is_decorative": False,
            "markdown": "- **OrderService** --[REST]--> **BillingService**",
        }

    fake_client = _FakeAsyncClient(_handler)
    monkeypatch.setattr(hv, "_create_async_client", lambda api_key: fake_client)

    out_md = tmp_path / "out" / "pega_like.md"
    convert_pdf(pdf_path, out_md, vision_mode="auto", api_key="test-key")

    # Only the 2 images on Page 3 should be sent (0 calls for Pages 1, 2, 4, 5)
    assert len(fake_client.models.calls) == 2
    assert all("Section: CrossAppComm" in c["prompt"] for c in fake_client.models.calls)
    assert any("Purpose: Exposes out-of-the-box Pega REST" in c["prompt"] for c in fake_client.models.calls)

    text = out_md.read_text(encoding="utf-8")
    assert text.startswith("[[pega_like.pdf]]")
    for p in range(1, 6):
        assert f"<!-- page: {p} -->" in text

    # Standalone image uses blockquote
    assert "> **Visual content:**\n> - **OrderService** --[REST]--> **BillingService**" in text
    # Table-cell image keeps GFM pipe row intact on a single line (no raw newlines inside row)
    purpose_lines = [ln for ln in text.splitlines() if "**Purpose**" in ln]
    assert len(purpose_lines) == 1
    assert "GET — /cases" in purpose_lines[0]
    assert "POST — /cases" in purpose_lines[0]
    assert "Pega API library" in purpose_lines[0]


def test_pdf_hybrid_pages_and_full_document_modes(tmp_path, monkeypatch):
    diagram_png = _make_detailed_png(320, 180)
    pdf_path = tmp_path / "mixed_deck.pdf"

    doc = pymupdf.open()
    try:
        # 4 pages: pages 2 and 4 have diagrams -> 2/4 = 50% -> hybrid_pages
        for p_idx in range(4):
            page = doc.new_page(width=595, height=842)
            page.insert_text(
                (56, 70),
                f"Page {p_idx + 1} native SharedDataModelLatest identifier.",
                fontsize=12,
            )
            if p_idx in (1, 3):
                page.insert_image(pymupdf.Rect(56, 120, 440, 360), stream=diagram_png)
        doc.save(pdf_path)
    finally:
        doc.close()

    def _page_handler(prompt: str) -> dict:
        for p in (1, 2, 3, 4):
            if f"Page number: {p}" in prompt:
                return {"markdown_body": f"## Vision Page {p}\n\nPreserved SharedDataModelLatest"}
        return {"markdown_body": "Fallback"}

    # 1. Default thresholds (30%..70%): 50% -> hybrid_pages (only pages 2 & 4 call vision)
    fake_hybrid = _FakeAsyncClient(_page_handler)
    monkeypatch.setattr(hv, "_create_async_client", lambda api_key: fake_hybrid)
    out_hybrid = tmp_path / "out_hybrid" / "mixed_deck.md"
    convert_pdf(pdf_path, out_hybrid, vision_mode="auto", api_key="test-key")

    assert len(fake_hybrid.models.calls) == 2
    assert all("SharedDataModelLatest" in c["prompt"] for c in fake_hybrid.models.calls)
    hybrid_text = out_hybrid.read_text(encoding="utf-8")
    assert "Page 1 native SharedDataModelLatest identifier." in hybrid_text
    assert "## Vision Page 2" in hybrid_text
    assert "Page 3 native SharedDataModelLatest identifier." in hybrid_text
    assert "## Vision Page 4" in hybrid_text

    # 2. Lower high_threshold=0.40: 50% > 40% -> full_document (all 4 pages call vision, still 1 file)
    fake_full = _FakeAsyncClient(_page_handler)
    monkeypatch.setattr(hv, "_create_async_client", lambda api_key: fake_full)
    out_full = tmp_path / "out_full" / "mixed_deck.md"
    convert_pdf(
        pdf_path,
        out_full,
        vision_mode="auto",
        api_key="test-key",
        low_threshold=0.20,
        high_threshold=0.40,
    )
    assert len(fake_full.models.calls) == 4
    full_text = out_full.read_text(encoding="utf-8")
    for p in (1, 2, 3, 4):
        assert f"<!-- page: {p} -->\n\n## Vision Page {p}" in full_text


def test_zero_candidate_auto_mode_needs_no_api_key(tmp_path, monkeypatch):
    monkeypatch.delenv("GEMINI_API_KEY", raising=False)
    monkeypatch.chdir(tmp_path)  # ensure no .env in cwd

    pdf_path = tmp_path / "plain_text.pdf"
    doc = pymupdf.open()
    try:
        page = doc.new_page()
        page.insert_text((72, 72), "Pure text document with zero visual candidates.")
        doc.save(pdf_path)
    finally:
        doc.close()

    out_md = tmp_path / "out" / "plain_text.md"
    convert_pdf(pdf_path, out_md, vision_mode="auto", api_key=None)
    assert "Pure text document with zero visual candidates." in out_md.read_text(encoding="utf-8")


def _write_docx_with_image(path: Path, png_bytes: bytes) -> None:
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""
    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""
    doc_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rIdImg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>"""
    styles = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>
</w:styles>"""
    # 4 pages via lastRenderedPageBreak, image on page 2 -> 1/4 = 25% <= 30% -> targeted
    doc_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>Page one overview text.</w:t></w:r></w:p>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:lastRenderedPageBreak/><w:t>Deployment Topology</w:t></w:r></w:p>
    <w:p>
      <w:r><w:t>Cluster routing diagram:</w:t></w:r>
      <w:r>
        <w:drawing>
          <wp:inline>
            <wp:extent cx="3048000" cy="1524000"/>
            <a:graphic><a:graphicData><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:blipFill><a:blip r:embed="rIdImg1"/></pic:blipFill>
            </pic:pic></a:graphicData></a:graphic>
          </wp:inline>
        </w:drawing>
      </w:r>
    </w:p>
    <w:p><w:r><w:lastRenderedPageBreak/><w:t>Page three details.</w:t></w:r></w:p>
    <w:p><w:r><w:lastRenderedPageBreak/><w:t>Page four appendix.</w:t></w:r></w:p>
  </w:body>
</w:document>"""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("word/_rels/document.xml.rels", doc_rels)
        zf.writestr("word/styles.xml", styles)
        zf.writestr("word/document.xml", doc_xml)
        zf.writestr("word/media/image1.png", png_bytes)


def test_docx_and_pptx_targeted_vision_single_file(tmp_path, monkeypatch):
    img_bytes = _make_detailed_png(280, 140)

    # 1. DOCX test (1/4 pages = 25% -> targeted)
    docx_path = tmp_path / "design.docx"
    _write_docx_with_image(docx_path, img_bytes)
    docx_report = screen_document(docx_path)
    assert docx_report.total_units == 4
    assert docx_report.candidate_units == 1
    assert docx_report.recommended_strategy == "targeted"

    fake_docx_client = _FakeAsyncClient(
        lambda prompt: {
            "is_decorative": False,
            "markdown": "- **Ingress** --> **PGPool** --> **PostgreSQL**",
        }
    )
    monkeypatch.setattr(hv, "_create_async_client", lambda api_key: fake_docx_client)
    out_docx_md = tmp_path / "out" / "design.md"
    convert_office(docx_path, out_docx_md, vision_mode="auto", api_key="test-key")

    assert len(fake_docx_client.models.calls) == 1
    assert "Deployment Topology" in fake_docx_client.models.calls[0]["prompt"]
    docx_text = out_docx_md.read_text(encoding="utf-8")
    assert docx_text.startswith("[[design.docx]]")
    assert docx_text.count("<!-- page:") == 4
    assert "![](design_images/" in docx_text
    assert "- **Ingress** --> **PGPool** --> **PostgreSQL**" in docx_text

    # 2. PPTX test (1/4 slides = 25% -> targeted, works even without LibreOffice)
    pptx_path = tmp_path / "arch_deck.pptx"
    prs = Presentation()
    for s_idx in range(4):
        layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(layout)
        slide.shapes.title.text = f"Slide {s_idx + 1} Architecture"
        slide.placeholders[1].text = f"Key bullet for slide {s_idx + 1}."
        if s_idx == 1:
            slide.shapes.add_picture(
                io.BytesIO(img_bytes), Inches(1.5), Inches(2.5), width=Inches(4.0), height=Inches(2.0)
            )
    prs.save(str(pptx_path))

    pptx_report = screen_document(pptx_path)
    assert pptx_report.total_units == 4
    assert pptx_report.candidate_units == 1
    assert pptx_report.recommended_strategy == "targeted"

    fake_pptx_client = _FakeAsyncClient(
        lambda prompt: {
            "is_decorative": False,
            "markdown": "- **Kafka** topic `audit-events` consumed by **AuditService**",
        }
    )
    monkeypatch.setattr(hv, "_create_async_client", lambda api_key: fake_pptx_client)
    out_pptx_md = tmp_path / "out" / "arch_deck.md"
    convert_office(pptx_path, out_pptx_md, vision_mode="auto", api_key="test-key")

    assert len(fake_pptx_client.models.calls) == 1
    assert "Slide 2 Architecture" in fake_pptx_client.models.calls[0]["prompt"]
    pptx_text = out_pptx_md.read_text(encoding="utf-8")
    assert pptx_text.startswith("[[arch_deck.pptx]]")
    assert pptx_text.count("<!-- page:") == 4
    assert "Key bullet for slide 4." in pptx_text
    assert "- **Kafka** topic `audit-events` consumed by **AuditService**" in pptx_text


def test_cli_screen_only_and_convert_tree_vision(tmp_path, monkeypatch, caplog):
    import pdf_to_md.main as main_module

    src_dir = tmp_path / "docs"
    src_dir.mkdir()
    pdf_path = src_dir / "note.pdf"
    doc = pymupdf.open()
    try:
        p = doc.new_page()
        p.insert_text((72, 72), "Simple text note.")
        doc.save(pdf_path)
    finally:
        doc.close()

    reports = screen_tree(src_dir)
    assert len(reports) == 1
    assert reports[0].recommended_strategy == "none"

    caplog.set_level("INFO")
    monkeypatch.setattr(sys, "argv", ["pdf-to-md", str(src_dir), "--screen-only"])
    main_module.main()
    assert any("note.pdf: 0/1 visual (0.0%) -> none" in rec.message for rec in caplog.records)

    out_dir = tmp_path / "extracts"
    counts = convert_tree(src_dir, out_dir, vision_mode="auto", api_key=None)
    assert counts["pdf"] == 1
    assert "Simple text note." in (out_dir / "note.md").read_text(encoding="utf-8")


def _make_detailed_bmp(width: int = 240, height: int = 120) -> bytes:
    """Create a multi-color BMP image (non-PNG/non-JPEG raster)."""
    img = Image.new("RGB", (width, height), (248, 250, 252))
    draw = ImageDraw.Draw(img)
    draw.rectangle((12, 12, width // 2 - 8, height - 12), fill=(30, 90, 200))
    draw.rectangle((width // 2 + 8, 20, width - 12, height - 20), fill=(40, 150, 90))
    buf = io.BytesIO()
    img.save(buf, format="BMP")
    return buf.getvalue()


def test_hybrid_model_and_thinking_level_forwarding_and_validation(tmp_path, monkeypatch):
    from pdf_to_md import DEFAULT_MODEL
    import pdf_to_md.main as main_module

    assert DEFAULT_MODEL == "gemini-4-hydrogen"

    pdf_path = tmp_path / "diagram.pdf"
    doc = pymupdf.open()
    try:
        p1 = doc.new_page(width=595, height=842)
        p1.insert_text((56, 72), "Page 1 prose.")
        p1.insert_image(pymupdf.Rect(56, 120, 420, 320), stream=_make_detailed_png(300, 160))
        p2 = doc.new_page(width=595, height=842)
        p2.insert_text((56, 72), "Page 2 prose only.")
        p3 = doc.new_page(width=595, height=842)
        p3.insert_text((56, 72), "Page 3 prose only.")
        p4 = doc.new_page(width=595, height=842)
        p4.insert_text((56, 72), "Page 4 prose only.")
        doc.save(pdf_path)
    finally:
        doc.close()

    # 1. Default model ("gemini-4-hydrogen") and thinking_level=None -> omits thinking_config
    default_client = _FakeAsyncClient(
        lambda prompt: {"is_decorative": False, "markdown": "- **ServiceA** -> **ServiceB**"}
    )
    monkeypatch.setattr(hv, "_create_async_client", lambda api_key: default_client)
    convert_pdf(pdf_path, tmp_path / "out_default" / "diagram.md", vision_mode="auto", api_key="k")
    assert len(default_client.models.calls) == 1
    assert default_client.models.calls[0]["model"] == "gemini-4-hydrogen"
    assert default_client.models.calls[0]["config"] == {"response_mime_type": "application/json"}

    # 2. Explicit model + case-insensitive thinking_level on convert_pdf
    custom_client = _FakeAsyncClient(
        lambda prompt: {"is_decorative": False, "markdown": "- **ServiceA** -> **ServiceB**"}
    )
    monkeypatch.setattr(hv, "_create_async_client", lambda api_key: custom_client)
    convert_pdf(
        pdf_path,
        tmp_path / "out_custom" / "diagram.md",
        vision_mode="auto",
        api_key="k",
        vision_model="gemini-3.8-flash",
        thinking_level="XHIGH",
    )
    assert custom_client.models.calls[0]["model"] == "gemini-3.8-flash"
    assert custom_client.models.calls[0]["config"] == {
        "response_mime_type": "application/json",
        "thinking_config": {"thinking_level": "XHIGH"},
    }

    # 3. Full-page mode + convert_office + convert_tree + CLI forwarding
    page_client = _FakeAsyncClient(lambda prompt: {"markdown_body": "# Full page"})
    monkeypatch.setattr(hv, "_create_async_client", lambda api_key: page_client)
    convert_pdf(
        pdf_path,
        tmp_path / "out_full" / "diagram.md",
        vision_mode="full",
        api_key="k",
        vision_model="gemini-4-hydrogen",
        thinking_level="low",
    )
    assert len(page_client.models.calls) == 4
    assert all(
        c["model"] == "gemini-4-hydrogen"
        and c["config"]
        == {
            "response_mime_type": "application/json",
            "thinking_config": {"thinking_level": "LOW"},
        }
        for c in page_client.models.calls
    )

    docx_path = tmp_path / "doc.docx"
    _write_docx_with_image(docx_path, _make_detailed_png(280, 140))
    office_client = _FakeAsyncClient(
        lambda prompt: {"is_decorative": False, "markdown": "- **Node**"}
    )
    monkeypatch.setattr(hv, "_create_async_client", lambda api_key: office_client)
    convert_office(
        docx_path,
        tmp_path / "out_docx" / "doc.md",
        vision_mode="auto",
        api_key="k",
        vision_model="gemini-3.8-flash",
        thinking_level="medium",
    )
    assert office_client.models.calls[0]["model"] == "gemini-3.8-flash"
    assert office_client.models.calls[0]["config"] == {
        "response_mime_type": "application/json",
        "thinking_config": {"thinking_level": "MEDIUM"},
    }

    tree_src = tmp_path / "tree_src"
    tree_src.mkdir()
    (tree_src / "diagram.pdf").write_bytes(pdf_path.read_bytes())
    tree_client = _FakeAsyncClient(
        lambda prompt: {"is_decorative": False, "markdown": "- **Tree**"}
    )
    monkeypatch.setattr(hv, "_create_async_client", lambda api_key: tree_client)
    convert_tree(
        tree_src,
        tmp_path / "tree_out",
        vision_mode="auto",
        api_key="k",
        vision_model="gemini-4-hydrogen",
        thinking_level="high",
    )
    assert tree_client.models.calls[0]["model"] == "gemini-4-hydrogen"
    assert tree_client.models.calls[0]["config"]["thinking_config"] == {
        "thinking_level": "HIGH"
    }

    cli_client = _FakeAsyncClient(
        lambda prompt: {"is_decorative": False, "markdown": "- **CLI**"}
    )
    monkeypatch.setattr(hv, "_create_async_client", lambda api_key: cli_client)
    monkeypatch.setattr(
        sys,
        "argv",
        [
            "pdf-to-md",
            str(pdf_path),
            "-o",
            str(tmp_path / "cli_out"),
            "--vision",
            "--api-key",
            "k",
            "--vision-model",
            "gemini-3.8-flash",
            "--vision-thinking-level",
            "low",
        ],
    )
    main_module.main()
    assert cli_client.models.calls[0]["model"] == "gemini-3.8-flash"
    assert cli_client.models.calls[0]["config"]["thinking_config"] == {
        "thinking_level": "LOW"
    }

    # 4. Validation failures
    with pytest.raises(ValueError, match="thinking_level"):
        convert_pdf(
            pdf_path,
            tmp_path / "bad.md",
            vision_mode="auto",
            api_key="k",
            thinking_level="ultra",
        )
    with pytest.raises(ValueError, match="thinking_level"):
        hv.run_hybrid_vision_tasks(
            image_tasks=[],
            page_tasks=[],
            api_key="k",
            thinking_level="",
        )
    with pytest.raises(ValueError, match="vision_model"):
        convert_pdf(
            pdf_path,
            tmp_path / "bad_model.md",
            vision_mode="auto",
            api_key="k",
            vision_model="   ",
        )


def test_pptx_non_png_raster_is_converted_to_real_png(tmp_path, monkeypatch):
    pptx_path = tmp_path / "bmp_deck.pptx"
    bmp_bytes = _make_detailed_bmp(240, 120)
    prs = Presentation()
    for s_idx in range(4):
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = f"Slide {s_idx + 1}"
        slide.placeholders[1].text = f"Text {s_idx + 1}"
        if s_idx == 1:
            slide.shapes.add_picture(
                io.BytesIO(bmp_bytes), Inches(1.5), Inches(2.5), width=Inches(4.0), height=Inches(2.0)
            )
    prs.save(str(pptx_path))

    fake_client = _FakeAsyncClient(
        lambda prompt: {"is_decorative": False, "markdown": "- **Diagram**"}
    )
    monkeypatch.setattr(hv, "_create_async_client", lambda api_key: fake_client)

    out_md = tmp_path / "out" / "bmp_deck.md"
    convert_office(pptx_path, out_md, vision_mode="auto", api_key="k")

    saved_pngs = list((tmp_path / "out" / "bmp_deck_images").glob("*.png"))
    assert len(saved_pngs) == 1
    raw_saved = saved_pngs[0].read_bytes()
    assert raw_saved.startswith(b"\x89PNG\r\n\x1a\n")
    with Image.open(io.BytesIO(raw_saved)) as im:
        assert im.format == "PNG"
    assert len(fake_client.models.calls) == 1
    assert fake_client.models.calls[0]["mime_type"] == "image/png"
