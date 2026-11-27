"""Tests for PDF-to-Markdown extraction in pdf_converter."""

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


def _draw_grid_row(
    page: pymupdf.Page,
    col_edges: list[float],
    y0: float,
    y1: float,
    *,
    fills: list[tuple[float, float, float] | None] | None = None,
    top_border: bool = True,
    bottom_border: bool = True,
) -> None:
    border = (0.85, 0.85, 0.85)
    for idx in range(len(col_edges) - 1):
        fill = fills[idx] if fills and idx < len(fills) else None
        if fill is not None:
            page.draw_rect(
                pymupdf.Rect(col_edges[idx], y0, col_edges[idx + 1], y1),
                color=None,
                fill=fill,
            )
    for x in col_edges:
        page.draw_line((x, y0), (x, y1), color=border, width=1)
    if top_border:
        page.draw_line((col_edges[0], y0), (col_edges[-1], y0), color=border, width=1)
    if bottom_border:
        page.draw_line((col_edges[0], y1), (col_edges[-1], y1), color=border, width=1)


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
    import pdf_to_md.pdf_converter as pdf_conv

    def _failing_extractor(*args, **kwargs):
        raise RuntimeError("forced structured extraction failure")

    monkeypatch.setattr(
        pdf_conv, "_extract_structured_pages", _failing_extractor, raising=False
    )

    pdf_path = tmp_path / "sample.pdf"
    _write_test_pdf(pdf_path)
    output_path = tmp_path / "out" / "sample.md"

    convert_pdf(pdf_path, output_path)

    _assert_page_markers(output_path.read_text(encoding="utf-8"), pdf_path.name)


def test_callout_panel_and_1row_table_preserve_all_text(tmp_path):
    pdf_path = tmp_path / "callout_and_stub.pdf"
    doc = pymupdf.open()
    try:
        p1 = doc.new_page(width=595, height=842)
        p1.insert_text((56, 60), "Architecture Overview", fontsize=18)
        p1.insert_text(
            (56, 95),
            "This body paragraph establishes the standard prose font size across the document.",
            fontsize=11,
        )
        # Shaded callout box (< 5% page height) with small icon shape inside
        p1.draw_rect(pymupdf.Rect(56, 120, 540, 148), color=(0.2, 0.5, 0.9), fill=(0.91, 0.95, 1.0))
        p1.draw_circle((66, 134), 4, color=(0.2, 0.5, 0.9), fill=(0.2, 0.5, 0.9))
        p1.insert_text((78, 138), "Our decision is to adopt a modular architecture.", fontsize=11)

        # Standalone subsection headings interleaved with two tables on the same page
        p1.insert_text((56, 430), "First Service", fontsize=13)
        edges = [56.0, 150.0, 540.0]
        stub_fills = [(0.94, 0.94, 0.95), (1.0, 1.0, 1.0)]
        _draw_grid_row(p1, edges, 450, 490, fills=stub_fills)
        p1.insert_text((62, 474), "Name", fontsize=11, fontname="helv")
        p1.insert_text((158, 474), "FirstSvc", fontsize=11, fontname="helv")

        p1.insert_text((56, 560), "Core Service", fontsize=13)
        _draw_grid_row(p1, edges, 580, 620, fills=stub_fills)
        p1.insert_text((62, 604), "Name", fontsize=11, fontname="helv")
        p1.insert_text((158, 604), "CoreService", fontsize=11, fontname="helv")
        _draw_grid_row(p1, edges, 620, 660, fills=stub_fills)
        p1.insert_text((62, 644), "Type", fontsize=11, fontname="helv")
        p1.insert_text((158, 644), "Application", fontsize=11, fontname="helv")

        # Page 2 has a 1-row continuation fragment of that stub table
        p2 = doc.new_page(width=595, height=842)
        _draw_grid_row(p2, edges, 40, 120, fills=stub_fills)
        p2.insert_text((62, 64), "Access groups", fontsize=11, fontname="helv")
        p2.insert_text((158, 64), "CoreService:Admins", fontsize=11, fontname="helv")
        p2.insert_text((158, 86), "CoreService:Users", fontsize=11, fontname="helv")
        doc.save(pdf_path)
    finally:
        doc.close()

    out_md = tmp_path / "out" / "callout_and_stub.md"
    convert_pdf(pdf_path, out_md)
    text = out_md.read_text(encoding="utf-8")

    assert "Our decision is to adopt a modular architecture." in text
    assert "## Core Service" in text or "### Core Service" in text
    assert (
        text.index("First Service")
        < text.index("FirstSvc")
        < text.index("Core Service")
        < text.index("CoreService")
    )
    assert "Col1" not in text and "Col2" not in text
    assert "CoreService:Admins" in text
    assert "CoreService:Users" in text
    # Stub Row 0 ("Name | CoreService") must be a data row, not swallowed above |---|---|
    sep_idx = text.index("| --- | --- |")
    assert text.index("CoreService") > sep_idx


def test_multipage_header_table_and_open_top_continuation(tmp_path):
    pdf_path = tmp_path / "multipage_table.pdf"
    doc = pymupdf.open()
    edges = [56.0, 160.0, 380.0, 540.0]
    hdr_fills = [(0.92, 0.92, 0.94)] * 3
    body_fills = [(1.0, 1.0, 1.0)] * 3
    try:
        # Page 1: header row + first row reaching bottom
        p1 = doc.new_page(width=595, height=842)
        _draw_grid_row(p1, edges, 600, 635, fills=hdr_fills)
        p1.insert_text((62, 622), "Component", fontsize=11, fontname="hebo")
        p1.insert_text((166, 622), "Description", fontsize=11, fontname="hebo")
        p1.insert_text((386, 622), "Rulesets", fontsize=11, fontname="hebo")

        _draw_grid_row(p1, edges, 635, 800, fills=body_fills, bottom_border=False)
        p1.insert_text((62, 658), "DataModel", fontsize=11)
        p1.insert_text((166, 658), "Defines shared logical entities.", fontsize=11)
        p1.insert_text((386, 658), "DMCore", fontsize=11)

        # Page 2: open-top AND open-bottom mid-row continuation (vertical borders only)
        p2 = doc.new_page(width=595, height=842)
        _draw_grid_row(p2, edges, 40, 800, fills=body_fills, top_border=False, bottom_border=False)
        p2.insert_text((166, 65), "1. Process-specific entities stay local.", fontsize=11)
        p2.insert_text((166, 95), "2. Shared entities live in enterprise layer.", fontsize=11)
        p2.insert_text((386, 65), "DMTriage", fontsize=11)
        p2.insert_text((386, 95), "DMRequest", fontsize=11)

        # Page 3: open-top closing fragment + a new complete row
        p3 = doc.new_page(width=595, height=842)
        _draw_grid_row(p3, edges, 40, 80, fills=body_fills, top_border=False, bottom_border=True)
        p3.insert_text((386, 62), "DMAudit", fontsize=11)
        _draw_grid_row(p3, edges, 80, 130, fills=body_fills)
        p3.insert_text((62, 108), "AuthLayer", fontsize=11)
        p3.insert_text((166, 108), "Handles SSO and OIDC.", fontsize=11)
        p3.insert_text((386, 108), "AuthSSO", fontsize=11)
        doc.save(pdf_path)
    finally:
        doc.close()

    out_md = tmp_path / "out" / "multipage_table.md"
    convert_pdf(pdf_path, out_md)
    text = out_md.read_text(encoding="utf-8")

    p1_idx = text.index("<!-- page: 1 -->")
    p2_idx = text.index("<!-- page: 2 -->")
    p3_idx = text.index("<!-- page: 3 -->")
    assert p1_idx < text.index("DataModel") < p2_idx
    assert p2_idx < text.index("DMTriage") < p3_idx
    assert p3_idx < text.index("DMAudit") < text.index("AuthLayer")
    # Real column headers should be carried onto continuation pages, never Col1/Col2
    assert "Col1" not in text
    assert text.count("| Component | Description | Rulesets |") == 3


def test_vector_bullets_preserve_4_level_hierarchy_in_table_cell(tmp_path):
    pdf_path = tmp_path / "nested_bullets.pdf"
    doc = pymupdf.open()
    try:
        page = doc.new_page(width=595, height=842)
        edges = [56.0, 135.0, 540.0]
        stub_fills = [(0.94, 0.94, 0.95), (1.0, 1.0, 1.0)]
        _draw_grid_row(page, edges, 50, 220, fills=stub_fills)
        page.insert_text((62, 72), "Case types", fontsize=11)

        items = [
            (0, "Indication"),
            (1, "Medical request"),
            (2, "Medical review"),
            (3, "Consultation"),
            (1, "Information request"),
        ]
        y = 72.0
        for depth, label in items:
            x_text = 150.0 + depth * 12.0
            # Draw small 2.5x2.5pt vector bullet 8pt to the left of text
            bx = x_text - 8.0
            page.draw_rect(
                pymupdf.Rect(bx, y - 6.0, bx + 2.5, y - 3.5),
                color=(0.16, 0.16, 0.18),
                fill=(0.16, 0.16, 0.18) if depth % 2 == 0 else None,
            )
            page.insert_text((x_text, y), label, fontsize=11)
            y += 22.0
        doc.save(pdf_path)
    finally:
        doc.close()

    out_md = tmp_path / "out" / "nested_bullets.md"
    convert_pdf(pdf_path, out_md)
    text = out_md.read_text(encoding="utf-8")

    expected = (
        "- Indication<br>"
        "  - Medical request<br>"
        "    - Medical review<br>"
        "      - Consultation<br>"
        "  - Information request"
    )
    assert expected in text


def test_conservative_identifier_unwrap_without_gluing_prose(tmp_path):
    pdf_path = tmp_path / "unwrap.pdf"
    doc = pymupdf.open()
    try:
        page = doc.new_page(width=595, height=842)
        # Both columns are narrow (~80pt) so both lines touch the right cell border!
        edges = [56.0, 140.0, 218.0]
        _draw_grid_row(page, edges, 50, 80, fills=[(0.92, 0.92, 0.94), (0.92, 0.92, 0.94)])
        page.insert_text((62, 70), "Name", fontsize=11, fontname="hebo")
        page.insert_text((146, 70), "Notes", fontsize=11, fontname="hebo")

        # Row 1: forced CamelCase overflow in col 0 vs normal prose word hitting right border in col 1
        _draw_grid_row(page, edges, 80, 160)
        page.insert_text((62, 100), "ClientProfileO", fontsize=11)
        page.insert_text((62, 118), "perations", fontsize=11)
        page.insert_text((146, 100), "Management", fontsize=11)
        page.insert_text((146, 118), "is centralized", fontsize=11)
        doc.save(pdf_path)
    finally:
        doc.close()

    out_md = tmp_path / "out" / "unwrap.md"
    convert_pdf(pdf_path, out_md)
    text = out_md.read_text(encoding="utf-8")

    assert "ClientProfileOperations" in text
    assert "Management is centralized" in text
    assert "Managementis" not in text


def test_vector_diagram_saved_while_callout_box_is_not(tmp_path):
    pdf_path = tmp_path / "diagram_vs_callout.pdf"
    doc = pymupdf.open()
    try:
        page = doc.new_page(width=595, height=842)
        # 1. Bordered callout box (should NOT create an image)
        page.draw_rect(pymupdf.Rect(56, 50, 520, 90), color=(0.2, 0.4, 0.8), fill=(0.94, 0.97, 1.0))
        page.insert_text((68, 74), "Important note inside a bordered panel.", fontsize=11)

        # 2. Multi-node vector flowchart diagram (SHOULD create an image and keep labels)
        page.draw_rect(pymupdf.Rect(60, 150, 500, 260), color=(0.7, 0.7, 0.7), fill=(0.99, 0.99, 0.99))
        page.draw_rect(pymupdf.Rect(80, 180, 180, 230), color=(0.1, 0.3, 0.7), fill=(0.85, 0.92, 1.0))
        page.insert_text((100, 208), "Gateway", fontsize=10)
        page.draw_line((180, 205), (240, 205), color=(0.1, 0.1, 0.1), width=1.5)
        page.draw_rect(pymupdf.Rect(240, 180, 340, 230), color=(0.1, 0.5, 0.3), fill=(0.85, 1.0, 0.90))
        page.insert_text((262, 208), "Worker", fontsize=10)
        page.draw_line((340, 205), (390, 205), color=(0.1, 0.1, 0.1), width=1.5)
        page.draw_rect(pymupdf.Rect(390, 180, 480, 230), color=(0.6, 0.3, 0.1), fill=(1.0, 0.93, 0.85))
        page.insert_text((412, 208), "Store", fontsize=10)
        doc.save(pdf_path)
    finally:
        doc.close()

    out_md = tmp_path / "out" / "diagram_vs_callout.md"
    convert_pdf(pdf_path, out_md)
    text = out_md.read_text(encoding="utf-8")
    img_dir = tmp_path / "out" / "diagram_vs_callout_images"

    assert "Important note inside a bordered panel." in text
    assert "Gateway" in text and "Worker" in text and "Store" in text
    assert img_dir.is_dir()
    saved_imgs = list(img_dir.glob("*.png"))
    assert len(saved_imgs) == 1
    assert f"![](diagram_vs_callout_images/{saved_imgs[0].name})" in text


def test_relative_image_paths_and_inline_hyperlinks(tmp_path):
    pdf_path = tmp_path / "My Report (2026).pdf"
    doc = pymupdf.open()
    try:
        page = doc.new_page(width=595, height=842)
        prefix = "See API documentation "
        prefix_w = pymupdf.get_text_length(prefix, fontname="helv", fontsize=11)
        link_x = 56.0 + prefix_w
        link_w = pymupdf.get_text_length("here", fontname="helv", fontsize=11)
        page.insert_text((56, 70), prefix, fontsize=11)
        page.insert_text((link_x, 70), "here", fontsize=11, color=(0.0, 0.3, 0.8))
        page.insert_text((link_x + link_w, 70), ".", fontsize=11)
        page.insert_link(
            {
                "kind": pymupdf.LINK_URI,
                "from": pymupdf.Rect(link_x - 1, 58, link_x + link_w + 1, 74),
                "uri": "https://example.com/api",
            }
        )
        # Insert a real 80x60 RGB raster image
        pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 80, 60), False)
        pix.clear_with(200)
        page.insert_image(pymupdf.Rect(56, 110, 216, 230), pixmap=pix)
        doc.save(pdf_path)
    finally:
        doc.close()

    out_md = tmp_path / "out" / "My Report (2026).md"
    convert_pdf(pdf_path, out_md)
    text = out_md.read_text(encoding="utf-8")

    assert "[here](https://example.com/api)" in text
    assert "![](My_Report_2026_images/" in text
    assert str(tmp_path) not in text


def test_two_column_reading_order_and_relative_headings(tmp_path):
    pdf_path = tmp_path / "two_col.pdf"
    doc = pymupdf.open()
    try:
        page = doc.new_page(width=595, height=842)
        page.insert_text((50, 60), "Paper Title", fontsize=18, fontname="hebo")
        page.insert_text((50, 95), "1 Introduction", fontsize=14, fontname="hebo")
        page.insert_textbox(
            pymupdf.Rect(50, 120, 275, 280),
            "Left column first sentence flows across lines naturally. "
            "Left column second sentence finishes the left side before right starts.",
            fontsize=10,
        )
        page.insert_textbox(
            pymupdf.Rect(315, 120, 545, 280),
            "Right column opening sentence begins only after left column. "
            "Right column closing sentence wraps up the page.",
            fontsize=10,
        )
        doc.save(pdf_path)
    finally:
        doc.close()

    out_md = tmp_path / "out" / "two_col.md"
    convert_pdf(pdf_path, out_md)
    text = out_md.read_text(encoding="utf-8")

    assert "# Paper Title" in text
    assert "## 1 Introduction" in text
    assert text.index("Left column second sentence") < text.index(
        "Right column opening sentence"
    )

