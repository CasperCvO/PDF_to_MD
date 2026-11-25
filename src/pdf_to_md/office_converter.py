"""Office-document-to-Markdown converter module.

Converts office files (docx, xlsx, pptx, xls, doc, …) to Markdown using
Microsoft's *markitdown* library — except ``.xlsx`` workbooks, which are
converted with openpyxl so every sheet gets a ``<!-- sheet: "Name"
range: A1:D4 -->`` locator and a row/column-addressed table for provenance.

``.docx`` files additionally get approximate ``<!-- page: N -->`` markers
(on by default, disable with ``docx_page_markers=False``): token runs are
injected into ``word/document.xml`` at Word's saved layout breaks
(``w:lastRenderedPageBreak``) or, when absent, at manual page breaks,
page-break-before paragraphs and page-starting section breaks, then
rewritten into markers after conversion.

Each output file is prepended with a Wikilink to the original document for
easy back-referencing in Obsidian.
"""

from __future__ import annotations

import copy
import datetime
import hashlib
import io
import logging
import re
import zipfile
from collections import Counter
from pathlib import Path
from typing import Any

LOGGER = logging.getLogger(__name__)

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_W_BODY = f"{{{_W_NS}}}body"
_W_P = f"{{{_W_NS}}}p"
_W_PPR = f"{{{_W_NS}}}pPr"
_W_R = f"{{{_W_NS}}}r"
_W_RPR = f"{{{_W_NS}}}rPr"
_W_T = f"{{{_W_NS}}}t"
_W_BR = f"{{{_W_NS}}}br"
_W_LRPB = f"{{{_W_NS}}}lastRenderedPageBreak"
_W_PBB = f"{{{_W_NS}}}pageBreakBefore"
_W_SECTPR = f"{{{_W_NS}}}sectPr"
_W_TYPE = f"{{{_W_NS}}}type"
_W_VAL = f"{{{_W_NS}}}val"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_SKIPPED_ANCESTORS = {f"{{{_W_NS}}}txbxContent", f"{{{_W_NS}}}del"}

_PAGE_TOKEN = r"PDFTOMDPAGE\d+END"
_PAGE_TOKEN_NUMBER = re.compile(r"PDFTOMDPAGE(\d+)END")
_PAGE_TOKEN_LINE = re.compile(
    rf"^(#{{1,6}} )?((?:{_PAGE_TOKEN}[ \t]*)+)([^\n]*)", re.MULTILINE
)
_PAGE_TOKEN_INLINE = re.compile(rf"[ \t]*(?:{_PAGE_TOKEN}[ \t]*)+")

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


def _has_skipped_ancestor(element) -> bool:
    """Return True when *element* sits inside a ``w:txbxContent`` or ``w:del``."""
    return any(
        ancestor.tag in _SKIPPED_ANCESTORS
        for ancestor in element.iterancestors()
    )


def _docx_page_break_candidates(body) -> list[tuple[str, Any]]:
    """Collect ``(kind, element)`` page-break candidates in document order.

    Rendered mode: when any usable ``w:lastRenderedPageBreak`` exists, only
    those are returned — Word already emitted one after every explicit break.
    Explicit mode: manual ``w:br`` page breaks, ``w:pageBreakBefore``
    paragraphs and page-starting ``w:sectPr`` section ends (whose break lands
    on the next usable paragraph).
    """
    rendered = [
        element
        for element in body.iter(_W_LRPB)
        if not _has_skipped_ancestor(element)
    ]
    if rendered:
        return [("break", element) for element in rendered]

    candidates: list[tuple[str, Any]] = []
    pending_section_break = False
    first_paragraph_seen = False
    for element in body.iter():
        if not isinstance(element.tag, str) or _has_skipped_ancestor(element):
            continue
        if element.tag == _W_P:
            is_first = not first_paragraph_seen
            first_paragraph_seen = True
            if pending_section_break:
                candidates.append(("start", element))
                pending_section_break = False
            else:
                properties = element.find(_W_PPR)
                page_break_before = (
                    properties.find(_W_PBB) if properties is not None else None
                )
                if (
                    page_break_before is not None
                    and page_break_before.get(_W_VAL) not in {"0", "false", "off"}
                    and not is_first
                ):
                    candidates.append(("start", element))
            properties = element.find(_W_PPR)
            section = properties.find(_W_SECTPR) if properties is not None else None
            if section is not None:
                section_type = section.find(_W_TYPE)
                value = (
                    section_type.get(_W_VAL) if section_type is not None else None
                )
                if value is None or value in {"nextPage", "oddPage", "evenPage"}:
                    pending_section_break = True
        elif element.tag == _W_BR and element.get(_W_TYPE) == "page":
            candidates.append(("break", element))
    return candidates


def _make_page_token_run(token: str):
    """Build an unstyled ``w:r`` run whose text is *token*."""
    from lxml import etree

    run = etree.Element(_W_R)
    text = etree.SubElement(run, _W_T)
    text.set(_XML_SPACE, "preserve")
    text.text = token
    return run


def _insert_token_at_break(break_element, token: str) -> bool:
    """Split the run containing *break_element* and put *token* between the parts."""
    from lxml import etree

    run = break_element.getparent()
    if run is None or run.tag != _W_R:
        return False

    children = list(run)
    index = children.index(break_element)
    run_properties = run.find(_W_RPR)

    second = etree.Element(_W_R)
    if run_properties is not None:
        second.append(copy.deepcopy(run_properties))
    for child in children[index + 1:]:
        second.append(child)
    run.remove(break_element)

    parent = run.getparent()
    position = list(parent).index(run)
    parent.insert(position + 1, _make_page_token_run(token))
    if any(child.tag != _W_RPR for child in second):
        parent.insert(position + 2, second)
    return True


def _insert_token_at_paragraph_start(paragraph, token: str) -> bool:
    """Insert *token* as the first run of *paragraph*."""
    properties = paragraph.find(_W_PPR)
    index = list(paragraph).index(properties) + 1 if properties is not None else 0
    paragraph.insert(index, _make_page_token_run(token))
    return True


def _insert_docx_page_tokens(docx_bytes: bytes) -> tuple[bytes, int]:
    """Inject ``PDFTOMDPAGE<n>END`` token runs at page-break positions.

    Returns the (possibly unchanged) document bytes and the number of tokens
    inserted.  Token *n* is the page that starts at that position, so the
    first break produces ``PDFTOMDPAGE2END``.
    """
    from lxml import etree

    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as source:
        try:
            document_xml = source.read("word/document.xml")
        except KeyError:
            return docx_bytes, 0
        root = etree.fromstring(document_xml)
        body = root.find(_W_BODY)
        if body is None:
            return docx_bytes, 0

        candidates = _docx_page_break_candidates(body)
        if not candidates:
            return docx_bytes, 0

        inserted = 0
        for kind, element in candidates:
            token = f"PDFTOMDPAGE{inserted + 2}END"
            if kind == "break":
                success = _insert_token_at_break(element, token)
            else:
                success = _insert_token_at_paragraph_start(element, token)
            if success:
                inserted += 1

        if not inserted:
            return docx_bytes, 0

        new_xml = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w") as target:
            for info in source.infolist():
                data = source.read(info.filename)
                if info.filename == "word/document.xml":
                    data = new_xml
                target.writestr(info, data)

    return buffer.getvalue(), inserted


def _replace_page_tokens(markdown: str) -> str:
    """Rewrite ``PDFTOMDPAGE<n>END`` tokens as ``<!-- page: n -->`` markers."""

    def line_sub(match: re.Match) -> str:
        heading = match.group(1) or ""
        numbers = _PAGE_TOKEN_NUMBER.findall(match.group(2))
        markers = "\n\n".join(f"<!-- page: {n} -->" for n in numbers)
        rest = match.group(3)
        if rest.strip():
            return f"{markers}\n\n{heading}{rest.lstrip()}"
        return markers

    def inline_sub(match: re.Match) -> str:
        numbers = _PAGE_TOKEN_NUMBER.findall(match.group(0))
        markers = " ".join(f"<!-- page: {n} -->" for n in numbers)
        text = match.string
        lead = "" if match.start() == 0 or text[match.start() - 1] == "\n" else " "
        trail = "" if match.end() == len(text) or text[match.end()] == "\n" else " "
        return f"{lead}{markers}{trail}"

    return _PAGE_TOKEN_INLINE.sub(
        inline_sub, _PAGE_TOKEN_LINE.sub(line_sub, markdown)
    )


def _docx_to_markdown(file_path: Path) -> str:
    """Return Markdown for a ``.docx`` with ``<!-- page: N -->`` markers."""
    try:
        from markitdown import MarkItDown, StreamInfo
    except ImportError as exc:
        raise RuntimeError(
            "Missing dependency. Install markitdown to enable office-document conversion: "
            "pip install 'markitdown[docx,xlsx,pptx]'"
        ) from exc

    docx_bytes, inserted = _insert_docx_page_tokens(file_path.read_bytes())
    converter = MarkItDown()
    result = converter.convert_stream(
        io.BytesIO(docx_bytes), stream_info=StreamInfo(extension=".docx")
    )
    markdown = _replace_page_tokens((result.text_content or "").strip())
    body = f"<!-- page: 1 -->\n\n{markdown}" if markdown else "<!-- page: 1 -->"

    found = body.count("<!-- page:")
    if found != inserted + 1:
        LOGGER.warning(
            "Page markers lost converting %s: expected %d, found %d",
            file_path,
            inserted + 1,
            found,
        )
    return body


def _extract_markdown(file_path: Path, *, docx_page_markers: bool = True) -> str:
    """Return Markdown text for *file_path*."""
    suffix = file_path.suffix.lower()
    if suffix == ".xlsx":
        return _xlsx_to_markdown(file_path)
    if suffix == ".docx" and docx_page_markers:
        return _docx_to_markdown(file_path)

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


def _docx_targeted_vision_markdown(
    file_path: Path,
    image_dir: Path,
    clean_stem: str,
    *,
    docx_page_markers: bool,
    api_key: str | None,
    vision_model: str,
    concurrency: int,
    max_retries: int,
) -> str:
    """Extract DOCX Markdown with inline contextual vision for substantive images."""
    from lxml import etree

    from .hybrid_vision import (
        VisionImageTask,
        format_vision_block_standalone,
        format_vision_for_table_cell,
        image_bytes_have_visual_detail,
        run_hybrid_vision_tasks,
    )

    try:
        from markitdown import MarkItDown, StreamInfo
    except ImportError as exc:
        raise RuntimeError(
            "Missing dependency. Install markitdown to enable office-document conversion: "
            "pip install 'markitdown[docx,xlsx,pptx]'"
        ) from exc

    r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    a_blip = "{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
    w_drawing = f"{{{_W_NS}}}drawing"
    w_tbl_cell = f"{{{_W_NS}}}tc"
    w_pstyle = f"{{{_W_NS}}}pStyle"
    wp_extent = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent"
    emu_per_pt = 12700.0

    raw_bytes = file_path.read_bytes()
    if docx_page_markers:
        working_bytes, inserted_pages = _insert_docx_page_tokens(raw_bytes)
    else:
        working_bytes, inserted_pages = raw_bytes, 0

    tasks: list[VisionImageTask] = []
    rel_refs_by_id: dict[str, str] = {}

    with zipfile.ZipFile(io.BytesIO(working_bytes)) as source:
        try:
            doc_xml = source.read("word/document.xml")
        except KeyError:
            return _extract_markdown(file_path, docx_page_markers=docx_page_markers)

        rels_map: dict[str, str] = {}
        try:
            rels_root = etree.fromstring(source.read("word/_rels/document.xml.rels"))
            for rel in rels_root:
                rid = rel.get("Id")
                target = rel.get("Target")
                if rid and target:
                    norm = target.lstrip("/")
                    if not norm.startswith("word/"):
                        norm = f"word/{norm}"
                    rels_map[rid] = norm
        except KeyError:
            pass

        root = etree.fromstring(doc_xml)
        body = root.find(_W_BODY)
        if body is None:
            return _extract_markdown(file_path, docx_page_markers=docx_page_markers)

        # Identify repeating media hashes across approximate pages
        media_cache: dict[str, bytes] = {}
        page_num = 1
        hash_pages: Counter[str] = Counter()
        seen_on_page: set[str] = set()
        for elem in body.iter():
            if not isinstance(elem.tag, str) or _has_skipped_ancestor(elem):
                continue
            if elem.tag == _W_T and elem.text and _PAGE_TOKEN_NUMBER.search(elem.text):
                for num_s in _PAGE_TOKEN_NUMBER.findall(elem.text):
                    page_num = int(num_s)
                    seen_on_page.clear()
            elif elem.tag == a_blip:
                target = rels_map.get(elem.get(f"{{{r_ns}}}embed") or "")
                if not target:
                    continue
                if target not in media_cache:
                    try:
                        media_cache[target] = source.read(target)
                    except KeyError:
                        media_cache[target] = b""
                if media_cache[target]:
                    h = hashlib.sha1(media_cache[target]).hexdigest()
                    if h not in seen_on_page:
                        seen_on_page.add(h)
                        hash_pages[h] += 1

        total_pages = max(1, page_num)
        rep_hashes = {
            h
            for h, cnt in hash_pages.items()
            if total_pages >= 3 and cnt >= max(3, int(total_pages * 0.40))
        }

        current_heading = ""
        cur_page = 1
        img_idx = 0
        drawings_to_replace: list[tuple[Any, Any, str]] = []

        for elem in body.iter():
            if not isinstance(elem.tag, str) or _has_skipped_ancestor(elem):
                continue
            if elem.tag == _W_P:
                p_text = "".join(t.text or "" for t in elem.iter(_W_T))
                for num_s in _PAGE_TOKEN_NUMBER.findall(p_text):
                    cur_page = int(num_s)
                clean_p = _PAGE_TOKEN_INLINE.sub(" ", p_text).strip()
                ppr = elem.find(_W_PPR)
                pstyle = ppr.find(w_pstyle) if ppr is not None else None
                sval = (pstyle.get(_W_VAL) or "").lower() if pstyle is not None else ""
                if clean_p and ("heading" in sval or "title" in sval):
                    current_heading = clean_p
            elif elem.tag == a_blip:
                target = rels_map.get(elem.get(f"{{{r_ns}}}embed") or "")
                blob = media_cache.get(target or "", b"")
                if not blob or hashlib.sha1(blob).hexdigest() in rep_hashes:
                    continue
                w_pt, h_pt = 0.0, 0.0
                drawing_anc = None
                para_anc = None
                cell_anc = None
                for anc in elem.iterancestors():
                    if drawing_anc is None and anc.tag == w_drawing:
                        drawing_anc = anc
                    if para_anc is None and anc.tag == _W_P:
                        para_anc = anc
                    if cell_anc is None and anc.tag == w_tbl_cell:
                        cell_anc = anc
                    ext_el = anc.find(wp_extent)
                    if w_pt == 0.0 and ext_el is not None:
                        try:
                            w_pt = int(ext_el.get("cx") or 0) / emu_per_pt
                            h_pt = int(ext_el.get("cy") or 0) / emu_per_pt
                        except ValueError:
                            pass
                has_detail, px_w, px_h, mime = image_bytes_have_visual_detail(blob)
                if not has_detail:
                    continue
                if not (
                    (w_pt >= 70.0 and h_pt >= 40.0)
                    or (w_pt == 0.0 and px_w >= 160 and px_h >= 80)
                ):
                    continue
                if drawing_anc is None or para_anc is None:
                    continue

                img_idx += 1
                tid = f"PDFTOMDIMG{img_idx}END"
                ext_out = "jpg" if mime == "image/jpeg" else "png"
                image_dir.mkdir(parents=True, exist_ok=True)
                fname = f"{clean_stem}-img{img_idx:02d}.{ext_out}"
                out_img_path = image_dir / fname
                if ext_out == "png" and not target.lower().endswith(".png"):
                    from PIL import Image

                    with Image.open(io.BytesIO(blob)) as pil_im:
                        pil_im.convert("RGB").save(out_img_path, format="PNG")
                    saved_bytes = out_img_path.read_bytes()
                    mime = "image/png"
                else:
                    out_img_path.write_bytes(blob)
                    saved_bytes = blob

                rel_refs_by_id[tid] = f"{image_dir.name}/{fname}"
                ctx_container = cell_anc if cell_anc is not None else para_anc
                surr_raw = "".join(t.text or "" for t in ctx_container.iter(_W_T))
                surr_clean = _PAGE_TOKEN_INLINE.sub(" ", surr_raw).strip()

                tasks.append(
                    VisionImageTask(
                        task_id=tid,
                        image_bytes=saved_bytes,
                        mime_type=mime,
                        document_name=file_path.name,
                        unit_number=cur_page if docx_page_markers else None,
                        section_heading=current_heading,
                        surrounding_text=surr_clean,
                    )
                )
                drawings_to_replace.append((para_anc, drawing_anc, tid))

        if not tasks:
            return _extract_markdown(file_path, docx_page_markers=docx_page_markers)

        for para_anc, drawing_anc, tid in drawings_to_replace:
            run_parent = drawing_anc.getparent()
            if run_parent is not None and run_parent.getparent() is para_anc:
                pos = list(para_anc).index(run_parent)
                run_parent.remove(drawing_anc)
                para_anc.insert(pos + 1, _make_page_token_run(f" {tid} "))
            else:
                para_anc.append(_make_page_token_run(f" {tid} "))

        mod_xml = etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as target_zf:
            for info in source.infolist():
                data = mod_xml if info.filename == "word/document.xml" else source.read(info.filename)
                target_zf.writestr(info, data)
        instrumented_bytes = buf.getvalue()

    img_md_map, _ = run_hybrid_vision_tasks(
        image_tasks=tasks,
        page_tasks=[],
        api_key=api_key,
        model=vision_model,
        concurrency=concurrency,
        max_retries=max_retries,
    )

    converter = MarkItDown()
    res = converter.convert_stream(
        io.BytesIO(instrumented_bytes), stream_info=StreamInfo(extension=".docx")
    )
    md_text = (res.text_content or "").strip()
    if docx_page_markers:
        md_text = _replace_page_tokens(md_text)
        md_text = f"<!-- page: 1 -->\n\n{md_text}" if md_text else "<!-- page: 1 -->"
        if md_text.count("<!-- page:") != inserted_pages + 1:
            LOGGER.warning(
                "Page markers lost converting %s", file_path
            )

    # Splice each image reference + vision extraction safely into prose or GFM table rows
    out_lines: list[str] = []
    for line in md_text.splitlines():
        if "PDFTOMDIMG" not in line:
            out_lines.append(line)
            continue
        is_table_row = line.lstrip().startswith("|")
        updated = line
        for tid, rel_ref in rel_refs_by_id.items():
            if tid not in updated:
                continue
            v_md = img_md_map.get(tid, "")
            if is_table_row:
                cell_parts = [
                    p.replace("|", "\\|")
                    for p in format_vision_for_table_cell(rel_ref, v_md)
                ]
                updated = updated.replace(tid, "<br>".join(cell_parts))
            else:
                block = format_vision_block_standalone(rel_ref, v_md)
                updated = updated.replace(tid, f"\n\n{block}\n\n")
        out_lines.append(updated)

    joined = "\n".join(out_lines)
    joined = re.sub(r"\n{3,}", "\n\n", joined).strip()
    return joined


def _pptx_targeted_vision_markdown(
    file_path: Path,
    image_dir: Path,
    clean_stem: str,
    *,
    api_key: str | None,
    vision_model: str,
    concurrency: int,
    max_retries: int,
) -> str:
    """Extract PPTX slides into one Markdown file with contextual vision on pictures."""
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    from .hybrid_vision import (
        VisionImageTask,
        _iter_pptx_shapes,
        format_vision_block_standalone,
        image_bytes_have_visual_detail,
        run_hybrid_vision_tasks,
    )

    emu_per_pt = 12700.0
    prs = Presentation(str(file_path))
    total_slides = len(prs.slides)
    if total_slides == 0:
        return ""

    sha_slides: Counter[str] = Counter()
    for slide in prs.slides:
        seen_sha: set[str] = set()
        for shape in _iter_pptx_shapes(slide.shapes):
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    seen_sha.add(shape.image.sha1)
                except Exception:
                    continue
        for s in seen_sha:
            sha_slides[s] += 1

    rep_shas = {
        s
        for s, cnt in sha_slides.items()
        if total_slides >= 3 and cnt >= max(3, int(total_slides * 0.40))
    }

    tasks: list[VisionImageTask] = []
    slide_entries: list[tuple[int, list[tuple[str, str]]]] = []

    for s_idx, slide in enumerate(prs.slides, start=1):
        items: list[tuple[str, str]] = []
        slide_title = ""
        if getattr(slide.shapes, "title", None) is not None and slide.shapes.title.text:
            slide_title = slide.shapes.title.text.strip()

        slide_text_parts: list[str] = []
        pic_counter = 0

        for shape in _iter_pptx_shapes(slide.shapes):
            if getattr(shape, "has_text_frame", False):
                is_title_shape = (
                    getattr(slide.shapes, "title", None) is not None
                    and shape == slide.shapes.title
                )
                for para in shape.text_frame.paragraphs:
                    txt = para.text.strip()
                    if not txt:
                        continue
                    slide_text_parts.append(txt)
                    if is_title_shape:
                        items.append(("text", f"## {txt}"))
                        is_title_shape = False
                    else:
                        items.append(("text", txt))
            elif getattr(shape, "has_table", False):
                rows = [
                    [cell.text.strip().replace("\n", "<br>").replace("|", "\\|") for cell in r.cells]
                    for r in shape.table.rows
                ]
                if rows:
                    hdr = "| " + " | ".join(rows[0]) + " |"
                    sep = "| " + " | ".join("---" for _ in rows[0]) + " |"
                    body = ["| " + " | ".join(r) + " |" for r in rows[1:]]
                    tbl_md = "\n".join([hdr, sep, *body])
                    items.append(("text", tbl_md))
                    slide_text_parts.append(tbl_md)
            elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                w_pt = float(shape.width or 0) / emu_per_pt
                h_pt = float(shape.height or 0) / emu_per_pt
                if w_pt < 70.0 or h_pt < 40.0:
                    continue
                try:
                    sha = shape.image.sha1
                    blob = shape.image.blob
                except Exception:
                    continue
                if sha in rep_shas:
                    continue
                has_detail, _, _, mime = image_bytes_have_visual_detail(blob)
                if not has_detail:
                    continue

                pic_counter += 1
                ext_out = "jpg" if mime == "image/jpeg" else "png"
                image_dir.mkdir(parents=True, exist_ok=True)
                fname = f"{clean_stem}-s{s_idx:03d}-img{pic_counter:02d}.{ext_out}"
                out_file = image_dir / fname
                if ext_out == "png" and mime != "image/png":
                    from PIL import Image

                    with Image.open(io.BytesIO(blob)) as pil_im:
                        pil_im.convert("RGB").save(out_file, format="PNG")
                    saved_bytes = out_file.read_bytes()
                    mime = "image/png"
                else:
                    out_file.write_bytes(blob)
                    saved_bytes = blob

                tid = f"s{s_idx}-p{pic_counter}"
                rel_ref = f"{image_dir.name}/{fname}"
                items.append(("img", f"{tid}|{rel_ref}"))
                tasks.append(
                    VisionImageTask(
                        task_id=tid,
                        image_bytes=saved_bytes,
                        mime_type=mime,
                        document_name=file_path.name,
                        unit_number=s_idx,
                        section_heading=slide_title,
                        surrounding_text="\n".join(slide_text_parts),
                    )
                )

        # Backfill full slide text context onto tasks created earlier on the same slide
        full_slide_ctx = "\n".join(slide_text_parts)
        for idx_t in range(len(tasks) - pic_counter, len(tasks)):
            t = tasks[idx_t]
            tasks[idx_t] = VisionImageTask(
                task_id=t.task_id,
                image_bytes=t.image_bytes,
                mime_type=t.mime_type,
                document_name=t.document_name,
                unit_number=t.unit_number,
                section_heading=t.section_heading,
                surrounding_text=full_slide_ctx,
            )

        slide_entries.append((s_idx, items))

    img_md_map, _ = run_hybrid_vision_tasks(
        image_tasks=tasks,
        page_tasks=[],
        api_key=api_key,
        model=vision_model,
        concurrency=concurrency,
        max_retries=max_retries,
    )

    slide_blocks: list[str] = []
    for s_idx, items in slide_entries:
        chunks: list[str] = []
        for kind, val in items:
            if kind == "text":
                chunks.append(val)
            else:
                tid, rel_ref = val.split("|", 1)
                chunks.append(
                    format_vision_block_standalone(rel_ref, img_md_map.get(tid, ""))
                )
        body = "\n\n".join(c for c in chunks if c.strip())
        marker = f"<!-- page: {s_idx} -->"
        slide_blocks.append(f"{marker}\n\n{body}" if body else marker)

    return "\n\n".join(slide_blocks)


def _extract_markdown_with_vision(
    file_path: Path,
    output_path: Path,
    *,
    docx_page_markers: bool,
    vision_mode: str,
    api_key: str | None,
    vision_model: str | None,
    low_threshold: float,
    high_threshold: float,
    dpi: int,
    concurrency: int,
    max_retries: int,
) -> str:
    from .hybrid_vision import (
        DEFAULT_MODEL,
        libreoffice_to_pdf,
        resolve_effective_strategy,
        screen_document,
        validate_vision_options,
    )
    from .pdf_converter import _extract_markdown as _extract_pdf_markdown

    validate_vision_options(
        vision_mode, low_threshold, high_threshold, dpi, concurrency, max_retries
    )
    effective_model = vision_model or DEFAULT_MODEL
    suffix = file_path.suffix.lower()
    if suffix not in (".docx", ".pptx"):
        return _extract_markdown(file_path, docx_page_markers=docx_page_markers)

    report = screen_document(
        file_path, low_threshold=low_threshold, high_threshold=high_threshold
    )
    strategy = resolve_effective_strategy(report, vision_mode)  # type: ignore[arg-type]
    LOGGER.info(
        "Screened %s: %d/%d unit(s) visual (%.0f%%) -> strategy=%s",
        file_path.name,
        report.candidate_units,
        report.total_units,
        report.vision_ratio * 100.0,
        strategy,
    )
    if strategy == "none":
        return _extract_markdown(file_path, docx_page_markers=docx_page_markers)

    clean_stem = re.sub(r"[\s\(\)\[\]]+", "_", file_path.stem).strip("_") or "document"
    image_dir = output_path.parent / f"{clean_stem}_images"

    if strategy in ("hybrid_pages", "full_document"):
        forced_pdf_mode = "pages" if strategy == "hybrid_pages" else "full"
        try:
            with libreoffice_to_pdf(file_path) as tmp_pdf:
                pdf_md = _extract_pdf_markdown(
                    tmp_pdf,
                    image_dir,
                    clean_stem,
                    vision_mode=forced_pdf_mode,
                    api_key=api_key,
                    vision_model=effective_model,
                    low_threshold=low_threshold,
                    high_threshold=high_threshold,
                    dpi=dpi,
                    concurrency=concurrency,
                    max_retries=max_retries,
                )
            if suffix == ".docx" and not docx_page_markers:
                pdf_md = re.sub(r"<!-- page: \d+ -->\n*", "", pdf_md).strip()
            return pdf_md
        except FileNotFoundError:
            LOGGER.info(
                "LibreOffice not found; falling back to embedded-image vision for %s.",
                file_path.name,
            )

    if suffix == ".docx":
        return _docx_targeted_vision_markdown(
            file_path,
            image_dir,
            clean_stem,
            docx_page_markers=docx_page_markers,
            api_key=api_key,
            vision_model=effective_model,
            concurrency=concurrency,
            max_retries=max_retries,
        )

    return _pptx_targeted_vision_markdown(
        file_path,
        image_dir,
        clean_stem,
        api_key=api_key,
        vision_model=effective_model,
        concurrency=concurrency,
        max_retries=max_retries,
    )


def convert_office(
    file_path: Path,
    output_path: Path,
    *,
    docx_page_markers: bool = True,
    vision_mode: str = "off",
    api_key: str | None = None,
    vision_model: str | None = None,
    low_threshold: float = 0.30,
    high_threshold: float = 0.70,
    dpi: int = 200,
    concurrency: int = 4,
    max_retries: int = 5,
) -> None:
    """Convert a single office document to Markdown and write to *output_path*.

    When *docx_page_markers* is true, ``.docx`` output gets approximate
    ``<!-- page: N -->`` markers.  When *vision_mode* is not ``"off"``,
    ``.docx`` and ``.pptx`` documents are screened and augmented with Gemini
    vision in a single unified Markdown file.
    """
    if vision_mode == "off":
        markdown_body = _extract_markdown(file_path, docx_page_markers=docx_page_markers)
    else:
        markdown_body = _extract_markdown_with_vision(
            file_path,
            output_path,
            docx_page_markers=docx_page_markers,
            vision_mode=vision_mode,
            api_key=api_key,
            vision_model=vision_model,
            low_threshold=low_threshold,
            high_threshold=high_threshold,
            dpi=dpi,
            concurrency=concurrency,
            max_retries=max_retries,
        )
    wikilink = f"[[{file_path.name}]]"
    content = f"{wikilink}\n\n{markdown_body}\n" if markdown_body else f"{wikilink}\n"

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(content, encoding="utf-8")
    LOGGER.info("Wrote %s", output_path)
