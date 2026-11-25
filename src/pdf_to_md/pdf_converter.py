"""PDF-to-Markdown converter module.

Converts a single PDF file to Markdown using document-relative PyMuPDF geometry
(tables, headings, lists, inline formatting, hyperlinks, raster images, and
vector diagrams), falling back to plain PyMuPDF text extraction on unexpected
errors.

Extracted images are saved lazily into a sibling ``<stem>_images/`` folder and
referenced via portable relative paths.  Each output file starts with a
Wikilink ``[[Original.pdf]]``, and every page starts with a ``<!-- page: N -->``
marker (including empty pages) so downstream consumers can trace content back
to its source page.
"""

from __future__ import annotations

import logging
import re
from collections import Counter
from collections.abc import Iterable
from dataclasses import dataclass, field
from pathlib import Path
from statistics import median
from typing import Literal

LOGGER = logging.getLogger(__name__)

SUPPORTED_EXTENSIONS: set[str] = {".pdf"}

_UNICODE_BULLETS: frozenset[str] = frozenset(
    {"•", "◦", "▪", "▫", "■", "□", "●", "○", "‣", "⁃", "∙", "⋅"}
)
_ORDERED_ITEM_RE = re.compile(r"^(?:\d{1,3}[\.\)]|\([a-zA-Z0-9]{1,3}\)|[a-zA-Z]\))\s+")
_HYPHEN_BULLET_RE = re.compile(r"^[-–—]\s+\S")
_MONO_KEYWORDS = ("mono", "courier", "consolas", "menlo", "firacode", "jetbrains", "dejavusansmono")
_BOLD_KEYWORDS = ("bold", "semibold", "demibold", "extrabold", "black", "heavy", "hebo", "tibo", "cobo")
_ITALIC_KEYWORDS = ("italic", "oblique", "it", "tiit", "coit", "heit")


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------


@dataclass
class _SpanItem:
    x0: float
    y0: float
    x1: float
    y1: float
    text: str
    size: float
    font: str
    flags: int
    color: int
    is_bold: bool
    is_italic: bool
    is_mono: bool
    link_uri: str | None = None


@dataclass
class _VisualLine:
    x0: float
    y0: float
    x1: float
    y1: float
    spans: list[_SpanItem]
    has_vector_bullet: bool = False

    @property
    def raw_text(self) -> str:
        return _render_line_spans(self.spans, plain=True).strip()

    @property
    def max_size(self) -> float:
        return max((s.size for s in self.spans), default=11.0)

    @property
    def dominant_font(self) -> str:
        counts: Counter[str] = Counter()
        for s in self.spans:
            stripped = s.text.strip()
            if stripped:
                counts[s.font] += len(stripped)
        return counts.most_common(1)[0][0] if counts else ""

    @property
    def is_all_bold(self) -> bool:
        non_ws = [s for s in self.spans if s.text.strip()]
        return bool(non_ws) and all(s.is_bold for s in non_ws)

    @property
    def is_all_mono(self) -> bool:
        non_ws = [s for s in self.spans if s.text.strip()]
        return bool(non_ws) and all(s.is_mono for s in non_ws)


@dataclass
class _DetectedTable:
    bbox: tuple[float, float, float, float]
    col_edges: list[float]
    rows: list[list[tuple[float, float, float, float] | None]]


@dataclass
class _ImageAsset:
    bbox: tuple[float, float, float, float]
    rel_md_path: str
    png_bytes: bytes | None = None
    is_vision_candidate: bool = False
    vision_markdown: str = ""
    source_kind: Literal["raster", "vector"] = "raster"


@dataclass
class _ActiveTableState:
    col_edges: list[float]
    header_mode: Literal["top_header", "stub_col", "plain_no_header"]
    header_cells: list[str] = field(default_factory=list)
    col_base_x0: dict[int, float] = field(default_factory=dict)


@dataclass
class _DocumentStyleProfile:
    body_size: float
    body_fonts: set[str]
    body_is_bold: bool
    heading_levels: dict[tuple[float, bool], int]
    repeating_margin_texts: set[str]


# ---------------------------------------------------------------------------
# Font & span helpers
# ---------------------------------------------------------------------------


def _font_is_bold(font_name: str, flags: int) -> bool:
    if flags & 16:
        return True
    low = font_name.lower()
    if low in ("hebo", "tibo", "cobo", "hebi", "tibi", "cobi"):
        return True
    return any(kw in low for kw in _BOLD_KEYWORDS)


def _font_is_italic(font_name: str, flags: int) -> bool:
    if flags & 2:
        return True
    low = font_name.lower()
    if low in ("heit", "tiit", "coit", "hebi", "tibi", "cobi"):
        return True
    return any(kw in low for kw in ("italic", "oblique"))


def _font_is_mono(font_name: str, flags: int) -> bool:
    if flags & 8:
        return True
    low = font_name.lower().replace(" ", "").replace("-", "")
    if low in ("cour", "cobo", "coit", "cobi"):
        return True
    return any(kw in low for kw in _MONO_KEYWORDS)


def _clean_text(text: str) -> str:
    return text.replace("\x00", "\ufffd").replace("\xa0", " ")


def _extract_page_spans(page) -> list[_SpanItem]:
    """Extract styled text spans with resolved URI links for *page*."""
    import pymupdf

    page_rect = page.rect
    links = [
        (pymupdf.Rect(lnk["from"]) & page_rect, lnk["uri"])
        for lnk in page.get_links()
        if lnk.get("kind") == pymupdf.LINK_URI and lnk.get("uri")
    ]

    spans: list[_SpanItem] = []
    text_dict = page.get_text("dict", flags=pymupdf.TEXTFLAGS_TEXT)
    for block in text_dict.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            # Skip vertical / upside-down rotated watermark text unless normal horizontal
            dir_vec = line.get("dir", (1.0, 0.0))
            if abs(dir_vec[0]) < 0.5:
                continue
            for sp in line.get("spans", []):
                raw = _clean_text(sp.get("text") or "")
                if not raw:
                    continue
                sr = pymupdf.Rect(sp["bbox"]) & page_rect
                if sr.is_empty:
                    continue
                uri = _match_link_uri(sr, links)
                font = sp.get("font") or ""
                flags = int(sp.get("flags") or 0)
                spans.append(
                    _SpanItem(
                        x0=sr.x0,
                        y0=sr.y0,
                        x1=sr.x1,
                        y1=sr.y1,
                        text=raw,
                        size=float(sp.get("size") or 11.0),
                        font=font,
                        flags=flags,
                        color=int(sp.get("color") or 0),
                        is_bold=_font_is_bold(font, flags),
                        is_italic=_font_is_italic(font, flags),
                        is_mono=_font_is_mono(font, flags),
                        link_uri=uri,
                    )
                )
    return spans


def _match_link_uri(span_rect, links: list[tuple[object, str]]) -> str | None:
    if not links or span_rect.width <= 0 or span_rect.height <= 0:
        return None
    best_uri: str | None = None
    best_ratio = 0.0
    for hot_rect, uri in links:
        inter = span_rect & hot_rect
        if inter.is_empty:
            continue
        v_ratio = inter.height / max(min(span_rect.height, hot_rect.height), 1.0)
        h_ratio = inter.width / max(span_rect.width, 1.0)
        if v_ratio >= 0.45 and h_ratio >= 0.5 and h_ratio > best_ratio:
            best_ratio = h_ratio
            best_uri = uri
    return best_uri


# ---------------------------------------------------------------------------
# Drawings: fills, borders, vector bullets, and vector diagrams
# ---------------------------------------------------------------------------


@dataclass
class _PageDrawings:
    page_bg: tuple[float, float, float]
    fill_rects: list[tuple[tuple[float, float, float, float], tuple[float, float, float]]]
    v_segments: list[tuple[float, float, float]]  # (x, y0, y1)
    h_segments: list[tuple[float, float, float]]  # (y, x0, x1)
    bullet_rects: list[tuple[float, float, float, float]]
    raw_drawings: list[dict]


def _quantize_rgb(rgb: Iterable[float] | None) -> tuple[float, float, float] | None:
    if rgb is None:
        return None
    vals = list(rgb)
    if len(vals) < 3:
        return None
    return (round(float(vals[0]), 2), round(float(vals[1]), 2), round(float(vals[2]), 2))


def _analyze_page_drawings(page) -> _PageDrawings:
    import pymupdf

    pr = page.rect
    raw_drawings = page.get_drawings()
    page_bg = (1.0, 1.0, 1.0)
    fill_rects: list[tuple[tuple[float, float, float, float], tuple[float, float, float]]] = []
    v_segments: list[tuple[float, float, float]] = []
    h_segments: list[tuple[float, float, float]] = []
    bullet_candidates: list[tuple[float, float, float, float]] = []

    for d in raw_drawings:
        raw_r = pymupdf.Rect(d["rect"])
        if (
            raw_r.x1 < pr.x0
            or raw_r.x0 > pr.x1
            or raw_r.y1 < pr.y0
            or raw_r.y0 > pr.y1
        ):
            continue
        x0 = max(pr.x0, raw_r.x0)
        y0 = max(pr.y0, raw_r.y0)
        x1 = min(pr.x1, raw_r.x1)
        y1 = min(pr.y1, raw_r.y1)
        w, h = max(0.0, x1 - x0), max(0.0, y1 - y0)
        fill = _quantize_rgb(d.get("fill"))
        stroke = _quantize_rgb(d.get("color"))

        if fill is not None and w >= 0.85 * pr.width and h >= 0.85 * pr.height:
            page_bg = fill
            continue

        # Small standalone geometric markers (circles, squares, hollow bullets)
        if 1.2 <= w <= 7.5 and 1.2 <= h <= 7.5 and 0.4 <= (w / max(h, 0.1)) <= 2.5:
            if fill != page_bg or stroke is not None:
                bullet_candidates.append((x0, y0, x1, y1))

        # Filled rectangles used as cell backgrounds or thin CSS border strips
        if fill is not None and w > 0.0 and h > 0.0:
            if w <= 3.5 and h >= 10.0 and fill != page_bg:
                v_segments.append(((x0 + x1) / 2.0, y0, y1))
            elif h <= 3.5 and w >= 10.0 and fill != page_bg:
                h_segments.append(((y0 + y1) / 2.0, x0, x1))
            elif w >= 12.0 and h >= 10.0:
                fill_rects.append(((x0, y0, x1, y1), fill))

        # Stroked lines and rectangles
        if stroke is not None:
            for item in d.get("items", []):
                kind = item[0]
                if kind == "l":
                    p1, p2 = item[1], item[2]
                    if abs(p1.x - p2.x) <= 2.0:
                        y0, y1 = max(pr.y0, min(p1.y, p2.y)), min(pr.y1, max(p1.y, p2.y))
                        if y1 - y0 >= 10.0:
                            v_segments.append(((p1.x + p2.x) / 2.0, y0, y1))
                    elif abs(p1.y - p2.y) <= 2.0:
                        x0, x1 = max(pr.x0, min(p1.x, p2.x)), min(pr.x1, max(p1.x, p2.x))
                        if x1 - x0 >= 10.0:
                            h_segments.append(((p1.y + p2.y) / 2.0, x0, x1))
                elif kind == "re":
                    rr = pymupdf.Rect(item[1]) & pr
                    if not rr.is_empty:
                        if rr.height >= 10.0:
                            v_segments.append((rr.x0, rr.y0, rr.y1))
                            v_segments.append((rr.x1, rr.y0, rr.y1))
                        if rr.width >= 10.0:
                            h_segments.append((rr.y0, rr.x0, rr.x1))
                            h_segments.append((rr.y1, rr.x0, rr.x1))

    return _PageDrawings(
        page_bg=page_bg,
        fill_rects=fill_rects,
        v_segments=v_segments,
        h_segments=h_segments,
        bullet_rects=bullet_candidates,
        raw_drawings=raw_drawings,
    )


def _cell_bg_color(
    cell_bbox: tuple[float, float, float, float],
    drawings: _PageDrawings,
) -> tuple[float, float, float] | None:
    cx0, cy0, cx1, cy1 = cell_bbox
    cw, ch = cx1 - cx0, cy1 - cy0
    if cw <= 0 or ch <= 0:
        return None
    mx, my = (cx0 + cx1) / 2.0, (cy0 + cy1) / 2.0
    matched: tuple[float, float, float] | None = None
    for (rx0, ry0, rx1, ry1), fill in drawings.fill_rects:
        if rx0 - 2.0 <= mx <= rx1 + 2.0 and ry0 - 2.0 <= my <= ry1 + 2.0:
            inter_w = min(cx1, rx1) - max(cx0, rx0)
            inter_h = min(cy1, ry1) - max(cy0, ry0)
            if inter_w >= 0.65 * cw and inter_h >= 0.65 * ch:
                if max(abs(fill[i] - drawings.page_bg[i]) for i in range(3)) > 0.02:
                    matched = fill
                else:
                    matched = None
    return matched


# ---------------------------------------------------------------------------
# Table detection & continuation synthesis
# ---------------------------------------------------------------------------


def _cluster_coords(values: Iterable[float], tol: float = 3.0) -> list[float]:
    sorted_vals = sorted(values)
    if not sorted_vals:
        return []
    groups: list[list[float]] = [[sorted_vals[0]]]
    for v in sorted_vals[1:]:
        if abs(v - groups[-1][-1]) <= tol:
            groups[-1].append(v)
        else:
            groups.append([v])
    return [sum(g) / len(g) for g in groups]


def _edges_match(edges_a: list[float], edges_b: list[float], tol: float = 5.0) -> bool:
    if len(edges_a) != len(edges_b) or len(edges_a) < 3:
        return False
    return all(abs(a - b) <= tol for a, b in zip(edges_a, edges_b))


def _merge_intervals(intervals: list[tuple[float, float]], gap_tol: float = 3.0) -> list[tuple[float, float]]:
    if not intervals:
        return []
    ordered = sorted(intervals)
    merged = [ordered[0]]
    for start, end in ordered[1:]:
        prev_s, prev_e = merged[-1]
        if start <= prev_e + gap_tol:
            merged[-1] = (prev_s, max(prev_e, end))
        else:
            merged.append((start, end))
    return merged


def _detect_page_tables(
    page,
    drawings: _PageDrawings,
    spans: list[_SpanItem],
    prev_bottom_edges: list[float] | None,
) -> list[_DetectedTable]:
    import pymupdf

    pr = page.rect
    tables: list[_DetectedTable] = []
    finder = page.find_tables(strategy="lines_strict")

    raw_candidates: list[tuple[object, list[float]]] = []
    for t in finder.tables:
        tb = pymupdf.Rect(t.bbox) & pr
        if tb.is_empty or t.col_count < 2 or t.row_count < 1:
            continue
        xs: list[float] = []
        for row in t.rows:
            for c in row.cells:
                if c is not None:
                    cr = pymupdf.Rect(c) & pr
                    if not cr.is_empty:
                        xs.extend([cr.x0, cr.x1])
        col_edges = _cluster_coords(xs, tol=3.5)
        if len(col_edges) - 1 < 2:
            continue
        # Reject full-page decorative border frame unless matching previous table columns
        if (
            t.row_count == 1
            and tb.width > 0.94 * pr.width
            and tb.height > 0.94 * pr.height
            and not (prev_bottom_edges and _edges_match(col_edges, prev_bottom_edges))
        ):
            continue
        raw_candidates.append((t, col_edges))

    for idx, (t, col_edges) in enumerate(raw_candidates):
        tb = pymupdf.Rect(t.bbox) & pr
        if t.row_count == 1:
            matches_neighbor = bool(
                prev_bottom_edges and _edges_match(col_edges, prev_bottom_edges)
            ) or any(
                j != idx and _edges_match(col_edges, other_edges)
                for j, (_, other_edges) in enumerate(raw_candidates)
            )
            has_content = any(
                tb.x0 <= (s.x0 + s.x1) / 2.0 <= tb.x1
                and tb.y0 <= (s.y0 + s.y1) / 2.0 <= tb.y1
                for s in spans
            )
            if not matches_neighbor and (tb.width < 0.25 * pr.width or not has_content):
                continue

        rows_cells: list[list[tuple[float, float, float, float] | None]] = []
        for row in t.rows:
            row_list: list[tuple[float, float, float, float] | None] = []
            for c in row.cells:
                if c is None:
                    row_list.append(None)
                else:
                    cr = pymupdf.Rect(c) & pr
                    row_list.append((cr.x0, cr.y0, cr.x1, cr.y1) if not cr.is_empty else None)
            if any(cell is not None for cell in row_list):
                rows_cells.append(row_list)

        if rows_cells:
            tables.append(
                _DetectedTable(
                    bbox=(tb.x0, tb.y0, tb.x1, tb.y1),
                    col_edges=col_edges,
                    rows=rows_cells,
                )
            )

    tables.sort(key=lambda item: (item.bbox[1], item.bbox[0]))

    # Extend last table downward if its final row is clipped open at the page bottom
    if tables:
        extended = _extend_open_bottom_row(pr, drawings, tables[-1])
        if extended is not None:
            tables[-1] = extended

    # Synthesize open-top / open-bottom split-row continuation at top of page
    if prev_bottom_edges is not None:
        y_cap = tables[0].bbox[1] if tables else pr.y1
        if y_cap > pr.y0 + 18.0:
            synth = _synthesize_open_top_continuation(
                pr, drawings, prev_bottom_edges, y_cap
            )
            if synth is not None:
                tables.insert(0, synth)

    # Merge vertically adjacent table fragments on the same page with matching columns
    merged_tables: list[_DetectedTable] = []
    for tb in tables:
        if (
            merged_tables
            and abs(tb.bbox[1] - merged_tables[-1].bbox[3]) <= 6.0
            and _edges_match(tb.col_edges, merged_tables[-1].col_edges)
        ):
            prev = merged_tables[-1]
            merged_tables[-1] = _DetectedTable(
                bbox=(
                    min(prev.bbox[0], tb.bbox[0]),
                    prev.bbox[1],
                    max(prev.bbox[2], tb.bbox[2]),
                    tb.bbox[3],
                ),
                col_edges=prev.col_edges,
                rows=[*prev.rows, *tb.rows],
            )
        else:
            merged_tables.append(tb)

    return merged_tables


def _extend_open_bottom_row(
    page_rect,
    drawings: _PageDrawings,
    table: _DetectedTable,
) -> _DetectedTable | None:
    """Append an open-bottom trailing row if all column borders continue below *table*."""
    tx0, ty0, tx1, ty1 = table.bbox
    y_ends: list[float] = []
    for edge_x in table.col_edges:
        segs = [
            (max(page_rect.y0, y0), min(page_rect.y1, y1))
            for x, y0, y1 in drawings.v_segments
            if abs(x - edge_x) <= 4.0 and y1 > ty1 + 10.0
        ]
        merged = _merge_intervals(segs, gap_tol=4.0)
        covering = [iv[1] for iv in merged if iv[0] <= ty1 + 4.0 and iv[1] > ty1 + 12.0]
        if not covering:
            return None
        y_ends.append(max(covering))

    if max(y_ends) - min(y_ends) > 8.0:
        return None
    new_y1 = sum(y_ends) / len(y_ends)
    extra_row: list[tuple[float, float, float, float] | None] = [
        (table.col_edges[i], ty1, table.col_edges[i + 1], new_y1)
        for i in range(len(table.col_edges) - 1)
    ]
    return _DetectedTable(
        bbox=(tx0, ty0, tx1, new_y1),
        col_edges=table.col_edges,
        rows=[*table.rows, extra_row],
    )


def _synthesize_open_top_continuation(
    page_rect,
    drawings: _PageDrawings,
    expected_edges: list[float],
    y_cap: float,
) -> _DetectedTable | None:
    """Detect an open-top (and possibly open-bottom) table continuation at page top."""
    top_intervals_per_edge: list[tuple[float, float]] = []
    for edge_x in expected_edges:
        segs = [
            (max(page_rect.y0, y0), min(y_cap, y1))
            for x, y0, y1 in drawings.v_segments
            if abs(x - edge_x) <= 4.0 and y0 < y_cap - 4.0 and y1 > page_rect.y0 + 8.0
        ]
        merged = _merge_intervals([iv for iv in segs if iv[1] - iv[0] >= 12.0], gap_tol=4.0)
        if not merged:
            return None
        first_y0, first_y1 = merged[0]
        if first_y0 > page_rect.y0 + 90.0:
            return None
        top_intervals_per_edge.append((first_y0, first_y1))

    y_start = min(iv[0] for iv in top_intervals_per_edge)
    y_end_min = min(iv[1] for iv in top_intervals_per_edge)
    y_end_max = max(iv[1] for iv in top_intervals_per_edge)
    if y_end_max - y_end_min > 8.0 or y_end_min - y_start < 14.0:
        return None
    y_end = (y_end_min + y_end_max) / 2.0

    x_left, x_right = expected_edges[0], expected_edges[-1]
    table_w = x_right - x_left
    internal_h = [
        y
        for y, hx0, hx1 in drawings.h_segments
        if y_start + 6.0 < y < y_end - 6.0
        and min(x_right, hx1) - max(x_left, hx0) >= 0.85 * table_w
    ]
    row_ys = [y_start, *_cluster_coords(internal_h, tol=3.0), y_end]
    rows: list[list[tuple[float, float, float, float] | None]] = []
    for r_idx in range(len(row_ys) - 1):
        ry0, ry1 = row_ys[r_idx], row_ys[r_idx + 1]
        if ry1 - ry0 < 8.0:
            continue
        rows.append(
            [
                (expected_edges[c_idx], ry0, expected_edges[c_idx + 1], ry1)
                for c_idx in range(len(expected_edges) - 1)
            ]
        )
    if not rows:
        return None
    return _DetectedTable(
        bbox=(x_left, y_start, x_right, y_end),
        col_edges=list(expected_edges),
        rows=rows,
    )


# ---------------------------------------------------------------------------
# Visual line grouping & span rendering
# ---------------------------------------------------------------------------


def _group_visual_lines(
    spans: list[_SpanItem],
    bullet_rects: list[tuple[float, float, float, float]],
) -> list[_VisualLine]:
    """Merge spans sharing the same horizontal baseline into sorted visual lines."""
    if not spans:
        return []

    ordered = sorted(spans, key=lambda s: ((s.y0 + s.y1) / 2.0, s.x0))
    lines: list[_VisualLine] = []

    for sp in ordered:
        placed = False
        sp_h = max(sp.y1 - sp.y0, 1.0)
        for vline in reversed(lines[-6:]):
            vl_h = max(vline.y1 - vline.y0, 1.0)
            overlap = min(sp.y1, vline.y1) - max(sp.y0, vline.y0)
            if overlap >= 0.48 * min(sp_h, vl_h):
                vline.spans.append(sp)
                vline.x0 = min(vline.x0, sp.x0)
                vline.y0 = min(vline.y0, sp.y0)
                vline.x1 = max(vline.x1, sp.x1)
                vline.y1 = max(vline.y1, sp.y1)
                placed = True
                break
        if not placed:
            lines.append(_VisualLine(x0=sp.x0, y0=sp.y0, x1=sp.x1, y1=sp.y1, spans=[sp]))

    for vline in lines:
        vline.spans.sort(key=lambda s: s.x0)
        vline.has_vector_bullet = _has_bullet_to_left(vline, bullet_rects)

    lines.sort(key=lambda vl: (vl.y0, vl.x0))
    return lines


def _has_bullet_to_left(
    vline: _VisualLine,
    bullet_rects: list[tuple[float, float, float, float]],
) -> bool:
    for bx0, by0, bx1, by1 in bullet_rects:
        bym = (by0 + by1) / 2.0
        if vline.y0 - 1.5 <= bym <= vline.y1 + 1.5 and 1.5 <= (vline.x0 - bx1) <= 18.0:
            return True
    return False


def _render_line_spans(
    spans: list[_SpanItem],
    *,
    plain: bool = False,
    suppress_bold: bool = False,
    suppress_mono: bool = False,
) -> str:
    """Render sorted spans on one line into inline-formatted Markdown."""
    if not spans:
        return ""

    # Coalesce adjacent spans sharing identical inline formatting/link state
    merged_runs: list[tuple[str, bool, bool, bool, str | None]] = []
    prev_end_x: float | None = None

    for sp in spans:
        txt = sp.text
        gap_space = ""
        if prev_end_x is not None:
            dx = sp.x0 - prev_end_x
            if dx > 0.22 * sp.size and not txt.startswith(" "):
                if merged_runs and not merged_runs[-1][0].endswith(" "):
                    gap_space = " "
        prev_end_x = max(prev_end_x or sp.x1, sp.x1)

        bold = False if (plain or suppress_bold) else sp.is_bold
        italic = False if plain else sp.is_italic
        mono = False if (plain or suppress_mono) else sp.is_mono
        uri = None if plain else sp.link_uri

        if merged_runs and merged_runs[-1][1:] == (bold, italic, mono, uri):
            prev_txt = merged_runs[-1][0]
            merged_runs[-1] = (prev_txt + gap_space + txt, bold, italic, mono, uri)
        else:
            if gap_space and merged_runs:
                p_txt, pb, pi, pm, pu = merged_runs[-1]
                merged_runs[-1] = (p_txt + gap_space, pb, pi, pm, pu)
            merged_runs.append((txt, bold, italic, mono, uri))

    parts: list[str] = []
    for txt, bold, italic, mono, uri in merged_runs:
        if not txt:
            continue
        lead = " " if txt[:1].isspace() and txt.strip() else ""
        trail = " " if txt[-1:].isspace() and txt.strip() else ""
        core = re.sub(r" {2,}", " ", txt.strip())
        if not core:
            parts.append(" ")
            continue
        if mono:
            core = f"`{core}`"
        else:
            if bold and italic:
                core = f"***{core}***"
            elif bold:
                core = f"**{core}**"
            elif italic:
                core = f"*{core}*"
        if uri:
            core = f"[{core}]({uri})"
        parts.append(f"{lead}{core}{trail}")

    rendered = "".join(parts)
    rendered = re.sub(r" {2,}", " ", rendered).strip()
    return rendered


# ---------------------------------------------------------------------------
# Document-relative heading analysis
# ---------------------------------------------------------------------------


def _bbox_contains_point(bbox: tuple[float, float, float, float], x: float, y: float) -> bool:
    return bbox[0] - 1.0 <= x <= bbox[2] + 1.0 and bbox[1] - 1.0 <= y <= bbox[3] + 1.0


def _analyze_document_styles(
    doc,
    per_page_spans: list[list[_SpanItem]],
    per_page_drawings: list[_PageDrawings],
    per_page_tables: list[list[_DetectedTable]],
) -> _DocumentStyleProfile:
    prose_size_chars: Counter[float] = Counter()
    all_size_chars: Counter[float] = Counter()
    prose_fonts_by_size: dict[float, Counter[str]] = {}
    prose_bold_chars = 0
    prose_total_chars = 0

    margin_text_pages: Counter[str] = Counter()
    page_count = len(per_page_spans)

    per_page_outside_lines: list[list[_VisualLine]] = []

    for p_idx in range(page_count):
        page_h = float(doc[p_idx].rect.height)
        tables = per_page_tables[p_idx]
        outside_spans = [
            sp
            for sp in per_page_spans[p_idx]
            if not any(
                _bbox_contains_point(t.bbox, (sp.x0 + sp.x1) / 2.0, (sp.y0 + sp.y1) / 2.0)
                for t in tables
            )
        ]
        vlines = _group_visual_lines(outside_spans, per_page_drawings[p_idx].bullet_rects)
        per_page_outside_lines.append(vlines)

        seen_margin_on_page: set[str] = set()
        for vl in vlines:
            raw = vl.raw_text
            if not raw:
                continue
            sz = round(vl.max_size, 2)
            all_size_chars[sz] += len(raw)

            if page_count >= 4 and len(raw) <= 60:
                if vl.y1 <= 0.07 * page_h or vl.y0 >= 0.93 * page_h:
                    norm_m = re.sub(r"\d+", "#", raw)
                    seen_margin_on_page.add(norm_m)

            if len(raw) >= 45:
                prose_size_chars[sz] += len(raw)
                prose_fonts_by_size.setdefault(sz, Counter())[vl.dominant_font] += len(raw)
                prose_total_chars += len(raw)
                if vl.is_all_bold:
                    prose_bold_chars += len(raw)

        for norm_m in seen_margin_on_page:
            margin_text_pages[norm_m] += 1

    # Also tally all non-table non-tiny text if long prose lines were absent
    if prose_size_chars:
        body_size = prose_size_chars.most_common(1)[0][0]
    elif all_size_chars:
        body_size = all_size_chars.most_common(1)[0][0]
    else:
        body_size = 11.0

    # Include fonts within +-0.15pt of body_size as normal prose fonts
    body_fonts: set[str] = set()
    for sz, font_counter in prose_fonts_by_size.items():
        if abs(sz - body_size) <= 0.15:
            body_fonts.update(font_counter.keys())
    if not body_fonts:
        for vlines in per_page_outside_lines:
            for vl in vlines:
                if abs(round(vl.max_size, 2) - body_size) <= 0.15:
                    body_fonts.add(vl.dominant_font)

    body_is_bold = bool(prose_total_chars and prose_bold_chars > 0.6 * prose_total_chars)
    repeating_margins = {
        txt for txt, cnt in margin_text_pages.items() if page_count >= 4 and cnt >= max(3, page_count // 2)
    }

    # Discover distinct heading tiers conservatively across the document
    heading_tiers: set[tuple[float, bool]] = set()
    for p_idx, vlines in enumerate(per_page_outside_lines):
        page_h = float(doc[p_idx].rect.height)
        for idx, vl in enumerate(vlines):
            tier = _candidate_heading_tier(
                vl,
                prev_line=vlines[idx - 1] if idx > 0 else None,
                next_line=vlines[idx + 1] if idx + 1 < len(vlines) else None,
                body_size=body_size,
                body_fonts=body_fonts,
                body_is_bold=body_is_bold,
                page_h=page_h,
                repeating_margins=repeating_margins,
            )
            if tier is not None:
                heading_tiers.add(tier)

    sorted_tiers = sorted(heading_tiers, key=lambda t: (t[0], t[1]), reverse=True)
    heading_levels = {tier: min(rank, 6) for rank, tier in enumerate(sorted_tiers, start=1)}

    return _DocumentStyleProfile(
        body_size=body_size,
        body_fonts=body_fonts,
        body_is_bold=body_is_bold,
        heading_levels=heading_levels,
        repeating_margin_texts=repeating_margins,
    )


def _line_has_list_prefix(vl: _VisualLine) -> bool:
    if vl.has_vector_bullet:
        return True
    raw = vl.raw_text
    if not raw:
        return False
    if raw[0] in _UNICODE_BULLETS or _HYPHEN_BULLET_RE.match(raw):
        return True
    # Only treat "1. Foo" (with period/paren) as ordered list, not "1 Introduction" heading
    return bool(_ORDERED_ITEM_RE.match(raw))


def _candidate_heading_tier(
    vl: _VisualLine,
    *,
    prev_line: _VisualLine | None,
    next_line: _VisualLine | None,
    body_size: float,
    body_fonts: set[str],
    body_is_bold: bool,
    page_h: float,
    repeating_margins: set[str],
) -> tuple[float, bool] | None:
    raw = vl.raw_text
    if not raw or len(raw) > 120 or _line_has_list_prefix(vl):
        return None
    if vl.is_all_mono:
        return None
    if (vl.y1 <= 0.07 * page_h or vl.y0 >= 0.93 * page_h) and re.sub(r"\d+", "#", raw) in repeating_margins:
        return None

    sz = vl.max_size
    tier_sz = round(sz, 1)

    # Tier A: clearly larger than body prose (>= 8% larger)
    if sz >= body_size * 1.08:
        return (tier_sz, vl.is_all_bold)

    # Standalone structural guards for subtle size/weight headings (Tiers B & C)
    if len(raw) > 80 or raw[-1] in (".", ",", ";", "?", "!"):
        return None
    if prev_line is not None and (vl.y0 - prev_line.y0) < 1.65 * sz and abs(prev_line.max_size - sz) < 0.2:
        return None
    if next_line is not None and (next_line.y0 - vl.y0) < 1.65 * sz and abs(next_line.max_size - sz) < 0.2:
        return None

    # Tier B: slightly larger (>= +0.20pt) AND distinct font resource or bold
    if sz >= body_size + 0.20 and (
        (vl.is_all_bold and not body_is_bold)
        or (body_fonts and vl.dominant_font not in body_fonts)
    ):
        return (tier_sz, vl.is_all_bold)

    # Tier C: body size AND explicitly bold when body prose is not bold
    if abs(sz - body_size) < 0.20 and vl.is_all_bold and not body_is_bold:
        if len(raw) <= 70 and not raw.endswith(":"):
            return (tier_sz, True)

    return None


# ---------------------------------------------------------------------------
# Line assembly: lists, paragraphs, and conservative mid-token unwrapping
# ---------------------------------------------------------------------------


def _strip_unicode_bullet(text: str) -> tuple[bool, str]:
    stripped = text.lstrip()
    if stripped and stripped[0] in _UNICODE_BULLETS:
        return True, stripped[1:].lstrip()
    if _HYPHEN_BULLET_RE.match(stripped):
        return True, stripped[1:].lstrip()
    return False, text


def _estimate_single_line_step(vlines: list[_VisualLine]) -> float:
    dys: list[float] = []
    for a, b in zip(vlines, vlines[1:]):
        if b.has_vector_bullet or _ORDERED_ITEM_RE.match(b.raw_text):
            continue
        sz = max(a.max_size, 8.0)
        dy = b.y0 - a.y0
        if 0.95 * sz <= dy <= 1.92 * sz and abs(a.x0 - b.x0) <= 6.0:
            dys.append(dy)
    if dys:
        dys.sort()
        # Take median of lower half so paragraph gaps don't skew single-line leading
        half = dys[: max(1, (len(dys) + 1) // 2)]
        return float(median(half))
    avg_sz = median([vl.max_size for vl in vlines]) if vlines else 11.0
    return float(avg_sz * 1.45)


def _should_join_mid_token(
    prev_line: _VisualLine,
    curr_line: _VisualLine,
    cell_bounds: tuple[float, float] | None,
) -> bool:
    """Conservatively detect a single long identifier forced to wrap mid-token."""
    if cell_bounds is None:
        return False
    cell_x0, cell_x1 = cell_bounds
    a = prev_line.raw_text
    b = curr_line.raw_text
    if not a or not b or " " in a or " " in b:
        return False
    if len(a) < 7 or not a[-1].isalnum() or not b.isalnum():
        return False
    if abs(prev_line.x0 - curr_line.x0) > 3.0:
        return False

    sz = prev_line.max_size
    left_pad = max(0.0, prev_line.x0 - cell_x0)
    right_gap = cell_x1 - prev_line.x1
    # Token must reach right up to the cell's right padding boundary
    if right_gap > max(left_pad + 1.0 * sz, 1.35 * sz):
        return False

    has_identifier_shape = any(ch.isupper() or ch.isdigit() or ch == "_" for ch in a[1:])
    if b[0].islower():
        is_single_trailing_consonant = len(b) == 1 and b not in ("a", "i")
        return has_identifier_shape or is_single_trailing_consonant
    return has_identifier_shape and len(b) <= 6


def _is_soft_continuation(
    prev_line: _VisualLine,
    curr_line: _VisualLine,
    single_step: float,
    region_x1: float,
    *,
    in_list_item: bool = False,
) -> bool:
    """Return True if *curr_line* is a soft-wrapped continuation of *prev_line*."""
    curr_raw = curr_line.raw_text
    if not curr_raw:
        return False
    if curr_line.has_vector_bullet or curr_raw[0] in _UNICODE_BULLETS or _HYPHEN_BULLET_RE.match(curr_raw):
        return False
    if _ORDERED_ITEM_RE.match(curr_raw):
        return False

    dy = curr_line.y0 - prev_line.y0
    sz = max(prev_line.max_size, 8.0)
    if dy < 0.0 or dy > max(single_step * 1.20, 1.92 * sz):
        return False

    if not in_list_item and abs(curr_line.x0 - prev_line.x0) > 18.0:
        return False
    if in_list_item and curr_line.x0 < prev_line.x0 - 4.0:
        return False

    # First-principles typesetter check: could the first word of curr_line have fit on prev_line?
    first_word = curr_raw.split(" ", 1)[0]
    first_word_w = len(first_word) * sz * 0.50
    remaining_room = region_x1 - prev_line.x1
    region_w = region_x1 - prev_line.x0
    if region_w > 4.5 * sz and remaining_room > first_word_w + 3.0 * sz:
        return False

    return True


def _compute_indent_levels(
    bullet_x0s: list[float],
    base_x0_hint: float | None = None,
) -> list[float]:
    if not bullet_x0s:
        return [base_x0_hint] if base_x0_hint is not None else []
    raw_clusters = _cluster_coords(bullet_x0s, tol=4.0)
    levels: list[float] = []
    if base_x0_hint is not None and raw_clusters[0] > base_x0_hint + 6.0:
        levels.append(base_x0_hint)
    for x in raw_clusters:
        if not levels or x - levels[-1] >= 6.0:
            levels.append(x)
    return levels


def _depth_for_x0(x0: float, levels: list[float]) -> int:
    if not levels:
        return 0
    best_idx = 0
    best_dist = abs(x0 - levels[0])
    for idx, lv in enumerate(levels[1:], start=1):
        dist = abs(x0 - lv)
        if dist < best_dist:
            best_dist = dist
            best_idx = idx
    return best_idx


# ---------------------------------------------------------------------------
# Table rendering
# ---------------------------------------------------------------------------


def _escape_table_cell(text: str) -> str:
    return text.replace("|", "\\|")


def _format_cell_content(
    vlines: list[_VisualLine],
    cell_bbox: tuple[float, float, float, float],
    inline_images: list[_ImageAsset],
    *,
    suppress_bold: bool = False,
    base_x0_hint: float | None = None,
) -> tuple[str, float | None]:
    """Format a table cell's lines and inline images into `<br>`-joined GFM text."""
    cx0, _, cx1, _ = cell_bbox

    # Determine indent levels for any bulleted lines in this cell
    bullet_xs = [
        vl.x0
        for vl in vlines
        if vl.has_vector_bullet
        or (vl.raw_text and (vl.raw_text[0] in _UNICODE_BULLETS or _HYPHEN_BULLET_RE.match(vl.raw_text)))
    ]
    indent_levels = _compute_indent_levels(bullet_xs, base_x0_hint=base_x0_hint)
    new_base_hint = indent_levels[0] if indent_levels else base_x0_hint

    single_step = _estimate_single_line_step(vlines)

    # Build chronological items (paragraphs/bullets and inline images) sorted by y0
    events: list[tuple[float, Literal["line", "img"], object]] = [
        (vl.y0, "line", vl) for vl in vlines if vl.raw_text
    ] + [(img.bbox[1], "img", img) for img in inline_images]
    events.sort(key=lambda item: item[0])

    logical_lines: list[str] = []
    prev_vl: _VisualLine | None = None
    prev_was_bullet = False

    for _, kind, payload in events:
        if kind == "img":
            img: _ImageAsset = payload  # type: ignore[assignment]
            if img.vision_markdown:
                from .hybrid_vision import format_vision_for_table_cell

                logical_lines.extend(
                    format_vision_for_table_cell(img.rel_md_path, img.vision_markdown)
                )
            else:
                logical_lines.append(f"![]({img.rel_md_path})")
            prev_vl = None
            prev_was_bullet = False
            continue

        vl: _VisualLine = payload  # type: ignore[assignment]
        rendered = _render_line_spans(vl.spans, suppress_bold=suppress_bold)
        has_uni_bullet, stripped_rendered = _strip_unicode_bullet(rendered)
        is_bullet = vl.has_vector_bullet or has_uni_bullet

        if is_bullet:
            depth = _depth_for_x0(vl.x0, indent_levels)
            prefix = ("  " * depth) + "- "
            logical_lines.append(prefix + stripped_rendered)
            prev_vl = vl
            prev_was_bullet = True
            continue

        if prev_vl is not None and _is_soft_continuation(
            prev_vl, vl, single_step, cx1, in_list_item=prev_was_bullet
        ):
            if not prev_was_bullet and _should_join_mid_token(prev_vl, vl, (cx0, cx1)):
                logical_lines[-1] = logical_lines[-1] + rendered
            else:
                sep = "" if logical_lines[-1].endswith("-") and rendered[:1].islower() else " "
                logical_lines[-1] = f"{logical_lines[-1]}{sep}{rendered}"
            prev_vl = vl
        else:
            logical_lines.append(rendered)
            prev_vl = vl
            prev_was_bullet = False

    cleaned = [_escape_table_cell(line.strip() if not line.startswith("  ") else line.rstrip()) for line in logical_lines if line.strip()]
    return "<br>".join(cleaned), new_base_hint


def _classify_new_table(
    table: _DetectedTable,
    cell_lines_grid: list[list[list[_VisualLine]]],
    drawings: _PageDrawings,
) -> Literal["top_header", "stub_col", "plain_no_header"]:
    row_count = len(table.rows)
    col_count = len(table.col_edges) - 1
    if row_count == 0 or col_count < 2:
        return "plain_no_header"

    def _row_bgs(r: int) -> list[tuple[float, float, float] | None]:
        return [
            _cell_bg_color(c, drawings) if c is not None else None
            for c in table.rows[r]
        ]

    row0_bgs = _row_bgs(0)
    row0_lines = cell_lines_grid[0]
    row0_non_empty = sum(1 for lines in row0_lines if any(vl.raw_text for vl in lines))

    # 1. Check explicit Row-0 header styling (uniform shading or bold/font distinction across row 0)
    if row0_non_empty >= max(2, col_count - 1):
        row0_all_shaded = all(bg is not None for bg in row0_bgs if bg is not None) and sum(
            1 for bg in row0_bgs if bg is not None
        ) == len(table.rows[0])
        if row0_all_shaded:
            if row_count == 1 or _row_bgs(1) != row0_bgs:
                return "top_header"

        if row_count >= 2:
            row0_bold = all(
                all(vl.is_all_bold for vl in lines if vl.raw_text)
                for lines in row0_lines
                if any(vl.raw_text for vl in lines)
            )
            row1_bold = all(
                all(vl.is_all_bold for vl in lines if vl.raw_text)
                for lines in cell_lines_grid[1]
                if any(vl.raw_text for vl in lines)
            )
            if row0_bold and not row1_bold:
                return "top_header"

            row0_fonts = {
                vl.dominant_font for lines in row0_lines for vl in lines if vl.raw_text
            }
            row1_fonts = {
                vl.dominant_font
                for lines in cell_lines_grid[1]
                for vl in lines
                if vl.raw_text
            }
            if len(row0_fonts) == 1 and row1_fonts and row0_fonts.isdisjoint(row1_fonts):
                return "top_header"

    # 2. Check Column-0 stub styling (col 0 shaded/distinct while col 1+ is unshaded/body)
    stub_shaded_rows = 0
    non_empty_rows = 0
    stub_font_distinct_rows = 0
    for r in range(row_count):
        bgs = _row_bgs(r)
        r_lines = cell_lines_grid[r]
        if any(vl.raw_text for col_l in r_lines for vl in col_l):
            non_empty_rows += 1
        if len(bgs) >= 2 and bgs[0] is not None and all(bg is None for bg in bgs[1:]):
            stub_shaded_rows += 1
        if len(r_lines) >= 2 and r_lines[0] and r_lines[1]:
            f0 = {vl.dominant_font for vl in r_lines[0] if vl.raw_text}
            f1 = {vl.dominant_font for vl in r_lines[1] if vl.raw_text}
            b0 = all(vl.is_all_bold for vl in r_lines[0] if vl.raw_text)
            b1 = all(vl.is_all_bold for vl in r_lines[1] if vl.raw_text)
            if (b0 and not b1) or (f0 and f1 and f0.isdisjoint(f1)):
                stub_font_distinct_rows += 1

    if non_empty_rows > 0 and (
        stub_shaded_rows >= max(1, non_empty_rows // 2)
        or stub_font_distinct_rows >= max(1, (non_empty_rows + 1) // 2)
    ):
        return "stub_col"

    # 3. Unstyled multi-row grid: treat concise label Row 0 as header
    if row_count >= 2 and row0_non_empty == col_count:
        concise_labels = True
        for lines in row0_lines:
            txt = " ".join(vl.raw_text for vl in lines).strip()
            has_bullets = any(_line_has_list_prefix(vl) for vl in lines)
            if not txt or len(txt) > 60 or len(lines) > 2 or has_bullets or txt.endswith("."):
                concise_labels = False
                break
        if concise_labels:
            return "top_header"

    return "plain_no_header"


def _render_table(
    table: _DetectedTable,
    page_spans: list[_SpanItem],
    drawings: _PageDrawings,
    table_images: list[_ImageAsset],
    active_state: _ActiveTableState | None,
) -> tuple[str, _ActiveTableState]:
    col_count = len(table.col_edges) - 1

    # Partition spans and inline images into (row, col) cells
    cell_lines_grid: list[list[list[_VisualLine]]] = []
    cell_imgs_grid: list[list[list[_ImageAsset]]] = []

    for row in table.rows:
        row_vlines: list[list[_VisualLine]] = []
        row_imgs: list[list[_ImageAsset]] = []
        # Normalize row length to col_count using col_edges fallback if needed
        norm_cells: list[tuple[float, float, float, float] | None] = []
        if len(row) == col_count:
            norm_cells = row
        else:
            ry0 = min((c[1] for c in row if c is not None), default=table.bbox[1])
            ry1 = max((c[3] for c in row if c is not None), default=table.bbox[3])
            norm_cells = [
                (table.col_edges[i], ry0, table.col_edges[i + 1], ry1)
                for i in range(col_count)
            ]

        for cell in norm_cells:
            if cell is None:
                row_vlines.append([])
                row_imgs.append([])
                continue
            c_spans = [
                sp
                for sp in page_spans
                if _bbox_contains_point(cell, (sp.x0 + sp.x1) / 2.0, (sp.y0 + sp.y1) / 2.0)
            ]
            row_vlines.append(_group_visual_lines(c_spans, drawings.bullet_rects))
            row_imgs.append(
                [
                    img
                    for img in table_images
                    if _bbox_contains_point(
                        cell,
                        (img.bbox[0] + img.bbox[2]) / 2.0,
                        (img.bbox[1] + img.bbox[3]) / 2.0,
                    )
                ]
            )
        cell_lines_grid.append(row_vlines)
        cell_imgs_grid.append(row_imgs)

    is_continuation = active_state is not None and _edges_match(
        table.col_edges, active_state.col_edges
    )

    if is_continuation and active_state is not None:
        state = active_state
    else:
        mode = _classify_new_table(table, cell_lines_grid, drawings)
        state = _ActiveTableState(col_edges=list(table.col_edges), header_mode=mode)

    formatted_rows: list[list[str]] = []
    for r_idx, row in enumerate(table.rows):
        norm_cells = (
            row
            if len(row) == col_count
            else [
                (
                    table.col_edges[i],
                    table.bbox[1],
                    table.col_edges[i + 1],
                    table.bbox[3],
                )
                for i in range(col_count)
            ]
        )
        out_row: list[str] = []
        for c_idx in range(col_count):
            cell_box = norm_cells[c_idx]
            if cell_box is None:
                out_row.append("")
                continue
            is_header_cell = (
                not is_continuation and r_idx == 0 and state.header_mode == "top_header"
            )
            is_stub_cell = c_idx == 0 and state.header_mode == "stub_col"
            text, updated_hint = _format_cell_content(
                cell_lines_grid[r_idx][c_idx],
                cell_box,
                cell_imgs_grid[r_idx][c_idx],
                suppress_bold=is_header_cell or is_stub_cell,
                base_x0_hint=state.col_base_x0.get(c_idx),
            )
            if updated_hint is not None:
                state.col_base_x0[c_idx] = updated_hint
            if is_stub_cell and text:
                text = f"**{text}**"
            out_row.append(text)
        formatted_rows.append(out_row)

    if not formatted_rows:
        return "", state

    if state.header_mode == "top_header":
        if not is_continuation:
            state.header_cells = formatted_rows[0]
            body_rows = formatted_rows[1:]
        else:
            # Skip duplicate header row if PDF producer repeated it at page top
            if [c.lower() for c in formatted_rows[0]] == [
                c.lower() for c in state.header_cells
            ]:
                body_rows = formatted_rows[1:]
            else:
                body_rows = formatted_rows
        header_row = state.header_cells
    else:
        header_row = [""] * col_count
        body_rows = formatted_rows

    md_lines = [
        "| " + " | ".join(header_row) + " |",
        "| " + " | ".join(["---"] * col_count) + " |",
    ]
    for r in body_rows:
        md_lines.append("| " + " | ".join(r) + " |")
    return "\n".join(md_lines), state


# ---------------------------------------------------------------------------
# Images: embedded rasters + genuine vector diagrams
# ---------------------------------------------------------------------------


def _extract_page_images(
    doc,
    page,
    page_number: int,
    drawings: _PageDrawings,
    tables: list[_DetectedTable],
    image_dir: Path,
    clean_stem: str,
    *,
    capture_for_vision: bool = False,
    repeating_xrefs: set[int] | None = None,
    repeating_digests: set[bytes] | None = None,
) -> list[_ImageAsset]:
    import pymupdf

    pr = page.rect
    assets: list[_ImageAsset] = []
    saved_rects: list[pymupdf.Rect] = []
    img_counter = 0
    rep_xrefs = repeating_xrefs or set()
    rep_digests = repeating_digests or set()

    def _next_paths() -> tuple[Path, str]:
        nonlocal img_counter
        img_counter += 1
        image_dir.mkdir(parents=True, exist_ok=True)
        fname = f"{clean_stem}-p{page_number:03d}-img{img_counter:02d}.png"
        return image_dir / fname, f"{image_dir.name}/{fname}"

    # 1. Embedded raster images
    for info in page.get_image_info(xrefs=True, hashes=capture_for_vision):
        ir = pymupdf.Rect(info["bbox"]) & pr
        if ir.is_empty or ir.width < 16.0 or ir.height < 16.0:
            continue
        if any(abs(ir & prev) >= 0.85 * abs(ir) for prev in saved_rects):
            continue

        xref = int(info.get("xref") or 0)
        out_file, rel_ref = _next_paths()
        try:
            if xref > 0:
                pix = pymupdf.Pixmap(doc, xref)
                if pix.colorspace is None or pix.n - pix.alpha > 3:
                    pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
                pix.save(str(out_file))
            else:
                pix = page.get_pixmap(clip=ir, dpi=150)
                pix.save(str(out_file))
            saved_rects.append(ir)

            is_cand = False
            png_bytes: bytes | None = None
            if capture_for_vision:
                from .hybrid_vision import is_substantive_pdf_raster

                is_cand = is_substantive_pdf_raster(
                    doc, page, info, rep_xrefs, rep_digests
                )
                if is_cand:
                    png_bytes = out_file.read_bytes()

            assets.append(
                _ImageAsset(
                    bbox=(ir.x0, ir.y0, ir.x1, ir.y1),
                    rel_md_path=rel_ref,
                    png_bytes=png_bytes,
                    is_vision_candidate=is_cand,
                    source_kind="raster",
                )
            )
        except Exception as exc:
            LOGGER.debug("Skipping unreadable image on page %d: %s", page_number, exc)

    # 2. Genuine vector diagrams outside tables (excluding callouts, borders, bullets)
    table_rects = [pymupdf.Rect(t.bbox) for t in tables]
    candidate_paths = []
    for d in drawings.raw_drawings:
        dr = pymupdf.Rect(d["rect"]) & pr
        if dr.is_empty:
            continue
        if dr.width > 0.90 * pr.width and dr.height > 0.90 * pr.height:
            continue
        if dr.width <= 8.0 and dr.height <= 8.0:
            continue
        if dr.height <= 3.0 and dr.width >= 10.0:
            continue
        if any(not (dr & tr).is_empty for tr in table_rects):
            continue
        candidate_paths.append(d)

    if candidate_paths:
        for cluster_rect in page.cluster_drawings(drawings=candidate_paths):
            cr = pymupdf.Rect(cluster_rect) & pr
            if cr.is_empty or cr.width < 60.0 or cr.height < 45.0:
                continue
            if any(not (cr & prev).is_empty for prev in saved_rects):
                continue
            if _is_genuine_vector_diagram(cr, candidate_paths):
                out_file, rel_ref = _next_paths()
                pix = page.get_pixmap(clip=cr, dpi=150)
                pix.save(str(out_file))
                saved_rects.append(cr)
                png_bytes = out_file.read_bytes() if capture_for_vision else None
                assets.append(
                    _ImageAsset(
                        bbox=(cr.x0, cr.y0, cr.x1, cr.y1),
                        rel_md_path=rel_ref,
                        png_bytes=png_bytes,
                        is_vision_candidate=capture_for_vision,
                        source_kind="vector",
                    )
                )

    assets.sort(key=lambda a: (a.bbox[1], a.bbox[0]))
    return assets


def _is_genuine_vector_diagram(cluster_rect, paths: list[dict]) -> bool:
    """Distinguish multi-shape diagrams/charts from callout panels and borders."""
    import pymupdf

    interior_rects: list[pymupdf.Rect] = []
    cluster_area = max(abs(cluster_rect), 1.0)

    for d in paths:
        r = pymupdf.Rect(d["rect"]) & cluster_rect
        if r.is_empty:
            continue
        # Ignore outer wrapper fill / border segments of the cluster box itself
        if abs(r) >= 0.78 * cluster_area:
            continue
        on_outer_edge = (
            (r.width <= 4.0 and (abs(r.x0 - cluster_rect.x0) <= 4.0 or abs(r.x1 - cluster_rect.x1) <= 4.0))
            or (r.height <= 4.0 and (abs(r.y0 - cluster_rect.y0) <= 4.0 or abs(r.y1 - cluster_rect.y1) <= 4.0))
        )
        if on_outer_edge:
            continue
        # Ignore small callout icons (e.g. 10x10 info/warning badge)
        if r.width <= 16.0 and r.height <= 16.0:
            continue
        interior_rects.append(r)

    if len(interior_rects) < 3:
        return False

    distinct_boxes: list[pymupdf.Rect] = []
    for r in interior_rects:
        if not any(abs(r & existing) >= 0.80 * min(abs(r), abs(existing)) for existing in distinct_boxes):
            distinct_boxes.append(r)
    return len(distinct_boxes) >= 2


# ---------------------------------------------------------------------------
# Outside-table prose, headings, lists, code blocks & multi-column ordering
# ---------------------------------------------------------------------------


def _split_vlines_around_obstacles(
    vlines: list[_VisualLine],
    col_x1: float,
    obstacle_ys: list[float],
) -> list[tuple[float, float, list[_VisualLine]]]:
    if not vlines:
        return []
    if not obstacle_ys:
        return [(vlines[0].y0, col_x1, vlines)]

    sub_bands: list[tuple[float, float, list[_VisualLine]]] = []
    current: list[_VisualLine] = [vlines[0]]
    for prev_vl, curr_vl in zip(vlines, vlines[1:]):
        if any(prev_vl.y0 < obs_y <= curr_vl.y0 for obs_y in obstacle_ys):
            sub_bands.append((current[0].y0, col_x1, current))
            current = [curr_vl]
        else:
            current.append(curr_vl)
    if current:
        sub_bands.append((current[0].y0, col_x1, current))
    return sub_bands


def _segment_prose_columns(
    page,
    tables: list[_DetectedTable],
    standalone_images: list[_ImageAsset],
    bullet_rects: list[tuple[float, float, float, float]],
    page_spans: list[_SpanItem],
) -> list[tuple[float, float, list[_VisualLine]]]:
    """Group non-table spans into reading-ordered column bands (y_sort_key, col_x1, vlines)."""
    import pymupdf

    pr = page.rect
    outside_spans = [
        sp
        for sp in page_spans
        if not any(
            _bbox_contains_point(t.bbox, (sp.x0 + sp.x1) / 2.0, (sp.y0 + sp.y1) / 2.0)
            for t in tables
        )
    ]
    if not outside_spans:
        return []

    obstacle_ys = [t.bbox[1] for t in tables] + [img.bbox[1] for img in standalone_images]

    # Use PyMuPDF text blocks first to detect genuine multi-line side-by-side columns
    raw_blocks: list[pymupdf.Rect] = []
    for b in page.get_text("dict", flags=pymupdf.TEXTFLAGS_TEXT).get("blocks", []):
        if b.get("type") != 0:
            continue
        br = pymupdf.Rect(b["bbox"]) & pr
        if br.is_empty:
            continue
        if any(
            _bbox_contains_point(t.bbox, (br.x0 + br.x1) / 2.0, (br.y0 + br.y1) / 2.0)
            for t in tables
        ):
            continue
        if len(b.get("lines", [])) >= 2 and br.width <= 0.52 * pr.width:
            raw_blocks.append(br)

    # Check if there are side-by-side multi-line blocks overlapping vertically with a gutter
    split_x: float | None = None
    for i, b1 in enumerate(raw_blocks):
        for b2 in raw_blocks[i + 1 :]:
            left, right = (b1, b2) if b1.x0 <= b2.x0 else (b2, b1)
            v_overlap = min(left.y1, right.y1) - max(left.y0, right.y0)
            if v_overlap >= 20.0 and right.x0 - left.x1 >= 12.0:
                split_x = (left.x1 + right.x0) / 2.0
                break
        if split_x is not None:
            break

    if split_x is None:
        vlines = _group_visual_lines(outside_spans, bullet_rects)
        col_x1 = max((vl.x1 for vl in vlines), default=pr.x1 - 36.0)
        return _split_vlines_around_obstacles(vlines, col_x1, obstacle_ys)

    # Partition spans BEFORE line grouping so side-by-side columns never merge across the gutter
    col_start_y = min(b.y0 for b in raw_blocks) - 2.0
    col_end_y = max(b.y1 for b in raw_blocks) + 2.0

    above_spans: list[_SpanItem] = []
    left_spans: list[_SpanItem] = []
    right_spans: list[_SpanItem] = []
    below_spans: list[_SpanItem] = []

    for sp in outside_spans:
        ym = (sp.y0 + sp.y1) / 2.0
        xm = (sp.x0 + sp.x1) / 2.0
        if ym < col_start_y:
            above_spans.append(sp)
        elif ym > col_end_y:
            below_spans.append(sp)
        elif sp.x0 < split_x - 10.0 and sp.x1 > split_x + 10.0:
            # Full-width banner line inside column vertical range
            above_spans.append(sp)
        elif xm < split_x:
            left_spans.append(sp)
        else:
            right_spans.append(sp)

    bands: list[tuple[float, float, list[_VisualLine]]] = []
    if above_spans:
        vls = _group_visual_lines(above_spans, bullet_rects)
        bands.extend(_split_vlines_around_obstacles(vls, pr.x1 - 36.0, obstacle_ys))
    col_anchor_y = col_start_y
    if left_spans:
        vls = _group_visual_lines(left_spans, bullet_rects)
        if vls:
            col_anchor_y = vls[0].y0
            bands.append((col_anchor_y, split_x - 6.0, vls))
    if right_spans:
        vls = _group_visual_lines(right_spans, bullet_rects)
        if vls:
            bands.append((col_anchor_y + 0.01, pr.x1 - 36.0, vls))
    if below_spans:
        vls = _group_visual_lines(below_spans, bullet_rects)
        bands.extend(_split_vlines_around_obstacles(vls, pr.x1 - 36.0, obstacle_ys))
    return bands


def _render_prose_band(
    vlines: list[_VisualLine],
    col_x1: float,
    page_h: float,
    style_profile: _DocumentStyleProfile,
) -> tuple[str, bool]:
    """Render a contiguous prose band into Markdown; return `(markdown, saw_non_footer_text)`."""
    if not vlines:
        return "", False

    single_step = _estimate_single_line_step(vlines)

    # Pre-compute bullet indent levels per contiguous bullet run
    blocks: list[str] = []
    saw_heading_or_prose = False
    i = 0
    n = len(vlines)

    while i < n:
        vl = vlines[i]
        raw = vl.raw_text
        if not raw:
            i += 1
            continue

        is_margin = (
            vl.y1 <= 0.07 * page_h or vl.y0 >= 0.93 * page_h
        ) and re.sub(r"\d+", "#", raw) in style_profile.repeating_margin_texts
        if not is_margin:
            saw_heading_or_prose = True

        # 1. Heading check
        tier = _candidate_heading_tier(
            vl,
            prev_line=vlines[i - 1] if i > 0 else None,
            next_line=vlines[i + 1] if i + 1 < n else None,
            body_size=style_profile.body_size,
            body_fonts=style_profile.body_fonts,
            body_is_bold=style_profile.body_is_bold,
            page_h=page_h,
            repeating_margins=style_profile.repeating_margin_texts,
        )
        if tier is not None and tier in style_profile.heading_levels:
            level = style_profile.heading_levels[tier]
            h_parts = [_render_line_spans(vl.spans, suppress_bold=True)]
            j = i + 1
            # Only join multi-line wrapped titles when clearly large (>= 1.12x body) and tight leading
            while j < n and vl.max_size >= style_profile.body_size * 1.12:
                nxt = vlines[j]
                nxt_tier = _candidate_heading_tier(
                    nxt,
                    prev_line=vlines[j - 1],
                    next_line=vlines[j + 1] if j + 1 < n else None,
                    body_size=style_profile.body_size,
                    body_fonts=style_profile.body_fonts,
                    body_is_bold=style_profile.body_is_bold,
                    page_h=page_h,
                    repeating_margins=style_profile.repeating_margin_texts,
                )
                if nxt_tier == tier and (nxt.y0 - vlines[j - 1].y0) <= 1.55 * vl.max_size:
                    h_parts.append(_render_line_spans(nxt.spans, suppress_bold=True))
                    j += 1
                else:
                    break
            blocks.append(f"{'#' * level} {' '.join(h_parts)}")
            i = j
            continue

        # 2. Multi-line monospace code block check
        if vl.is_all_mono and i + 1 < n and vlines[i + 1].is_all_mono:
            code_lines: list[_VisualLine] = [vl]
            j = i + 1
            while j < n and vlines[j].is_all_mono and (vlines[j].y0 - vlines[j - 1].y0) <= 2.2 * vl.max_size:
                code_lines.append(vlines[j])
                j += 1
            min_x = min(cl.x0 for cl in code_lines)
            char_w = max(vl.max_size * 0.55, 4.0)
            rendered_code = [
                (" " * int(round(max(0.0, cl.x0 - min_x) / char_w)))
                + _render_line_spans(cl.spans, plain=True)
                for cl in code_lines
            ]
            blocks.append("```\n" + "\n".join(rendered_code) + "\n```")
            i = j
            continue

        # 3. Bulleted or ordered list run
        has_uni_b, _ = _strip_unicode_bullet(raw)
        if vl.has_vector_bullet or has_uni_b or _ORDERED_ITEM_RE.match(raw):
            run_lines: list[_VisualLine] = []
            j = i
            while j < n:
                cur = vlines[j]
                cur_raw = cur.raw_text
                if not cur_raw:
                    j += 1
                    continue
                cur_tier = _candidate_heading_tier(
                    cur,
                    prev_line=vlines[j - 1] if j > 0 else None,
                    next_line=vlines[j + 1] if j + 1 < n else None,
                    body_size=style_profile.body_size,
                    body_fonts=style_profile.body_fonts,
                    body_is_bold=style_profile.body_is_bold,
                    page_h=page_h,
                    repeating_margins=style_profile.repeating_margin_texts,
                )
                if cur_tier is not None and cur_tier in style_profile.heading_levels:
                    break
                c_uni, _ = _strip_unicode_bullet(cur_raw)
                is_item = cur.has_vector_bullet or c_uni or bool(_ORDERED_ITEM_RE.match(cur_raw))
                if is_item:
                    if run_lines and (cur.y0 - run_lines[-1].y0) > 3.2 * cur.max_size:
                        break
                    run_lines.append(cur)
                    j += 1
                elif run_lines and _is_soft_continuation(
                    run_lines[-1], cur, single_step, col_x1, in_list_item=True
                ):
                    run_lines.append(cur)
                    j += 1
                else:
                    break

            bullet_xs = [
                rl.x0
                for rl in run_lines
                if rl.has_vector_bullet
                or _strip_unicode_bullet(rl.raw_text)[0]
                or _ORDERED_ITEM_RE.match(rl.raw_text)
            ]
            levels = _compute_indent_levels(bullet_xs)
            list_out: list[str] = []
            for rl in run_lines:
                rendered_rl = _render_line_spans(rl.spans)
                u_b, stripped_rl = _strip_unicode_bullet(rendered_rl)
                ord_m = _ORDERED_ITEM_RE.match(rl.raw_text)
                if rl.has_vector_bullet or u_b:
                    depth = _depth_for_x0(rl.x0, levels)
                    list_out.append(("  " * depth) + "- " + stripped_rl)
                elif ord_m:
                    depth = _depth_for_x0(rl.x0, levels)
                    list_out.append(("  " * depth) + rendered_rl)
                elif list_out:
                    list_out[-1] = f"{list_out[-1]} {rendered_rl}"
            blocks.append("\n".join(list_out))
            i = j
            continue

        # 4. Standard paragraph (join soft-wrapped lines)
        para_parts = [_render_line_spans(vl.spans)]
        j = i + 1
        while j < n:
            nxt = vlines[j]
            nxt_tier = _candidate_heading_tier(
                nxt,
                prev_line=vlines[j - 1],
                next_line=vlines[j + 1] if j + 1 < n else None,
                body_size=style_profile.body_size,
                body_fonts=style_profile.body_fonts,
                body_is_bold=style_profile.body_is_bold,
                page_h=page_h,
                repeating_margins=style_profile.repeating_margin_texts,
            )
            if nxt_tier is not None and nxt_tier in style_profile.heading_levels:
                break
            if not _is_soft_continuation(vlines[j - 1], nxt, single_step, col_x1):
                break
            rendered_nxt = _render_line_spans(nxt.spans)
            if para_parts[-1].endswith("-") and rendered_nxt[:1].islower():
                para_parts[-1] = para_parts[-1] + rendered_nxt
            else:
                para_parts.append(rendered_nxt)
            j += 1
        blocks.append(" ".join(para_parts))
        i = j

    return "\n\n".join(b for b in blocks if b.strip()), saw_heading_or_prose


# ---------------------------------------------------------------------------
# Full document pipeline & fallback
# ---------------------------------------------------------------------------


def _gather_image_context(
    img: _ImageAsset,
    page_spans: list[_SpanItem],
    tables: list[_DetectedTable],
    active_heading: str,
) -> tuple[str, str]:
    """Collect ``(section_heading, surrounding_text)`` around *img* on its page."""
    cx, cy = (img.bbox[0] + img.bbox[2]) / 2.0, (img.bbox[1] + img.bbox[3]) / 2.0

    # Check if image sits inside a table cell
    for tb in tables:
        if not _bbox_contains_point(tb.bbox, cx, cy):
            continue
        for row in tb.rows:
            for c_idx, cell in enumerate(row):
                if cell is not None and _bbox_contains_point(cell, cx, cy):
                    row_label = ""
                    if c_idx > 0 and row[0] is not None:
                        stub_spans = [
                            sp
                            for sp in page_spans
                            if _bbox_contains_point(
                                row[0], (sp.x0 + sp.x1) / 2.0, (sp.y0 + sp.y1) / 2.0
                            )
                        ]
                        row_label = " ".join(
                            vl.raw_text for vl in _group_visual_lines(stub_spans, [])
                        ).strip()
                    cell_spans = [
                        sp
                        for sp in page_spans
                        if _bbox_contains_point(
                            cell, (sp.x0 + sp.x1) / 2.0, (sp.y0 + sp.y1) / 2.0
                        )
                    ]
                    cell_text = "\n".join(
                        vl.raw_text for vl in _group_visual_lines(cell_spans, []) if vl.raw_text
                    )
                    prefix = f"{row_label}: " if row_label else ""
                    return active_heading, f"{prefix}{cell_text}".strip()

    # Standalone image: collect nearby non-table lines within 150pt above/below
    nearby_spans = [
        sp
        for sp in page_spans
        if img.bbox[1] - 150.0 <= sp.y0 <= img.bbox[3] + 150.0
        and not any(
            _bbox_contains_point(tb.bbox, (sp.x0 + sp.x1) / 2.0, (sp.y0 + sp.y1) / 2.0)
            for tb in tables
        )
    ]
    nearby_text = "\n".join(
        vl.raw_text for vl in _group_visual_lines(nearby_spans, []) if vl.raw_text
    )
    return active_heading, nearby_text


def _render_pdf_page_png(page, dpi: int) -> bytes:
    import pymupdf

    matrix = pymupdf.Matrix(dpi / 72.0, dpi / 72.0)
    pix = page.get_pixmap(matrix=matrix, alpha=False)
    return pix.tobytes("png")


def _extract_structured_pages(
    doc,
    image_dir: Path,
    clean_stem: str,
    *,
    document_name: str = "",
    vision_mode: str = "off",
    api_key: str | None = None,
    vision_model: str | None = None,
    low_threshold: float = 0.30,
    high_threshold: float = 0.70,
    dpi: int = 200,
    concurrency: int = 4,
    max_retries: int = 5,
) -> list[tuple[int, str]]:
    per_page_spans: list[list[_SpanItem]] = []
    per_page_drawings: list[_PageDrawings] = []
    per_page_tables: list[list[_DetectedTable]] = []

    prev_bottom_edges: list[float] | None = None

    for p_idx in range(doc.page_count):
        page = doc[p_idx]
        page.remove_rotation()
        spans = _extract_page_spans(page)
        drawings = _analyze_page_drawings(page)
        tables = _detect_page_tables(page, drawings, spans, prev_bottom_edges)

        per_page_spans.append(spans)
        per_page_drawings.append(drawings)
        per_page_tables.append(tables)

        if tables:
            last_tb = tables[-1]
            page_h = float(page.rect.height)
            has_content_below = any(
                sp.y0 > last_tb.bbox[3] + 12.0
                and sp.y0 < 0.93 * page_h
                and not _bbox_contains_point(
                    last_tb.bbox, (sp.x0 + sp.x1) / 2.0, (sp.y0 + sp.y1) / 2.0
                )
                for sp in spans
            )
            prev_bottom_edges = None if has_content_below else last_tb.col_edges
        else:
            prev_bottom_edges = None

    style_profile = _analyze_document_styles(
        doc, per_page_spans, per_page_drawings, per_page_tables
    )

    use_vision = vision_mode != "off"
    rep_xrefs: set[int] = set()
    rep_digests: set[bytes] = set()
    if use_vision:
        from .hybrid_vision import find_repeating_pdf_image_keys

        rep_xrefs, rep_digests = find_repeating_pdf_image_keys(doc)

    per_page_images: list[list[_ImageAsset]] = []
    for p_idx in range(doc.page_count):
        per_page_images.append(
            _extract_page_images(
                doc,
                doc[p_idx],
                p_idx + 1,
                per_page_drawings[p_idx],
                per_page_tables[p_idx],
                image_dir,
                clean_stem,
                capture_for_vision=use_vision,
                repeating_xrefs=rep_xrefs,
                repeating_digests=rep_digests,
            )
        )

    # Track nearest preceding section heading per page for vision context
    per_page_heading_before: list[str] = []
    running_heading = ""
    for p_idx in range(doc.page_count):
        per_page_heading_before.append(running_heading)
        page_h = float(doc[p_idx].rect.height)
        outside = [
            sp
            for sp in per_page_spans[p_idx]
            if not any(
                _bbox_contains_point(
                    tb.bbox, (sp.x0 + sp.x1) / 2.0, (sp.y0 + sp.y1) / 2.0
                )
                for tb in per_page_tables[p_idx]
            )
        ]
        vls = _group_visual_lines(outside, per_page_drawings[p_idx].bullet_rects)
        for idx, vl in enumerate(vls):
            tier = _candidate_heading_tier(
                vl,
                prev_line=vls[idx - 1] if idx > 0 else None,
                next_line=vls[idx + 1] if idx + 1 < len(vls) else None,
                body_size=style_profile.body_size,
                body_fonts=style_profile.body_fonts,
                body_is_bold=style_profile.body_is_bold,
                page_h=page_h,
                repeating_margins=style_profile.repeating_margin_texts,
            )
            if tier is not None and tier in style_profile.heading_levels:
                running_heading = vl.raw_text

    strategy = "none"
    page_vision_overrides: dict[int, str] = {}
    doc_label = document_name or clean_stem

    if use_vision:
        from .hybrid_vision import (
            DEFAULT_MODEL,
            VisionImageTask,
            VisionPageTask,
            resolve_effective_strategy,
            run_hybrid_vision_tasks,
            screen_pdf_doc,
        )

        vec_counts = [
            sum(1 for a in imgs if a.source_kind == "vector" and a.is_vision_candidate)
            for imgs in per_page_images
        ]
        report = screen_pdf_doc(
            doc,
            Path(doc_label),
            low_threshold=low_threshold,
            high_threshold=high_threshold,
            precomputed_vector_counts=vec_counts,
        )
        strategy = resolve_effective_strategy(report, vision_mode)  # type: ignore[arg-type]
        LOGGER.info(
            "Screened %s: %d/%d page(s) visual (%.0f%%) -> strategy=%s",
            doc_label,
            report.candidate_units,
            report.total_units,
            report.vision_ratio * 100.0,
            strategy,
        )

        effective_model = vision_model or DEFAULT_MODEL
        if strategy == "targeted":
            image_tasks: list[VisionImageTask] = []
            page_tasks: list[VisionPageTask] = []
            asset_by_id: dict[str, _ImageAsset] = {}

            for p_idx, unit in enumerate(report.units):
                if not unit.is_candidate:
                    continue
                pno = p_idx + 1
                if unit.needs_full_page:
                    page_tasks.append(
                        VisionPageTask(
                            unit_number=pno,
                            image_bytes=_render_pdf_page_png(doc[p_idx], dpi),
                            mime_type="image/png",
                            document_name=doc_label,
                            native_text_hint=doc[p_idx].get_text("text").strip(),
                            preceding_context=per_page_heading_before[p_idx],
                        )
                    )
                else:
                    # Find nearest heading above each image on this page
                    page_h = float(doc[p_idx].rect.height)
                    outside = [
                        sp
                        for sp in per_page_spans[p_idx]
                        if not any(
                            _bbox_contains_point(
                                tb.bbox, (sp.x0 + sp.x1) / 2.0, (sp.y0 + sp.y1) / 2.0
                            )
                            for tb in per_page_tables[p_idx]
                        )
                    ]
                    vls = _group_visual_lines(
                        outside, per_page_drawings[p_idx].bullet_rects
                    )
                    page_headings: list[tuple[float, str]] = []
                    for i_vl, vl in enumerate(vls):
                        tier = _candidate_heading_tier(
                            vl,
                            prev_line=vls[i_vl - 1] if i_vl > 0 else None,
                            next_line=vls[i_vl + 1] if i_vl + 1 < len(vls) else None,
                            body_size=style_profile.body_size,
                            body_fonts=style_profile.body_fonts,
                            body_is_bold=style_profile.body_is_bold,
                            page_h=page_h,
                            repeating_margins=style_profile.repeating_margin_texts,
                        )
                        if tier is not None and tier in style_profile.heading_levels:
                            page_headings.append((vl.y0, vl.raw_text))

                    for a_idx, asset in enumerate(per_page_images[p_idx]):
                        if not asset.is_vision_candidate or not asset.png_bytes:
                            continue
                        active_h = per_page_heading_before[p_idx]
                        for hy, htxt in page_headings:
                            if hy <= asset.bbox[1]:
                                active_h = htxt
                        sec_h, surr = _gather_image_context(
                            asset,
                            per_page_spans[p_idx],
                            per_page_tables[p_idx],
                            active_h,
                        )
                        tid = f"p{pno}-a{a_idx}"
                        asset_by_id[tid] = asset
                        image_tasks.append(
                            VisionImageTask(
                                task_id=tid,
                                image_bytes=asset.png_bytes,
                                mime_type="image/png",
                                document_name=doc_label,
                                unit_number=pno,
                                section_heading=sec_h,
                                surrounding_text=surr,
                            )
                        )

            img_res, page_vision_overrides = run_hybrid_vision_tasks(
                image_tasks=image_tasks,
                page_tasks=page_tasks,
                api_key=api_key,
                model=effective_model,
                concurrency=concurrency,
                max_retries=max_retries,
            )
            for tid, md in img_res.items():
                asset_by_id[tid].vision_markdown = md

    pages_out: list[tuple[int, str]] = []
    active_table_state: _ActiveTableState | None = None

    for p_idx in range(doc.page_count):
        page_number = p_idx + 1
        page = doc[p_idx]
        page_h = float(page.rect.height)
        spans = per_page_spans[p_idx]
        drawings = per_page_drawings[p_idx]
        tables = per_page_tables[p_idx]
        images = per_page_images[p_idx]

        # Separate images inside tables from standalone page-level images
        standalone_images: list[_ImageAsset] = []
        images_by_table: list[list[_ImageAsset]] = [[] for _ in tables]
        for img in images:
            cx, cy = (img.bbox[0] + img.bbox[2]) / 2.0, (img.bbox[1] + img.bbox[3]) / 2.0
            matched_t: int | None = None
            for t_idx, tb in enumerate(tables):
                if _bbox_contains_point(tb.bbox, cx, cy):
                    matched_t = t_idx
                    break
            if matched_t is not None:
                images_by_table[matched_t].append(img)
            else:
                standalone_images.append(img)

        prose_bands = _segment_prose_columns(
            page, tables, standalone_images, drawings.bullet_rects, spans
        )

        # Interleave prose bands, tables, and standalone images in reading order
        elements: list[tuple[float, Literal["prose", "table", "img"], object]] = []
        for y_key, col_x1, band_lines in prose_bands:
            elements.append((y_key, "prose", (col_x1, band_lines)))
        for t_idx, tb in enumerate(tables):
            elements.append((tb.bbox[1], "table", (tb, images_by_table[t_idx])))
        for img in standalone_images:
            elements.append((img.bbox[1], "img", img))

        elements.sort(key=lambda e: e[0])

        rendered_chunks: list[str] = []
        for _, kind, payload in elements:
            if kind == "prose":
                col_x1, band_lines = payload  # type: ignore[misc]
                prose_md, breaks_table = _render_prose_band(
                    band_lines, col_x1, page_h, style_profile
                )
                if breaks_table:
                    active_table_state = None
                if prose_md:
                    rendered_chunks.append(prose_md)
            elif kind == "table":
                tb, t_imgs = payload  # type: ignore[misc]
                table_md, active_table_state = _render_table(
                    tb, spans, drawings, t_imgs, active_table_state
                )
                if table_md:
                    rendered_chunks.append(table_md)
            elif kind == "img":
                img_asset: _ImageAsset = payload  # type: ignore[assignment]
                if img_asset.vision_markdown:
                    from .hybrid_vision import format_vision_block_standalone

                    rendered_chunks.append(
                        format_vision_block_standalone(
                            img_asset.rel_md_path, img_asset.vision_markdown
                        )
                    )
                else:
                    rendered_chunks.append(f"![]({img_asset.rel_md_path})")

        if page_number in page_vision_overrides:
            pages_out.append((page_number, page_vision_overrides[page_number]))
        else:
            pages_out.append((page_number, "\n\n".join(rendered_chunks).strip()))

    # Execute page-level or full-document vision using structured local text as spelling hint
    if use_vision and strategy in ("hybrid_pages", "full_document"):
        from .hybrid_vision import DEFAULT_MODEL, VisionPageTask, run_hybrid_vision_tasks

        effective_model = vision_model or DEFAULT_MODEL
        target_page_nums = (
            {u.unit_number for u in report.units if u.is_candidate}
            if strategy == "hybrid_pages"
            else {
                u.unit_number
                for u in report.units
                if u.is_candidate or u.native_char_count > 0
            }
        )
        p_tasks = [
            VisionPageTask(
                unit_number=pno,
                image_bytes=_render_pdf_page_png(doc[pno - 1], dpi),
                mime_type="image/png",
                document_name=doc_label,
                native_text_hint=local_md,
                preceding_context=per_page_heading_before[pno - 1],
            )
            for pno, local_md in pages_out
            if pno in target_page_nums
        ]
        _, full_page_res = run_hybrid_vision_tasks(
            image_tasks=[],
            page_tasks=p_tasks,
            api_key=api_key,
            model=effective_model,
            concurrency=concurrency,
            max_retries=max_retries,
        )
        pages_out = [
            (pno, full_page_res.get(pno, local_md)) for pno, local_md in pages_out
        ]

    return pages_out


def _extract_markdown(
    pdf_path: Path,
    image_dir: Path,
    clean_stem: str,
    *,
    vision_mode: str = "off",
    api_key: str | None = None,
    vision_model: str | None = None,
    low_threshold: float = 0.30,
    high_threshold: float = 0.70,
    dpi: int = 200,
    concurrency: int = 4,
    max_retries: int = 5,
) -> str:
    """Return Markdown text for *pdf_path*, saving images lazily to *image_dir*."""
    try:
        import pymupdf
    except ImportError as exc:
        raise RuntimeError(
            "Missing dependency. Install pymupdf to enable PDF extraction."
        ) from exc

    if vision_mode != "off":
        from .hybrid_vision import validate_vision_options

        validate_vision_options(
            vision_mode, low_threshold, high_threshold, dpi, concurrency, max_retries
        )

    pdf_bytes = pdf_path.read_bytes()

    try:
        with pymupdf.open(stream=pdf_bytes, filetype="pdf") as doc:
            if vision_mode == "off":
                pages = _extract_structured_pages(doc, image_dir, clean_stem)
            else:
                pages = _extract_structured_pages(
                    doc,
                    image_dir,
                    clean_stem,
                    document_name=pdf_path.name,
                    vision_mode=vision_mode,
                    api_key=api_key,
                    vision_model=vision_model,
                    low_threshold=low_threshold,
                    high_threshold=high_threshold,
                    dpi=dpi,
                    concurrency=concurrency,
                    max_retries=max_retries,
                )
        return _join_pages(pages)
    except Exception as exc:
        if vision_mode != "off":
            raise
        LOGGER.warning(
            "Structured PDF extraction failed for %s (%s). Falling back to PyMuPDF text.",
            pdf_path,
            exc,
        )

    with pymupdf.open(stream=pdf_bytes, filetype="pdf") as doc:
        pages = [
            (index, page.get_text("text").strip())
            for index, page in enumerate(doc, start=1)
        ]

    return _join_pages(pages)


def _join_pages(pages: Iterable[tuple[int, str]]) -> str:
    """Join ``(page_number, text)`` pairs into page-marker-delimited Markdown."""
    blocks = []
    for number, text in pages:
        marker = f"<!-- page: {number} -->"
        blocks.append(f"{marker}\n\n{text}" if text else marker)
    return "\n\n".join(blocks)


def convert_pdf(
    pdf_path: Path,
    output_path: Path,
    *,
    vision_mode: str = "off",
    api_key: str | None = None,
    vision_model: str | None = None,
    low_threshold: float = 0.30,
    high_threshold: float = 0.70,
    dpi: int = 200,
    concurrency: int = 4,
    max_retries: int = 5,
) -> None:
    """Convert a single PDF to Markdown and write the result to *output_path*.

    Images are extracted lazily into a sibling ``<stem>_images/`` folder next
    to *output_path* and referenced using relative paths.  When *vision_mode*
    is ``"auto"`` (or ``"images"``/``"pages"``/``"full"``), the PDF is screened
    and augmented with Gemini vision while remaining a single Markdown file.
    """
    clean_stem = re.sub(r"[\s\(\)\[\]]+", "_", pdf_path.stem).strip("_") or "document"
    image_dir = output_path.parent / f"{clean_stem}_images"

    markdown_body = _extract_markdown(
        pdf_path,
        image_dir,
        clean_stem,
        vision_mode=vision_mode,
        api_key=api_key,
        vision_model=vision_model,
        low_threshold=low_threshold,
        high_threshold=high_threshold,
        dpi=dpi,
        concurrency=concurrency,
        max_retries=max_retries,
    )
    wikilink = f"[[{pdf_path.name}]]"
    content = f"{wikilink}\n\n{markdown_body}\n" if markdown_body else f"{wikilink}\n"

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(content, encoding="utf-8")
    LOGGER.info("Wrote %s", output_path)
