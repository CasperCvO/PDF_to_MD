"""Document vision screening and hybrid local + Gemini vision orchestration.

Screens documents (PDF, DOCX, PPTX) to determine whether and how Gemini vision
should augment local deterministic extraction, while always keeping the output
as a **single unified Markdown file** with ``[[Source.ext]]`` and page markers:

- ``0%`` visual candidates (``"none"``): pure local extraction (0 API calls).
- ``<= low_threshold`` (default ``<= 30%``, ``"targeted"``): local extraction
  runs across the whole document; only substantive embedded images/diagrams are
  sent to Gemini *together with surrounding heading/paragraph/table context* and
  spliced inline next to their ``![](...)`` references.  (Individual full-page
  scans inside a mostly native PDF are upgraded to full-page vision.)
- ``low_threshold < ratio <= high_threshold`` (default ``30%–70%``,
  ``"hybrid_pages"``): native text pages use local extraction; visual candidate
  pages use full-page vision anchored by the native text layer.
- ``> high_threshold`` (default ``> 70%``, ``"full_document"``): all non-empty
  pages use full-page vision anchored by native text, merged into one ``.md``
  file with ``<!-- page: N -->`` provenance markers.
"""

from __future__ import annotations

import asyncio
import hashlib
import io
import logging
import random
import re
import subprocess
import tempfile
import zipfile
from collections import Counter
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import Iterator, Literal

from .vision_converter import (
    API_KEY_ENV_VAR,
    DEFAULT_CONCURRENCY,
    DEFAULT_MAX_RETRIES,
    DEFAULT_MODEL,
    VALID_THINKING_LEVELS,
    ThinkingLevel,
    _build_generate_config,
    _create_async_client,
    _describe_error,
    _find_soffice,
    _InvalidModelOutput,
    _is_retryable,
    _normalize_image_bytes,
    _normalize_thinking_level,
    _parse_model_json,
    _resolve_api_key,
)

LOGGER = logging.getLogger(__name__)

VisionMode = Literal["off", "auto", "images", "pages", "full"]
VisionStrategy = Literal["none", "targeted", "hybrid_pages", "full_document"]

DEFAULT_LOW_THRESHOLD = 0.30
DEFAULT_HIGH_THRESHOLD = 0.70
DEFAULT_HYBRID_DPI = 200
VALID_VISION_MODES: tuple[str, ...] = ("off", "auto", "images", "pages", "full")

_CONTEXT_HINT_MAX_CHARS = 3000
_NATIVE_PAGE_HINT_MAX_CHARS = 6000


IMAGE_IN_CONTEXT_PROMPT = """\
You are an expert software architecture and technical documentation analyst.
You are given ONE cropped image (diagram, UI screenshot, API spec snippet, chart,
or terminal dump) extracted from a larger technical document, together with its
surrounding document context.

Return ONLY a single raw JSON object (no Markdown fences) with these exact keys:
- "is_decorative": boolean — true ONLY if the image is a logo, stock graphic,
  empty spacer, or icon carrying no technical/informational content.
- "markdown": string — structured GitHub-flavored Markdown capturing ALL
  information inside the image itself (empty string if "is_decorative" is true).

Rules for "markdown":
1. Use the provided surrounding context ONLY to understand domain terms,
   abbreviations, and component names. Do NOT repeat paragraphs or labels that
   already appear verbatim in the surrounding text outside the image.
2. For architecture / layer / sequence / data-flow diagrams: list the layers,
   components/containers, and directed relationships or protocols clearly using
   concise bullets (e.g. "- **OrderService** --[REST API]--> **CustomerProfile**").
3. For UI / Swagger / API / configuration screenshots: transcribe visible
   endpoints, HTTP methods, fields, statuses, or settings accurately.
4. For tables inside the image: reconstruct as a clean Markdown table or bullet
   list.
5. For charts / dashboards: summarize metrics, numerical values, and trends.
6. For terminal / code screenshots: transcribe verbatim inside a fenced code block.
7. Never invent information not visible in the image.
"""


FULL_PAGE_HYBRID_PROMPT = """\
You are an expert technical documentation analyst converting ONE page of a
document into clean GitHub-flavored Markdown that will be merged into a multi-page
Markdown export.

Return ONLY a single raw JSON object (no Markdown fences) with this exact key:
- "markdown_body": string — the complete Markdown content of this page.

Rules for "markdown_body":
1. Extract ALL visible content on the page in reading order: headings, prose,
   lists, tables, architecture diagrams, UI screenshots, charts, and code dumps.
2. EXACT IDENTIFIER FIDELITY: when native text is provided below, preserve the
   exact spelling and CamelCase/snake_case of technical identifiers, class names,
   rulesets, and URLs from the native text layer (never insert spaces inside
   identifiers such as "SharedDataModelLatest" or "ClientProfileOperations").
3. VALID GFM TABLES: format tables with a header row and "| --- |" separator.
   NEVER put blank lines or raw newlines inside table cells — join multi-line
   content inside a table cell using "<br>" and escape literal pipes as "\\|".
4. Preserve nested bullet hierarchies using 2-space indentation ("- ", "  - ",
   "    - ").
5. Describe architecture diagrams, flowcharts, and embedded screenshots inline at
   their logical position on the page so text-only consumers get complete context.
6. Do NOT include page numbers, running headers/footers, YAML frontmatter, or
   "<!-- page: N -->" comments (those are added by the caller).
"""


# ---------------------------------------------------------------------------
# Screening & task data models
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class UnitVisionAssessment:
    """Screening result for a single page (PDF/DOCX) or slide (PPTX)."""

    unit_number: int
    is_candidate: bool
    needs_full_page: bool = False
    raster_count: int = 0
    vector_diagram_count: int = 0
    native_char_count: int = 0
    reasons: tuple[str, ...] = ()


@dataclass(frozen=True)
class VisionScreenReport:
    """Document-level screening summary and recommended vision strategy."""

    file_path: Path
    total_units: int
    candidate_units: int
    vision_ratio: float
    recommended_strategy: VisionStrategy
    units: tuple[UnitVisionAssessment, ...]

    @property
    def candidate_unit_numbers(self) -> list[int]:
        return [u.unit_number for u in self.units if u.is_candidate]


@dataclass(frozen=True)
class VisionImageTask:
    """A cropped image + surrounding document context to analyze."""

    task_id: str
    image_bytes: bytes
    mime_type: str
    document_name: str
    unit_number: int | None = None
    section_heading: str = ""
    surrounding_text: str = ""


@dataclass(frozen=True)
class VisionPageTask:
    """A rendered full page/slide + native text hint to analyze."""

    unit_number: int
    image_bytes: bytes
    mime_type: str
    document_name: str
    native_text_hint: str = ""
    preceding_context: str = ""


# ---------------------------------------------------------------------------
# Option validation & strategy routing
# ---------------------------------------------------------------------------


def validate_vision_options(
    vision_mode: str,
    low_threshold: float,
    high_threshold: float,
    dpi: int = DEFAULT_HYBRID_DPI,
    concurrency: int = DEFAULT_CONCURRENCY,
    max_retries: int = DEFAULT_MAX_RETRIES,
    vision_model: str | None = None,
    thinking_level: str | None = None,
) -> None:
    if vision_mode not in VALID_VISION_MODES:
        raise ValueError(
            f"vision_mode must be one of {VALID_VISION_MODES} (got {vision_mode!r})"
        )
    if not (0.0 <= low_threshold <= 1.0):
        raise ValueError(f"low_threshold must be between 0.0 and 1.0 (got {low_threshold})")
    if not (0.0 <= high_threshold <= 1.0):
        raise ValueError(f"high_threshold must be between 0.0 and 1.0 (got {high_threshold})")
    if low_threshold > high_threshold:
        raise ValueError(
            f"low_threshold ({low_threshold}) must not exceed high_threshold ({high_threshold})"
        )
    if not 72 <= dpi <= 600:
        raise ValueError(f"dpi must be between 72 and 600 (got {dpi})")
    if concurrency < 1:
        raise ValueError(f"concurrency must be >= 1 (got {concurrency})")
    if max_retries < 0:
        raise ValueError(f"max_retries must be >= 0 (got {max_retries})")
    if vision_model is not None and not vision_model.strip():
        raise ValueError("vision_model must not be empty")
    _normalize_thinking_level(thinking_level)


def choose_strategy_from_ratio(
    total_units: int,
    candidate_units: int,
    *,
    low_threshold: float = DEFAULT_LOW_THRESHOLD,
    high_threshold: float = DEFAULT_HIGH_THRESHOLD,
) -> VisionStrategy:
    """Map ``candidate_units / total_units`` onto a :data:`VisionStrategy`."""
    if total_units <= 0 or candidate_units <= 0:
        return "none"
    ratio = candidate_units / total_units
    if ratio > high_threshold or (high_threshold == 0.0 and candidate_units > 0):
        return "full_document"
    if ratio <= low_threshold:
        return "targeted"
    return "hybrid_pages"


def resolve_effective_strategy(
    report: VisionScreenReport,
    vision_mode: VisionMode,
) -> VisionStrategy:
    """Resolve the execution strategy given *report* and explicit *vision_mode*."""
    if vision_mode == "off":
        return "none"
    if vision_mode == "full":
        return "full_document" if report.total_units > 0 else "none"
    if vision_mode == "pages":
        return "hybrid_pages" if report.candidate_units > 0 else "none"
    if vision_mode == "images":
        return "targeted" if report.candidate_units > 0 else "none"
    return report.recommended_strategy


# ---------------------------------------------------------------------------
# Visual complexity & image helpers
# ---------------------------------------------------------------------------


def pixmap_has_visual_detail(pix) -> bool:
    """Return False for solid single-color spacer/banner images."""
    if pix.width < 16 or pix.height < 16:
        return False
    samples = pix.samples
    n = pix.n
    alpha = pix.alpha
    color_channels = max(1, n - alpha)
    total_pixels = pix.width * pix.height
    step = max(1, total_pixels // 400) * n

    buckets: set[tuple[int, ...]] = set()
    for offset in range(0, len(samples) - n + 1, step):
        if alpha and samples[offset + n - 1] < 16:
            continue
        bucket = tuple(samples[offset + c] // 32 for c in range(min(3, color_channels)))
        buckets.add(bucket)
        if len(buckets) >= 2:
            return True
    return False


def image_bytes_have_visual_detail(data: bytes) -> tuple[bool, int, int, str]:
    """Inspect encoded image bytes; return ``(has_detail, width, height, mime_type)``."""
    try:
        from PIL import Image
    except ImportError:
        return True, 0, 0, "image/png"

    try:
        with Image.open(io.BytesIO(data)) as img:
            fmt = (img.format or "PNG").upper()
            if fmt not in ("PNG", "JPEG", "GIF", "BMP", "WEBP", "TIFF"):
                return False, 0, 0, "image/png"
            width, height = img.size
            if width < 24 or height < 24:
                return False, width, height, "image/png"
            rgb = img.convert("RGB")
            raw = rgb.tobytes()
            total_px = width * height
            step = max(1, total_px // 400) * 3
            buckets: set[tuple[int, int, int]] = set()
            for off in range(0, len(raw) - 2, step):
                buckets.add((raw[off] // 32, raw[off + 1] // 32, raw[off + 2] // 32))
                if len(buckets) >= 2:
                    mime = "image/jpeg" if fmt == "JPEG" else "image/png"
                    return True, width, height, mime
            return False, width, height, "image/png"
    except Exception:
        return False, 0, 0, "image/png"


def prepare_png_or_jpeg_for_api(data: bytes, mime_type: str = "image/png") -> tuple[bytes, str]:
    """Ensure image bytes are PNG/JPEG and normalized within Gemini payload limits."""
    try:
        from PIL import Image

        with Image.open(io.BytesIO(data)) as img:
            width, height = img.size
            fmt = (img.format or "").upper()
            if fmt not in ("PNG", "JPEG"):
                buf = io.BytesIO()
                img.convert("RGB").save(buf, format="PNG")
                data = buf.getvalue()
                mime_type = "image/png"
            else:
                mime_type = "image/jpeg" if fmt == "JPEG" else "image/png"
            return _normalize_image_bytes(data, mime_type, width, height)
    except Exception:
        return data, mime_type


# ---------------------------------------------------------------------------
# Document screening (PDF, DOCX, PPTX)
# ---------------------------------------------------------------------------


def screen_document(
    file_path: Path | str,
    *,
    low_threshold: float = DEFAULT_LOW_THRESHOLD,
    high_threshold: float = DEFAULT_HIGH_THRESHOLD,
) -> VisionScreenReport:
    """Screen *file_path* to assess whether and how vision extraction applies."""
    validate_vision_options("auto", low_threshold, high_threshold)
    resolved = Path(file_path).resolve()
    if not resolved.is_file():
        raise FileNotFoundError(f"Document not found: {resolved}")

    ext = resolved.suffix.lower()
    if ext == ".pdf":
        return screen_pdf_path(
            resolved, low_threshold=low_threshold, high_threshold=high_threshold
        )
    if ext == ".docx":
        return screen_docx_path(
            resolved, low_threshold=low_threshold, high_threshold=high_threshold
        )
    if ext == ".pptx":
        return screen_pptx_path(
            resolved, low_threshold=low_threshold, high_threshold=high_threshold
        )

    return VisionScreenReport(
        file_path=resolved,
        total_units=1,
        candidate_units=0,
        vision_ratio=0.0,
        recommended_strategy="none",
        units=(UnitVisionAssessment(unit_number=1, is_candidate=False),),
    )


def find_repeating_pdf_image_keys(doc) -> tuple[set[int], set[bytes]]:
    """Identify template/logo image xrefs and digests repeated across pages."""
    page_count = doc.page_count
    if page_count < 3:
        return set(), set()

    xref_pages: Counter[int] = Counter()
    digest_pages: Counter[bytes] = Counter()

    for p_idx in range(page_count):
        page = doc[p_idx]
        seen_xrefs: set[int] = set()
        seen_digests: set[bytes] = set()
        for info in page.get_image_info(xrefs=True, hashes=True):
            xref = int(info.get("xref") or 0)
            digest = info.get("digest")
            if xref > 0:
                seen_xrefs.add(xref)
            if isinstance(digest, (bytes, bytearray)) and digest:
                seen_digests.add(bytes(digest))
        for x in seen_xrefs:
            xref_pages[x] += 1
        for dg in seen_digests:
            digest_pages[dg] += 1

    min_repeats = max(3, int(page_count * 0.40))
    rep_xrefs = {x for x, cnt in xref_pages.items() if cnt >= min_repeats}
    rep_digests = {dg for dg, cnt in digest_pages.items() if cnt >= min_repeats}
    return rep_xrefs, rep_digests


def is_substantive_pdf_raster(
    doc,
    page,
    info: dict,
    repeating_xrefs: set[int],
    repeating_digests: set[bytes],
) -> bool:
    """Return True if an image_info entry represents substantive content (not chrome/logo)."""
    import pymupdf

    pr = page.rect
    page_area = max(abs(pr), 1.0)
    ir = pymupdf.Rect(info["bbox"]) & pr
    if ir.is_empty:
        return False

    xref = int(info.get("xref") or 0)
    digest = bytes(info["digest"]) if info.get("digest") else b""
    if (xref > 0 and xref in repeating_xrefs) or (digest and digest in repeating_digests):
        return False

    px_w = int(info.get("width") or 0)
    px_h = int(info.get("height") or 0)
    large_enough = (
        ir.width >= 70.0 and ir.height >= 40.0 and abs(ir) >= 0.012 * page_area
    ) or (px_w >= 220 and px_h >= 100 and ir.width >= 50.0 and ir.height >= 30.0)
    if not large_enough:
        return False

    try:
        if xref > 0:
            pix = pymupdf.Pixmap(doc, xref)
        else:
            pix = page.get_pixmap(clip=ir, dpi=96)
        return pixmap_has_visual_detail(pix)
    except Exception:
        return False


def screen_pdf_doc(
    doc,
    source_path: Path,
    *,
    low_threshold: float = DEFAULT_LOW_THRESHOLD,
    high_threshold: float = DEFAULT_HIGH_THRESHOLD,
    precomputed_vector_counts: list[int] | None = None,
) -> VisionScreenReport:
    """Screen an already-open ``pymupdf.Document``."""
    import pymupdf

    from .pdf_converter import (
        _analyze_page_drawings,
        _detect_page_tables,
        _extract_page_spans,
        _is_genuine_vector_diagram,
    )

    rep_xrefs, rep_digests = find_repeating_pdf_image_keys(doc)
    units: list[UnitVisionAssessment] = []
    prev_bottom_edges: list[float] | None = None

    for p_idx in range(doc.page_count):
        page = doc[p_idx]
        pr = page.rect
        page_area = max(abs(pr), 1.0)
        native_text = page.get_text("text").strip()
        native_chars = len(native_text)

        raster_count = 0
        visual_area = 0.0
        seen_rects: list[pymupdf.Rect] = []
        for info in page.get_image_info(xrefs=True, hashes=True):
            ir = pymupdf.Rect(info["bbox"]) & pr
            if ir.is_empty or any(abs(ir & prev) >= 0.85 * abs(ir) for prev in seen_rects):
                continue
            if is_substantive_pdf_raster(doc, page, info, rep_xrefs, rep_digests):
                seen_rects.append(ir)
                raster_count += 1
                visual_area += abs(ir)

        if precomputed_vector_counts is not None:
            vec_count = precomputed_vector_counts[p_idx]
        else:
            spans = _extract_page_spans(page)
            drawings = _analyze_page_drawings(page)
            tables = _detect_page_tables(page, drawings, spans, prev_bottom_edges)
            if tables:
                last_tb = tables[-1]
                has_below = any(
                    sp.y0 > last_tb.bbox[3] + 12.0 and sp.y0 < 0.93 * float(pr.height)
                    for sp in spans
                )
                prev_bottom_edges = None if has_below else last_tb.col_edges
            else:
                prev_bottom_edges = None

            table_rects = [pymupdf.Rect(t.bbox) for t in tables]
            cand_paths = []
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
                cand_paths.append(d)

            vec_count = 0
            if cand_paths:
                for cr_raw in page.cluster_drawings(drawings=cand_paths):
                    cr = pymupdf.Rect(cr_raw) & pr
                    if cr.is_empty or cr.width < 60.0 or cr.height < 45.0:
                        continue
                    if any(not (cr & prev).is_empty for prev in seen_rects):
                        continue
                    if _is_genuine_vector_diagram(cr, cand_paths):
                        vec_count += 1
                        visual_area += abs(cr)

        reasons: list[str] = []
        if raster_count > 0:
            reasons.append("embedded_image")
        if vec_count > 0:
            reasons.append("vector_diagram")

        is_scanned_like = native_chars < 40 and (raster_count > 0 or vec_count > 0)
        if is_scanned_like:
            reasons.append("scanned_or_low_text")

        is_candidate = bool(reasons)
        needs_full_page = is_candidate and (
            is_scanned_like
            or visual_area >= 0.45 * page_area
            or (raster_count + vec_count >= 3 and native_chars < 400)
        )

        units.append(
            UnitVisionAssessment(
                unit_number=p_idx + 1,
                is_candidate=is_candidate,
                needs_full_page=needs_full_page,
                raster_count=raster_count,
                vector_diagram_count=vec_count,
                native_char_count=native_chars,
                reasons=tuple(reasons),
            )
        )

    total = len(units)
    cand = sum(1 for u in units if u.is_candidate)
    ratio = cand / total if total > 0 else 0.0
    strategy = choose_strategy_from_ratio(
        total, cand, low_threshold=low_threshold, high_threshold=high_threshold
    )
    return VisionScreenReport(
        file_path=source_path,
        total_units=total,
        candidate_units=cand,
        vision_ratio=ratio,
        recommended_strategy=strategy,
        units=tuple(units),
    )


def screen_pdf_path(
    pdf_path: Path,
    *,
    low_threshold: float = DEFAULT_LOW_THRESHOLD,
    high_threshold: float = DEFAULT_HIGH_THRESHOLD,
) -> VisionScreenReport:
    import pymupdf

    with pymupdf.open(pdf_path) as doc:
        return screen_pdf_doc(
            doc, pdf_path, low_threshold=low_threshold, high_threshold=high_threshold
        )


def screen_docx_path(
    docx_path: Path,
    *,
    low_threshold: float = DEFAULT_LOW_THRESHOLD,
    high_threshold: float = DEFAULT_HIGH_THRESHOLD,
) -> VisionScreenReport:
    """Screen a ``.docx`` package by inspecting embedded drawings per approximate page."""
    from lxml import etree

    from .office_converter import (
        _SKIPPED_ANCESTORS,
        _W_BODY,
        _W_T,
        _docx_page_break_candidates,
    )

    r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    a_blip = "{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
    wp_extent = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent"
    emu_per_pt = 12700.0

    with zipfile.ZipFile(docx_path) as zf:
        try:
            doc_xml = zf.read("word/document.xml")
        except KeyError:
            return VisionScreenReport(
                file_path=docx_path,
                total_units=1,
                candidate_units=0,
                vision_ratio=0.0,
                recommended_strategy="none",
                units=(UnitVisionAssessment(unit_number=1, is_candidate=False),),
            )

        rels_map: dict[str, str] = {}
        try:
            rels_root = etree.fromstring(zf.read("word/_rels/document.xml.rels"))
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
            return VisionScreenReport(
                file_path=docx_path,
                total_units=1,
                candidate_units=0,
                vision_ratio=0.0,
                recommended_strategy="none",
                units=(UnitVisionAssessment(unit_number=1, is_candidate=False),),
            )

        breaks = {id(elem) for _, elem in _docx_page_break_candidates(body)}
        total_pages = 1 + len(breaks)

        # First pass: locate all blips and detect repeating logo hashes across pages
        page_num = 1
        page_chars: Counter[int] = Counter()
        page_blips: dict[int, list[tuple[str, float, float]]] = {
            p: [] for p in range(1, total_pages + 1)
        }
        hash_pages: Counter[str] = Counter()
        media_cache: dict[str, bytes] = {}

        for elem in body.iter():
            if not isinstance(elem.tag, str):
                continue
            if any(anc.tag in _SKIPPED_ANCESTORS for anc in elem.iterancestors()):
                continue
            if id(elem) in breaks:
                page_num = min(total_pages, page_num + 1)
            if elem.tag == _W_T and elem.text:
                page_chars[page_num] += len(elem.text.strip())
            elif elem.tag == a_blip:
                rid = elem.get(f"{{{r_ns}}}embed")
                target = rels_map.get(rid or "")
                if not target:
                    continue
                if target not in media_cache:
                    try:
                        media_cache[target] = zf.read(target)
                    except KeyError:
                        media_cache[target] = b""
                blob = media_cache[target]
                if not blob:
                    continue
                # Find nearest ancestor drawing extent if present
                w_pt, h_pt = 0.0, 0.0
                for anc in elem.iterancestors():
                    ext_el = anc.find(wp_extent)
                    if ext_el is not None:
                        try:
                            w_pt = int(ext_el.get("cx") or 0) / emu_per_pt
                            h_pt = int(ext_el.get("cy") or 0) / emu_per_pt
                        except ValueError:
                            pass
                        break
                page_blips[page_num].append((target, w_pt, h_pt))

        for p, blips in page_blips.items():
            seen_h: set[str] = set()
            for target, _, _ in blips:
                seen_h.add(hashlib.sha1(media_cache[target]).hexdigest())
            for h in seen_h:
                hash_pages[h] += 1

        rep_hashes = {
            h
            for h, cnt in hash_pages.items()
            if total_pages >= 3 and cnt >= max(3, int(total_pages * 0.40))
        }

        units: list[UnitVisionAssessment] = []
        for p in range(1, total_pages + 1):
            substantive = 0
            for target, w_pt, h_pt in page_blips[p]:
                blob = media_cache[target]
                if hashlib.sha1(blob).hexdigest() in rep_hashes:
                    continue
                has_detail, px_w, px_h, _ = image_bytes_have_visual_detail(blob)
                if not has_detail:
                    continue
                if (w_pt >= 70.0 and h_pt >= 40.0) or (w_pt == 0.0 and px_w >= 160 and px_h >= 80):
                    substantive += 1
            chars = page_chars[p]
            reasons = ("embedded_image",) if substantive > 0 else ()
            units.append(
                UnitVisionAssessment(
                    unit_number=p,
                    is_candidate=substantive > 0,
                    needs_full_page=substantive > 0 and chars < 40,
                    raster_count=substantive,
                    native_char_count=chars,
                    reasons=reasons,
                )
            )

    cand = sum(1 for u in units if u.is_candidate)
    ratio = cand / total_pages if total_pages > 0 else 0.0
    strategy = choose_strategy_from_ratio(
        total_pages, cand, low_threshold=low_threshold, high_threshold=high_threshold
    )
    return VisionScreenReport(
        file_path=docx_path,
        total_units=total_pages,
        candidate_units=cand,
        vision_ratio=ratio,
        recommended_strategy=strategy,
        units=tuple(units),
    )


def _iter_pptx_shapes(shapes):
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    for shape in shapes:
        yield shape
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _iter_pptx_shapes(shape.shapes)


def screen_pptx_path(
    pptx_path: Path,
    *,
    low_threshold: float = DEFAULT_LOW_THRESHOLD,
    high_threshold: float = DEFAULT_HIGH_THRESHOLD,
) -> VisionScreenReport:
    """Screen a ``.pptx`` presentation slide by slide."""
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    emu_per_pt = 12700.0
    prs = Presentation(str(pptx_path))
    total_slides = len(prs.slides)
    if total_slides == 0:
        return VisionScreenReport(
            file_path=pptx_path,
            total_units=0,
            candidate_units=0,
            vision_ratio=0.0,
            recommended_strategy="none",
            units=(),
        )

    sha_slides: Counter[str] = Counter()
    per_slide_pics: list[list[tuple[str, bytes, float, float]]] = []
    per_slide_charts: list[int] = []
    per_slide_diagrams: list[int] = []
    per_slide_chars: list[int] = []

    for slide in prs.slides:
        pics: list[tuple[str, bytes, float, float]] = []
        charts = 0
        auto_shapes = 0
        chars = 0
        seen_sha: set[str] = set()

        for shape in _iter_pptx_shapes(slide.shapes):
            if getattr(shape, "has_text_frame", False):
                for para in shape.text_frame.paragraphs:
                    chars += len(para.text.strip())
            if getattr(shape, "has_table", False):
                for row in shape.table.rows:
                    for cell in row.cells:
                        chars += len(cell.text.strip())
            if getattr(shape, "has_chart", False):
                charts += 1
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                w_pt = float(shape.width or 0) / emu_per_pt
                h_pt = float(shape.height or 0) / emu_per_pt
                try:
                    sha = shape.image.sha1
                    blob = shape.image.blob
                except Exception:
                    continue
                pics.append((sha, blob, w_pt, h_pt))
                seen_sha.add(sha)
            elif (
                shape.shape_type in (MSO_SHAPE_TYPE.AUTO_SHAPE, MSO_SHAPE_TYPE.LINE)
                and not getattr(shape, "is_placeholder", False)
            ):
                auto_shapes += 1

        for sha in seen_sha:
            sha_slides[sha] += 1
        per_slide_pics.append(pics)
        per_slide_charts.append(charts)
        per_slide_diagrams.append(1 if auto_shapes >= 4 else 0)
        per_slide_chars.append(chars)

    rep_shas = {
        sha
        for sha, cnt in sha_slides.items()
        if total_slides >= 3 and cnt >= max(3, int(total_slides * 0.40))
    }

    units: list[UnitVisionAssessment] = []
    for idx in range(total_slides):
        substantive_pics = 0
        for sha, blob, w_pt, h_pt in per_slide_pics[idx]:
            if sha in rep_shas:
                continue
            if w_pt < 70.0 or h_pt < 40.0:
                continue
            has_detail, _, _, _ = image_bytes_have_visual_detail(blob)
            if has_detail:
                substantive_pics += 1

        reasons: list[str] = []
        if substantive_pics > 0:
            reasons.append("embedded_image")
        if per_slide_charts[idx] > 0:
            reasons.append("chart")
        if per_slide_diagrams[idx] > 0:
            reasons.append("vector_diagram")

        is_cand = bool(reasons)
        needs_full = is_cand and (
            per_slide_charts[idx] > 0
            or per_slide_diagrams[idx] > 0
            or substantive_pics >= 2
            or per_slide_chars[idx] < 40
        )
        units.append(
            UnitVisionAssessment(
                unit_number=idx + 1,
                is_candidate=is_cand,
                needs_full_page=needs_full,
                raster_count=substantive_pics,
                vector_diagram_count=per_slide_charts[idx] + per_slide_diagrams[idx],
                native_char_count=per_slide_chars[idx],
                reasons=tuple(reasons),
            )
        )

    cand = sum(1 for u in units if u.is_candidate)
    ratio = cand / total_slides
    strategy = choose_strategy_from_ratio(
        total_slides, cand, low_threshold=low_threshold, high_threshold=high_threshold
    )
    return VisionScreenReport(
        file_path=pptx_path,
        total_units=total_slides,
        candidate_units=cand,
        vision_ratio=ratio,
        recommended_strategy=strategy,
        units=tuple(units),
    )


# ---------------------------------------------------------------------------
# Concurrent Gemini execution for image-in-context and full-page tasks
# ---------------------------------------------------------------------------


def run_hybrid_vision_tasks(
    *,
    image_tasks: list[VisionImageTask],
    page_tasks: list[VisionPageTask],
    api_key: str | None,
    model: str = DEFAULT_MODEL,
    thinking_level: str | None = None,
    concurrency: int = DEFAULT_CONCURRENCY,
    max_retries: int = DEFAULT_MAX_RETRIES,
) -> tuple[dict[str, str], dict[int, str]]:
    """Execute image-level and page-level Gemini calls in one bounded async pool.

    Returns ``(image_markdown_by_task_id, page_markdown_by_unit_number)``.
    """
    normalized_thinking = _normalize_thinking_level(thinking_level)
    if not image_tasks and not page_tasks:
        return {}, {}

    resolved_key = _resolve_api_key(api_key)
    if not resolved_key:
        raise RuntimeError(
            f"No API key available for vision extraction. Set the {API_KEY_ENV_VAR} "
            "environment variable or pass api_key / --api-key."
        )

    return asyncio.run(
        _run_tasks_async(
            image_tasks=image_tasks,
            page_tasks=page_tasks,
            api_key=resolved_key,
            model=model,
            thinking_level=normalized_thinking,
            concurrency=concurrency,
            max_retries=max_retries,
        )
    )


async def _run_tasks_async(
    *,
    image_tasks: list[VisionImageTask],
    page_tasks: list[VisionPageTask],
    api_key: str,
    model: str,
    thinking_level: str | None = None,
    concurrency: int,
    max_retries: int,
) -> tuple[dict[str, str], dict[int, str]]:
    client = _create_async_client(api_key)
    semaphore = asyncio.Semaphore(concurrency)

    async def _run_img(task: VisionImageTask) -> tuple[str, str]:
        async with semaphore:
            md = await _call_with_retry(
                label=f"image {task.task_id}",
                max_retries=max_retries,
                coro_factory=lambda: _extract_image_in_context_once(
                    client, model, task, thinking_level=thinking_level
                ),
            )
            return task.task_id, md

    async def _run_page(task: VisionPageTask) -> tuple[int, str]:
        async with semaphore:
            md = await _call_with_retry(
                label=f"page {task.unit_number}",
                max_retries=max_retries,
                coro_factory=lambda: _extract_full_page_once(
                    client, model, task, thinking_level=thinking_level
                ),
            )
            return task.unit_number, md

    try:
        img_coros = [_run_img(t) for t in image_tasks]
        page_coros = [_run_page(t) for t in page_tasks]
        all_results = await asyncio.gather(*img_coros, *page_coros)
    finally:
        await client.aclose()

    img_results: dict[str, str] = {}
    page_results: dict[int, str] = {}
    for idx, item in enumerate(all_results):
        if idx < len(image_tasks):
            k, v = item
            img_results[str(k)] = v
        else:
            pno, v = item
            page_results[int(pno)] = v
    return img_results, page_results


async def _call_with_retry(label: str, max_retries: int, coro_factory) -> str:
    last_error: BaseException | None = None
    attempts_made = 0
    for attempt in range(max_retries + 1):
        attempts_made = attempt + 1
        try:
            return await coro_factory()
        except Exception as exc:  # noqa: BLE001
            last_error = exc
            if not _is_retryable(exc) or attempt == max_retries:
                break
            delay = min(60.0, 2**attempt) + random.uniform(0.0, 0.5)
            LOGGER.warning(
                "Vision %s attempt %d/%d failed (%s). Retrying in %.1fs.",
                label,
                attempt + 1,
                max_retries + 1,
                _describe_error(exc),
                delay,
            )
            await asyncio.sleep(delay)
    raise RuntimeError(
        f"Vision {label} failed after {attempts_made} attempt(s) "
        f"(max_retries={max_retries}): {_describe_error(last_error)}"
    ) from last_error


async def _extract_image_in_context_once(
    client,
    model: str,
    task: VisionImageTask,
    *,
    thinking_level: str | None = None,
) -> str:
    ctx_lines = [f"Document: {task.document_name}"]
    if task.unit_number is not None:
        ctx_lines.append(f"Page/Slide: {task.unit_number}")
    if task.section_heading:
        ctx_lines.append(f"Section: {task.section_heading}")
    if task.surrounding_text.strip():
        hint = task.surrounding_text.strip()[:_CONTEXT_HINT_MAX_CHARS]
        ctx_lines.append(
            "\n--- surrounding document text (for context only; do not duplicate) ---\n"
            f"{hint}\n"
            "--- end surrounding text ---"
        )
    prompt = IMAGE_IN_CONTEXT_PROMPT + "\n\n" + "\n".join(ctx_lines)

    data, mime = prepare_png_or_jpeg_for_api(task.image_bytes, task.mime_type)
    contents = [
        {
            "role": "user",
            "parts": [
                {"text": prompt},
                {"inline_data": {"mime_type": mime, "data": data}},
            ],
        }
    ]
    response = await asyncio.wait_for(
        client.models.generate_content(
            model=model,
            contents=contents,
            config=_build_generate_config(thinking_level),
        ),
        timeout=180,
    )
    text = getattr(response, "text", None)
    if not text:
        feedback = getattr(response, "prompt_feedback", None)
        reason = getattr(feedback, "block_reason", None) if feedback else None
        raise _InvalidModelOutput(f"Model returned no text (block reason: {reason})")

    payload = _parse_model_json(text)
    if bool(payload.get("is_decorative")):
        return ""
    return str(payload.get("markdown", "") or "").strip()


async def _extract_full_page_once(
    client,
    model: str,
    task: VisionPageTask,
    *,
    thinking_level: str | None = None,
) -> str:
    ctx_lines = [
        f"Document: {task.document_name}",
        f"Page number: {task.unit_number}",
    ]
    if task.preceding_context.strip():
        ctx_lines.append(f"Preceding context: {task.preceding_context.strip()[:600]}")
    if task.native_text_hint.strip():
        hint = task.native_text_hint.strip()[:_NATIVE_PAGE_HINT_MAX_CHARS]
        ctx_lines.append(
            "\nNative text extracted from this page is provided below as a spelling and "
            "structure anchor. Preserve exact technical identifier names and URLs while "
            "enriching with visual/diagram/screenshot content:\n\n"
            f"--- native text ---\n{hint}\n--- end native text ---"
        )
    prompt = FULL_PAGE_HYBRID_PROMPT + "\n\n" + "\n".join(ctx_lines)

    data, mime = prepare_png_or_jpeg_for_api(task.image_bytes, task.mime_type)
    contents = [
        {
            "role": "user",
            "parts": [
                {"text": prompt},
                {"inline_data": {"mime_type": mime, "data": data}},
            ],
        }
    ]
    response = await asyncio.wait_for(
        client.models.generate_content(
            model=model,
            contents=contents,
            config=_build_generate_config(thinking_level),
        ),
        timeout=180,
    )
    text = getattr(response, "text", None)
    if not text:
        feedback = getattr(response, "prompt_feedback", None)
        reason = getattr(feedback, "block_reason", None) if feedback else None
        raise _InvalidModelOutput(f"Model returned no text (block reason: {reason})")

    payload = _parse_model_json(text)
    body = str(payload.get("markdown_body", "") or payload.get("markdown", "") or "").strip()
    return body


# ---------------------------------------------------------------------------
# Markdown formatting helpers for inline integration
# ---------------------------------------------------------------------------


def format_vision_block_standalone(image_md_ref: str | None, vision_md: str) -> str:
    """Format a standalone image reference paired with its vision extraction."""
    cleaned = vision_md.strip()
    if not cleaned:
        return f"![]({image_md_ref})" if image_md_ref else ""

    quoted_lines = [
        f"> {line}" if line.strip() else ">"
        for line in cleaned.splitlines()
    ]
    blockquote = "> **Visual content:**\n" + "\n".join(quoted_lines)
    if image_md_ref:
        return f"![]({image_md_ref})\n\n{blockquote}"
    return blockquote


def format_vision_for_table_cell(image_md_ref: str | None, vision_md: str) -> list[str]:
    """Convert multi-line vision Markdown into GFM-table-safe cell lines."""
    parts: list[str] = []
    if image_md_ref:
        parts.append(f"![]({image_md_ref})")

    cleaned = vision_md.strip()
    if not cleaned:
        return parts

    cell_lines: list[str] = ["*[Visual content]*:"]
    in_fence = False
    for raw_line in cleaned.splitlines():
        stripped = raw_line.strip()
        if stripped.startswith("```"):
            in_fence = not in_fence
            continue
        if not stripped:
            continue
        if in_fence:
            safe_code = stripped.replace("`", "'")
            cell_lines.append(f"`{safe_code}`")
            continue
        # Convert ATX headings into bold inline labels so they never break table cells
        h_match = re.match(r"^#{1,6}\s+(.*)$", stripped)
        if h_match:
            cell_lines.append(f"**{h_match.group(1).strip()}**")
            continue
        # Convert inner pipe-table separator lines; format pipe rows compactly
        if re.match(r"^\|?[\s:-]+\|[\s|:-]*$", stripped):
            continue
        if stripped.startswith("|") and stripped.endswith("|"):
            cols = [c.strip() for c in stripped.strip("|").split("|")]
            cell_lines.append("- " + " — ".join(c for c in cols if c))
            continue
        # Preserve 2-space list indentation while escaping literal pipes
        indent_len = len(raw_line) - len(raw_line.lstrip(" "))
        indent = " " * min(indent_len, 6) if stripped.startswith("-") else ""
        cell_lines.append(indent + stripped)

    parts.extend(cell_lines)
    return parts


@contextmanager
def libreoffice_to_pdf(office_path: Path) -> Iterator[Path]:
    """Convert an Office file (.docx/.pptx) to a temporary PDF using LibreOffice."""
    soffice = _find_soffice()
    if soffice is None:
        raise FileNotFoundError("LibreOffice (soffice) is not installed.")
    with tempfile.TemporaryDirectory(prefix="pdf_to_md_hybrid_") as tmp_dir:
        result = subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp_dir, str(office_path)],
            capture_output=True,
            text=True,
            timeout=600,
        )
        if result.returncode != 0:
            raise RuntimeError(
                f"LibreOffice conversion failed: {(result.stderr or result.stdout).strip()}"
            )
        pdf_path = Path(tmp_dir) / f"{office_path.stem}.pdf"
        if not pdf_path.is_file():
            raise RuntimeError("LibreOffice reported success but produced no PDF.")
        yield pdf_path
