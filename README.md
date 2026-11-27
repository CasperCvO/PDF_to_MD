# pdf-to-md

Batch-convert PDFs and office documents to Markdown while preserving tables, headers, and images.

## Features

- **Recursive scanning** — walks all directories and subdirectories of a source folder.
- **Customizable output** — outputs mirror the source directory layout into a specified directory (defaults to `Extracts/`).
- **Multi-format support** — converts PDFs, Word documents, Excel spreadsheets, and PowerPoint presentations.
- **Table, heading & list preservation** — geometry-driven PyMuPDF extraction for PDFs (multi-page split tables, stub vs. top-header tables, relative heading hierarchy, nested vector/Unicode bullet lists, two-column prose, inline hyperlinks) and [markitdown](https://pypi.org/project/markitdown/) / `openpyxl` for office documents.
- **Image extraction** — embedded raster images and genuine vector diagrams from PDFs are saved lazily to per-document `<name>_images/` folders using portable relative paths.
- **Hybrid vision screening & tandem extraction** *(optional, `--vision`)* — screens each PDF, `.docx`, or `.pptx` to measure visual page density (filtering out repeating header/footer logos and solid spacers) and routes automatically while keeping output as **one unified `.md` file** with `<!-- page: N -->` markers:
  - `0%` visual pages → pure local extraction (`0` API calls).
  - `<= 30%` visual pages (`--vision-low-threshold`) → **targeted image-in-context** vision: local extraction handles the document while only substantive images/diagrams are sent to Gemini alongside surrounding section/paragraph/table context and spliced inline.
  - `30%–70%` visual pages → **hybrid page-level** vision: native text pages stay local; visual candidate pages use full-page vision anchored by the native text layer.
  - `> 70%` visual pages (`--vision-high-threshold`) → **full-document** vision anchored by native text, merged into one `.md` file.
- **Wikilink header** — each Markdown file starts with `[[OriginalFile.ext]]` for easy back-referencing (e.g. in Obsidian).
- **PDF page markers** — every PDF page starts with a `<!-- page: N -->` marker (including empty pages) so downstream consumers can trace content back to its source page.
- **XLSX sheet/cell locators** — `.xlsx` workbooks are converted with openpyxl into `## Sheet: Name` sections with `<!-- sheet: "Name" range: A1:D4 -->` markers and row/column-addressed tables, so every value stays addressable as `Sheet!B7`.
- **DOCX page markers** *(approximate)* — `.docx` output gets `<!-- page: N -->` markers taken from Word's saved layout (`w:lastRenderedPageBreak`). When the file wasn't last saved by Word (e.g. LibreOffice, Google Docs, python-docx), it falls back to manual page breaks, page-break-before paragraphs, and page-starting section breaks — so positions are approximate. Disable with `docx_page_markers=False` / `--no-docx-page-markers`.
- **Fallback** — if structured PDF extraction fails for a PDF, plain text is extracted via PyMuPDF.
- **PDF Merger** — merge multiple PDF files in alphabetical order into a single PDF.
- **Vision-based slide extraction** *(add-on)* — extract structured Markdown from
  presentations containing native text, embedded UI screenshots, terminal dumps, and
  metric dashboards using the Gemini vision API (`gemini-4-hydrogen`). Slide-by-slide
  multimodal extraction with YAML frontmatter, Markdown table/code-block reconstruction,
  unified JSON export, concurrency, and exponential-backoff retries.

## Supported Formats

| Extension | Type                | Converter             |
| --------- | ------------------- | --------------------- |
| `.pdf`    | PDF                 | PyMuPDF               |
| `.docx`   | Word                | markitdown (+ page markers) |
| `.doc`    | Word (legacy)       | markitdown            |
| `.xlsx`   | Excel               | openpyxl (cell locators) |
| `.xls`    | Excel (legacy)      | markitdown            |
| `.pptx`   | PowerPoint          | markitdown            |
| `.ppt`    | PowerPoint (legacy) | markitdown            |
| `.pdf`    | Presentation (vision) | `convert_presentation_vision` |
| `.pptx`   | Presentation (vision) | `convert_presentation_vision` |
| `.png`    | Slide image (vision)  | `convert_presentation_vision` |
| `.jpg`/`.jpeg` | Slide image (vision) | `convert_presentation_vision` |

## Installation

Requires **Python 3.13+**.

### Use in Another Project via `uv`

Add this repository directly to your target project with `uv`:

```bash
uv add git+https://github.com/CasperCvO/PDF_to_MD.git
```

Or declare it in your target project's `pyproject.toml`:

```toml
[project]
dependencies = [
    "pdf-to-md @ git+https://github.com/CasperCvO/PDF_to_MD.git"
]
```

### Local Development Installation

```bash
# Clone the repository
git clone https://github.com/CasperCvO/PDF_to_MD.git
cd PDF_to_MD

# Install dependencies and editable package
uv sync
```

## Python API Usage

Import functions directly into your Python code:

```python
from pathlib import Path
from pdf_to_md import (
    convert_tree,
    convert_pdf,
    convert_office,
    screen_document,
    screen_tree,
    merge_pdfs,
)

# 1. Batch convert an entire directory tree (local deterministic extraction)
# Output defaults to <source_dir>/Extracts if output_root is omitted
counts = convert_tree("path/to/docs", output_root="path/to/output_markdown")
print(counts)  # e.g. {'pdf': 5, 'office': 2, 'skipped': 0}

# 2. Screen a document (or tree) to see if/how vision extraction applies
report = screen_document("path/to/architecture.pdf")
print(report.vision_ratio, report.recommended_strategy, report.candidate_unit_numbers)
# e.g. 0.03125 'targeted' [24]

# 3. Hybrid local + Gemini vision conversion (remains a single unified .md file)
convert_pdf(
    Path("path/to/architecture.pdf"),
    Path("path/to/architecture.md"),
    vision_mode="auto",              # "off" (default) | "auto" | "images" | "pages" | "full"
    vision_model="gemini-4-hydrogen", # default model
    thinking_level="low",            # None (model default) | "minimal" | "low" | "medium" | "high" | "xhigh"
    low_threshold=0.30,              # <= 30%: targeted image + context extraction
    high_threshold=0.70,             # > 70%: full-document vision; 30-70%: hybrid pages
    api_key="YOUR_GEMINI_KEY",       # or set GEMINI_API_KEY / .env (not needed if 0% visual)
)

convert_tree("path/to/docs", output_root="out", vision_mode="auto")

# 4. Convert a single Office file (.docx, .xlsx, .pptx, etc.)
convert_office(Path("path/to/presentation.pptx"), Path("path/to/output.md"))
convert_office(Path("path/to/spec.docx"), Path("path/to/spec.md"), vision_mode="auto")

# Disable the approximate page markers for .docx files
convert_office(Path("notes.docx"), Path("notes.md"), docx_page_markers=False)
convert_tree("path/to/docs", output_root="out", docx_page_markers=False)

# 5. Merge all PDFs in a folder into one PDF
merge_pdfs(Path("path/to/pdf_folder"), Path("path/to/merged.pdf"))

# 6. Vision-based slide extraction (Gemini API; add-on, does not affect 1-5)
from pdf_to_md import convert_presentation_vision

written = convert_presentation_vision(
    "path/to/presentation.pdf",   # .pdf, .pptx, .png/.jpg/.jpeg, or a directory
    "path/to/output_markdown",
    api_key="YOUR_GEMINI_API_KEY",  # or set the GEMINI_API_KEY env var
    model="gemini-4-hydrogen",
    thinking_level="low",  # None (model default) | "minimal" | "low" | "medium" | "high" | "xhigh"
    dpi=300,
    export_format="both",  # "md" (default) | "json" | "both"
    concurrency=4,
    max_retries=5,
    page_start=1,
    page_end=20,
)
print(written)  # per-slide .md files and/or a unified .json
```

## CLI Usage

When installed as a package, CLI commands are available directly:

```bash
# Batch convert documents (pure local)
pdf-to-md [source_dir_or_file] [-o output_dir]

# Screen documents for visual density without converting or calling the API
pdf-to-md [source_dir_or_file] --screen-only

# Hybrid local + vision extraction into single unified .md files
pdf-to-md [source_dir_or_file] --vision [-o output_dir]
pdf-to-md [source_dir_or_file] --vision-mode auto --vision-low-threshold 0.30 --vision-high-threshold 0.70

# Merge PDFs
merge-pdfs [source_dir] [-o output.pdf]

# Per-slide vision extraction add-on (one .md per slide + YAML frontmatter)
pdf-to-md-vision --input-path slides.pdf --output-dir out --export-format both
```

You can also run directly from the repo root:

```bash
python main.py [source_dir] [-o output_dir]
python merge_pdfs.py [source_dir] [-o output.pdf]
```

### Options

- **`source_dir`** — Root directory (or single file) to scan. Defaults to the current directory (`.`) if omitted.
- **`-o`, `--output-dir`** — Directory where converted Markdown files will be written. Defaults to `<source_dir>/Extracts`.
- **`--docx-page-markers` / `--no-docx-page-markers`** — Enable/disable approximate `<!-- page: N -->` markers in `.docx` output. Enabled by default.
- **`--vision`** — Shorthand for `--vision-mode auto` (hybrid screening + Gemini vision while keeping one `.md` file per source document).
- **`--vision-mode {off,auto,images,pages,full}`** — Vision strategy (`off` by default).
- **`--vision-low-threshold FLOAT`** / **`--vision-high-threshold FLOAT`** — Visual page ratio boundaries for `auto` routing (`0.30` and `0.70` by default).
- **`--screen-only`** — Screen supported files and log each document's visual page ratio and recommended strategy without writing files or calling the API.
- **`--api-key`**, **`--vision-model`** (default: `gemini-4-hydrogen`), **`--vision-thinking-level`** (`minimal`, `low`, `medium`, `high`, `xhigh`; default: model default, `gemini-4-hydrogen` supports `low`, `medium`, `high`, `xhigh`), **`--vision-dpi`**, **`--vision-concurrency`**, **`--vision-max-retries`** — Gemini API configuration for hybrid vision runs.

### Directory Tree Example

```bash
pdf-to-md "C:\Users\me\Documents\Reports" -o "C:\Users\me\Documents\Reports_MD"
```

Produces:

```
Reports_MD/
  subdir/
    Report.md
    Report_images/
      Report-p001-img01.png
    Meeting_Notes.md
  AnnualReview.md
  AnnualReview_images/
    AnnualReview-p001-img01.png
  Budget.md
```

Each generated `.md` file starts with a Wikilink to the original document:

```markdown
[[Report.pdf]]

# Report Title
...
```

PDF output also marks the start of every page with a `<!-- page: N -->`
comment (empty pages get a marker too, so the page sequence has no gaps):

```markdown
[[Report.pdf]]

<!-- page: 1 -->

# Report Title

<!-- page: 2 -->

<!-- page: 3 -->

Page-three content
```

`.xlsx` workbooks produce one `## Sheet: Name` section per worksheet, with a
sheet/range locator comment and a table whose first column holds the Excel row
number and whose header holds the column letters (only non-empty rows and
columns are included):

```markdown
[[Budget.xlsx]]

## Sheet: Data

<!-- sheet: "Data" range: A1:D4 -->

| Row | A | B | D |
| --- | --- | --- | --- |
| 1 | Name | Amount |  |
| 2 | Alice | 100 |  |
| 4 | Bob\|Smith | =SUM(B2:B2) | x |

## Sheet: Notes

_(empty sheet)_
```

`.docx` output gets approximate `<!-- page: N -->` markers (Word's saved
layout when present, else manual/section page breaks):

```markdown
[[Notes.docx]]

<!-- page: 1 -->

First page text

<!-- page: 2 -->

# Chapter Two
```

## Vision-Based Slide Extraction (Add-on)

`convert_presentation_vision` / `pdf-to-md-vision` processes presentations and
pre-rendered slide images that mix **native text, embedded UI screenshots,
terminal dumps, and metric dashboards** into machine-readable Markdown for an
LLM context repository. It is a pure add-on: the local converters above are
unchanged and still work without an API key.

### How it works

1. **Input handling** — `.pdf` pages are rendered to high-resolution images
   (300 DPI default) with PyMuPDF. `.pptx` decks are converted headlessly via
   LibreOffice CLI when available; without LibreOffice, native slide text and
   tables are extracted with `python-pptx` (no vision, flagged in the output).
   Raw `.png`/`.jpg`/`.jpeg` slide images are validated (mime type, resolution,
   color depth) before being sent.
2. **Slide-by-slide extraction** — each slide image is sent to
   `gemini-4-hydrogen` together with a structured extraction prompt
   (JSON mode: `topic`, `technologies`, `markdown_body`). Native PDF text is
   included as a transcription hint.
3. **Structured output** — one self-contained Markdown file per slide with YAML
   frontmatter (`document_source`, `slide_number`, `topic`, `technologies`,
   `vision_analyzed`), containing headings, bullets, GitHub-flavored tables,
   numeric chart summaries, and `bash`/`text` code blocks for terminal dumps.
4. **Resilience** — configurable concurrency (`--concurrency`), exponential
   backoff with jitter for HTTP 429/5xx, network errors, and malformed model
   output (`--max-retries`). Images are automatically downscaled/re-encoded to
   stay within Gemini's inline payload limits.

### CLI

```bash
pdf-to-md-vision --input-path deck.pdf --output-dir out_md
pdf-to-md-vision --input-path deck.pptx --export-format both
pdf-to-md-vision --input-path slide.png --api-key $GEMINI_API_KEY
pdf-to-md-vision --input-path presentations_dir --page-start 1 --page-end 10
```

| Option | Description |
| ------ | ----------- |
| `input_path` / `--input-path` | Presentation, slide image, or directory of files to process. |
| `-o`, `--output-dir` | Output directory (default: `<input_path>/vision_extracts`). |
| `--dpi` | Render resolution for PDF/PPTX pages (default: 300). |
| `--export-format` | `md` (per-slide Markdown, default), `json` (unified array), or `both`. |
| `--api-key` | Gemini API key (default: `GEMINI_API_KEY` environment variable, or a `.env` file in the working directory). |
| `--model` | Gemini model id (default: `gemini-4-hydrogen`). |
| `--thinking-level` | Gemini thinking level: `minimal`, `low`, `medium`, `high`, or `xhigh` (default: model default; `gemini-4-hydrogen` supports `low`/`medium`/`high`/`xhigh`). |
| `--concurrency` | Max simultaneous API requests (default: 4). |
| `--max-retries` | Retries per slide for transient failures (default: 5). |
| `--page-start`, `--page-end` | 1-based inclusive slide range to process. |

### Example output (`deck_slide_018.md`)

````markdown
---
document_source: "Apoyo_Prestatiemeting.pdf"
slide_number: 18
topic: "5. Gebruik van de database-connectiepool"
technologies: ["Pega", "Alfresco", "PostgreSQL", "PGPool"]
vision_analyzed: true
---

# 5. Gebruik van de database-connectiepool

The PEGA and Alfresco applications connect to the PostgreSQL database via an Azure Load Balancer routing to individual PGPool HA nodes.

## Pega Database Metrics
- **PGPool connection pool:** 800 (2 x 400)
- **PGPool connections in use:** ~50 (2 x ~25)
- **Active PGPool -> PostgreSQL connections:** ~130

### Bash Commands & Terminal Outputs
```bash
[postgres@cizprd1ppsql01 ~]$ ps -ef | grep pgpool | grep "ciz_pega_prod" | wc -l
18
```
````

The unified JSON export (`deck.json`) contains the same data as an array of
objects (`document_source`, `slide_number`, `topic`, `technologies`,
`markdown`, `vision_analyzed`) for programmatic ingestion.

> **Note:** a Gemini API key is required for vision analysis. Get one from
> [Google AI Studio](https://aistudio.google.com/apikey). LibreOffice is only
> needed for `.pptx` rendering; without it, `.pptx` falls back to native text.

## Project Structure

```
PDF_to_MD/
├── pyproject.toml              # Build backend, package metadata & CLI entry points
├── src/
│   └── pdf_to_md/
│       ├── __init__.py         # Package exports (convert_tree, convert_pdf, screen_document, etc.)
│       ├── main.py             # CLI and directory batch conversion (convert_tree, screen_tree)
│       ├── pdf_converter.py    # Single PDF conversion (convert_pdf)
│       ├── office_converter.py # Single Office conversion (convert_office)
│       ├── hybrid_vision.py    # Document vision screening & single-file hybrid orchestration
│       ├── vision_converter.py # Per-slide Gemini vision extraction (convert_presentation_vision)
│       └── merge_pdfs.py       # PDF merger (merge_pdfs)
├── tests/                      # pytest suite (PDF, Office, hybrid vision, slide vision, and PDF merger)
├── main.py                     # Convenience root launcher
└── merge_pdfs.py               # Convenience root launcher
```
