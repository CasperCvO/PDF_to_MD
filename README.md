# pdf-to-md

Batch-convert PDFs and office documents to Markdown while preserving tables, headers, and images.

## Features

- **Recursive scanning** — walks all directories and subdirectories of a source folder.
- **Customizable output** — outputs mirror the source directory layout into a specified directory (defaults to `Extracts/`).
- **Multi-format support** — converts PDFs, Word documents, Excel spreadsheets, and PowerPoint presentations.
- **Table & header preservation** — uses [pymupdf4llm](https://pypi.org/project/pymupdf4llm/) for high-fidelity PDF conversion and [markitdown](https://pypi.org/project/markitdown/) for office documents.
- **Image extraction** — embedded images from PDFs are saved to per-document `<name>_images/` folders and referenced in the Markdown.
- **Wikilink header** — each Markdown file starts with `[[OriginalFile.ext]]` for easy back-referencing (e.g. in Obsidian).
- **Fallback** — if pymupdf4llm fails for a PDF, plain text is extracted via PyMuPDF.
- **PDF Merger** — merge multiple PDF files in alphabetical order into a single PDF.

## Supported Formats

| Extension | Type                | Converter             |
| --------- | ------------------- | --------------------- |
| `.pdf`    | PDF                 | pymupdf4llm / PyMuPDF |
| `.docx`   | Word                | markitdown            |
| `.doc`    | Word (legacy)       | markitdown            |
| `.xlsx`   | Excel               | markitdown            |
| `.xls`    | Excel (legacy)      | markitdown            |
| `.pptx`   | PowerPoint          | markitdown            |
| `.ppt`    | PowerPoint (legacy) | markitdown            |

## Installation

Requires **Python 3.13+**.

### Use in Another Project via `uv`

Add this repository directly to your target project with `uv`:

```bash
uv add git+https://github.com/<username>/PDF_to_MD.git
```

Or declare it in your target project's `pyproject.toml`:

```toml
[project]
dependencies = [
    "pdf-to-md @ git+https://github.com/<username>/PDF_to_MD.git"
]
```

### Local Development Installation

```bash
# Clone the repository
git clone <repo-url>
cd PDF_to_MD

# Install dependencies and editable package
uv sync
```

## Python API Usage

Import functions directly into your Python code:

```python
from pathlib import Path
from pdf_to_md import convert_tree, convert_pdf, convert_office, merge_pdfs

# 1. Batch convert an entire directory tree
# Output defaults to <source_dir>/Extracts if output_root is omitted
counts = convert_tree("path/to/docs", output_root="path/to/output_markdown")
print(counts)  # e.g. {'pdf': 5, 'office': 2, 'skipped': 0}

# 2. Convert a single PDF file
convert_pdf(Path("path/to/input.pdf"), Path("path/to/output.md"))

# 3. Convert a single Office file (.docx, .xlsx, .pptx, etc.)
convert_office(Path("path/to/presentation.pptx"), Path("path/to/output.md"))

# 4. Merge all PDFs in a folder into one PDF
merge_pdfs(Path("path/to/pdf_folder"), Path("path/to/merged.pdf"))
```

## CLI Usage

When installed as a package, CLI commands are available directly:

```bash
# Batch convert documents
pdf-to-md [source_dir] [-o output_dir]

# Merge PDFs
merge-pdfs [source_dir] [-o output.pdf]
```

You can also run directly from the repo root:

```bash
python main.py [source_dir] [-o output_dir]
python merge_pdfs.py [source_dir] [-o output.pdf]
```

### Options

- **`source_dir`** — Root directory to scan. Defaults to the current directory (`.`) if omitted.
- **`-o`, `--output-dir`** — Directory where converted Markdown files will be written. Defaults to `<source_dir>/Extracts`.

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
      img-0001.png
    Meeting_Notes.md
  AnnualReview.md
  AnnualReview_images/
    img-0001.png
  Budget.md
```

Each generated `.md` file starts with a Wikilink to the original document:

```markdown
[[Report.pdf]]

# Report Title
...
```

## Project Structure

```
PDF_to_MD/
├── pyproject.toml              # Build backend, package metadata & CLI entry points
├── src/
│   └── pdf_to_md/
│       ├── __init__.py         # Package exports (convert_tree, convert_pdf, etc.)
│       ├── main.py             # CLI and directory batch conversion (convert_tree)
│       ├── pdf_converter.py    # Single PDF conversion (convert_pdf)
│       ├── office_converter.py # Single Office conversion (convert_office)
│       └── merge_pdfs.py       # PDF merger (merge_pdfs)
├── main.py                     # Convenience root launcher
└── merge_pdfs.py               # Convenience root launcher
```
