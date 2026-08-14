from __future__ import annotations

import sys
from pathlib import Path

# Allow direct execution from repo root without prior installation
sys.path.insert(0, str(Path(__file__).parent / "src"))

from pdf_to_md.pdf_converter import SUPPORTED_EXTENSIONS, convert_pdf  # noqa: E402

__all__ = ["SUPPORTED_EXTENSIONS", "convert_pdf"]
