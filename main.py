from __future__ import annotations

import sys
from pathlib import Path

# Allow direct execution from repo root without prior installation
sys.path.insert(0, str(Path(__file__).parent / "src"))

from pdf_to_md.main import build_output_path, convert_tree, main  # noqa: E402

if __name__ == "__main__":
    main()
