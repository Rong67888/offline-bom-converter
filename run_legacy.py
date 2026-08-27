from __future__ import annotations

import argparse
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent
SRC = PROJECT_ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from bom_converter.gui import main  # noqa: E402


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="BOM 转换器原界面回退入口")
    parser.add_argument("--mode", choices=("quick", "audit"), default="audit")
    args = parser.parse_args()
    main(args.mode)
