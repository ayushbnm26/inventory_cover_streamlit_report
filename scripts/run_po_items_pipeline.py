"""Run the Amazon PO Items consolidation pipeline from the project checkout."""

from __future__ import annotations

from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from inventory_cover.cli import add_po_items_args, run_po_items_from_args  # noqa: E402
import argparse  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(description="Run PO Items consolidation.")
    add_po_items_args(parser)
    args = parser.parse_args()
    return run_po_items_from_args(args)


if __name__ == "__main__":
    raise SystemExit(main())
