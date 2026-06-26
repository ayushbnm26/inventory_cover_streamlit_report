"""Run the B2B Dispatch Tracker pipeline from the project checkout."""

from __future__ import annotations

from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))

from inventory_cover.cli import add_b2b_dispatch_args, run_b2b_dispatch_from_args  # noqa: E402
import argparse  # noqa: E402


def main() -> int:
    parser = argparse.ArgumentParser(description="Run B2B Dispatch Tracker pipeline.")
    add_b2b_dispatch_args(parser)
    args = parser.parse_args()
    return run_b2b_dispatch_from_args(args)


if __name__ == "__main__":
    raise SystemExit(main())
