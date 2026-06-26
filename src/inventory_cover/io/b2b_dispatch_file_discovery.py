"""File discovery for B2B Dispatch Tracker workbooks."""

from __future__ import annotations

from pathlib import Path

from inventory_cover.config import B2BDispatchPipelineConfig
from inventory_cover.exceptions import CatastrophicPipelineError


def discover_b2b_dispatch_files(config: B2BDispatchPipelineConfig) -> list[Path]:
    """Discover dispatch tracker Excel files in the configured input folder."""

    input_dir = Path(config.input_dir)
    if not input_dir.exists():
        raise CatastrophicPipelineError(f"Input folder does not exist: {input_dir}")
    if not input_dir.is_dir():
        raise CatastrophicPipelineError(f"Input path is not a folder: {input_dir}")

    files = sorted(
        path
        for path in input_dir.iterdir()
        if path.is_file()
        and path.suffix.lower() == ".xlsx"
        and not path.name.startswith("~$")
    )
    if not files:
        raise CatastrophicPipelineError(f"No .xlsx files found in {input_dir}")
    if len(files) > 1 and not config.allow_multiple_files:
        raise CatastrophicPipelineError(
            f"Found {len(files)} Excel files in {input_dir}. "
            "Use --allow-multiple-files only after confirming this is intentional."
        )
    return files
