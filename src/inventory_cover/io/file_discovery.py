"""File discovery for PO Items workbooks."""

from __future__ import annotations

from pathlib import Path
import re

from inventory_cover.config import PipelineConfig
from inventory_cover.exceptions import CatastrophicPipelineError


def discover_po_item_files(config: PipelineConfig) -> tuple[list[Path], list[str]]:
    """Discover candidate PO Items Excel files in the configured input folder."""

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
    files = sorted(
        files,
        key=_file_sort_key,
    )
    if not files:
        raise CatastrophicPipelineError(f"No .xlsx files found in {input_dir}")

    warnings: list[str] = []
    suspicious = [path.name for path in files if not is_likely_po_items_name(path.name)]
    if suspicious:
        warnings.append(
            "Some Excel files have non-standard names and will be validated by headers: "
            + ", ".join(suspicious)
        )

    if len(files) < config.min_files:
        if len(files) == 1 and config.allow_single_file:
            warnings.append("Only one PO Items file found; processing because allow_single_file=True.")
        else:
            raise CatastrophicPipelineError(
                f"Found {len(files)} Excel file(s), but min_files={config.min_files}. "
                "Use --allow-single-file for an intentional one-file run."
            )

    if len(files) > config.max_files and not config.allow_more_than_max_files:
        raise CatastrophicPipelineError(
            f"Found {len(files)} Excel files, which exceeds max_files={config.max_files}. "
            "Raise --max-files or use --allow-more-than-max-files after reviewing the folder."
        )
    if len(files) > config.max_files:
        warnings.append(
            f"Found {len(files)} files, exceeding max_files={config.max_files}; processing because allowed."
        )

    return files, warnings


def is_likely_po_items_name(file_name: str) -> bool:
    normalized = re.sub(r"[^a-z0-9]+", "", file_name.lower())
    return (
        "purchaseorderitems" in normalized
        or "poitems" in normalized
        or ("purchase" in normalized and "order" in normalized and "item" in normalized)
    )


def _file_sort_key(path: Path) -> tuple[str, int, str]:
    suffix_match = re.search(r"\((\d+)\)$", path.stem)
    sequence = int(suffix_match.group(1)) if suffix_match else 0
    base_name = re.sub(r"\s*\(\d+\)$", "", path.stem).lower()
    return (base_name, sequence, path.name.lower())
