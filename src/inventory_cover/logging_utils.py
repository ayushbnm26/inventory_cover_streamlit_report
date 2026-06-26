"""Logging and metadata helpers."""

from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Any

from inventory_cover.schemas import json_safe


def setup_run_logger(
    log_file: Path,
    log_level: str = "INFO",
    logger_name: str = "inventory_cover.po_items",
) -> logging.Logger:
    """Create a run-scoped logger with file and console handlers."""

    logger = logging.getLogger(logger_name)
    logger.handlers.clear()
    logger.setLevel(getattr(logging, log_level.upper(), logging.INFO))
    logger.propagate = False

    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    log_file.parent.mkdir(parents=True, exist_ok=True)

    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    return logger


def write_json_file(path: Path, payload: dict[str, Any]) -> None:
    """Write JSON atomically enough for local reporting runs."""

    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_suffix(path.suffix + ".tmp")
    temp_path.write_text(
        json.dumps(json_safe(payload), indent=2, sort_keys=True),
        encoding="utf-8",
    )
    temp_path.replace(path)
