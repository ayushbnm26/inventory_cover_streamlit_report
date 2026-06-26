"""Checkout shim for running the src-layout package without installation."""

from __future__ import annotations

from pathlib import Path

_SRC_PACKAGE = Path(__file__).resolve().parent.parent / "src" / "inventory_cover"
if _SRC_PACKAGE.exists():
    __path__.append(str(_SRC_PACKAGE))

__version__ = "0.1.0"
