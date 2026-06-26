"""Create a pasteable project context dump for LLM review."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
import platform
import re
import sys


PROJECT_ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = PROJECT_ROOT / "context_dumps"

EXCLUDED_DIR_NAMES = {
    ".git",
    ".venv",
    "venv",
    "env",
    "__pycache__",
    ".pytest_cache",
    ".mypy_cache",
    ".ruff_cache",
    "runs",
    "context_dumps",
}

EXCLUDED_PREFIXES = {
    Path("data/incoming"),
    Path("data/processed"),
    Path("data/archive"),
}

EXCLUDED_SUFFIXES = {
    ".xlsx",
    ".xlsm",
    ".xls",
    ".log",
    ".pyc",
    ".pyo",
}

INCLUDED_CONFIG_SUFFIXES = {".toml", ".yaml", ".yml", ".json"}
INCLUDED_SPECIAL_FILES = {"README.md", "PROJECT_AMBITION.md", ".gitignore"}


def main() -> int:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = OUTPUT_DIR / f"project_context_dump_{timestamp}.txt"
    content = build_dump()
    output_path.write_text(content, encoding="utf-8")
    print(f"Context dump created: {output_path}")
    return 0


def build_dump() -> str:
    lines: list[str] = []
    lines.append("# Inventory Cover Project Context Dump")
    lines.append("")
    lines.append(f"Generated: {datetime.now().isoformat(timespec='seconds')}")
    lines.append(f"Project root: {PROJECT_ROOT}")
    lines.append(f"Python executable: {sys.executable}")
    lines.append(f"Python version: {platform.python_version()}")
    lines.append("")
    lines.append("## Command Examples")
    lines.append("")
    lines.append("python scripts/run_po_items_pipeline.py --input-dir data/incoming/po_items")
    lines.append("python -m pytest")
    lines.append("python scripts/dump_py_to_txt.py")
    lines.append("")
    lines.append("## Folder Tree")
    lines.append("")
    lines.extend(render_tree())
    lines.append("")

    for path in files_to_include():
        rel = path.relative_to(PROJECT_ROOT).as_posix()
        lines.append("")
        lines.append(f"### FILE: {rel}")
        lines.append("")
        lines.append(read_text_safely(path))
    return "\n".join(lines).rstrip() + "\n"


def render_tree() -> list[str]:
    lines: list[str] = [PROJECT_ROOT.name + "/"]
    for path in sorted(PROJECT_ROOT.rglob("*")):
        if should_exclude(path):
            continue
        rel = path.relative_to(PROJECT_ROOT)
        depth = len(rel.parts)
        prefix = "  " * depth
        suffix = "/" if path.is_dir() else ""
        lines.append(f"{prefix}{path.name}{suffix}")
    return lines


def files_to_include() -> list[Path]:
    candidates: list[Path] = []
    for path in sorted(PROJECT_ROOT.rglob("*")):
        if path.is_dir() or should_exclude(path):
            continue
        if path.name in INCLUDED_SPECIAL_FILES or path.suffix == ".py" or path.suffix in INCLUDED_CONFIG_SUFFIXES:
            candidates.append(path)
        elif path.name.startswith(".env"):
            candidates.append(path)
    return candidates


def should_exclude(path: Path) -> bool:
    try:
        rel = path.relative_to(PROJECT_ROOT)
    except ValueError:
        return True
    if any(part in EXCLUDED_DIR_NAMES for part in rel.parts):
        return True
    if any(rel == prefix or rel.is_relative_to(prefix) for prefix in EXCLUDED_PREFIXES):
        return True
    if path.is_file() and path.suffix.lower() in EXCLUDED_SUFFIXES:
        return True
    return False


def read_text_safely(path: Path) -> str:
    try:
        text = path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        text = path.read_text(encoding="utf-8", errors="replace")
    if path.name.startswith(".env"):
        return redact_env(text)
    return text.rstrip()


def redact_env(text: str) -> str:
    redacted_lines = []
    for line in text.splitlines():
        if not line.strip() or line.lstrip().startswith("#") or "=" not in line:
            redacted_lines.append(line)
            continue
        key, _value = line.split("=", 1)
        if re.search(r"(key|secret|token|password|credential)", key, re.IGNORECASE):
            redacted_lines.append(f"{key}=<REDACTED>")
        else:
            redacted_lines.append(line)
    return "\n".join(redacted_lines)


if __name__ == "__main__":
    raise SystemExit(main())
