from __future__ import annotations

from datetime import date
from pathlib import Path

from core.constants import DEFAULT_OUTPUT_EXTENSION, DEFAULT_OUTPUT_PREFIX


def build_output_filename(run_date: date | None = None) -> str:
    run_date = run_date or date.today()
    return f"{DEFAULT_OUTPUT_PREFIX}{run_date.strftime('%Y%m%d')}{DEFAULT_OUTPUT_EXTENSION}"


def build_output_path(output_dir: str | Path, filename: str | None = None) -> Path:
    output_dir = Path(output_dir)
    filename = filename or build_output_filename()
    return output_dir / filename


def build_renamed_output_path(path: str | Path) -> Path:
    path = Path(path)
    base = path.stem
    suffix = path.suffix

    counter = 2
    candidate = path
    while candidate.exists():
        candidate = path.with_name(f"{base} ({counter}){suffix}")
        counter += 1
    return candidate
