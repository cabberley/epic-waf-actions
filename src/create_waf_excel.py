"""Create an Excel workbook summarizing Epic WAF YAML files."""

from __future__ import annotations

import argparse
import json
from datetime import datetime, date
from pathlib import Path
from typing import Any, Dict, Iterable, List, cast


try:  # openpyxl is a runtime dependency; guide the user if missing.
    from openpyxl import Workbook
    from openpyxl.styles import Alignment
    from openpyxl.worksheet.worksheet import Worksheet
except ModuleNotFoundError as exc:  # pragma: no cover - import guard
    raise SystemExit("openpyxl is required. Install it with 'pip install openpyxl'.") from exc


try:  # PyYAML parses the WAF definition files.
    import yaml
except ModuleNotFoundError as exc:  # pragma: no cover - import guard
    raise SystemExit("PyYAML is required. Install it with 'pip install pyyaml'.") from exc


DEFAULT_ROOT = Path("")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--root",
        type=Path,
        default=DEFAULT_ROOT,
        help="Base directory that contains the WAF folder and output location",
    )
    parser.add_argument(
        "--waf-dir",
        type=Path,
        default=Path("WAF"),
        help="Relative path under --root where WAF YAML files are stored (default: WAF)",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("excel"),
        help="Directory that will receive the generated workbook (default: --root)",
    )
    parser.add_argument(
        "--pattern",
        type=str,
        default="*.yml,*.yaml",
        help="Comma-separated glob patterns for YAML files (default: *.yml,*.yaml)",
    )
    return parser.parse_args()


def iter_yaml_files(waf_dir: Path, patterns: Iterable[str]) -> List[Path]:
    files: List[Path] = []
    for pattern in patterns:
        files.extend(waf_dir.glob(pattern.strip()))
    unique_files = [path.resolve() for path in files if path.is_file()]
    return sorted(unique_files, key=lambda path: path.name.lower())


def load_documents(files: Iterable[Path]) -> List[Dict[str, Any]]:
    documents: List[Dict[str, Any]] = []
    for file_path in files:
        data = yaml.safe_load(file_path.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            raise ValueError(f"File {file_path} does not contain a top-level mapping")
        if is_falsey(data.get("active")):
            continue
        documents.append(dict(data))
    return documents


def is_falsey(value: Any) -> bool:
    if value is False:
        return True
    if isinstance(value, str):
        return value.strip().lower() in {"false", "0", "no"}
    return False


def collect_columns(documents: Iterable[Dict[str, Any]]) -> List[str]:
    keys: List[str] = []
    seen = set()
    for document in documents:
        for key in document.keys():
            if key not in seen:
                seen.add(key)
                keys.append(key)
    return keys


def serialize_value(value: Any, column: str | None = None) -> Any:
    if column == "labels" and isinstance(value, list):
        entries = [entry.strip() for entry in value if isinstance(entry, str) and entry.strip()]
        return "\n".join(entries)
    if value is None:
        return ""
    if isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, list) and not value:
        return "[]"
    if isinstance(value, dict) and not value:
        return "{}"
    try:
        return json.dumps(value, ensure_ascii=False)
    except TypeError:
        return str(value)


def build_workbook(columns: List[str], documents: List[Dict[str, Any]]) -> Workbook:
    workbook = Workbook()
    sheet = cast(Worksheet, workbook.active)
    sheet.title = "WAF"
    sheet.append(columns)
    for document in documents:
        row = [serialize_value(document.get(column, ""), column) for column in columns]
        sheet.append(row)

    if "labels" in columns:
        labels_index = columns.index("labels") + 1  # openpyxl is 1-based
        for cell in sheet.iter_rows(min_col=labels_index, max_col=labels_index):
            cell[0].alignment = Alignment(wrap_text=True)

    return workbook


def main() -> None:
    args = parse_args()

    root = args.root.expanduser().resolve()
    waf_dir = (root / args.waf_dir).resolve()
    output_dir = (args.output_dir or root).expanduser().resolve()

    if not waf_dir.is_dir():
        raise SystemExit(f"WAF directory not found: {waf_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)

    patterns = [pattern.strip() for pattern in args.pattern.split(",") if pattern.strip()]
    files = iter_yaml_files(waf_dir, patterns)
    if not files:
        raise SystemExit(f"No YAML files found in {waf_dir}")

    documents = load_documents(files)
    if not documents:
        raise SystemExit("No active WAF files to export. All entries have active=false.")
    columns = collect_columns(documents)
    workbook = build_workbook(columns, documents)

    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    output_path = output_dir / f"WAF-version-{timestamp}.xlsx"
    workbook.save(output_path)
    print(f"Workbook created: {output_path}")


if __name__ == "__main__":
    main()
