"""Validate Epic WAF YAML files against the provided template."""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Set, Tuple


try:  # PyYAML is not part of the stdlib, give a helpful message if missing.
    import yaml
except ModuleNotFoundError as exc:  # pragma: no cover - import guard
    sys.stderr.write("PyYAML is required. Install it with 'pip install pyyaml'.\n")
    raise exc


STATUS_PATTERN = re.compile(r"^([A-Za-z0-9_]+)\s*:\s*(mandatory|optional)\s*$", re.IGNORECASE)
OPTIONAL_WARN_KEYS = {"labels", "specs", "epic_resources"}


def parse_template(template_path: Path) -> Tuple[Dict[str, str], Dict[str, Dict[str, str]]]:
    """Return dictionaries describing required keys based on the template file."""

    if not template_path.is_file():
        raise FileNotFoundError(f"Template not found: {template_path}")

    top_level: Dict[str, str] = {}
    nested: Dict[str, Dict[str, str]] = {}
    current_parent: str | None = None
    current_list_parent: str | None = None

    for raw_line in template_path.read_text(encoding="utf-8").splitlines():
        if not raw_line.strip():
            continue

        indent = len(raw_line) - len(raw_line.lstrip())
        stripped = raw_line.strip()

        match = STATUS_PATTERN.match(stripped)
        if indent == 0 and match:
            key, status = match.group(1), match.group(2).lower()
            top_level[key] = status
            current_parent = key
            current_list_parent = None
            continue

        parent = current_list_parent or current_parent
        if parent is None:
            continue

        if stripped.startswith("- "):
            inner = stripped[2:].strip()
            inner_match = STATUS_PATTERN.match(inner)
            if inner_match:
                nested.setdefault(parent, {})[inner_match.group(1)] = inner_match.group(2).lower()
                current_list_parent = parent
            continue

        nested_match = STATUS_PATTERN.match(stripped)
        if nested_match:
            nested.setdefault(parent, {})[nested_match.group(1)] = nested_match.group(2).lower()

    return top_level, nested


def load_yaml(path: Path) -> Any:
    with path.open("r", encoding="utf-8") as handle:
        return yaml.safe_load(handle)


def has_value(value: Any) -> bool:
    if value is None:
        return False
    if isinstance(value, str):
        return bool(value.strip())
    if isinstance(value, (list, tuple, set, dict)):
        return bool(value)
    return True


def allows_blank_value(key: str, value: Any) -> bool:
    if key != "end_date":
        return False
    if value is None:
        return True
    if isinstance(value, str):
        stripped = value.strip().lower()
        return stripped == "" or stripped == "null"
    return False


def load_allowed_labels(labels_path: Path) -> Set[str]:
    if not labels_path.is_file():
        raise FileNotFoundError(f"Labels file not found: {labels_path}")

    data = load_yaml(labels_path)
    if isinstance(data, dict):
        raw_labels = data.get("labels")
    else:
        raw_labels = data

    if not isinstance(raw_labels, list):
        raise ValueError("labels.yml must contain a list (optionally under 'labels')")

    allowed: Set[str] = set()
    for entry in raw_labels:
        if isinstance(entry, str) and entry.strip():
            allowed.add(entry.strip())
    if not allowed:
        raise ValueError("labels.yml does not define any labels")
    return allowed


def load_epic_resources(resources_path: Path) -> Set[str]:
    if not resources_path.is_file():
        raise FileNotFoundError(f"Epic resources file not found: {resources_path}")

    data = load_yaml(resources_path)
    if data is None:
        raise ValueError("epic_resources.yml is empty")

    resources: Set[str] = set()

    def _collect(node: Any) -> None:
        if isinstance(node, str):
            cleaned = node.strip()
            if cleaned:
                resources.add(cleaned)
            return
        if isinstance(node, dict):
            for key, value in node.items():
                _collect(key)
                _collect(value)
            return
        if isinstance(node, list):
            for entry in node:
                _collect(entry)

    _collect(data)
    if not resources:
        raise ValueError("epic_resources.yml does not define any resources")
    return resources


def load_validations(validations_path: Path, required_keys: List[str]) -> Dict[str, Set[str]]:
    if not validations_path.is_file():
        raise FileNotFoundError(f"Validations file not found: {validations_path}")

    data = load_yaml(validations_path)
    if not isinstance(data, dict):
        raise ValueError("validations.yml must be a mapping of keys to allowed values")

    validations: Dict[str, Set[str]] = {}
    for key, values in data.items():
        if not isinstance(values, list) or not values:
            raise ValueError(f"Validation list for '{key}' must be a non-empty list")
        normalized = {normalize_scalar(entry) for entry in values if has_value(entry) or entry is None}
        if not normalized:
            raise ValueError(f"Validation list for '{key}' does not contain usable entries")
        validations[key] = normalized

    for req_key in required_keys:
        if req_key not in validations:
            raise ValueError(f"validations.yml must define allowed values for '{req_key}'")

    return validations


def normalize_scalar(value: Any) -> str:
    if isinstance(value, str):
        return value.strip().lower()
    if isinstance(value, bool):
        return str(value).lower()
    if isinstance(value, (int, float)):
        return str(value)
    if value is None:
        return "null"
    return str(value).strip().lower()


def validate_string_list(
    field_name: str,
    value: Any,
    allowed_values: Set[str],
    source_description: str,
) -> List[str]:
    if value is None:
        return [f"Missing or empty mandatory key '{field_name}'"]
    if not isinstance(value, list):
        return [f"Key '{field_name}' must be a list"]

    errors: List[str] = []
    for idx, entry in enumerate(value):
        if not isinstance(entry, str) or not entry.strip():
            errors.append(f"'{field_name}' entry {idx} must be a non-empty string")
            continue
        normalized = entry.strip()
        if normalized not in allowed_values:
            errors.append(
                f"'{field_name}' entry '{normalized}' (index {idx}) is not defined in {source_description}"
            )
    return errors


def validate_allowed_value(key: str, value: Any, allowed_values: Set[str]) -> List[str]:
    if value is None:
        return []  # handled by mandatory validation elsewhere
    if isinstance(value, (list, tuple, set, dict)):
        return [f"Value for '{key}' must be a scalar"]

    normalized = normalize_scalar(value)
    if normalized not in allowed_values:
        pretty = ", ".join(sorted(allowed_values))
        return [f"Value '{value}' for '{key}' must be one of: {pretty}"]
    return []


def validate_document(
    content: Any,
    top_level_rules: Dict[str, str],
    nested_rules: Dict[str, Dict[str, str]],
) -> Tuple[List[str], List[str]]:
    errors: List[str] = []
    warnings: List[str] = []

    if not isinstance(content, dict):
        return ["Document root must be a mapping"], warnings

    for key, status in top_level_rules.items():
        value = content.get(key)
        value_present = key in content and has_value(value)

        if status == "mandatory" and not value_present:
            if key in OPTIONAL_WARN_KEYS:
                warnings.append(f"Optional key '{key}' is missing or empty")
                continue
            if key not in content or not allows_blank_value(key, value):
                errors.append(f"Missing or empty mandatory key '{key}'")
                continue
            continue

        if key in content and key in nested_rules:
            errors.extend(validate_nested(key, content[key], nested_rules[key]))

    return errors, warnings


def validate_nested(parent: str, value: Any, rules: Dict[str, str]) -> List[str]:
    errors: List[str] = []

    if value is None:
        return [f"Key '{parent}' must not be empty"]

    if isinstance(value, list):
        if not value:
            return [f"Key '{parent}' must contain at least one entry"]
        for idx, item in enumerate(value):
            if not isinstance(item, dict):
                errors.append(f"Entry {idx} under '{parent}' must be an object")
                continue
            for nested_key, nested_status in rules.items():
                if nested_status != "mandatory":
                    continue
                nested_value = item.get(nested_key)
                if not has_value(nested_value):
                    errors.append(
                        f"Entry {idx} under '{parent}' is missing mandatory key '{nested_key}'"
                    )
        return errors

    if isinstance(value, dict):
        for nested_key, nested_status in rules.items():
            if nested_status != "mandatory":
                continue
            nested_value = value.get(nested_key)
            if not has_value(nested_value):
                errors.append(f"Key '{parent}' is missing nested key '{nested_key}'")
        return errors

    return [f"Key '{parent}' must be a list or mapping"]


def iter_yaml_files(waf_dir: Path) -> List[Path]:
    return sorted(list(waf_dir.glob("*.yml")) + list(waf_dir.glob("*.yaml")))


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--root",
        type=Path,
        #required=True,
        default=Path("/"),
        help="Base directory containing the template and WAF folder",
    )
    parser.add_argument(
        "--template",
        type=Path,
        default=Path("waf_check_template.yml"),
        help="Path to the template file relative to --root (default: waf_check_template.yml)",
    )
    parser.add_argument(
        "--waf-dir",
        type=Path,
        default=Path("WAF"),
        help="Relative path to the folder with WAF YAML definitions (default: WAF)",
    )
    parser.add_argument(
        "--labels",
        type=Path,
        default=Path("labels.yml"),
        help="Relative path to labels.yml defining the allowed label names",
    )
    parser.add_argument(
        "--validations",
        type=Path,
        default=Path("validations.yml"),
        help="Relative path to validations.yml describing allowed scalar values",
    )
    parser.add_argument(
        "--epic-resources",
        type=Path,
        default=Path("epic_resources.yml"),
        help="Relative path to epic_resources.yml listing allowed resource names",
    )
    args = parser.parse_args()

    root = args.root.expanduser().resolve()
    template_path = (root / args.template).resolve()
    waf_dir = (root / args.waf_dir).resolve()
    labels_path = (root / args.labels).resolve()
    validations_path = (root / args.validations).resolve()
    epic_resources_path = (root / args.epic_resources).resolve()

    top_level_rules, nested_rules = parse_template(template_path)
    allowed_labels = load_allowed_labels(labels_path)
    validations = load_validations(validations_path, ["impact", "active", "kql_check"])
    allowed_epic_resources = load_epic_resources(epic_resources_path)

    yaml_files = iter_yaml_files(waf_dir)
    if not yaml_files:
        print(f"No YAML files found in {waf_dir}")
        sys.exit(1)

    total_errors = 0
    for file_path in yaml_files:
        try:
            document = load_yaml(file_path)
        except (yaml.YAMLError, OSError) as exc:  # pragma: no cover - PyYAML error formatting
            print(f"[FAIL] {file_path}: unable to load YAML ({exc})")
            total_errors += 1
            continue

        failures, warnings = validate_document(document, top_level_rules, nested_rules)

        labels_value = document.get("labels")
        if has_value(labels_value):
            failures.extend(
                validate_string_list("labels", labels_value, allowed_labels, "labels.yml")
            )

        for scoped_key in ("impact", "active", "kql_check"):
            if scoped_key in document:
                failures.extend(
                    validate_allowed_value(
                        scoped_key,
                        document.get(scoped_key),
                        validations[scoped_key],
                    )
                )

        epic_resources_value = document.get("epic_resources")
        if has_value(epic_resources_value):
            failures.extend(
                validate_string_list(
                    "epic_resources",
                    epic_resources_value,
                    allowed_epic_resources,
                    "epic_resources.yml",
                )
            )

        if failures:
            total_errors += len(failures)
            print(f"[FAIL] {file_path}")
            for issue in failures:
                print(f"  - {issue}")
            for warning in warnings:
                print(f"  - WARNING: {warning}")
        elif warnings:
            print(f"[WARN] {file_path}")
            for warning in warnings:
                print(f"  - WARNING: {warning}")
        else:
            print(f"[OK]   {file_path}")

    if total_errors:
        sys.exit(1)


if __name__ == "__main__":
    main()
