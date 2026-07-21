from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any


PROJECT_ROOT = Path(__file__).resolve().parent


def _default_memory_path() -> Path:
    mapped_path = PROJECT_ROOT / "Mappings" / "normalization_memory.json"
    flat_path = PROJECT_ROOT / "normalization_memory.json"
    if mapped_path.exists() or not flat_path.exists():
        return mapped_path
    return flat_path


DEFAULT_MEMORY_PATH = _default_memory_path()


def normalise_memory_key(value: object) -> str | None:
    if not isinstance(value, str):
        return None
    value = value.strip()
    if not value:
        return None
    return " ".join(value.lower().replace("\n", " ").split())


def _blank_memory() -> dict[str, Any]:
    return {
        "version": 1,
        "aliases": {},
        "sheet_aliases": {},
        "history": [],
    }


def load_normalization_memory(path: Path = DEFAULT_MEMORY_PATH) -> dict[str, Any]:
    if not path.exists():
        return _blank_memory()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return _blank_memory()
    if not isinstance(data, dict):
        return _blank_memory()
    data.setdefault("version", 1)
    data.setdefault("aliases", {})
    data.setdefault("sheet_aliases", {})
    data.setdefault("history", [])
    if not isinstance(data["aliases"], dict):
        data["aliases"] = {}
    if not isinstance(data["sheet_aliases"], dict):
        data["sheet_aliases"] = {}
    if not isinstance(data["history"], list):
        data["history"] = []
    return data


def save_normalization_memory(data: dict[str, Any], path: Path = DEFAULT_MEMORY_PATH) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2, sort_keys=True), encoding="utf-8")


def memory_aliases_for_profile(profile_name: str, path: Path = DEFAULT_MEMORY_PATH) -> dict[str, str]:
    data = load_normalization_memory(path)
    aliases = data.get("aliases", {}).get(profile_name, {})
    if not isinstance(aliases, dict):
        return {}

    clean_aliases: dict[str, str] = {}
    for report_label, source_label in aliases.items():
        report_key = normalise_memory_key(report_label)
        source_key = normalise_memory_key(source_label)
        if report_key and source_key:
            clean_aliases[report_key] = source_key
    return clean_aliases


def memory_sheet_aliases_for_profile(profile_name: str, path: Path = DEFAULT_MEMORY_PATH) -> dict[str, str]:
    data = load_normalization_memory(path)
    aliases = data.get("sheet_aliases", {}).get(profile_name, {})
    if not isinstance(aliases, dict):
        return {}

    clean_aliases: dict[str, str] = {}
    for expected_sheet, actual_sheet in aliases.items():
        expected_key = normalise_memory_key(expected_sheet)
        if expected_key and isinstance(actual_sheet, str) and actual_sheet.strip():
            clean_aliases[expected_key] = actual_sheet.strip()
    return clean_aliases


def save_memory_alias(
    profile_name: str,
    report_label: str,
    source_label: str,
    source_ref: str = "",
    confidence: int | None = None,
    path: Path = DEFAULT_MEMORY_PATH,
) -> bool:
    report_key = normalise_memory_key(report_label)
    source_key = normalise_memory_key(source_label)
    if report_key is None or source_key is None:
        return False

    data = load_normalization_memory(path)
    aliases = data.setdefault("aliases", {})
    profile_aliases = aliases.setdefault(profile_name, {})
    if not isinstance(profile_aliases, dict):
        profile_aliases = {}
        aliases[profile_name] = profile_aliases

    previous = profile_aliases.get(report_key)
    profile_aliases[report_key] = source_key
    data.setdefault("history", []).append(
        {
            "profile": profile_name,
            "report_label": report_key,
            "source_label": source_key,
            "source_ref": source_ref,
            "confidence": confidence,
            "previous_source_label": previous,
            "approved_at": datetime.now().isoformat(timespec="seconds"),
        }
    )
    save_normalization_memory(data, path)
    return True


def save_memory_sheet_alias(
    profile_name: str,
    expected_sheet_name: str,
    actual_sheet_name: str,
    confidence: int | None = None,
    path: Path = DEFAULT_MEMORY_PATH,
) -> bool:
    expected_key = normalise_memory_key(expected_sheet_name)
    if expected_key is None or not isinstance(actual_sheet_name, str) or not actual_sheet_name.strip():
        return False

    actual_sheet_name = actual_sheet_name.strip()
    data = load_normalization_memory(path)
    aliases = data.setdefault("sheet_aliases", {})
    profile_aliases = aliases.setdefault(profile_name, {})
    if not isinstance(profile_aliases, dict):
        profile_aliases = {}
        aliases[profile_name] = profile_aliases

    previous = profile_aliases.get(expected_key)
    profile_aliases[expected_key] = actual_sheet_name
    data.setdefault("history", []).append(
        {
            "profile": profile_name,
            "expected_sheet_name": expected_key,
            "actual_sheet_name": actual_sheet_name,
            "confidence": confidence,
            "previous_actual_sheet_name": previous,
            "approved_at": datetime.now().isoformat(timespec="seconds"),
        }
    )
    save_normalization_memory(data, path)
    return True


def list_memory_entries(path: Path = DEFAULT_MEMORY_PATH) -> list[dict[str, str]]:
    """All approved mappings as flat rows for display/management in the UI."""
    data = load_normalization_memory(path)
    approved_at: dict[tuple[str, str, str], str] = {}
    for event in data.get("history", []):
        if not isinstance(event, dict):
            continue
        profile = str(event.get("profile", ""))
        when = str(event.get("approved_at", ""))
        if "report_label" in event:
            approved_at[("Source row", profile, str(event.get("report_label", "")))] = when
        elif "expected_sheet_name" in event:
            approved_at[("Workbook tab", profile, str(event.get("expected_sheet_name", "")))] = when

    entries: list[dict[str, str]] = []
    for profile, aliases in data.get("aliases", {}).items():
        if not isinstance(aliases, dict):
            continue
        for report_label, source_label in aliases.items():
            entries.append(
                {
                    "Type": "Source row",
                    "Profile": profile,
                    "Expected": report_label,
                    "Mapped to": str(source_label),
                    "Approved": approved_at.get(("Source row", profile, report_label), ""),
                }
            )
    for profile, sheet_aliases in data.get("sheet_aliases", {}).items():
        if not isinstance(sheet_aliases, dict):
            continue
        for expected_sheet, actual_sheet in sheet_aliases.items():
            entries.append(
                {
                    "Type": "Workbook tab",
                    "Profile": profile,
                    "Expected": expected_sheet,
                    "Mapped to": str(actual_sheet),
                    "Approved": approved_at.get(("Workbook tab", profile, expected_sheet), ""),
                }
            )
    return entries


def remove_memory_alias(
    profile_name: str,
    report_label: str,
    path: Path = DEFAULT_MEMORY_PATH,
) -> bool:
    report_key = normalise_memory_key(report_label)
    if report_key is None:
        return False
    data = load_normalization_memory(path)
    profile_aliases = data.get("aliases", {}).get(profile_name)
    if not isinstance(profile_aliases, dict) or report_key not in profile_aliases:
        return False
    removed = profile_aliases.pop(report_key)
    data.setdefault("history", []).append(
        {
            "profile": profile_name,
            "report_label": report_key,
            "removed_source_label": removed,
            "removed_at": datetime.now().isoformat(timespec="seconds"),
        }
    )
    save_normalization_memory(data, path)
    return True


def remove_memory_sheet_alias(
    profile_name: str,
    expected_sheet_name: str,
    path: Path = DEFAULT_MEMORY_PATH,
) -> bool:
    expected_key = normalise_memory_key(expected_sheet_name)
    if expected_key is None:
        return False
    data = load_normalization_memory(path)
    profile_aliases = data.get("sheet_aliases", {}).get(profile_name)
    if not isinstance(profile_aliases, dict) or expected_key not in profile_aliases:
        return False
    removed = profile_aliases.pop(expected_key)
    data.setdefault("history", []).append(
        {
            "profile": profile_name,
            "expected_sheet_name": expected_key,
            "removed_actual_sheet_name": removed,
            "removed_at": datetime.now().isoformat(timespec="seconds"),
        }
    )
    save_normalization_memory(data, path)
    return True


def source_label_from_note(note: str) -> str | None:
    match = re.search(r"Possible renamed source label:\s*(.+?)\.$", note)
    return match.group(1) if match else None


def source_sheet_from_note(note: str) -> tuple[str, str] | None:
    match = re.search(
        r"Possible renamed source sheet:\s*expected '(.+?)' -> actual '(.+?)'\.$",
        note,
    )
    if not match:
        return None
    return match.group(1), match.group(2)
