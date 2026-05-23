"""Check the local Google Drive course folder setup.

This script reads the shared and local course configuration files, verifies
that expected Google Drive folders exist, and confirms that the original slide
folder contains at least one PowerPoint deck.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

from lecture_files import discover_lecture_decks

try:
    import yaml
    from yaml import YAMLError
except ImportError:  # pragma: no cover - depends on the local Python environment
    yaml = None

    class YAMLError(Exception):
        """Placeholder when PyYAML is not installed."""


REQUIRED_FOLDER_KEYS = [
    ("manifest", "00_Manifest"),
    ("original_slides", "01_Original_Slides"),
    ("revised_slides", "02_Revised_Slides"),
    ("scripts", "03_Scripts"),
    ("notes_quizzes", "04_Notes_and_Quizzes"),
    ("figures", "05_Figures"),
    ("exports", "06_Exports"),
    ("videos", "07_Videos"),
    ("archive", "99_Archive"),
]


def load_yaml(path: Path) -> dict[str, Any]:
    if yaml is None:
        raise RuntimeError(
            "PyYAML is required. Install dependencies with: pip install -r requirements.txt"
        )

    with path.open("r", encoding="utf-8") as handle:
        data = yaml.safe_load(handle)

    if not isinstance(data, dict):
        raise ValueError(f"{path.name} must contain a YAML mapping.")

    return data


def folder_status(path: Path) -> tuple[str, bool, str]:
    try:
        if path.is_dir():
            return "OK", True, ""
        return "MISSING", False, ""
    except OSError as exc:
        return "INACCESSIBLE", False, str(exc)


def status_line(status: str, label: str, detail: str, reason: str = "") -> str:
    line = f"[{status}] {label:<24} {detail}"
    if reason:
        line = f"{line} ({reason})"
    return line


def main() -> int:
    project_root = Path(__file__).resolve().parents[1]
    shared_config_path = project_root / "course_config.yaml"
    local_config_path = project_root / "course_config.local.yaml"

    try:
        shared_config = load_yaml(shared_config_path)
        local_config = load_yaml(local_config_path)
    except (OSError, RuntimeError, ValueError, YAMLError) as exc:
        print("Environment check failed while loading configuration.")
        print(f"Error: {exc}")
        return 1

    course_drive_root_value = local_config.get("course_drive_root")
    folders = shared_config.get("folders")

    if not isinstance(course_drive_root_value, str) or not course_drive_root_value.strip():
        print("Environment check failed: course_config.local.yaml is missing course_drive_root.")
        return 1

    if not isinstance(folders, dict):
        print("Environment check failed: course_config.yaml is missing the folders mapping.")
        return 1

    course_drive_root = Path(course_drive_root_value).expanduser()

    print("ECE 591 course environment check")
    print("=" * 34)
    print(f"Project root:       {project_root}")
    print(f"Course drive root:  {course_drive_root}")
    print()

    print("Folder checks")
    print("-" * 13)

    missing_items: list[str] = []
    resolved_folders: dict[str, Path] = {}

    for key, default_name in REQUIRED_FOLDER_KEYS:
        folder_name = folders.get(key, default_name)
        if not isinstance(folder_name, str) or not folder_name.strip():
            missing_items.append(f"Invalid folder config for {key}")
            print(status_line("MISSING", default_name, f"invalid config key: folders.{key}"))
            continue

        folder_path = course_drive_root / folder_name
        resolved_folders[key] = folder_path
        status, exists, reason = folder_status(folder_path)

        if not exists:
            issue = f"{folder_path} ({status})"
            if reason:
                issue = f"{issue}: {reason}"
            missing_items.append(issue)

        print(status_line(status, folder_name, str(folder_path), reason))

    print()
    print("Original slide deck check")
    print("-" * 25)

    original_slides_path = resolved_folders.get("original_slides")
    lecture_decks = []
    pptx_check_failed = False

    if original_slides_path:
        status, accessible, reason = folder_status(original_slides_path)
        if accessible:
            try:
                lecture_decks = discover_lecture_decks(original_slides_path)
            except ValueError as exc:
                message = f"Invalid lecture PowerPoint filename pattern in {original_slides_path}: {exc}"
                missing_items.append(message)
                print(f"[INVALID] {message}")
                pptx_check_failed = True
            except OSError as exc:
                message = f"Could not list valid lecture .pptx files in {original_slides_path}: {exc}"
                missing_items.append(message)
                print(f"[INACCESSIBLE] {message}")
                pptx_check_failed = True
        elif status == "INACCESSIBLE":
            message = f"Could not access {original_slides_path}: {reason}"
            missing_items.append(message)
            print(f"[INACCESSIBLE] {message}")
            pptx_check_failed = True
        else:
            message = f"Original slides folder is missing: {original_slides_path}"
            missing_items.append(message)
            print(f"[MISSING] {message}")
            pptx_check_failed = True

    if lecture_decks:
        print(f"[OK] Found {len(lecture_decks)} lecture .pptx file(s) in {original_slides_path}")
        for deck in lecture_decks[:12]:
            print(f"     - Lecture {deck.lecture_number}: {deck.path.name}")
        if len(lecture_decks) > 12:
            print(f"     ... and {len(lecture_decks) - 12} more")
        if any(deck.lecture_number == 0 for deck in lecture_decks):
            print("     Lecture 0 is a special case; batch processing starts at Lecture 1.")
    elif original_slides_path and not pptx_check_failed:
        message = f"No .pptx files found in {original_slides_path}"
        missing_items.append(message)
        print(f"[MISSING] {message}")
    elif not original_slides_path:
        message = "Original slides folder could not be resolved."
        missing_items.append(message)
        print(f"[MISSING] {message}")

    print()
    print("PowerPoint safety")
    print("-" * 17)
    print("No PowerPoint files were opened, edited, or overwritten by this check.")
    print()

    if missing_items:
        print(f"Result: FAILED ({len(missing_items)} issue(s) found)")
        return 1

    print("Result: PASSED")
    return 0


if __name__ == "__main__":
    sys.exit(main())
