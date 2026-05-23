"""Build a lecture-level manifest for original PowerPoint decks.

The script reads the course configuration files, locates the original slides
folder, discovers PowerPoint files named "{number} - {title}.pptx", counts
slides read-only, and writes manifests/lecture_manifest.csv.
"""

from __future__ import annotations

import csv
import sys
from pathlib import Path
from typing import Any

from lecture_files import LectureDeck, discover_lecture_decks
from pptx_reader import PptxReadError, open_presentation_readonly

try:
    import yaml
    from yaml import YAMLError
except ImportError:  # pragma: no cover - depends on the local Python environment
    yaml = None

    class YAMLError(Exception):
        """Placeholder when PyYAML is not installed."""

OUTPUT_CSV = Path("manifests") / "lecture_manifest.csv"
CSV_FIELDS = [
    "lecture_number",
    "lecture_title",
    "original_pptx",
    "total_slides",
    "processing_status",
    "slide_count_error",
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


def get_project_root() -> Path:
    return Path(__file__).resolve().parents[1]


def get_original_slides_folder(project_root: Path) -> Path:
    shared_config = load_yaml(project_root / "course_config.yaml")
    local_config = load_yaml(project_root / "course_config.local.yaml")

    course_drive_root_value = local_config.get("course_drive_root")
    folders = shared_config.get("folders")

    if not isinstance(course_drive_root_value, str) or not course_drive_root_value.strip():
        raise ValueError("course_config.local.yaml is missing course_drive_root.")

    if not isinstance(folders, dict):
        raise ValueError("course_config.yaml is missing the folders mapping.")

    original_slides_name = folders.get("original_slides")
    if not isinstance(original_slides_name, str) or not original_slides_name.strip():
        raise ValueError("course_config.yaml is missing folders.original_slides.")

    return Path(course_drive_root_value).expanduser() / original_slides_name


def count_slides(deck: LectureDeck) -> tuple[int | None, str]:
    try:
        with open_presentation_readonly(deck.path) as (presentation, _repair_info):
            return len(presentation.slides), ""
    except PptxReadError as exc:
        return None, str(exc)


def processing_status(lecture_number: int) -> str:
    if lecture_number == 0:
        return "special_case_already_30min"
    return "ready_for_inventory"


def build_rows(decks: list[LectureDeck]) -> list[dict[str, str | int]]:
    rows: list[dict[str, str | int]] = []

    for deck in decks:
        total_slides, slide_count_error = count_slides(deck)
        rows.append(
            {
                "lecture_number": deck.lecture_number,
                "lecture_title": deck.lecture_title,
                "original_pptx": deck.path.name,
                "total_slides": "" if total_slides is None else total_slides,
                "processing_status": processing_status(deck.lecture_number),
                "slide_count_error": slide_count_error,
            }
        )

    return rows


def write_manifest(rows: list[dict[str, str | int]], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=CSV_FIELDS)
        writer.writeheader()
        writer.writerows(rows)


def main() -> int:
    project_root = get_project_root()
    output_path = project_root / OUTPUT_CSV

    try:
        original_slides_folder = get_original_slides_folder(project_root)
        decks = discover_lecture_decks(original_slides_folder)
        rows = build_rows(decks)
        write_manifest(rows, output_path)
    except (OSError, RuntimeError, ValueError, YAMLError) as exc:
        print("Lecture manifest build failed.")
        print(f"Error: {exc}")
        return 1

    special_case_count = sum(
        1 for row in rows if row["processing_status"] == "special_case_already_30min"
    )
    ready_count = sum(1 for row in rows if row["processing_status"] == "ready_for_inventory")
    count_error_rows = [row for row in rows if row["slide_count_error"]]
    total_counted_slides = sum(
        int(row["total_slides"]) for row in rows if row["total_slides"] != ""
    )

    print("Lecture manifest build complete")
    print(f"Original slides folder: {original_slides_folder}")
    print(f"Lectures found:         {len(rows)}")
    print(f"Special cases:          {special_case_count}")
    print(f"Ready for inventory:    {ready_count}")
    print(f"Slides counted:         {total_counted_slides}")
    print(f"Slide count errors:     {len(count_error_rows)}")
    for row in rows:
        slide_count = row["total_slides"] if row["total_slides"] != "" else "ERROR"
        print(
            f"  - Lecture {row['lecture_number']}: {row['original_pptx']} "
            f"({slide_count} slides, {row['processing_status']})"
        )
    if count_error_rows:
        print("Files with slide count errors:")
        for row in count_error_rows:
            print(f"  - Lecture {row['lecture_number']}: {row['original_pptx']}")
    print(f"Output CSV:             {output_path}")
    print("PowerPoint safety: original decks were read only; no PowerPoint files were modified.")
    return 1 if count_error_rows else 0


if __name__ == "__main__":
    sys.exit(main())
