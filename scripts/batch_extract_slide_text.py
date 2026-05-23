"""Batch extract visible slide text for inventory-ready lectures.

This script reads manifests/lecture_manifest.csv, processes only lectures 1+
marked ready_for_inventory, extracts visible slide text from each original deck,
and writes one CSV per lecture plus a combined inventory.
"""

from __future__ import annotations

import csv
import sys
from pathlib import Path
from typing import Any

from build_lecture_manifest import get_original_slides_folder, get_project_root
from extract_slide_text import possible_slide_title, slide_visible_text
from pptx_reader import PptxReadError, open_presentation_readonly


INPUT_MANIFEST = Path("manifests") / "lecture_manifest.csv"
OUTPUT_DIR = Path("manifests") / "slide_text"
COMBINED_OUTPUT_CSV = Path("manifests") / "all_slide_text_inventory.csv"
READY_STATUS = "ready_for_inventory"
CSV_FIELDS = [
    "lecture_number",
    "lecture_title",
    "original_pptx",
    "slide_number",
    "possible_slide_title",
    "visible_text",
]


def read_lecture_manifest(manifest_path: Path) -> list[dict[str, str]]:
    with manifest_path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def selected_lectures(rows: list[dict[str, str]]) -> list[dict[str, str]]:
    selected: list[dict[str, str]] = []

    for row in rows:
        lecture_number = parse_lecture_number(row)
        processing_status = row.get("processing_status", "")
        if lecture_number >= 1 and processing_status == READY_STATUS:
            selected.append(row)

    return sorted(selected, key=parse_lecture_number)


def parse_lecture_number(row: dict[str, str]) -> int:
    value = row.get("lecture_number", "")
    try:
        return int(value)
    except ValueError as exc:
        raise ValueError(f"Invalid lecture_number in manifest: {value!r}") from exc


def original_pptx_path(original_slides_folder: Path, row: dict[str, str]) -> Path:
    original_pptx = row.get("original_pptx", "")
    if not original_pptx:
        raise ValueError(f"Missing original_pptx for Lecture {row.get('lecture_number', '')}")

    filename = Path(original_pptx).name
    if filename != original_pptx:
        raise ValueError(
            f"original_pptx must be a filename, not a path: {original_pptx!r}"
        )

    return original_slides_folder / filename


def extract_rows_for_lecture(
    row: dict[str, str],
    original_slides_folder: Path,
) -> list[dict[str, str | int]]:
    lecture_number = parse_lecture_number(row)
    lecture_title = row.get("lecture_title", "")
    original_pptx = row.get("original_pptx", "")
    pptx_path = original_pptx_path(original_slides_folder, row)

    if not pptx_path.is_file():
        raise FileNotFoundError(f"Original PowerPoint not found: {pptx_path}")

    try:
        with open_presentation_readonly(pptx_path) as (presentation, _repair_info):
            rows: list[dict[str, str | int]] = []
            for slide_index, slide in enumerate(presentation.slides, start=1):
                visible_text = slide_visible_text(slide)
                rows.append(
                    {
                        "lecture_number": lecture_number,
                        "lecture_title": lecture_title,
                        "original_pptx": original_pptx,
                        "slide_number": slide_index,
                        "possible_slide_title": possible_slide_title(slide, visible_text),
                        "visible_text": visible_text,
                    }
                )
            return rows
    except PptxReadError as exc:
        raise RuntimeError(
            f"Could not read Lecture {lecture_number} PowerPoint file {original_pptx}: {exc}"
        ) from exc


def lecture_output_path(output_dir: Path, lecture_number: int) -> Path:
    return output_dir / f"{lecture_number}_slide_text.csv"


def write_rows(rows: list[dict[str, str | int]], output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=CSV_FIELDS)
        writer.writeheader()
        writer.writerows(rows)


def main() -> int:
    project_root = get_project_root()
    manifest_path = project_root / INPUT_MANIFEST
    output_dir = project_root / OUTPUT_DIR
    combined_output_path = project_root / COMBINED_OUTPUT_CSV

    try:
        original_slides_folder = get_original_slides_folder(project_root)
        manifest_rows = read_lecture_manifest(manifest_path)
        lectures = selected_lectures(manifest_rows)
    except (OSError, RuntimeError, ValueError) as exc:
        print("Batch slide text extraction failed.")
        print(f"Error: {exc}")
        return 1

    all_rows: list[dict[str, str | int]] = []
    slides_by_lecture: dict[int, int] = {}

    try:
        for lecture_row in lectures:
            lecture_number = parse_lecture_number(lecture_row)
            lecture_rows = extract_rows_for_lecture(lecture_row, original_slides_folder)
            write_rows(lecture_rows, lecture_output_path(output_dir, lecture_number))
            all_rows.extend(lecture_rows)
            slides_by_lecture[lecture_number] = len(lecture_rows)

        write_rows(all_rows, combined_output_path)
    except (OSError, RuntimeError, ValueError) as exc:
        print("Batch slide text extraction failed.")
        print(f"Error: {exc}")
        return 1

    print("Batch slide text extraction complete")
    print(f"Lecture manifest:        {manifest_path}")
    print(f"Original slides folder:  {original_slides_folder}")
    print(f"Lectures processed:      {len(lectures)}")
    print(f"Slides processed:        {len(all_rows)}")
    for lecture_number in sorted(slides_by_lecture):
        print(f"  - Lecture {lecture_number}: {slides_by_lecture[lecture_number]} slides")
    print(f"Per-lecture CSV folder:  {output_dir}")
    print(f"Combined CSV:            {combined_output_path}")
    print("PowerPoint safety: original decks were read only; no PowerPoint files were modified.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
