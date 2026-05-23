"""Audit generated segment PowerPoint decks.

The script inspects generated decks under the configured 02_Revised_Slides
folder and writes CSV/Markdown reports under manifests/deck_audit/.

PowerPoint files are opened read-only. Original PowerPoint files are never
modified.
"""

from __future__ import annotations

import csv
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass
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


VIDEO_SEGMENTS_DIR = Path("manifests") / "video_segments"
AUDIT_DIR = Path("manifests") / "deck_audit"
AUDIT_CSV = AUDIT_DIR / "generated_deck_audit.csv"
SUMMARY_MD = AUDIT_DIR / "generated_deck_audit_summary.md"
APPROVED_STATUS = "approved_for_pptx"
OUTPUT_FIELD = "output_pptx"
SCRIPT_MARKER = "Generated speaker-script outline:"


@dataclass(frozen=True)
class SegmentRecord:
    segment_csv: Path
    row_number: int
    lecture_number: int
    lecture_title: str
    video_number: int
    video_title: str
    slide_start: int
    slide_end: int
    review_status: str
    output_pptx: str


@dataclass(frozen=True)
class FileSnapshot:
    exists: bool
    size: int | None
    mtime_ns: int | None


def project_root() -> Path:
    return Path(__file__).resolve().parents[1]


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


def configured_folder(project_root_path: Path, folder_key: str) -> Path:
    shared_config = load_yaml(project_root_path / "course_config.yaml")
    local_config = load_yaml(project_root_path / "course_config.local.yaml")

    course_drive_root_value = local_config.get("course_drive_root")
    folders = shared_config.get("folders")

    if not isinstance(course_drive_root_value, str) or not course_drive_root_value.strip():
        raise ValueError("course_config.local.yaml is missing course_drive_root.")
    if not isinstance(folders, dict):
        raise ValueError("course_config.yaml is missing the folders mapping.")

    folder_name = folders.get(folder_key)
    if not isinstance(folder_name, str) or not folder_name.strip():
        raise ValueError(f"course_config.yaml is missing folders.{folder_key}.")

    return Path(course_drive_root_value).expanduser() / folder_name


def speaker_notes_expected(project_root_path: Path) -> bool:
    shared_config = load_yaml(project_root_path / "course_config.yaml")
    slide_policy = shared_config.get("slide_policy", {})
    if not isinstance(slide_policy, dict):
        return True
    return bool(slide_policy.get("generate_speaker_notes", True))


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def write_csv(path: Path, fieldnames: list[str], rows: list[dict[str, object]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def bool_text(value: bool) -> str:
    return "yes" if value else "no"


def path_key(path: Path) -> str:
    return path.as_posix()


def snapshot(path: Path) -> FileSnapshot:
    if not path.exists():
        return FileSnapshot(exists=False, size=None, mtime_ns=None)
    stat = path.stat()
    return FileSnapshot(exists=True, size=stat.st_size, mtime_ns=stat.st_mtime_ns)


def snapshots_match(before: FileSnapshot, after: FileSnapshot) -> bool:
    return (
        before.exists == after.exists
        and before.size == after.size
        and before.mtime_ns == after.mtime_ns
    )


def segment_files(root: Path) -> list[Path]:
    files = sorted(
        (root / VIDEO_SEGMENTS_DIR).glob("*_video_segments.csv"),
        key=lambda path: int(path.name.split("_", 1)[0]),
    )
    if not files:
        raise FileNotFoundError(f"No video segment CSV files found in {root / VIDEO_SEGMENTS_DIR}")
    return files


def parse_int(value: str, field_name: str, source: Path, row_number: int) -> int:
    try:
        return int(value)
    except ValueError as exc:
        raise ValueError(f"{source} row {row_number} has invalid {field_name}: {value}") from exc


def load_segment_records(root: Path) -> list[SegmentRecord]:
    records: list[SegmentRecord] = []
    for path in segment_files(root):
        rows = read_csv(path)
        for row_index, row in enumerate(rows, start=2):
            output_pptx = row.get(OUTPUT_FIELD, "").strip()
            if not output_pptx:
                continue
            records.append(
                SegmentRecord(
                    segment_csv=path,
                    row_number=row_index,
                    lecture_number=parse_int(row["lecture_number"], "lecture_number", path, row_index),
                    lecture_title=row.get("lecture_title", ""),
                    video_number=parse_int(row["video_number"], "video_number", path, row_index),
                    video_title=row.get("video_title", ""),
                    slide_start=parse_int(
                        row["original_slide_start"],
                        "original_slide_start",
                        path,
                        row_index,
                    ),
                    slide_end=parse_int(
                        row["original_slide_end"],
                        "original_slide_end",
                        path,
                        row_index,
                    ),
                    review_status=row.get("review_status", ""),
                    output_pptx=output_pptx.replace("\\", "/"),
                )
            )
    return records


def discover_generated_decks(revised_slides_folder: Path) -> list[Path]:
    if not revised_slides_folder.is_dir():
        raise FileNotFoundError(f"Revised slides folder not found: {revised_slides_folder}")
    return sorted(
        path
        for path in revised_slides_folder.rglob("*.pptx")
        if not path.name.startswith("~$")
    )


def text_from_slide(slide: object) -> str:
    parts: list[str] = []
    for shape in slide.shapes:
        text = getattr(shape, "text", "")
        if text:
            parts.append(text)
    return "\n".join(parts).strip()


def slide_title(slide: object) -> str:
    title_shape = slide.shapes.title
    if title_shape is not None:
        return title_shape.text.strip()
    return ""


def has_named_slide(presentation: object, names: set[str]) -> bool:
    lowered = {name.lower() for name in names}
    for slide in presentation.slides:
        title = slide_title(slide).lower()
        text = text_from_slide(slide).lower()
        if any(name in title or name in text for name in lowered):
            return True
    return False


def notes_text(slide: object) -> str:
    try:
        notes_slide = slide.notes_slide
        text_frame = notes_slide.notes_text_frame
    except Exception:
        return ""
    if text_frame is None:
        return ""
    return text_frame.text.strip()


def audit_presentation(path: Path, record: SegmentRecord | None, notes_expected: bool) -> dict[str, object]:
    result: dict[str, object] = {
        "can_open": "no",
        "open_error": "",
        "slide_count": "",
        "expected_slide_count": "",
        "slide_count_reasonable": "no",
        "has_title_slide": "no",
        "title_slide_text": "",
        "has_recap_slide": "no",
        "has_quiz_checkpoint_slide": "no",
        "speaker_notes_expected": bool_text(notes_expected),
        "speaker_notes_slide_count": "",
        "speaker_notes_ok": "no" if notes_expected else "not_expected",
        "used_temporary_repair": "no",
    }

    try:
        with open_presentation_readonly(path) as (presentation, repair_info):
            slide_count = len(presentation.slides)
            result["can_open"] = "yes"
            result["slide_count"] = slide_count
            result["used_temporary_repair"] = bool_text(repair_info.used_repair)
            if record is not None:
                expected_slide_count = record.slide_end - record.slide_start + 1 + 3
                result["expected_slide_count"] = expected_slide_count
                result["slide_count_reasonable"] = bool_text(slide_count == expected_slide_count)
            else:
                result["slide_count_reasonable"] = bool_text(slide_count >= 3)

            if slide_count:
                first_slide_title = slide_title(presentation.slides[0])
                first_slide_text = text_from_slide(presentation.slides[0])
                result["title_slide_text"] = first_slide_title or first_slide_text.splitlines()[0]
                result["has_title_slide"] = bool_text(bool(first_slide_title or first_slide_text))

            result["has_recap_slide"] = bool_text(has_named_slide(presentation, {"recap"}))
            result["has_quiz_checkpoint_slide"] = bool_text(
                has_named_slide(presentation, {"checkpoint", "quiz"})
            )

            notes_count = 0
            marker_count = 0
            for slide in presentation.slides:
                notes = notes_text(slide)
                if notes:
                    notes_count += 1
                if SCRIPT_MARKER in notes:
                    marker_count += 1
            result["speaker_notes_slide_count"] = notes_count
            if notes_expected:
                result["speaker_notes_ok"] = bool_text(slide_count > 0 and notes_count == slide_count)
                result["speaker_notes_have_generated_marker"] = bool_text(
                    slide_count > 0 and marker_count == slide_count
                )
            else:
                result["speaker_notes_have_generated_marker"] = "not_expected"
    except (OSError, PptxReadError, ValueError, KeyError) as exc:
        result["open_error"] = str(exc)
        result["speaker_notes_have_generated_marker"] = "no"

    return result


def status_from_checks(row: dict[str, object], checks: list[str]) -> str:
    passing_values = {"yes", "not_required", "not_expected"}
    return "pass" if all(row.get(check) in passing_values for check in checks) else "fail"


def is_lecture_zero_manual_deck(output_key: str) -> bool:
    return output_key.startswith("0/") and output_key.lower().endswith(".pptx")


def build_audit_rows(
    root: Path,
    original_slides_folder: Path,
    revised_slides_folder: Path,
    notes_expected: bool,
) -> tuple[list[dict[str, object]], dict[str, object]]:
    records = load_segment_records(root)
    records_by_output: dict[str, list[SegmentRecord]] = defaultdict(list)
    for record in records:
        records_by_output[record.output_pptx].append(record)

    generated_decks = discover_generated_decks(revised_slides_folder)
    actual_outputs = {
        path.relative_to(revised_slides_folder).as_posix(): path for path in generated_decks
    }
    output_counter = Counter(record.output_pptx for record in records)
    actual_basename_counter = Counter(path.name for path in generated_decks)

    lecture_decks = discover_lecture_decks(original_slides_folder)
    lecture_by_number = {deck.lecture_number: deck for deck in lecture_decks}
    original_before = {deck.path: snapshot(deck.path) for deck in lecture_decks}

    output_keys = sorted(
        set(actual_outputs) | set(records_by_output),
        key=lambda value: tuple(int(part) if part.isdigit() else part for part in value.split("/")),
    )

    rows: list[dict[str, object]] = []
    for output_key in output_keys:
        path = actual_outputs.get(output_key, revised_slides_folder / output_key)
        record = records_by_output[output_key][0] if records_by_output.get(output_key) else None
        special_case = "lecture_0_manual" if record is None and is_lecture_zero_manual_deck(output_key) else ""
        duplicate_output = output_counter[output_key] > 1 or actual_basename_counter[path.name] > 1
        original_deck: LectureDeck | None = None
        if record is not None:
            original_deck = lecture_by_number.get(record.lecture_number)
        elif special_case == "lecture_0_manual":
            original_deck = lecture_by_number.get(0)

        audit_result: dict[str, object]
        if path.exists():
            audit_result = audit_presentation(path, record, notes_expected)
        else:
            audit_result = {
                "can_open": "no",
                "open_error": "Referenced output file does not exist.",
                "slide_count": "",
                "expected_slide_count": (
                    record.slide_end - record.slide_start + 1 + 3 if record is not None else ""
                ),
                "slide_count_reasonable": "no",
                "has_title_slide": "no",
                "title_slide_text": "",
                "has_recap_slide": "no",
                "has_quiz_checkpoint_slide": "no",
                "speaker_notes_expected": bool_text(notes_expected),
                "speaker_notes_slide_count": "",
                "speaker_notes_ok": "no" if notes_expected else "not_expected",
                "speaker_notes_have_generated_marker": "no" if notes_expected else "not_expected",
                "used_temporary_repair": "no",
            }

        original_after = snapshot(original_deck.path) if original_deck is not None else None
        original_unchanged = (
            snapshots_match(original_before[original_deck.path], original_after)
            if original_deck is not None and original_after is not None
            else False
        )

        original_file_exists = (
            bool_text(original_deck.path.exists()) if original_deck is not None else "not_mapped"
        )
        original_unmodified = (
            bool_text(original_unchanged) if original_deck is not None else "not_mapped"
        )

        recorded_in_segment_csv = bool_text(output_key in records_by_output)
        lecture_number: int | str = record.lecture_number if record is not None else ""
        lecture_title = record.lecture_title if record is not None else ""
        video_number: int | str = record.video_number if record is not None else ""
        video_title = record.video_title if record is not None else ""
        review_status = record.review_status if record is not None else ""
        if special_case == "lecture_0_manual":
            recorded_in_segment_csv = "not_required"
            lecture_number = 0
            lecture_title = original_deck.lecture_title if original_deck is not None else "Intro"
            video_number = 1
            video_title = "Introduction"
            review_status = "special_case_already_30min"
            audit_result["has_recap_slide"] = "not_required"
            audit_result["has_quiz_checkpoint_slide"] = "not_required"
            audit_result["speaker_notes_expected"] = "not_required"
            audit_result["speaker_notes_ok"] = "not_required"
            audit_result["speaker_notes_have_generated_marker"] = "not_required"

        row: dict[str, object] = {
            "special_case": special_case,
            "output_pptx": output_key,
            "generated_pptx_path": str(path),
            "file_exists": bool_text(path.exists()),
            "recorded_in_segment_csv": recorded_in_segment_csv,
            "duplicate_output_filename": bool_text(duplicate_output),
            "segment_csv": str(record.segment_csv) if record is not None else "",
            "segment_csv_row": record.row_number if record is not None else "",
            "lecture_number": lecture_number,
            "lecture_title": lecture_title,
            "video_number": video_number,
            "video_title": video_title,
            "review_status": review_status,
            "original_pptx": str(original_deck.path) if original_deck is not None else "",
            "original_file_exists": original_file_exists,
            "original_unmodified_during_audit": original_unmodified,
            **audit_result,
        }
        row["overall_status"] = status_from_checks(
            row,
            [
                "file_exists",
                "can_open",
                "slide_count_reasonable",
                "has_title_slide",
                "has_recap_slide",
                "has_quiz_checkpoint_slide",
                "recorded_in_segment_csv",
                "original_file_exists",
                "original_unmodified_during_audit",
            ],
        )
        if (
            notes_expected
            and row["overall_status"] == "pass"
            and row["speaker_notes_ok"] not in {"yes", "not_required", "not_expected"}
        ):
            row["overall_status"] = "fail"
        if row["overall_status"] == "pass" and row["duplicate_output_filename"] != "no":
            row["overall_status"] = "fail"
        rows.append(row)

    original_after_all = {deck.path: snapshot(deck.path) for deck in lecture_decks}
    original_changed = [
        str(path)
        for path, before in original_before.items()
        if not snapshots_match(before, original_after_all[path])
    ]

    summary = {
        "records": records,
        "generated_decks": generated_decks,
        "duplicate_csv_outputs": {
            key: count for key, count in output_counter.items() if count > 1
        },
        "duplicate_actual_basenames": {
            key: count for key, count in actual_basename_counter.items() if count > 1
        },
        "original_changed": original_changed,
    }
    return rows, summary


def write_summary(path: Path, rows: list[dict[str, object]], summary: dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    total = len(rows)
    passing = sum(1 for row in rows if row["overall_status"] == "pass")
    failing = total - passing
    special_cases = [row for row in rows if row["special_case"]]
    missing_files = [row for row in rows if row["file_exists"] == "no"]
    cannot_open = [row for row in rows if row["can_open"] == "no"]
    unreasonable_counts = [row for row in rows if row["slide_count_reasonable"] == "no"]
    missing_title = [row for row in rows if row["has_title_slide"] == "no"]
    missing_recap = [row for row in rows if row["has_recap_slide"] == "no"]
    missing_checkpoint = [row for row in rows if row["has_quiz_checkpoint_slide"] == "no"]
    missing_notes = [row for row in rows if row["speaker_notes_ok"] == "no"]
    unrecorded = [row for row in rows if row["recorded_in_segment_csv"] == "no"]
    duplicate_rows = [row for row in rows if row["duplicate_output_filename"] == "yes"]
    original_changed = summary["original_changed"]

    lines = [
        "# Generated Deck Audit Summary",
        "",
        "| Check | Count |",
        "|---|---:|",
        f"| Generated/expected decks audited | {total} |",
        f"| Manual special-case decks | {len(special_cases)} |",
        f"| Passing decks | {passing} |",
        f"| Failing decks | {failing} |",
        f"| Missing output files | {len(missing_files)} |",
        f"| Files that could not be opened | {len(cannot_open)} |",
        f"| Unreasonable slide counts | {len(unreasonable_counts)} |",
        f"| Missing title slide | {len(missing_title)} |",
        f"| Missing recap slide | {len(missing_recap)} |",
        f"| Missing quiz/checkpoint slide | {len(missing_checkpoint)} |",
        f"| Missing speaker notes | {len(missing_notes)} |",
        f"| Not recorded in segment CSV | {len(unrecorded)} |",
        f"| Duplicate output filenames | {len(duplicate_rows)} |",
        f"| Original PPTX files changed during audit | {len(original_changed)} |",
        "",
        "Original-file safety check compares file size and modification timestamp before and after this audit run.",
    ]

    if failing:
        lines.extend(["", "## Items Needing Attention", ""])
        for row in rows:
            if row["overall_status"] == "pass":
                continue
            problems = [
                field
                for field in [
                    "file_exists",
                    "can_open",
                    "slide_count_reasonable",
                    "has_title_slide",
                    "has_recap_slide",
                    "has_quiz_checkpoint_slide",
                    "speaker_notes_ok",
                    "recorded_in_segment_csv",
                    "duplicate_output_filename",
                    "original_file_exists",
                    "original_unmodified_during_audit",
                ]
                if row.get(field) not in {"yes", "not_required", "not_expected", "not_mapped", "no"}
            ]
            if row.get("duplicate_output_filename") == "yes":
                problems.append("duplicate_output_filename")
            failed_yes_checks = [
                field
                for field in [
                    "file_exists",
                    "can_open",
                    "slide_count_reasonable",
                    "has_title_slide",
                    "has_recap_slide",
                    "has_quiz_checkpoint_slide",
                    "speaker_notes_ok",
                    "recorded_in_segment_csv",
                    "original_file_exists",
                    "original_unmodified_during_audit",
                ]
                if row.get(field) == "no"
            ]
            all_problems = failed_yes_checks + problems
            lines.append(f"- `{row['output_pptx']}`: {', '.join(sorted(set(all_problems)))}")

    duplicate_csv_outputs = summary["duplicate_csv_outputs"]
    duplicate_actual_basenames = summary["duplicate_actual_basenames"]
    if duplicate_csv_outputs or duplicate_actual_basenames:
        lines.extend(["", "## Duplicate Details", ""])
        for name, count in sorted(duplicate_csv_outputs.items()):
            lines.append(f"- CSV output reference `{name}` appears {count} times.")
        for name, count in sorted(duplicate_actual_basenames.items()):
            lines.append(f"- Generated filename `{name}` appears {count} times on disk.")

    if original_changed:
        lines.extend(["", "## Original Files Changed During Audit", ""])
        for item in original_changed:
            lines.append(f"- `{item}`")

    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> int:
    root = project_root()
    try:
        original_slides_folder = configured_folder(root, "original_slides")
        revised_slides_folder = configured_folder(root, "revised_slides")
        notes_expected = speaker_notes_expected(root)
        rows, summary = build_audit_rows(
            root,
            original_slides_folder,
            revised_slides_folder,
            notes_expected,
        )

        fieldnames = [
            "overall_status",
            "special_case",
            "output_pptx",
            "generated_pptx_path",
            "file_exists",
            "can_open",
            "open_error",
            "slide_count",
            "expected_slide_count",
            "slide_count_reasonable",
            "has_title_slide",
            "title_slide_text",
            "has_recap_slide",
            "has_quiz_checkpoint_slide",
            "speaker_notes_expected",
            "speaker_notes_slide_count",
            "speaker_notes_ok",
            "speaker_notes_have_generated_marker",
            "used_temporary_repair",
            "recorded_in_segment_csv",
            "duplicate_output_filename",
            "segment_csv",
            "segment_csv_row",
            "lecture_number",
            "lecture_title",
            "video_number",
            "video_title",
            "review_status",
            "original_pptx",
            "original_file_exists",
            "original_unmodified_during_audit",
        ]
        write_csv(root / AUDIT_CSV, fieldnames, rows)
        write_summary(root / SUMMARY_MD, rows, summary)

    except (OSError, RuntimeError, ValueError, KeyError, YAMLError, PptxReadError) as exc:
        print("Generated deck audit failed.")
        print(str(exc))
        return 1

    passing = sum(1 for row in rows if row["overall_status"] == "pass")
    print("Generated deck audit complete")
    print(f"Original slides folder: {original_slides_folder}")
    print(f"Revised slides folder:  {revised_slides_folder}")
    print(f"Decks audited:          {len(rows)}")
    print(f"Passing decks:          {passing}")
    print(f"Failing decks:          {len(rows) - passing}")
    print(f"CSV report:             {root / AUDIT_CSV}")
    print(f"Summary report:         {root / SUMMARY_MD}")
    print("PowerPoint safety: presentations were opened read-only and never saved.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
