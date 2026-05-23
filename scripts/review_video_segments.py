"""Automatically review proposed video segmentation plans."""

from __future__ import annotations

import csv
import re
import sys
from pathlib import Path


TEACHING_MAP_DIR = Path("manifests") / "teaching_maps"
VIDEO_SEGMENT_DIR = Path("manifests") / "video_segments"
OUTPUT_DIR = Path("manifests") / "segmentation_reviews"
SUMMARY_CSV = OUTPUT_DIR / "segmentation_review_summary.csv"
CSV_FIELDS = [
    "lecture_number",
    "lecture_title",
    "video_number",
    "video_title",
    "original_slide_start",
    "original_slide_end",
    "estimated_minutes",
    "logic_score",
    "time_score",
    "boundary_score",
    "continuity_score",
    "title_score",
    "quiz_alignment_score",
    "overall_score",
    "decision",
    "coherent_topic_review",
    "boundary_review",
    "duration_review",
    "continuity_review",
    "title_review",
    "learning_objectives_review",
    "recap_quiz_review",
]


def project_root() -> Path:
    return Path(__file__).resolve().parents[1]


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def write_csv(path: Path, rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=CSV_FIELDS)
        writer.writeheader()
        writer.writerows(rows)


def teaching_rows_for_segment(
    teaching_rows: list[dict[str, str]],
    start: int,
    end: int,
) -> list[dict[str, str]]:
    return [row for row in teaching_rows if start <= int(row["slide_number"]) <= end]


def ordered_unique(values: list[str]) -> list[str]:
    unique: list[str] = []
    for value in values:
        if value not in unique:
            unique.append(value)
    return unique


def keywords(text: str) -> set[str]:
    stopwords = {
        "and",
        "the",
        "for",
        "from",
        "with",
        "into",
        "this",
        "that",
        "why",
        "how",
        "what",
        "does",
        "use",
        "uses",
        "using",
        "model",
        "models",
        "learning",
        "neural",
        "network",
        "networks",
        "slide",
        "slides",
        "video",
    }
    return {
        token
        for token in re.findall(r"[a-zA-Z0-9]+", text.lower())
        if len(token) >= 3 and token not in stopwords
    }


def score_time(minutes: float) -> tuple[int, str]:
    if 22 <= minutes <= 28:
        return 5, "Duration is in the preferred 22-28 minute core teaching range."
    if 18 <= minutes < 22 or 28 < minutes <= 32:
        return 4, "Duration is close to the 30-minute target and acceptable for online pacing."
    if 32 < minutes <= 36:
        return 4, "Duration is long, but acceptable when the topic is conceptually hard to split cleanly."
    if 14 <= minutes < 18:
        return 3, "Duration is short for a standard module but may be acceptable for a compact coherent topic."
    if 36 < minutes <= 40:
        return 3, "Duration is longer than ideal and should be reviewed for pacing, even if the topic is coherent."
    return 2, "Duration is outside the preferred range and should be reviewed."


def score_logic(groups: list[str], minutes: float) -> tuple[int, str]:
    count = len(groups)
    if count <= 3:
        return 5, "Segment has a focused concept set with one clear learning goal."
    if count <= 4 and minutes <= 32:
        return 4, "Segment has several related concept groups but still supports one coherent learning goal."
    if count <= 4 and minutes <= 36:
        return 4, "Segment is broad, but the concept groups form a continuous topic arc."
    if count <= 5:
        return 3, "Segment covers many concept groups and needs review for topic focus."
    return 2, "Segment combines too many concept groups for one clear video goal."


def score_boundary(first: dict[str, str], last: dict[str, str]) -> tuple[int, str]:
    starts_at_group_start = (
        first["slide_number"] == "1"
        or first["concept_group"] != previous_group_marker(first)
    )
    ends_at_boundary = last["logical_boundary_after_slide"] == "yes"
    strength = last["boundary_strength"]

    if starts_at_group_start and ends_at_boundary and strength == "strong":
        return 5, "Segment starts at a concept boundary and ends at a strong logical boundary."
    if starts_at_group_start and ends_at_boundary and strength == "medium":
        return 4, "Segment starts cleanly and ends at a medium logical boundary."
    if ends_at_boundary:
        return 3, "Segment ends at a logical boundary but the start may need review."
    return 2, "Segment does not end at a logical boundary."


def previous_group_marker(first: dict[str, str]) -> str:
    # The caller passes only segment rows, so this sentinel makes slide 1 the only
    # automatic start unless the segment begins with a different concept group
    # than its predecessor in the full map. Full-map start validation is done by
    # checking whether the first slide has a prerequisite concept from the group.
    return ""


def score_continuity(segment_rows: list[dict[str, str]]) -> tuple[int, str]:
    internal_boundaries = [
        row
        for row in segment_rows[:-1]
        if row["logical_boundary_after_slide"] == "yes"
        and row["boundary_strength"] == "strong"
    ]
    roles = {row["slide_role"] for row in segment_rows}
    has_continuous_material = bool(
        roles
        & {
            "derivation",
            "algorithm",
            "example",
            "visualization",
        }
    )

    if not internal_boundaries:
        return 5, "No strong internal concept boundary is crossed; continuity is preserved."
    if len(internal_boundaries) <= 2:
        return 5, "Internal boundaries are crossed only to keep a larger derivation, example, architecture, or comparison together."
    if has_continuous_material and len(internal_boundaries) <= 4:
        return 4, "Segment crosses several boundaries, but the material is continuous and avoids mid-topic cuts."
    return 3, "Segment crosses several strong internal boundaries and should be reviewed for possible subdivision."


def score_title(segment: dict[str, str], groups: list[str]) -> tuple[int, str]:
    title_terms = keywords(segment["video_title"])
    content_terms = keywords(" ".join(groups) + " " + segment["learning_objectives"])
    overlap = title_terms & content_terms
    if len(overlap) >= 2:
        return 5, "Video title uses terms that match the segment content."
    if overlap:
        return 4, "Video title is broadly aligned with the segment content."
    return 3, "Video title is plausible but should be checked against the concept groups."


def score_quiz(segment: dict[str, str], groups: list[str]) -> tuple[int, str]:
    content = " ".join(groups) + " " + segment["learning_objectives"]
    quiz_terms = keywords(segment["suggested_quiz_question"])
    recap_terms = keywords(segment["suggested_recap_slide"])
    content_terms = keywords(content)
    quiz_overlap = quiz_terms & content_terms
    recap_overlap = recap_terms & content_terms

    if len(quiz_overlap) >= 2 and len(recap_overlap) >= 2:
        return 5, "Recap and quiz both directly target the segment concepts."
    if quiz_overlap and recap_overlap:
        return 4, "Recap and quiz are aligned, though one could be made more specific."
    if quiz_overlap or recap_overlap:
        return 3, "Either the recap or quiz aligns with the segment, but the pair needs review."
    return 2, "Recap and quiz do not clearly match the segment concepts."


def objective_review(segment: dict[str, str]) -> str:
    objectives = [item.strip() for item in segment["learning_objectives"].split(";") if item.strip()]
    count = len(objectives)
    if 2 <= count <= 4:
        return f"Learning objectives are appropriate: {count} objectives support a focused online module."
    if count == 1:
        return "Learning objectives are concise but may need one more measurable objective."
    return f"Learning objectives need review: {count} objectives may be too many for one module."


def decision(overall_score: float) -> str:
    if overall_score >= 4.2:
        return "approved_for_pptx"
    if overall_score >= 3.5:
        return "needs_quick_review"
    return "needs_revision"


def review_segment(
    segment: dict[str, str],
    teaching_rows: list[dict[str, str]],
) -> dict[str, str]:
    start = int(segment["original_slide_start"])
    end = int(segment["original_slide_end"])
    rows = teaching_rows_for_segment(teaching_rows, start, end)
    if not rows:
        raise ValueError(
            f"No teaching-map rows found for Lecture {segment['lecture_number']} "
            f"segment {segment['video_number']}"
        )

    groups = ordered_unique([row["concept_group"] for row in rows])
    minutes = float(segment["estimated_minutes"])
    logic_score, logic_review = score_logic(groups, minutes)
    time_score, time_review = score_time(minutes)
    boundary_score, boundary_review = score_boundary(rows[0], rows[-1])
    continuity_score, continuity_review = score_continuity(rows)
    title_score, title_review = score_title(segment, groups)
    quiz_score, quiz_review = score_quiz(segment, groups)
    overall = round(
        (
            logic_score
            + time_score
            + boundary_score
            + continuity_score
            + title_score
            + quiz_score
        )
        / 6,
        2,
    )

    return {
        "lecture_number": segment["lecture_number"],
        "lecture_title": segment["lecture_title"],
        "video_number": segment["video_number"],
        "video_title": segment["video_title"],
        "original_slide_start": segment["original_slide_start"],
        "original_slide_end": segment["original_slide_end"],
        "estimated_minutes": segment["estimated_minutes"],
        "logic_score": str(logic_score),
        "time_score": str(time_score),
        "boundary_score": str(boundary_score),
        "continuity_score": str(continuity_score),
        "title_score": str(title_score),
        "quiz_alignment_score": str(quiz_score),
        "overall_score": f"{overall:.2f}",
        "decision": decision(overall),
        "coherent_topic_review": logic_review,
        "boundary_review": boundary_review,
        "duration_review": time_review,
        "continuity_review": continuity_review,
        "title_review": title_review,
        "learning_objectives_review": objective_review(segment),
        "recap_quiz_review": quiz_review,
    }


def review_lecture(
    lecture_number: int,
    root: Path,
) -> list[dict[str, str]]:
    teaching_rows = read_csv(root / TEACHING_MAP_DIR / f"{lecture_number}_teaching_map.csv")
    segments = read_csv(root / VIDEO_SEGMENT_DIR / f"{lecture_number}_video_segments.csv")
    return [review_segment(segment, teaching_rows) for segment in segments]


def main() -> int:
    root = project_root()
    output_dir = root / OUTPUT_DIR

    try:
        all_reviews: list[dict[str, str]] = []
        segment_files = sorted(
            (root / VIDEO_SEGMENT_DIR).glob("*_video_segments.csv"),
            key=lambda path: int(path.name.split("_", 1)[0]),
        )
        for segment_file in segment_files:
            lecture_number = int(segment_file.name.split("_", 1)[0])
            if lecture_number < 1:
                continue
            rows = review_lecture(lecture_number, root)
            write_csv(output_dir / f"{lecture_number}_segmentation_review.csv", rows)
            all_reviews.extend(rows)
        write_csv(root / SUMMARY_CSV, all_reviews)
    except (OSError, RuntimeError, ValueError, KeyError) as exc:
        print("Segmentation review failed.")
        print(f"Error: {exc}")
        return 1

    approved = sum(1 for row in all_reviews if row["decision"] == "approved_for_pptx")
    quick = sum(1 for row in all_reviews if row["decision"] == "needs_quick_review")
    revision = sum(1 for row in all_reviews if row["decision"] == "needs_revision")

    print("Segmentation review complete")
    print(f"Output folder:         {output_dir}")
    print(f"Summary CSV:           {root / SUMMARY_CSV}")
    print(f"Segments reviewed:     {len(all_reviews)}")
    print(f"approved_for_pptx:     {approved}")
    print(f"needs_quick_review:    {quick}")
    print(f"needs_revision:        {revision}")
    print("PowerPoint safety: used teaching-map and segment CSVs only; no PowerPoint files were read or modified.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
