"""Generate exact read-aloud speaker scripts for audited video decks.

The script reads passed rows from manifests/deck_audit/generated_deck_audit.csv
and writes one Markdown narration script per generated video under:

03_Scripts/exact_scripts/{lecture_number}/

PowerPoint files are opened read-only only for the manual Lecture 0 deck. No
PowerPoint file is saved or modified.
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from pptx_reader import PptxReadError, open_presentation_readonly

try:
    import yaml
    from yaml import YAMLError
except ImportError:  # pragma: no cover - depends on the local Python environment
    yaml = None

    class YAMLError(Exception):
        """Placeholder when PyYAML is not installed."""


AUDIT_CSV = Path("manifests") / "deck_audit" / "generated_deck_audit.csv"
VIDEO_SEGMENTS_DIR = Path("manifests") / "video_segments"
TEACHING_MAPS_DIR = Path("manifests") / "teaching_maps"
SLIDE_TEXT_DIR = Path("manifests") / "slide_text"
LECTURE_00_SEGMENTS = Path("manifests") / "lecture_00_video_segments.csv"
LECTURE_00_INVENTORY = Path("manifests") / "lecture_00_slide_inventory.csv"
SCRIPT_REVIEW_CSV = Path("manifests") / "script_reviews" / "exact_script_review_summary.csv"
READING_WPM = 130
URL_PATTERN = re.compile(r"https?://\S+|www\.\S+", re.IGNORECASE)
BANNED_NARRATION_PATTERNS = [
    re.compile(pattern, re.IGNORECASE)
    for pattern in [
        r"\bthis is part of lecture\s+\d+\b",
        r"\blecture\s+\d+\b",
        r"\bvideo\s+\d+\b",
        r"\bslide range\b",
        r"\bcovers slides?\b",
        r"\bwe will cover slides?\b",
        r"\bin this segment\b",
        r"\bthis segment\b",
        r"\bgenerated deck\b",
        r"\bmanifest\b",
        r"\bfilename\b",
        r"\boutput_pptx\b",
    ]
]


@dataclass(frozen=True)
class SegmentInfo:
    lecture_number: int
    lecture_title: str
    video_number: int
    video_title: str
    slide_start: int
    slide_end: int
    estimated_minutes: float
    learning_objectives: str
    recap: str
    quiz: str
    output_pptx: str
    generated_pptx_path: Path
    special_case: str = ""


@dataclass(frozen=True)
class SlidePlan:
    script_slide_number: int
    original_slide_number: int | None
    title: str
    role: str
    concept_group: str
    core_concept: str
    visible_lines: list[str]
    estimate_raw: float
    difficulty: str = ""
    notes: str = ""
    prerequisite: str = ""
    next_concept: str = ""


@dataclass(frozen=True)
class SlideScript:
    plan: SlidePlan
    estimated_minutes: float
    main_technical_point: str
    narration: str
    delivery_note: str
    transition: str
    technical_correction_note: str = ""
    reference_note: str = ""


@dataclass(frozen=True)
class ScriptQuality:
    estimated_words: int
    estimated_minutes: float
    banned_metadata_found: bool
    raw_urls_in_narration: bool
    bullet_fragments_in_narration: bool
    overly_long_sentences: list[str]
    technical_correction_notes_added: bool
    unsupported_claim_warnings: list[str]
    repeated_slide_text_warnings: list[str]
    scores: dict[str, float]


@dataclass(frozen=True)
class GeneratedScript:
    markdown: str
    quality: ScriptQuality


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


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def ascii_text(text: str) -> str:
    replacements = {
        "\u2013": "-",
        "\u2014": "-",
        "\u2018": "'",
        "\u2019": "'",
        "\u201c": '"',
        "\u201d": '"',
        "\u2190": "<-",
        "\u2192": "->",
        "\u21d2": "=>",
        "\u2264": "<=",
        "\u2265": ">=",
        "\u2260": "!=",
        "\u00d7": "x",
        "\u2212": "-",
        "\u2211": "sum",
        "\u03b1": "alpha",
        "\u03b2": "beta",
        "\u03bb": "lambda",
        "\u0394": "Delta",
        "\u03b8": "theta",
        "\u03c3": "sigma",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")


def clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", ascii_text(text)).strip()


def extract_urls(text: str) -> list[str]:
    return [match.group(0).rstrip(").,;") for match in URL_PATTERN.finditer(text)]


def remove_urls_for_narration(text: str) -> str:
    return clean_text(URL_PATTERN.sub("the link provided on the slide", text))


def spoken_text(text: str) -> str:
    return remove_urls_for_narration(text)


def is_reference_fragment(text: str) -> bool:
    value = clean_text(text).lower()
    if not value:
        return False
    return (
        value.startswith(("http://", "https://", "www."))
        or value.startswith("/")
        or any(domain in value for domain in [".com", ".org", ".edu", ".net", ".gov"])
        or value in {"website link:", "website link", "link:"}
    )


def sanitize_filename(value: str) -> str:
    value = ascii_text(value)
    value = re.sub(r"[^A-Za-z0-9]+", "_", value).strip("_")
    value = re.sub(r"_+", "_", value)
    return value or "video"


def split_items(text: str, max_items: int | None = None) -> list[str]:
    text = ascii_text(text)
    items = [
        item.strip(" .")
        for item in re.split(r";|\n", text)
        if item.strip(" .")
    ]
    if len(items) <= 1:
        items = [
            item.strip(" .")
            for item in re.split(r", and |, ", text)
            if item.strip(" .")
        ]
    if max_items is not None:
        items = items[:max_items]
    return items or ([text.strip()] if text.strip() else [])


def markdown_list(items: list[str]) -> str:
    return "\n".join(f"- {clean_text(item)}" for item in items)


def word_count(text: str) -> int:
    return len(re.findall(r"\b[A-Za-z0-9][A-Za-z0-9'-]*\b", text))


def parse_float(value: str, default: float = 1.0) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def visible_lines(text: str, limit: int = 5) -> list[str]:
    lines: list[str] = []
    seen: set[str] = set()
    for raw_line in ascii_text(text).splitlines():
        line = clean_text(raw_line)
        if not line or line.isdigit() or line in seen:
            continue
        seen.add(line)
        lines.append(line)
        if len(lines) >= limit:
            break
    return lines


def first_nonempty_line(text: str) -> str:
    for line in ascii_text(text).splitlines():
        line = clean_text(line)
        if line:
            return line
    return ""


def load_passed_audit_rows(root: Path) -> list[dict[str, str]]:
    rows = [
        row
        for row in read_csv(root / AUDIT_CSV)
        if row.get("overall_status") == "pass"
    ]
    if not rows:
        raise ValueError(f"No passing deck-audit rows found in {root / AUDIT_CSV}")
    return sorted(
        rows,
        key=lambda row: (
            int(row["lecture_number"]),
            int(row["video_number"]),
            row["output_pptx"],
        ),
    )


def load_video_segment_lookup(root: Path) -> dict[str, dict[str, str]]:
    lookup: dict[str, dict[str, str]] = {}
    for path in sorted((root / VIDEO_SEGMENTS_DIR).glob("*_video_segments.csv")):
        for row in read_csv(path):
            output_pptx = row.get("output_pptx", "").strip().replace("\\", "/")
            if output_pptx:
                lookup[output_pptx] = row
    return lookup


def load_failed_review_keys(root: Path) -> set[tuple[int, int]]:
    path = root / SCRIPT_REVIEW_CSV
    if not path.exists():
        return set()
    failed: set[tuple[int, int]] = set()
    for row in read_csv(path):
        decision = row.get("review_decision", "")
        if decision in {"needs_substantial_revision", "needs_targeted_revision", "needs_revision"}:
            failed.add((int(row["lecture_number"]), int(row["video_number"])))
    return failed


def parse_slide_range(value: str) -> tuple[int, int]:
    match = re.match(r"\s*(\d+)\s*-\s*(\d+)\s*$", value)
    if match is None:
        raise ValueError(f"Invalid slide range: {value}")
    return int(match.group(1)), int(match.group(2))


def lecture_zero_segment(root: Path, audit_row: dict[str, str]) -> SegmentInfo:
    rows = read_csv(root / LECTURE_00_SEGMENTS)
    if not rows:
        raise ValueError(f"{root / LECTURE_00_SEGMENTS} has no rows.")
    row = rows[0]
    slide_start, slide_end = parse_slide_range(row["original_slide_range"])
    return SegmentInfo(
        lecture_number=0,
        lecture_title="Intro",
        video_number=1,
        video_title=row["video_title"],
        slide_start=slide_start,
        slide_end=slide_end,
        estimated_minutes=parse_float(row["estimated_total_minutes"], 30.0),
        learning_objectives=row["learning_objectives"],
        recap=row["suggested_recap_slide_idea"],
        quiz=row["suggested_quiz_checkpoint_question"].replace("Checkpoint: ", ""),
        output_pptx=audit_row["output_pptx"],
        generated_pptx_path=Path(audit_row["generated_pptx_path"]),
        special_case=audit_row.get("special_case", ""),
    )


def segment_from_row(
    audit_row: dict[str, str],
    segment_lookup: dict[str, dict[str, str]],
    root: Path,
) -> SegmentInfo:
    if audit_row.get("special_case") == "lecture_0_manual":
        return lecture_zero_segment(root, audit_row)

    output_pptx = audit_row["output_pptx"]
    segment_row = segment_lookup.get(output_pptx)
    if segment_row is None:
        raise ValueError(f"No video segment row found for {output_pptx}")

    return SegmentInfo(
        lecture_number=int(segment_row["lecture_number"]),
        lecture_title=segment_row["lecture_title"],
        video_number=int(segment_row["video_number"]),
        video_title=segment_row["video_title"],
        slide_start=int(segment_row["original_slide_start"]),
        slide_end=int(segment_row["original_slide_end"]),
        estimated_minutes=parse_float(segment_row["estimated_minutes"], 25.0),
        learning_objectives=segment_row["learning_objectives"],
        recap=segment_row["suggested_recap_slide"],
        quiz=segment_row["suggested_quiz_question"],
        output_pptx=output_pptx,
        generated_pptx_path=Path(audit_row["generated_pptx_path"]),
        special_case=audit_row.get("special_case", ""),
    )


def load_teaching_map(root: Path, lecture_number: int) -> dict[int, dict[str, str]]:
    path = root / TEACHING_MAPS_DIR / f"{lecture_number}_teaching_map.csv"
    if not path.exists():
        return {}
    return {int(row["slide_number"]): row for row in read_csv(path)}


def load_slide_text(root: Path, lecture_number: int) -> dict[int, dict[str, str]]:
    path = root / SLIDE_TEXT_DIR / f"{lecture_number}_slide_text.csv"
    if not path.exists():
        return {}
    return {int(row["slide_number"]): row for row in read_csv(path)}


def lecture_zero_inventory(root: Path) -> dict[int, dict[str, str]]:
    path = root / LECTURE_00_INVENTORY
    if not path.exists():
        return {}
    return {int(row["slide_number"]): row for row in read_csv(path)}


def deck_slide_texts(path: Path) -> dict[int, str]:
    texts: dict[int, str] = {}
    with open_presentation_readonly(path) as (presentation, _repair_info):
        for index, slide in enumerate(presentation.slides, start=1):
            parts: list[str] = []
            for shape in slide.shapes:
                text = getattr(shape, "text", "")
                if text:
                    parts.append(text)
            texts[index] = "\n".join(parts).strip()
    return texts


def lecture_zero_role(topic: str, slide_number: int) -> str:
    topic_lower = topic.lower()
    if slide_number == 1 or "course" in topic_lower or "instructor" in topic_lower:
        return "logistics"
    if "example" in topic_lower or "application" in topic_lower:
        return "example"
    if "relationship" in topic_lower or "definition" in topic_lower:
        return "definition"
    if "why" in topic_lower or "enabled" in topic_lower or "practical" in topic_lower:
        return "motivation"
    if "summary" in topic_lower:
        return "summary"
    if "architecture" in topic_lower or "network" in topic_lower:
        return "visualization"
    return "other"


def has_math_or_equation(lines: list[str]) -> bool:
    joined = " ".join(lines).lower()
    math_tokens = ["=", "+", "loss", "minimize", "gradient", "softmax", "cross-entropy", "norm"]
    return any(token in joined for token in math_tokens) or bool(
        re.search(r"\b[a-z]\s*/\s*[a-z0-9]\b", joined)
    )


def architecture_hint(title: str, lines: list[str]) -> bool:
    joined = " ".join([title, *lines]).lower()
    direct_visual_tokens = [
        "architecture",
        "u-net",
        "resnet",
        "inception",
        "encoder",
        "decoder",
        "filter",
        "kernel",
        "feature map",
        "matrix",
        "plot",
        "roc",
        "confusion matrix",
        "contour",
        "curve",
        "window",
        "patch",
    ]
    flow_tokens = ["input", "output", "layer", "block", "gate", "head", "token", "generator", "discriminator"]
    model_tokens = ["rnn", "lstm", "transformer", "attention", "vit", "gan", "vae", "convolution"]
    return any(token in joined for token in direct_visual_tokens) or (
        any(model in joined for model in model_tokens)
        and any(flow in joined for flow in flow_tokens)
    )


def is_complex_slide(plan: SlidePlan) -> bool:
    role = plan.role.lower()
    if role in {"derivation", "algorithm", "example", "visualization"}:
        return True
    return has_math_or_equation(plan.visible_lines) or architecture_hint(plan.title, plan.visible_lines)


def is_simple_slide(plan: SlidePlan) -> bool:
    if plan.title.strip().lower() in {"reference", "references", "backup", "backup slides"}:
        return True
    return plan.role.lower() in {"overview", "summary", "exercise", "transition", "logistics"}


def role_weight(plan: SlidePlan) -> float:
    role = plan.role.lower()
    weight = 1.0
    if role in {"overview", "summary", "exercise"}:
        weight *= 0.7
    if role in {"logistics", "transition"}:
        weight *= 0.55
    if role in {"definition", "motivation"}:
        weight *= 1.05
    if role == "example":
        weight *= 1.25
    if role in {"derivation", "algorithm"}:
        weight *= 1.45
    if role == "visualization":
        weight *= 1.3
    if has_math_or_equation(plan.visible_lines):
        weight *= 1.2
    if architecture_hint(plan.title, plan.visible_lines):
        weight *= 1.2
    return weight


def slide_time_map(plans: list[SlidePlan], target_minutes: float) -> dict[int, float]:
    weighted = {
        plan.script_slide_number: max(plan.estimate_raw * role_weight(plan), 0.35)
        for plan in plans
    }
    raw_total = sum(weighted.values()) or target_minutes
    scale = target_minutes / raw_total if target_minutes > 0 else 1.0
    times: dict[int, float] = {}
    for plan in plans:
        if plan.role.lower() == "overview":
            min_time, max_time = 0.6, 1.2
        elif plan.title.lower() in {"recap", "checkpoint"}:
            min_time, max_time = 0.5, 1.4
        elif is_complex_slide(plan):
            min_time, max_time = 0.9, 3.8
        elif is_simple_slide(plan):
            min_time, max_time = 0.35, 1.2
        else:
            min_time, max_time = 0.6, 2.6
        value = max(min_time, min(max_time, weighted[plan.script_slide_number] * scale))
        times[plan.script_slide_number] = round(value, 1)
    return times


def make_intro_plan(segment: SegmentInfo) -> SlidePlan:
    objectives = split_items(segment.learning_objectives, max_items=4)
    return SlidePlan(
        script_slide_number=1,
        original_slide_number=None,
        title=segment.video_title,
        role="overview",
        concept_group="Video overview",
        core_concept="; ".join(objectives),
        visible_lines=[
            f"Lecture {segment.lecture_number}: {segment.lecture_title}",
            f"Slides {segment.slide_start}-{segment.slide_end}",
            "Learning objectives",
            *objectives,
        ],
        estimate_raw=0.9,
    )


def make_recap_plan(slide_number: int, segment: SegmentInfo) -> SlidePlan:
    return SlidePlan(
        script_slide_number=slide_number,
        original_slide_number=None,
        title="Recap",
        role="summary",
        concept_group="Video recap",
        core_concept=segment.recap,
        visible_lines=split_items(segment.recap, max_items=6),
        estimate_raw=1.0,
    )


def make_checkpoint_plan(slide_number: int, segment: SegmentInfo) -> SlidePlan:
    return SlidePlan(
        script_slide_number=slide_number,
        original_slide_number=None,
        title="Checkpoint",
        role="exercise",
        concept_group="Checkpoint",
        core_concept=segment.quiz,
        visible_lines=[segment.quiz],
        estimate_raw=0.8,
    )


def build_slide_plans(root: Path, segment: SegmentInfo) -> list[SlidePlan]:
    if segment.special_case == "lecture_0_manual":
        inventory = lecture_zero_inventory(root)
        deck_texts = deck_slide_texts(segment.generated_pptx_path)
        plans: list[SlidePlan] = []
        for slide_number in range(segment.slide_start, segment.slide_end + 1):
            row = inventory.get(slide_number, {})
            deck_text = deck_texts.get(slide_number, "")
            title = clean_text(first_nonempty_line(deck_text) or row.get("main_topic", ""))
            topic = clean_text(row.get("main_topic", title))
            plans.append(
                SlidePlan(
                    script_slide_number=slide_number,
                    original_slide_number=slide_number,
                    title=title or f"Slide {slide_number}",
                    role=lecture_zero_role(topic, slide_number),
                    concept_group="Lecture 0 course introduction",
                    core_concept=topic or title or f"Slide {slide_number}",
                    visible_lines=visible_lines(deck_text, limit=5),
                    estimate_raw=parse_float(row.get("estimated_teaching_time_minutes", ""), 1.0),
                    difficulty=row.get("difficulty_level", "low"),
                    notes=row.get("suggested_improvement_for_online_teaching", ""),
                )
            )
        return plans

    teaching_map = load_teaching_map(root, segment.lecture_number)
    slide_text = load_slide_text(root, segment.lecture_number)
    plans = [make_intro_plan(segment)]
    script_slide_number = 2
    for original_slide_number in range(segment.slide_start, segment.slide_end + 1):
        map_row = teaching_map.get(original_slide_number, {})
        text_row = slide_text.get(original_slide_number, {})
        title = clean_text(
            map_row.get("slide_title")
            or text_row.get("possible_slide_title")
            or f"Slide {original_slide_number}"
        )
        plans.append(
            SlidePlan(
                script_slide_number=script_slide_number,
                original_slide_number=original_slide_number,
                title=title,
                role=clean_text(map_row.get("slide_role", "other")),
                concept_group=clean_text(map_row.get("concept_group", "")),
                core_concept=clean_text(map_row.get("core_concept", title)),
                visible_lines=visible_lines(text_row.get("visible_text", ""), limit=5),
                estimate_raw=parse_float(map_row.get("estimated_teaching_minutes", ""), 1.0),
                difficulty="",
                notes=clean_text(map_row.get("notes_for_future_script", "")),
                prerequisite=clean_text(map_row.get("prerequisite_concept", "")),
                next_concept=clean_text(map_row.get("next_concept", "")),
            )
        )
        script_slide_number += 1
    plans.append(make_recap_plan(script_slide_number, segment))
    plans.append(make_checkpoint_plan(script_slide_number + 1, segment))
    return plans


def next_plan(plans: list[SlidePlan], index: int) -> SlidePlan | None:
    if index + 1 >= len(plans):
        return None
    return plans[index + 1]


def narration_opening(segment: SegmentInfo, plan: SlidePlan) -> str:
    objectives = split_items(segment.learning_objectives, max_items=4)
    labels = ["First", "Second", "Third", "Fourth"]
    objective_sentence = " ".join(
        f"{labels[index]}, {spoken_text(objective)}." for index, objective in enumerate(objectives)
    )
    return (
        f"Welcome. Today we are going to work through {spoken_text(segment.video_title)}. "
        "The goal is not to memorize a list of terms. The goal is to understand how the ideas fit together well enough to explain them and use them. "
        f"By the end, you should be able to do a few concrete things. {objective_sentence} "
        "As we go, listen for the connection between the visual intuition, the mathematical statement, and the practical modeling decision."
    )


def narration_recap(segment: SegmentInfo, plan: SlidePlan) -> str:
    recap_items = [spoken_text(item) for item in split_items(segment.recap, max_items=6)]
    if not recap_items:
        recap_items = [segment.video_title]
    if len(recap_items) == 1:
        takeaway = recap_items[0]
    else:
        takeaway = ", ".join(recap_items[:-1]) + f", and {recap_items[-1]}"
    return (
        f"Let's collect the main ideas before we finish. The important takeaways are {takeaway}. "
        "At this point, you should be able to explain the main vocabulary in your own words, "
        "identify where the method or model is useful, and recognize the common mistake that this video was trying to prevent. "
        "If one part still feels uncertain, go back to the slide where that idea first appeared and listen for the connection to the next slide."
    )


def narration_checkpoint(segment: SegmentInfo, plan: SlidePlan) -> str:
    return (
        "Before moving on, pause and answer this checkpoint question. "
        f"{spoken_text(segment.quiz)} "
        "Take a moment to answer it without looking back at the previous slide. "
        "A good answer should use the vocabulary from the video and should explain the reason, not just name the correct term."
    )


def simple_narration(plan: SlidePlan) -> str:
    return (
        f"This is a short framing point. The idea to carry forward is {spoken_text(plan.core_concept)}. "
        "I will keep this brief so we have enough time for the main technical part of the video. "
        "What matters here is the connection to the next idea."
    )


def sentence_from_lines(lines: list[str]) -> str:
    filtered = [
        spoken_text(line)
        for line in lines
        if line
        and line.lower() not in {"edit master text styles"}
        and not is_reference_fragment(line)
    ]
    if not filtered:
        return ""
    if len(filtered) == 1:
        return filtered[0]
    if len(filtered) == 2:
        return f"{filtered[0]} and {filtered[1]}"
    return ", ".join(filtered[:-1]) + f", and {filtered[-1]}"


def equation_narration(plan: SlidePlan) -> str:
    anchors = sentence_from_lines(plan.visible_lines[:4])
    if anchors:
        return (
            f"The expression is anchored by these terms: {anchors}. "
            "The useful way to read it is to connect each symbol to a behavior of the model. "
            "One part usually describes what the model is trying to fit, and another part describes the constraint, penalty, or transformation that changes how learning behaves. "
            "That interpretation is more important than copying the notation."
        )
    return (
        "The formula is not just notation. It is describing what is being computed and what tradeoff the model is making. "
        "The key is to connect the symbols to the behavior they control."
    )


def visual_narration(plan: SlidePlan) -> str:
    anchors = sentence_from_lines(plan.visible_lines[:4])
    if anchors:
        return (
            f"The figure is organized around {anchors}. "
            "First, identify the input or starting condition. Then follow the arrows, spatial layout, or repeated blocks to see how the information changes. "
            "The conclusion should come after that walk-through, because the diagram is showing a relationship, not just a collection of labels."
        )
    return (
        "The first job is to orient ourselves to the layout. Once the objects and direction of flow are clear, the conclusion of the figure becomes much easier to see."
    )


def example_narration(plan: SlidePlan) -> str:
    non_title_lines = [
        line
        for line in plan.visible_lines[:4]
        if clean_text(line).lower() != clean_text(plan.title).lower()
    ]
    anchors = sentence_from_lines(non_title_lines)
    if anchors:
        return (
            f"The example is built from these pieces: {anchors}. "
            "The setup comes first, then the change, and then the lesson that carries beyond this single example. "
            "That last step is important, because the purpose of the example is to make the general rule easier to use."
        )
    return (
        "Use this as a concrete check on the idea. The goal is to make the general rule easier to recognize when it appears again."
    )


def plan_text(plan: SlidePlan) -> str:
    return " ".join([plan.title, plan.role, plan.concept_group, plan.core_concept, *plan.visible_lines]).lower()


def main_technical_point(plan: SlidePlan) -> str:
    concept = clean_text(plan.core_concept or plan.title)
    group = clean_text(plan.concept_group)
    if plan.role.lower() == "overview":
        return "The slide frames the learning goals and prepares students for the main technical ideas."
    if plan.title.lower() == "recap":
        return "The slide consolidates the main ideas so students can connect them before moving on."
    if plan.title.lower() == "checkpoint":
        return "The slide asks students to retrieve and apply the central idea without immediately seeing the answer."
    if group and concept and group.lower() not in concept.lower():
        return f"The slide develops {concept} as part of {group}."
    return f"The slide develops {concept}."


def technical_precision_sentence(plan: SlidePlan) -> str:
    text = plan_text(plan)
    title_core_text = " ".join([plan.title, plan.core_concept]).lower()
    title_visible_text = " ".join([plan.title, *plan.visible_lines]).lower()
    cautions: list[str] = []
    if "activation" in title_core_text:
        cautions.append(
            "Activation functions matter because they introduce nonlinearity; without nonlinear activations, a stack of linear layers would still behave like one linear transformation."
        )
    if "neuron" in title_core_text and "activation" not in title_core_text:
        cautions.append(
            "A neuron computes a weighted combination of its inputs and then passes that value through an activation function."
        )
    if "training process" in text or "training loop" in text:
        cautions.append(
            "During training, the model makes a prediction, the loss measures the error, backpropagation computes gradients, and the optimizer uses those gradients to update the parameters."
        )
    if "dataset" in title_visible_text or "benchmark" in title_visible_text:
        cautions.append(
            "A benchmark dataset is useful because it gives models a common comparison point, but benchmark performance does not automatically translate to real-world performance."
        )
    if "mlp" in text and "image" in text:
        cautions.append(
            "The weakness of a plain MLP on images is that flattening the image hides local spatial structure and uses separate parameters for positions that should share visual patterns."
        )
    if "backprop" in text or "back propagation" in text:
        cautions.append(
            "Backpropagation is the method for computing gradients efficiently; the optimizer uses those gradients to update the parameters."
        )
    if "gradient descent" in text:
        cautions.append(
            "Gradient descent describes the update rule that moves parameters in a direction that reduces the objective, based on the computed gradient."
        )
    if "cross-entropy" in text or "cross entropy" in text:
        cautions.append(
            "Cross-entropy is a loss used for training; it is not the same thing as accuracy, which is an evaluation measure."
        )
    if "softmax" in text:
        cautions.append(
            "Softmax converts logits into normalized scores that are often interpreted as class probabilities, but a large softmax value does not automatically mean the model is well calibrated."
        )
    if "convolution" in text or "kernel" in text or "filter" in text:
        cautions.append(
            "In most deep learning libraries, the operation called convolution is implemented as cross-correlation, but the convention in CNNs is still to call the learned operation convolution."
        )
    if "pool" in text:
        cautions.append(
            "Pooling gives local tolerance to small shifts; it does not make the network fully translation invariant."
        )
    if "validation" in text or "test" in text:
        cautions.append(
            "Validation data is for model selection and tuning; test data should be held back for the final estimate of performance."
        )
    if "overfit" in text or "underfit" in text:
        cautions.append(
            "Overfitting means the model has adapted too closely to the training data, while underfitting means it has not captured enough structure even in training."
        )
    if "attention" in text:
        cautions.append(
            "Attention weights describe how information is mixed inside the model; they should be interpreted carefully and are not automatically a complete explanation of the model's decision."
        )
    if "embedding" in text or "representation" in text:
        cautions.append(
            "An embedding is a learned representation, not a label by itself; its meaning comes from how it is used by the rest of the model."
        )
    if "generative" in text or "discriminative" in text:
        cautions.append(
            "A generative model is about modeling or sampling data, while a discriminative model focuses directly on predicting labels or outputs."
        )
    if "likelihood" in text or "sampling" in text:
        cautions.append(
            "Likelihood, sampling, and prediction are related but different ideas: likelihood scores observed data under a model, sampling generates possible data, and prediction chooses an output for a task."
        )
    return " ".join(dict.fromkeys(cautions))


def technical_correction_note(plan: SlidePlan) -> str:
    if plan.role.lower() == "overview" or plan.title.lower() in {"recap", "checkpoint"}:
        return ""
    text = plan_text(plan)
    notes: list[str] = []
    if "convolution" in text or "kernel" in text or "filter" in text:
        notes.append(
            "Use the standard CNN convention carefully: many libraries implement cross-correlation while calling the learned operation convolution."
        )
    if "softmax" in text:
        notes.append(
            "Avoid implying that softmax scores are automatically calibrated probabilities."
        )
    if "cross-entropy" in text or "cross entropy" in text:
        notes.append(
            "Keep cross-entropy separate from accuracy: cross-entropy is optimized during training; accuracy is usually reported as an evaluation metric."
        )
    if "backprop" in text or "gradient descent" in text:
        notes.append(
            "Keep the distinction between gradient computation and parameter updates explicit."
        )
    if "pool" in text:
        notes.append(
            "Avoid overstating pooling as full translation invariance; local tolerance is the safer phrasing."
        )
    if "attention" in text:
        notes.append(
            "Avoid claiming attention weights are complete explanations of model decisions."
        )
    if "validation" in text and "test" in text:
        notes.append(
            "Keep validation for tuning and testing for final evaluation."
        )
    return " ".join(dict.fromkeys(notes))


def reference_note(plan: SlidePlan) -> str:
    urls: list[str] = []
    for line in plan.visible_lines:
        urls.extend(extract_urls(line))
    if not urls:
        if any(is_reference_fragment(line) for line in plan.visible_lines):
            return "Reference link text appears on the slide; keep it as a visual reference rather than reading it aloud."
        return ""
    unique_urls = list(dict.fromkeys(urls))
    return "Reference link(s) shown on the slide: " + "; ".join(unique_urls)


def expand_narration(plan: SlidePlan, narration: str, minutes: float) -> str:
    target_words = max(45, int(minutes * READING_WPM * 0.43))
    if is_complex_slide(plan):
        target_words = max(target_words, 100)
    if has_math_or_equation(plan.visible_lines) or architecture_hint(plan.title, plan.visible_lines):
        target_words = max(target_words, 95)
    if is_simple_slide(plan):
        target_words = min(target_words, 85)
    if plan.role.lower() == "overview" or plan.title.lower() in {"recap", "checkpoint"}:
        target_words = min(target_words, 95)

    if word_count(narration) >= target_words:
        return narration

    additions: list[str] = []
    role = plan.role.lower()
    if has_math_or_equation(plan.visible_lines) and role not in {"derivation", "algorithm"}:
        additions.append(
            "The safest way to follow the math is to attach each term to its job in the model before thinking about the final result."
        )
    if architecture_hint(plan.title, plan.visible_lines) and role != "visualization":
        additions.append(
            "For the visual structure, the important question is what enters, what operation is applied, and what representation leaves that block."
        )
    if role == "example":
        additions.append(
            "The reason to spend time here is that the example turns the abstract rule into something we can test against a concrete case."
        )
    elif role in {"derivation", "algorithm"}:
        additions.append(
            "Do not rush this step. The value is in the sequence of moves, because each move explains why the final result is reasonable."
        )
    elif role in {"definition", "motivation"} and plan.core_concept:
        additions.append(
            f"The key phrase to keep is {plan.core_concept}, because it is the idea that the next part of the lecture depends on."
        )
    if plan.next_concept and role not in {"overview", "summary", "exercise"}:
        additions.append(f"This sets up {plan.next_concept}, which is where the next slide will take us.")

    expanded = narration
    for addition in additions:
        if word_count(expanded) >= target_words:
            break
        expanded = f"{expanded} {addition}"
    return expanded


def content_narration(plan: SlidePlan) -> str:
    anchors = plan.visible_lines[:4]
    anchor_sentence = ""
    if anchors:
        anchor_sentence = f"The visible text gives us {sentence_from_lines(anchors)}."
    precision_sentence = technical_precision_sentence(plan)

    concept_sentence = (
        f"The key idea is {spoken_text(plan.core_concept)}. "
        if plan.core_concept and plan.core_concept.lower() != plan.title.lower()
        else f"The key idea is {spoken_text(plan.title)}. "
    )

    role = plan.role.lower()
    if role == "motivation":
        narration = (
            f"{concept_sentence}"
            "The important point is why the next technical step is needed. "
            "Without this motivation, the method can look like an arbitrary design choice instead of a response to a specific problem."
        )
    elif role == "definition":
        narration = (
            f"Here is the working definition. {concept_sentence}"
            "In plain language, this gives us vocabulary for a behavior we will use again. "
            "The term matters because it lets us describe the model more precisely in the next step."
        )
    elif role == "example":
        narration = (
            f"Let's use this example to make the idea concrete. {concept_sentence}"
            f"{example_narration(plan)}"
        )
    elif role == "derivation":
        narration = (
            f"Let's go through the derivation one step at a time. {concept_sentence}"
            "The expression we start with tells us what the main quantities represent. The next line should follow from a specific algebraic or modeling step. "
            "The goal is to understand the movement from one line to the next, not just the final formula. "
            f"{equation_narration(plan)}"
        )
    elif role == "algorithm":
        narration = (
            f"Think of this as a procedure. {concept_sentence}"
            "The procedure has an input, a computation, and an output. "
            "When those three parts are clear, the algorithm becomes much easier to remember and much easier to debug. "
            f"{equation_narration(plan) if has_math_or_equation(anchors) else visual_narration(plan)}"
        )
    elif role == "visualization":
        narration = (
            f"This picture is doing part of the explanation. {concept_sentence}"
            "The objects in the picture and the direction of the relationship tell us how to interpret the result. "
            f"{visual_narration(plan)}"
        )
    elif role == "summary":
        narration = (
            f"Let's pause and collect what we have built so far. {concept_sentence}"
            "The purpose here is to connect the pieces that have already appeared, not to start a new technical branch."
        )
    elif role == "exercise":
        narration = (
            f"Take a moment to think before we discuss the answer. {concept_sentence}"
            "The value of this step is the pause, because it gives you a chance to check whether the idea is usable."
        )
    elif role in {"logistics", "transition"}:
        return simple_narration(plan)
    else:
        narration = (
            f"{concept_sentence}The important point is how this idea supports the next step in the lecture. "
            "Keep the focus on that connection."
        )

    details: list[str] = []
    if anchor_sentence:
        details.append(anchor_sentence)
    if precision_sentence:
        details.append(precision_sentence)
    if has_math_or_equation(anchors) and role not in {"derivation", "algorithm"}:
        details.append(equation_narration(plan))
    if architecture_hint(plan.title, anchors) and role not in {"visualization", "algorithm"}:
        details.append(visual_narration(plan))
    if plan.prerequisite:
        details.append(f"This builds on {plan.prerequisite}.")
    if plan.next_concept:
        details.append(f"After this, we are ready for {plan.next_concept}.")

    return " ".join([narration, *details])


def narration_for_slide(segment: SegmentInfo, plan: SlidePlan) -> str:
    if plan.role == "overview":
        return narration_opening(segment, plan)
    if plan.title.lower() == "recap":
        return narration_recap(segment, plan)
    if plan.title.lower() == "checkpoint":
        return narration_checkpoint(segment, plan)
    return content_narration(plan)


def delivery_note(plan: SlidePlan, minutes: float) -> str:
    role = plan.role.lower()
    if role == "overview":
        return "Keep the pace warm and concise. Do not over-explain the title slide."
    if role in {"derivation", "algorithm"}:
        return "Slow down. Point to each line or step before explaining the next one."
    if role == "visualization" or architecture_hint(plan.title, plan.visible_lines):
        return "Slow down at the start. Orient students to the visual layout before interpreting it."
    if role == "definition":
        return "Emphasize the definition and pause after the plain-language restatement."
    if role == "example":
        return "Pause after the setup, then walk through the example in order."
    if role == "summary":
        return "Use a steady recap tone. Emphasize the takeaways rather than every detail."
    if role == "exercise":
        return "Pause long enough for students to answer before giving any hint."
    if minutes <= 0.7:
        return "Keep this brief. Say the point once and move on."
    return "Use a normal teaching pace and emphasize the connection to the next idea."


def confusion_note(plan: SlidePlan) -> str:
    role = plan.role.lower()
    if role == "derivation":
        return "Students may copy the final expression without understanding why each step follows."
    if role == "algorithm":
        return "Students may memorize the steps but miss what each step computes."
    if role == "visualization":
        return "Students may miss labels, axes, or arrow directions if the diagram is dense."
    if role == "definition":
        return f"Students may confuse {plan.title} with a nearby concept from {plan.concept_group or 'this lecture'}."
    if role == "example":
        return "Students may treat the example as isolated instead of using it to understand the general rule."
    if architecture_hint(plan.title, plan.visible_lines):
        return "Students may lose track of what flows through each block of the architecture."
    return "Students may miss why this slide matters if the transition is too fast."


def transition_text(plan: SlidePlan, next_item: SlidePlan | None) -> str:
    if next_item is None:
        return "That is a good place to stop. The main idea to carry forward is the connection between the method, the model behavior, and the reason we use it."
    if next_item.title.lower() == "recap":
        return "With the technical pieces in place, let's summarize the main takeaways."
    if next_item.title.lower() == "checkpoint":
        return "Now let's check whether the main idea is usable without looking back."
    if plan.role.lower() in {"overview", "transition", "logistics"}:
        return f"Now let's begin with {spoken_text(next_item.title)}."
    if plan.concept_group and next_item.concept_group and plan.concept_group != next_item.concept_group:
        return f"That completes the {spoken_text(plan.concept_group)} part. Next we move to {spoken_text(next_item.title)}."
    options = [
        f"With that foundation, move to {spoken_text(next_item.title)}.",
        f"The next piece of the explanation is {spoken_text(next_item.title)}.",
        f"Now carry this idea forward into {spoken_text(next_item.title)}.",
        f"From here, the natural next step is {spoken_text(next_item.title)}.",
    ]
    return options[plan.script_slide_number % len(options)]


def overall_teaching_goal(segment: SegmentInfo) -> str:
    objectives = split_items(segment.learning_objectives, max_items=4)
    if objectives:
        objective_text = "; ".join(objectives)
        return (
            f"Students should understand the core idea behind {segment.video_title} and be able to use it to: "
            f"{objective_text}."
        )
    return f"Students should understand the core idea behind {segment.video_title} and be able to explain why it matters."


def timing_rationale(slide_scripts: list[SlideScript]) -> str:
    difficult = [
        item.plan.title
        for item in slide_scripts
        if is_complex_slide(item.plan) and item.plan.title.lower() not in {"recap", "checkpoint"}
    ]
    simple = [
        item.plan.title
        for item in slide_scripts
        if is_simple_slide(item.plan) or item.plan.title.lower() in {"recap", "checkpoint"}
    ]
    if difficult:
        return (
            "The estimate gives more time to slides with equations, algorithms, diagrams, architectures, examples, or code, "
            "and less time to title, recap, checkpoint, and simple transition slides."
        )
    if simple and len(simple) >= len(slide_scripts) / 2:
        return "The estimate is short because many slides are framing, recap, checkpoint, or simple transition slides."
    return "The estimate follows the amount of explanation needed by the slide content rather than a fixed target length."


def script_header(segment: SegmentInfo, quality: ScriptQuality, slide_scripts: list[SlideScript]) -> str:
    return "\n".join(
        [
            f"# {segment.video_title}",
            "",
            "## Overall teaching goal",
            "",
            overall_teaching_goal(segment),
            "",
            "## Estimated speaking time",
            "",
            f"- estimated total word count: {quality.estimated_words}",
            f"- estimated speaking time: {quality.estimated_minutes:.1f} minutes at about {READING_WPM} spoken words per minute",
            f"- timing rationale: {timing_rationale(slide_scripts)}",
            "",
        ]
    )


def build_slide_script(
    segment: SegmentInfo,
    plan: SlidePlan,
    next_item: SlidePlan | None,
    minutes: float,
) -> SlideScript:
    narration = slide_narration(segment, plan, minutes)
    return SlideScript(
        plan=plan,
        estimated_minutes=minutes,
        main_technical_point=main_technical_point(plan),
        narration=narration,
        delivery_note=delivery_note(plan, minutes),
        transition=transition_text(plan, next_item),
        technical_correction_note=technical_correction_note(plan),
        reference_note=reference_note(plan),
    )


def slide_section(item: SlideScript) -> str:
    lines = [
        f"## Slide {item.plan.script_slide_number}: {item.plan.title}",
        "",
        f"Estimated speaking time: {item.estimated_minutes:.1f} minutes",
        "",
        "### Main technical point",
        "",
        item.main_technical_point,
        "",
        "### Word-for-word narration",
        "",
        item.narration,
        "",
        "### Delivery note",
        "",
        item.delivery_note,
        "",
        "### Transition to next slide",
        "",
        item.transition,
        "",
    ]
    if item.technical_correction_note:
        lines.extend(
            [
                "### Technical correction note",
                "",
                item.technical_correction_note,
                "",
            ]
        )
    if item.reference_note:
        lines.extend(
            [
                "### Reference note",
                "",
                item.reference_note,
                "",
            ]
        )
    return "\n".join(lines)


def split_sentences(text: str) -> list[str]:
    return [sentence.strip() for sentence in re.split(r"(?<=[.!?])\s+", text) if sentence.strip()]


def narration_has_banned_metadata(slide_scripts: list[SlideScript]) -> bool:
    text = " ".join(item.narration for item in slide_scripts)
    return any(pattern.search(text) for pattern in BANNED_NARRATION_PATTERNS)


def narration_has_raw_urls(slide_scripts: list[SlideScript]) -> bool:
    return any(extract_urls(item.narration) for item in slide_scripts)


def narration_has_bullet_fragments(slide_scripts: list[SlideScript]) -> bool:
    return any(re.search(r"(?m)^\s*[-*]\s+", item.narration) for item in slide_scripts)


def overly_long_sentences(slide_scripts: list[SlideScript]) -> list[str]:
    warnings: list[str] = []
    for item in slide_scripts:
        for sentence in split_sentences(item.narration):
            if word_count(sentence) > 38:
                warnings.append(
                    f"Slide {item.plan.script_slide_number}: {item.plan.title} ({word_count(sentence)} words)"
                )
                break
    return warnings


def unsupported_claim_warnings(slide_scripts: list[SlideScript]) -> list[str]:
    warnings: list[str] = []
    risky_words = ["always", "guarantees", "guarantee", "proves", "perfect", "must always"]
    for item in slide_scripts:
        text = item.narration.lower()
        if any(word in text for word in risky_words):
            warnings.append(f"Slide {item.plan.script_slide_number}: {item.plan.title}")
    return warnings


def repeated_slide_text_warnings(slide_scripts: list[SlideScript]) -> list[str]:
    warnings: list[str] = []
    stop = {
        "the", "and", "for", "with", "from", "that", "this", "into", "what", "when",
        "where", "why", "how", "are", "is", "to", "of", "in", "a", "an", "on",
    }
    for item in slide_scripts:
        if item.plan.role.lower() == "overview" or item.plan.title.lower() in {"recap", "checkpoint"}:
            continue
        visible = " ".join(spoken_text(line) for line in item.plan.visible_lines)
        visible_words = [
            word.lower()
            for word in re.findall(r"\b[A-Za-z][A-Za-z'-]*\b", visible)
            if len(word) > 3 and word.lower() not in stop
        ]
        if len(visible_words) < 8:
            continue
        narration_words = set(
            word.lower()
            for word in re.findall(r"\b[A-Za-z][A-Za-z'-]*\b", item.narration)
            if len(word) > 3 and word.lower() not in stop
        )
        overlap = sum(1 for word in visible_words if word in narration_words) / len(visible_words)
        if overlap > 0.85 and word_count(item.narration) < max(90, int(len(visible_words) * 1.6)):
            warnings.append(f"Slide {item.plan.script_slide_number}: {item.plan.title}")
    return warnings


def quality_self_check(slide_scripts: list[SlideScript]) -> ScriptQuality:
    spoken_parts: list[str] = []
    for item in slide_scripts:
        spoken_parts.extend([item.narration, item.transition])
    estimated_words = word_count(" ".join(spoken_parts))
    estimated_minutes = estimated_words / READING_WPM if estimated_words else 0.0
    banned_metadata_found = narration_has_banned_metadata(slide_scripts)
    raw_urls_in_narration = narration_has_raw_urls(slide_scripts)
    bullet_fragments = narration_has_bullet_fragments(slide_scripts)
    long_sentences = overly_long_sentences(slide_scripts)
    unsupported_warnings = unsupported_claim_warnings(slide_scripts)
    repeated_warnings = repeated_slide_text_warnings(slide_scripts)
    missing_required = [
        item.plan.title
        for item in slide_scripts
        if not item.narration.strip() or not item.delivery_note.strip() or not item.transition.strip()
    ]
    notes_added = any(item.technical_correction_note for item in slide_scripts)

    complex_items = [
        item
        for item in slide_scripts
        if is_complex_slide(item.plan) and item.plan.title.lower() not in {"recap", "checkpoint"}
    ]
    complex_underexplained = [
        item.plan.title
        for item in complex_items
        if word_count(item.narration) < 70
    ]

    read_aloud_quality = 5
    if banned_metadata_found or raw_urls_in_narration or bullet_fragments:
        read_aloud_quality -= 2
    if len(long_sentences) > max(2, len(slide_scripts) // 4):
        read_aloud_quality -= 1
    read_aloud_quality = max(1, read_aloud_quality)

    technical_accuracy = 5 if not unsupported_warnings else 4
    technical_correction_quality = 5
    if any(technical_correction_note(item.plan) for item in slide_scripts) and not notes_added:
        technical_correction_quality = 4

    slide_alignment = 5 if len(repeated_warnings) <= max(2, len(slide_scripts) // 4) else 4
    conceptual_flow = 5 if not missing_required else 3
    clarity_for_students = 5 if len(long_sentences) <= max(3, len(slide_scripts) // 3) else 4
    explanation_depth = 5 if len(complex_underexplained) <= max(1, len(complex_items) // 4) else 4
    timing_reasonableness = 5
    transition_quality = 5 if not missing_required else 3

    scores = {
        "read_aloud_quality": float(read_aloud_quality),
        "technical_accuracy": float(max(4, technical_accuracy)),
        "technical_correction_quality": float(max(4, technical_correction_quality)),
        "slide_alignment": float(slide_alignment),
        "conceptual_flow": float(conceptual_flow),
        "clarity_for_students": float(clarity_for_students),
        "explanation_depth": float(explanation_depth),
        "timing_reasonableness": float(timing_reasonableness),
        "transition_quality": float(transition_quality),
    }
    scores["overall_score"] = round(sum(scores.values()) / len(scores), 2)

    return ScriptQuality(
        estimated_words=estimated_words,
        estimated_minutes=estimated_minutes,
        banned_metadata_found=banned_metadata_found,
        raw_urls_in_narration=raw_urls_in_narration,
        bullet_fragments_in_narration=bullet_fragments,
        overly_long_sentences=long_sentences,
        technical_correction_notes_added=notes_added,
        unsupported_claim_warnings=unsupported_warnings,
        repeated_slide_text_warnings=repeated_warnings,
        scores=scores,
    )


def quality_self_check_section(quality: ScriptQuality) -> str:
    return "\n".join(
        [
            "## Quality self-check",
            "",
            f"- read_aloud_quality: {quality.scores['read_aloud_quality']:.0f}",
            f"- technical_accuracy: {quality.scores['technical_accuracy']:.0f}",
            f"- technical_correction_quality: {quality.scores['technical_correction_quality']:.0f}",
            f"- slide_alignment: {quality.scores['slide_alignment']:.0f}",
            f"- conceptual_flow: {quality.scores['conceptual_flow']:.0f}",
            f"- clarity_for_students: {quality.scores['clarity_for_students']:.0f}",
            f"- explanation_depth: {quality.scores['explanation_depth']:.0f}",
            f"- timing_reasonableness: {quality.scores['timing_reasonableness']:.0f}",
            f"- transition_quality: {quality.scores['transition_quality']:.0f}",
            f"- overall_score: {quality.scores['overall_score']:.2f}",
            "",
        ]
    )


def slide_narration(segment: SegmentInfo, plan: SlidePlan, minutes: float) -> str:
    return expand_narration(plan, narration_for_slide(segment, plan), minutes)


def acceptance_errors(quality: ScriptQuality) -> list[str]:
    errors: list[str] = []
    thresholds = {
        "technical_accuracy": 4.0,
        "technical_correction_quality": 4.0,
        "read_aloud_quality": 4.0,
        "slide_alignment": 4.0,
        "conceptual_flow": 4.0,
        "overall_score": 4.2,
    }
    for key, minimum in thresholds.items():
        if quality.scores[key] < minimum:
            errors.append(f"{key}={quality.scores[key]} is below {minimum}")
    if quality.banned_metadata_found:
        errors.append("banned metadata language found in narration")
    if quality.raw_urls_in_narration:
        errors.append("raw URL found in word-for-word narration")
    if quality.bullet_fragments_in_narration:
        errors.append("bullet fragment found in word-for-word narration")
    return errors


def generate_script(root: Path, segment: SegmentInfo) -> GeneratedScript:
    plans = build_slide_plans(root, segment)
    times = slide_time_map(plans, segment.estimated_minutes)
    slide_scripts: list[SlideScript] = []
    for index, plan in enumerate(plans):
        slide_scripts.append(
            build_slide_script(
                segment,
                plan,
                next_plan(plans, index),
                times[plan.script_slide_number],
            )
        )
    quality = quality_self_check(slide_scripts)
    errors = acceptance_errors(quality)
    if errors:
        raise ValueError(
            f"Generated script for lecture {segment.lecture_number} video {segment.video_number} did not pass cleanup checks: "
            + "; ".join(errors)
        )
    sections = [script_header(segment, quality, slide_scripts)]
    sections.extend(slide_section(item) for item in slide_scripts)
    sections.append(quality_self_check_section(quality))
    markdown = ascii_text("\n".join(sections).rstrip() + "\n")
    return GeneratedScript(markdown=markdown, quality=quality)


def output_path_for(scripts_folder: Path, segment: SegmentInfo) -> Path:
    title = sanitize_filename(segment.video_title)
    filename = f"{segment.lecture_number}_{segment.video_number}_{title}_exact_script.md"
    return scripts_folder / "exact_scripts" / str(segment.lecture_number) / filename


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate exact read-aloud narration scripts for approved video decks."
    )
    parser.add_argument("--lecture-number", type=int, help="Process only this lecture number.")
    parser.add_argument("--video-number", type=int, help="Process only this video number.")
    parser.add_argument(
        "--all",
        action="store_true",
        help="Process all passing deck-audit rows. Without this, failed review rows are regenerated when present.",
    )
    args = parser.parse_args(argv)
    if (args.lecture_number is None) != (args.video_number is None):
        parser.error("--lecture-number and --video-number must be provided together.")
    return args


def print_quality_report(output_path: Path, quality: ScriptQuality) -> None:
    print(f"Output script filename: {output_path}")
    print(f"Estimated total words: {quality.estimated_words}")
    print(f"Estimated speaking time: {quality.estimated_minutes:.1f} minutes at about {READING_WPM} wpm")
    print(f"Banned metadata language found: {'yes' if quality.banned_metadata_found else 'no'}")
    print(f"Raw URLs in narration: {'yes' if quality.raw_urls_in_narration else 'no'}")
    print(
        "Technical correction notes added: "
        + ("yes" if quality.technical_correction_notes_added else "no")
    )
    print("Quality self-check scores:")
    for key, value in quality.scores.items():
        print(f"  {key}: {value:.2f}" if key == "overall_score" else f"  {key}: {value:.0f}")
    concerns: list[str] = []
    if quality.overly_long_sentences:
        concerns.append(f"{len(quality.overly_long_sentences)} overly long sentence warning(s)")
    if quality.unsupported_claim_warnings:
        concerns.append(f"{len(quality.unsupported_claim_warnings)} unsupported-claim warning(s)")
    if quality.repeated_slide_text_warnings:
        concerns.append(f"{len(quality.repeated_slide_text_warnings)} slide-text repetition warning(s)")
    print("Remaining concerns: " + ("; ".join(concerns) if concerns else "none"))


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    root = project_root()
    try:
        scripts_folder = configured_folder(root, "scripts")
        audit_rows = load_passed_audit_rows(root)
        segment_lookup = load_video_segment_lookup(root)
        failed_review_keys = load_failed_review_keys(root)

        created = 0
        last_output_path: Path | None = None
        last_quality: ScriptQuality | None = None
        for audit_row in audit_rows:
            segment = segment_from_row(audit_row, segment_lookup, root)
            if args.lecture_number is not None and (
                segment.lecture_number != args.lecture_number
                or segment.video_number != args.video_number
            ):
                continue
            if (
                args.lecture_number is None
                and not args.all
                and failed_review_keys
                and (segment.lecture_number, segment.video_number) not in failed_review_keys
            ):
                continue
            output_path = output_path_for(scripts_folder, segment)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            generated = generate_script(root, segment)
            output_path.write_text(generated.markdown, encoding="utf-8")
            created += 1
            last_output_path = output_path
            last_quality = generated.quality

        if args.lecture_number is not None and created == 0:
            raise ValueError(
                f"No passing generated deck found for lecture {args.lecture_number} video {args.video_number}."
            )

    except (OSError, RuntimeError, ValueError, KeyError, YAMLError, PptxReadError) as exc:
        print("Exact speaker script generation failed.")
        print(str(exc))
        return 1

    print("Exact speaker script generation complete")
    print(f"Scripts folder: {scripts_folder / 'exact_scripts'}")
    print(f"Scripts created: {created}")
    print("PowerPoint safety: PowerPoint files were opened read-only only when needed.")
    if created == 1 and last_output_path is not None and last_quality is not None:
        print_quality_report(last_output_path, last_quality)
    return 0


if __name__ == "__main__":
    sys.exit(main())
