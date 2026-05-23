"""Generate Markdown speaker scripts for audited segment decks.

The script reads passed rows from manifests/deck_audit/generated_deck_audit.csv
and creates one Markdown speaker script per deck in the configured 03_Scripts
course folder.

PowerPoint files are opened read-only only when needed for the manual Lecture 0
special case. No PowerPoint file is saved or modified.
"""

from __future__ import annotations

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


@dataclass(frozen=True)
class SegmentInfo:
    lecture_number: int
    lecture_title: str
    video_number: int
    video_title: str
    slide_start: int
    slide_end: int
    estimated_minutes: str
    learning_objectives: str
    recap: str
    quiz: str
    output_pptx: str
    generated_pptx_path: Path
    special_case: str = ""


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


def sanitize_filename(value: str) -> str:
    value = ascii_text(value)
    value = re.sub(r"[^A-Za-z0-9]+", "_", value).strip("_")
    value = re.sub(r"_+", "_", value)
    return value or "video"


def markdown_list(items: list[str]) -> str:
    return "\n".join(f"- {ascii_text(item)}" for item in items)


def first_nonempty_line(text: str) -> str:
    text = ascii_text(text)
    for line in text.splitlines():
        line = line.strip()
        if line:
            return line
    return ""


def text_lines(text: str, limit: int = 5) -> list[str]:
    lines: list[str] = []
    seen: set[str] = set()
    for raw_line in text.splitlines():
        line = clean_text(raw_line)
        if not line or line.isdigit() or line in seen:
            continue
        seen.add(line)
        lines.append(line)
        if len(lines) >= limit:
            break
    return lines


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
    title = row["video_title"].strip()
    if not title.lower().startswith("introduction"):
        video_title = title
    else:
        video_title = "Introduction"
    return SegmentInfo(
        lecture_number=0,
        lecture_title="Intro",
        video_number=1,
        video_title=video_title,
        slide_start=slide_start,
        slide_end=slide_end,
        estimated_minutes=row["estimated_total_minutes"],
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
        estimated_minutes=segment_row["estimated_minutes"],
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


def slide_role_for_lecture_zero(topic: str, slide_number: int) -> str:
    topic_lower = topic.lower()
    if slide_number == 1 or "instructor" in topic_lower or "course" in topic_lower:
        return "logistics"
    if "example" in topic_lower or "application" in topic_lower:
        return "example"
    if "relationship" in topic_lower or "definition" in topic_lower:
        return "definition"
    if "why" in topic_lower or "practical" in topic_lower or "enabled" in topic_lower:
        return "motivation"
    if "summary" in topic_lower:
        return "summary"
    if "architecture" in topic_lower or "network" in topic_lower:
        return "visualization"
    return "other"


def slow_down_note(role: str, difficulty: str, title: str) -> str:
    if difficulty == "high" or role in {"derivation", "algorithm"}:
        return f"Slow down when explaining {title}; students need the sequence, not just the final result."
    if role in {"visualization", "model architecture"}:
        return "Slow down to orient students to the diagram before interpreting it."
    if role in {"definition", "comparison"}:
        return "Slow down for the vocabulary and the distinction from nearby concepts."
    if role == "example":
        return "Pause briefly after the example so students can connect it to the general idea."
    return "Keep this slide concise; slow down only if the terminology is new."


def confusion_note(role: str, title: str, concept_group: str) -> str:
    if role == "derivation":
        return "Students may follow the algebra locally but miss why the derivation is being done."
    if role == "algorithm":
        return "Students may memorize the steps without understanding what each step computes."
    if role == "visualization":
        return "Students watching on a small screen may miss labels or axes in the figure."
    if role == "comparison":
        return "Students may remember both labels but not the practical tradeoff between them."
    if role == "definition":
        return f"Students may confuse {title} with a related idea from {concept_group}."
    if role == "example":
        return "Students may treat the example as a special case instead of evidence for the broader pattern."
    return "The main risk is moving too quickly before students know why this slide matters."


def role_script_sentence(role: str, title: str, core_concept: str) -> str:
    if role == "motivation":
        return f"Use this slide to answer why {core_concept} matters before adding details."
    if role == "definition":
        return f"Define {core_concept} carefully, then restate it in plain language."
    if role == "example":
        return f"Treat {title} as the concrete example that makes the previous idea less abstract."
    if role == "derivation":
        return f"Walk through the derivation for {core_concept} one step at a time."
    if role == "algorithm":
        return f"Frame {core_concept} as a procedure: what goes in, what is computed, and what comes out."
    if role == "visualization":
        return f"Describe the visual first, then explain what it reveals about {core_concept}."
    if role == "summary":
        return f"Use this slide to consolidate the main takeaway about {core_concept}."
    if role == "transition":
        return f"Use this as a bridge into {core_concept}; keep the explanation short."
    if role == "exercise":
        return f"Give students a moment to attempt the exercise before explaining {core_concept}."
    if role == "logistics":
        return "Keep this slide brief and focus on what students need to do next."
    return f"Explain the central idea: {core_concept}."


def source_slide_section(
    generated_slide_number: int,
    original_slide_number: int,
    row: dict[str, str],
    slide_text_row: dict[str, str] | None,
    next_title: str,
) -> str:
    title = ascii_text(row.get("slide_title", "").strip() or row.get("main_topic", "").strip())
    if not title and slide_text_row is not None:
        title = ascii_text(slide_text_row.get("possible_slide_title", "").strip())
    title = title or f"Slide {original_slide_number}"
    role = row.get("slide_role", "").strip() or "other"
    concept_group = ascii_text(row.get("concept_group", "").strip() or row.get("main_topic", "").strip())
    core_concept = ascii_text(row.get("core_concept", "").strip() or row.get("main_topic", "").strip() or title)
    difficulty = row.get("difficulty_level", "").strip().lower()
    minutes = row.get("estimated_teaching_minutes") or row.get("estimated_teaching_time_minutes") or ""
    notes = ascii_text(row.get("notes_for_future_script") or row.get("suggested_improvement_for_online_teaching") or "")
    visible_text = slide_text_row.get("visible_text", "") if slide_text_row is not None else ""
    visible_lines = text_lines(visible_text, limit=4)

    simple_slide = role in {"logistics", "transition", "summary"} or difficulty == "low"
    script_lines = [
        role_script_sentence(role, title, core_concept),
    ]
    if notes:
        script_lines.append(clean_text(notes))
    if not simple_slide and visible_lines:
        script_lines.append(
            "Point students to the on-screen wording, especially: "
            + "; ".join(visible_lines[:3])
            + "."
        )
    if row.get("prerequisite_concept"):
        script_lines.append(f"Connect this back to {ascii_text(row['prerequisite_concept'])}.")
    if row.get("next_concept"):
        script_lines.append(f"Set up the next idea: {ascii_text(row['next_concept'])}.")

    transition = (
        f"Now that {core_concept} is established, move to {next_title}."
        if next_title
        else "This closes the content sequence and sets up the recap."
    )
    key_points = [
        f"Core concept: {core_concept}",
        f"Role in video: {role}",
    ]
    if concept_group:
        key_points.append(f"Concept group: {concept_group}")
    if visible_lines:
        key_points.append("On-screen anchors: " + "; ".join(visible_lines[:3]))
    if minutes:
        key_points.append(f"Approximate time: {minutes} minutes")

    return "\n".join(
        [
            f"### Slide {generated_slide_number}: {title}",
            "",
            f"Original slide: {original_slide_number}",
            "",
            "**Speaker script**",
            "",
            " ".join(script_lines),
            "",
            "**Transition**",
            "",
            transition,
            "",
            "**Key explanation points**",
            "",
            markdown_list(key_points),
            "",
            "**Slow down here**",
            "",
            slow_down_note(role, difficulty, title),
            "",
            "**Possible student confusion**",
            "",
            confusion_note(role, title, concept_group or core_concept),
            "",
        ]
    )


def overview_section(segment: SegmentInfo) -> str:
    objectives = split_items(segment.learning_objectives, max_items=4)
    return "\n".join(
        [
            f"### Slide 1: {segment.video_title}",
            "",
            "**Speaker script**",
            "",
            (
                f"Welcome to this video on {segment.video_title}. "
                f"This is Lecture {segment.lecture_number}: {segment.lecture_title}, "
                f"covering slides {segment.slide_start}-{segment.slide_end}. "
                "Frame the goal before going into details: students should know what they are trying to learn and why it matters."
            ),
            "",
            "**Learning objectives to say out loud**",
            "",
            markdown_list(objectives),
            "",
            "**Transition**",
            "",
            "With the goals in place, move into the first content slide and start building the idea step by step.",
            "",
        ]
    )


def recap_section(slide_number: int, segment: SegmentInfo) -> str:
    recap_items = split_items(segment.recap, max_items=6)
    return "\n".join(
        [
            f"### Slide {slide_number}: Recap",
            "",
            "**Speaker script**",
            "",
            "Use this slide to compress the video into a few durable takeaways. Do not re-teach every detail; name the main ideas and connect them back to the learning objectives.",
            "",
            "**Key explanation points**",
            "",
            markdown_list(recap_items),
            "",
            "**Transition**",
            "",
            "After the recap, ask students to test the idea with a short checkpoint question.",
            "",
        ]
    )


def checkpoint_section(slide_number: int, segment: SegmentInfo) -> str:
    return "\n".join(
        [
            f"### Slide {slide_number}: Checkpoint",
            "",
            "**Speaker script**",
            "",
            "Pause here and give students a moment to answer before continuing. The goal is not a long calculation; it is a quick check that the main idea is usable.",
            "",
            "**Checkpoint question**",
            "",
            segment.quiz,
            "",
            "**Transition**",
            "",
            "Use the answer to close the loop and summarize what students should carry into the next video.",
            "",
        ]
    )


def closing_summary(segment: SegmentInfo) -> str:
    objectives = split_items(segment.learning_objectives, max_items=4)
    return "\n".join(
        [
            "## Closing Summary",
            "",
            (
                f"Close by restating that this video was about {segment.video_title}. "
                "Name the two or three most important ideas one more time, then connect them to the next topic in the course."
            ),
            "",
            "**Final takeaways**",
            "",
            markdown_list(objectives[:3]),
            "",
            "**Checkpoint question**",
            "",
            segment.quiz,
            "",
        ]
    )


def script_header(segment: SegmentInfo) -> str:
    objectives = split_items(segment.learning_objectives, max_items=4)
    return "\n".join(
        [
            f"# {segment.video_title}",
            "",
            f"Lecture: {segment.lecture_number} - {segment.lecture_title}",
            f"Video: {segment.video_number}",
            f"Original slide range: {segment.slide_start}-{segment.slide_end}",
            f"Estimated recording time: {segment.estimated_minutes} minutes",
            "",
            "## Learning Objectives",
            "",
            markdown_list(objectives),
            "",
            "## Recording Notes",
            "",
            "- Use a conversational online teaching pace.",
            "- Do not read every bullet on the slide.",
            "- Slow down for definitions, algorithms, derivations, architecture diagrams, and comparisons.",
            "- Keep simple transition or logistics slides brief.",
            "",
            "## Slide-by-Slide Script",
            "",
        ]
    )


def next_slide_title(
    slide_number: int,
    slide_end: int,
    teaching_map: dict[int, dict[str, str]],
    inventory: dict[int, dict[str, str]],
) -> str:
    next_number = slide_number + 1
    if next_number > slide_end:
        return ""
    row = teaching_map.get(next_number) or inventory.get(next_number) or {}
    return (
        row.get("slide_title", "")
        or row.get("main_topic", "")
        or f"slide {next_number}"
    )


def generate_segment_script(root: Path, segment: SegmentInfo) -> str:
    sections = [script_header(segment)]

    if segment.special_case == "lecture_0_manual":
        inventory = lecture_zero_inventory(root)
        deck_texts = deck_slide_texts(segment.generated_pptx_path)
        for slide_number in range(segment.slide_start, segment.slide_end + 1):
            inventory_row = inventory.get(slide_number, {})
            deck_text = deck_texts.get(slide_number, "")
            title = first_nonempty_line(deck_text) or inventory_row.get("main_topic", "")
            row = {
                "main_topic": inventory_row.get("main_topic", title),
                "difficulty_level": inventory_row.get("difficulty_level", "low"),
                "estimated_teaching_time_minutes": inventory_row.get(
                    "estimated_teaching_time_minutes",
                    "",
                ),
                "suggested_improvement_for_online_teaching": inventory_row.get(
                    "suggested_improvement_for_online_teaching",
                    "",
                ),
                "slide_role": slide_role_for_lecture_zero(
                    inventory_row.get("main_topic", title),
                    slide_number,
                ),
                "concept_group": "Lecture 0 course introduction",
            }
            slide_text_row = {"visible_text": deck_text, "possible_slide_title": title}
            sections.append(
                source_slide_section(
                    slide_number,
                    slide_number,
                    row,
                    slide_text_row,
                    next_slide_title(slide_number, segment.slide_end, {}, inventory),
                )
            )
        sections.append(closing_summary(segment))
        return ascii_text("\n".join(sections).rstrip() + "\n")

    teaching_map = load_teaching_map(root, segment.lecture_number)
    slide_text = load_slide_text(root, segment.lecture_number)
    sections.append(overview_section(segment))

    generated_slide_number = 2
    for original_slide_number in range(segment.slide_start, segment.slide_end + 1):
        row = teaching_map.get(original_slide_number, {})
        slide_text_row = slide_text.get(original_slide_number)
        sections.append(
            source_slide_section(
                generated_slide_number,
                original_slide_number,
                row,
                slide_text_row,
                next_slide_title(original_slide_number, segment.slide_end, teaching_map, {}),
            )
        )
        generated_slide_number += 1

    sections.append(recap_section(generated_slide_number, segment))
    sections.append(checkpoint_section(generated_slide_number + 1, segment))
    sections.append(closing_summary(segment))
    return ascii_text("\n".join(sections).rstrip() + "\n")


def output_path_for(scripts_folder: Path, segment: SegmentInfo) -> Path:
    title = sanitize_filename(segment.video_title)
    filename = f"{segment.lecture_number}_{segment.video_number}_{title}_script.md"
    return scripts_folder / str(segment.lecture_number) / filename


def main() -> int:
    root = project_root()
    try:
        scripts_folder = configured_folder(root, "scripts")
        audit_rows = load_passed_audit_rows(root)
        segment_lookup = load_video_segment_lookup(root)

        created = 0
        for audit_row in audit_rows:
            segment = segment_from_row(audit_row, segment_lookup, root)
            output_path = output_path_for(scripts_folder, segment)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            output_path.write_text(generate_segment_script(root, segment), encoding="utf-8")
            created += 1

    except (OSError, RuntimeError, ValueError, KeyError, YAMLError, PptxReadError) as exc:
        print("Speaker script generation failed.")
        print(str(exc))
        return 1

    print("Speaker script generation complete")
    print(f"Scripts folder: {scripts_folder}")
    print(f"Scripts created: {created}")
    print("PowerPoint safety: PowerPoint files were opened read-only only when needed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
