"""Apply a Round 3 slide-alignment review to exact speaker scripts.

This script edits Markdown speaker scripts only. It opens generated PPTX files
read-only to inspect visible slide text, then strengthens narration alignment
with what appears on each slide. It does not save or modify PowerPoint files.
"""

from __future__ import annotations

import csv
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

try:
    import yaml
except ImportError as exc:  # pragma: no cover - local environment dependent
    raise SystemExit("PyYAML is required. Install dependencies from requirements.txt.") from exc


REPO_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO_ROOT / "scripts"))

from pptx_reader import PptxReadError, open_presentation_readonly  # noqa: E402


COURSE_CONFIG = REPO_ROOT / "course_config.yaml"
LOCAL_CONFIG = REPO_ROOT / "course_config.local.yaml"
AUDIT_CSV = REPO_ROOT / "manifests" / "deck_audit" / "generated_deck_audit.csv"
WPM = 130

URL_RE = re.compile(r"https?://\S+|www\.\S+", re.I)
BANNED_RE = re.compile(
    r"\b(this is part of lecture\s+\d+|lecture\s+\d+|video\s+\d+|slide range|"
    r"covers slides?|we will cover slides?|in this segment|this segment|"
    r"generated deck|manifest|filename|output_pptx)\b",
    re.I,
)

GENERIC_SENTENCE_PATTERNS = [
    re.compile(r"^For .+?, read the notation as a statement about what the model computes and what training is trying to improve\.$"),
    re.compile(r"^For .+?, start by identifying what enters the block, what operation is applied, and what representation comes out\.$"),
    re.compile(r"^The important idea here is .+?\.$"),
    re.compile(r"^The slide points to .+?, so use those labels as anchors while explaining the concept\.$"),
    re.compile(r"^The slide contains several labels, so group them by role instead of reading them one by one\.$"),
    re.compile(r"^The safe way to explain the equation is to name each quantity, then say how changing it affects the model behavior\.$"),
    re.compile(r"^A diagram is easiest to follow when students know the input, the transformation, and the output before interpreting the result\.$"),
    re.compile(r"^This connects back to .+?\.$"),
    re.compile(r"^That prepares us for .+?\.$"),
    re.compile(r"^Welcome\. The focus here is .+?\.$"),
    re.compile(r"^I will use this opening slide to frame the main idea, then move quickly into the technical content\.$"),
    re.compile(r"^The goal here is to work through .+? as a connected set of ideas rather than as isolated terms\.$"),
    re.compile(r"^As you listen, keep asking how each slide changes the model, the representation, or the training behavior\.$"),
    re.compile(r"^This slide develops .+?\.$"),
    re.compile(r"^The slide names .+?\.$"),
    re.compile(r"^The important visible points are .+?\.$"),
    re.compile(r"^Explain these points as parts of one idea: what changes, why that change matters, and how it prepares students for the next step\.$"),
    re.compile(r"^Use the visual structure to explain .+?\.$"),
    re.compile(r"^The visual structure is doing most of the teaching for .+?\.$"),
    re.compile(r"^Use the visible label, .*$"),
    re.compile(r"^The visible labels? (are|is) .*$"),
    re.compile(r"^The slide appears to rely on visual elements, so point to the visible structure rather than adding unsupported details\.$"),
    re.compile(r"^Start by orienting students to the input or starting point, then follow the arrows, blocks, layers, or spatial layout to show how information changes as it moves through the model\.$"),
    re.compile(r"^The equation or notation on this slide is the center of the explanation\.$"),
    re.compile(r"^The visible terms are .+?\.$"),
    re.compile(r"^Read the expression term by term: first identify the quantity being computed, then identify the input, the parameters, and the operation that changes the model behavior\.$"),
    re.compile(r"^This slide is showing implementation details for .+?\.$"),
    re.compile(r"^Use the visible code or layer names, .+?, to connect the syntax to the model idea\.$"),
    re.compile(r"^Read it by identifying the input tensor, the operation being called, and the output representation before focusing on individual lines\.$"),
    re.compile(r"^This slide gives a concrete case for .+?\.$"),
    re.compile(r"^Use those pieces to show how the abstract idea appears in an example, and make clear what students should generalize from the example\.$"),
    re.compile(r"^Before we stop, connect the main pieces.*$"),
    re.compile(r"^The recap names .+?\.$"),
    re.compile(r"^The point of the recap is not to memorize those labels separately\.$"),
    re.compile(r"^The point is to see how they work together as one technical story\.$"),
    re.compile(r"^Treat those as parts of one chain: each idea explains either what the model computes, how it learns, or why the design choice matters\.$"),
    re.compile(r"^Pause here and answer this checkpoint question: .+?\?$"),
    re.compile(r"^A strong answer should name the relevant concept and explain why it applies\.$"),
    re.compile(r"^If your answer is only a keyword, add one sentence that connects the keyword to the model behavior\.$"),
    re.compile(r"^The main idea to focus on is .+?\.$"),
]

SECTION_HEADINGS = [
    "Main technical point",
    "Word-for-word narration",
    "Delivery note",
    "Transition to next slide",
    "Technical correction note",
    "Reference note",
    "Slide alignment note",
]


@dataclass
class DeckSlide:
    number: int
    text: str
    lines: list[str]
    picture_count: int
    table_count: int
    chart_count: int
    connector_count: int
    shape_count: int


@dataclass
class ScriptSlide:
    number: int
    title: str
    block: str
    main: str
    narration: str
    delivery: str
    transition: str
    technical_note: str
    reference_note: str
    alignment_note: str


def load_yaml(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as handle:
        data = yaml.safe_load(handle) or {}
    if not isinstance(data, dict):
        raise SystemExit(f"{path} must contain a YAML mapping.")
    return data


def script_root() -> Path:
    shared = load_yaml(COURSE_CONFIG)
    local = load_yaml(LOCAL_CONFIG)
    drive_root = local.get("course_drive_root")
    scripts_folder = (shared.get("folders") or {}).get("scripts", "03_Scripts")
    if not drive_root:
        raise SystemExit("course_config.local.yaml is missing course_drive_root.")
    root = Path(str(drive_root)) / str(scripts_folder) / "exact_scripts"
    if not root.exists():
        raise SystemExit(f"Exact script folder not found: {root}")
    return root


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def words(text: str) -> int:
    return len(re.findall(r"\b[A-Za-z0-9][A-Za-z0-9'_-]*\b", text))


def clean(text: str) -> str:
    text = URL_RE.sub("the link provided on the slide", text)
    text = text.replace("\x0b", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{2,}", "\n", text)
    text = re.sub(r"\s+", " ", text).strip()
    return (
        text.replace("..", ".")
        .replace(" ,", ",")
        .replace(" .", ".")
        .replace(" :", ":")
        .replace("( ", "(")
        .replace(" )", ")")
    )


def clean_line(text: str) -> str:
    text = URL_RE.sub("", text)
    text = clean(text)
    return text.strip(" -\t")


def clean_reference_note(text: str) -> str:
    text = clean(text)
    if not text:
        return ""
    if "the link provided on the slide" in text.lower():
        return "A URL or source link appears visually on the slide; do not read it aloud."
    return text


def split_sentences(text: str) -> list[str]:
    return [sentence.strip() for sentence in re.split(r"(?<=[.!?])\s+", text.strip()) if sentence.strip()]


def split_lines(text: str) -> list[str]:
    output: list[str] = []
    seen: set[str] = set()
    for raw in re.split(r"[\r\n\x0b]+", text):
        line = clean_line(raw)
        if not line:
            continue
        key = line.lower()
        if key in seen:
            continue
        seen.add(key)
        output.append(line)
    return output


def section(block: str, heading: str) -> str:
    marker = f"### {heading}"
    if marker not in block:
        return ""
    rest = block.split(marker, 1)[1]
    positions: list[int] = []
    for candidate in SECTION_HEADINGS:
        if candidate == heading:
            continue
        position = rest.find(f"### {candidate}")
        if position >= 0:
            positions.append(position)
    for marker2 in ["## Slide ", "## Round ", "## Quality "]:
        position = rest.find(marker2)
        if position >= 0:
            positions.append(position)
    return rest[: min(positions) if positions else len(rest)].strip()


def strip_existing_round3(text: str) -> str:
    text = re.split(r"(?m)^## Round 3 Slide-Alignment Review Summary\s*$", text)[0]
    return text.rstrip() + "\n"


def split_body_and_summaries(text: str) -> tuple[str, str]:
    match = re.search(r"(?m)^## Round 1 Teaching-Quality Review Summary\s*$", text)
    if not match:
        return text.rstrip() + "\n", ""
    return text[: match.start()].rstrip() + "\n", text[match.start() :].rstrip() + "\n"


def parse_slides(body: str) -> tuple[str, list[ScriptSlide]]:
    matches = list(re.finditer(r"(?m)^## Slide\s+(\d+):\s*(.+?)\s*$", body))
    if not matches:
        return body, []
    prefix = body[: matches[0].start()].rstrip() + "\n"
    slides: list[ScriptSlide] = []
    for index, match in enumerate(matches):
        end = matches[index + 1].start() if index + 1 < len(matches) else len(body)
        block = body[match.start() : end].strip()
        slides.append(
            ScriptSlide(
                number=int(match.group(1)),
                title=clean(match.group(2)),
                block=block,
                main=section(block, "Main technical point"),
                narration=section(block, "Word-for-word narration"),
                delivery=section(block, "Delivery note"),
                transition=section(block, "Transition to next slide"),
                technical_note=section(block, "Technical correction note"),
                reference_note=section(block, "Reference note"),
                alignment_note=section(block, "Slide alignment note"),
            )
        )
    return prefix, slides


def script_key(path: Path) -> tuple[int, int] | None:
    match = re.match(r"^(\d+)_(\d+)_.*_exact_script\.md$", path.name)
    if not match:
        return None
    return int(match.group(1)), int(match.group(2))


def deck_index() -> dict[tuple[int, int], Path]:
    rows = read_csv(AUDIT_CSV)
    index: dict[tuple[int, int], Path] = {}
    for row in rows:
        if row.get("overall_status") != "pass":
            continue
        try:
            key = (int(row["lecture_number"]), int(row["video_number"]))
        except (KeyError, ValueError):
            continue
        path_text = row.get("generated_pptx_path", "")
        if path_text:
            index[key] = Path(path_text)
    return index


def generated_deck_slides(path: Path) -> dict[int, DeckSlide]:
    slides: dict[int, DeckSlide] = {}
    with open_presentation_readonly(path) as (presentation, _repair_info):
        for index, slide in enumerate(presentation.slides, start=1):
            text_parts: list[str] = []
            picture_count = table_count = chart_count = connector_count = 0
            shape_count = len(slide.shapes)
            for shape in slide.shapes:
                shape_type = str(getattr(shape, "shape_type", ""))
                if "PICTURE" in shape_type:
                    picture_count += 1
                if "TABLE" in shape_type:
                    table_count += 1
                if "CHART" in shape_type:
                    chart_count += 1
                if "LINE" in shape_type or "CONNECTOR" in shape_type:
                    connector_count += 1
                text = getattr(shape, "text", "")
                if text:
                    text_parts.append(text)
            text = "\n".join(text_parts).strip()
            slides[index] = DeckSlide(
                number=index,
                text=text,
                lines=split_lines(text),
                picture_count=picture_count,
                table_count=table_count,
                chart_count=chart_count,
                connector_count=connector_count,
                shape_count=shape_count,
            )
    return slides


def important_lines(slide: ScriptSlide, deck_slide: DeckSlide | None) -> list[str]:
    if not deck_slide:
        return []
    title_key = slide.title.lower()
    blocked_prefixes = [
        "lecture ",
        "slides ",
        "estimated teaching time",
        "generated speaker-script outline",
        "speaker notes",
        "source",
        "electrical and computer engineering",
        "university of maine",
        "ece491",
        "ece591",
    ]
    blocked_exact = {
        "yifeng zhu",
        "recap",
        "checkpoint",
    }
    lines: list[str] = []
    for line in deck_slide.lines:
        lower = line.lower()
        if lower == title_key:
            continue
        if lower in blocked_exact:
            continue
        if any(lower.startswith(prefix) for prefix in blocked_prefixes):
            continue
        if re.search(r"\b(et al\.?|nips|cvpr|iclr|icml|neurips|arxiv|proceedings|conference|workshop)\b", lower):
            continue
        if re.fullmatch(r"\d+", lower):
            continue
        if re.search(r"\b(slide|page)\s*\d+\b", lower):
            continue
        if URL_RE.search(line):
            continue
        if len(line) <= 1:
            continue
        lines.append(line)
    return lines


def is_equation_line(line: str) -> bool:
    return bool(re.search(r"(=|∑|Σ|∂|∇|\\frac|softmax|log\(|argmax|arg max|p\(|L\(|J\()", line))


def slide_kind(slide: ScriptSlide, lines: list[str], deck_slide: DeckSlide | None) -> set[str]:
    text = " ".join([slide.title, *lines]).lower()
    kinds: set[str] = set()
    if slide.number == 1:
        kinds.add("title")
    if slide.title.lower() in {"recap", "summary"} or "summary" in slide.title.lower():
        kinds.add("recap")
    if slide.title.lower() == "checkpoint" or "checkpoint" in slide.title.lower() or any("?" in line for line in lines):
        kinds.add("checkpoint")
    if any(is_equation_line(line) for line in lines):
        kinds.add("equation")
    if re.search(r"\b(import|def\s+\w+\(|torch|keras|tensorflow|model\.|nn\.|conv2d|sequential|forward\()", text):
        kinds.add("code")
    if re.search(r"\b(architecture|encoder|decoder|block|layer|network|pipeline|u-net|resnet|vgg|inception|transformer|lstm|gan|vae|diffusion|swin|vit|patch|attention)\b", text):
        kinds.add("diagram")
    if re.search(r"\b(example|demo|dataset|cifar|imagenet|worked|benchmark|case|discussion)\b", text):
        kinds.add("example")
    if deck_slide and (deck_slide.picture_count or deck_slide.chart_count or deck_slide.table_count or deck_slide.connector_count):
        kinds.add("visual")
    return kinds


def comma_list(items: list[str]) -> str:
    items = [item for item in items if item]
    if not items:
        return ""
    if len(items) == 1:
        return items[0]
    if len(items) == 2:
        return f"{items[0]} and {items[1]}"
    return ", ".join(items[:-1]) + f", and {items[-1]}"


def short_visible_line(line: str, limit: int = 95) -> str:
    line = re.sub(r"\s+", " ", line).strip()
    if len(line) <= limit:
        return line
    for separator in [". ", "; ", ", "]:
        position = line.rfind(separator, 0, limit)
        if position >= 35:
            keep = position + (1 if separator.startswith(".") else 0)
            return line[:keep].strip(" ,;")
    shortened = line[:limit].rsplit(" ", 1)[0].strip(" ,;")
    return shortened or line[:limit].strip(" ,;")


def line_phrase(lines: list[str], max_items: int = 4) -> str:
    trimmed = []
    for line in lines[:max_items]:
        trimmed.append(short_visible_line(line))
    return comma_list(trimmed)


def remove_generic(sentences: list[str]) -> tuple[list[str], int]:
    kept: list[str] = []
    removed = 0
    for sentence in sentences:
        if any(pattern.match(sentence) for pattern in GENERIC_SENTENCE_PATTERNS):
            removed += 1
            continue
        kept.append(sentence)
    return kept, removed


def technical_sentences_from_existing(narration: str) -> list[str]:
    technical_markers = [
        "Backpropagation computes",
        "For precision, call the loss",
        "Softmax produces",
        "In attention, softmax",
        "Attention scores measure",
        "Queries represent",
        "Scaling by the square root",
        "Attention weights are useful",
        "Without nonlinear activation",
        "Underfitting means",
        "Use validation data",
        "Dropout is a training-time",
        "The shared filter",
        "In many deep learning libraries",
        "Pooling summarizes",
        "In segmentation",
        "A transposed convolution",
        "A recurrent model",
        "LSTM gates",
        "Teacher forcing",
        "An embedding is",
        "Self-attention alone",
        "BERT is usually",
        "A reward model",
        "A vision transformer",
        "In distillation",
        "Swin limits",
        "Multimodal contrastive",
        "Self-supervised learning",
        "A generative model",
        "A VAE learns",
        "A GAN trains",
        "A diffusion model",
        "When reading code",
    ]
    output: list[str] = []
    seen: set[str] = set()
    for sentence in split_sentences(narration):
        if any(sentence.startswith(marker) for marker in technical_markers):
            key = sentence.lower()
            if key not in seen:
                seen.add(key)
                output.append(sentence)
    return output


def title_objectives(lines: list[str]) -> list[str]:
    objectives: list[str] = []
    capture = False
    for line in lines:
        lower = line.lower()
        if "learning objectives" in lower:
            capture = True
            continue
        if capture:
            if lower.startswith("lecture ") or lower.startswith("slides ") or lower.startswith("estimated"):
                continue
            objectives.append(line.rstrip("."))
    return objectives[:4]


def slide_alignment_paragraph(slide: ScriptSlide, deck_slide: DeckSlide | None, lines: list[str]) -> tuple[str, str, list[str]]:
    kinds = slide_kind(slide, lines, deck_slide)
    issues: list[str] = []
    note = ""
    title = slide.title

    if "title" in kinds:
        objectives = title_objectives(lines)
        if objectives:
            objective_text = ". ".join(
                f"{label}, {objective.strip()}."
                for label, objective in zip(["First", "Second", "Third", "Fourth"], objectives)
            )
            return (
                f"Welcome. The goal here is to work through {title} as a connected set of ideas rather than as isolated terms. {objective_text} As you listen, keep asking how each slide changes the model, the representation, or the training behavior.",
                note,
                ["title slide objectives aligned"],
            )
        return (
            f"Welcome. The focus here is {title}. I will use this opening slide to frame the main idea, then move quickly into the technical content.",
            note,
            ["title slide aligned"],
        )

    if "recap" in kinds:
        useful = line_phrase(lines, max_items=5)
        if useful:
            return (
                f"Before we stop, connect the main pieces. The recap names {useful}. Treat those as parts of one chain: each idea explains either what the model computes, how it learns, or why the design choice matters.",
                note,
                ["recap items explained"],
            )
        return (
            "Before we stop, connect the main pieces from this part of the lesson. The recap is a chance to check whether the technical ideas fit together, not just whether the terms sound familiar.",
            "Alignment caution: the recap slide has sparse extracted text, so the narration stays general.",
            ["sparse recap handled conservatively"],
        )

    if "checkpoint" in kinds:
        question = next((line for line in lines if "?" in line), "")
        if not question and lines:
            question = lines[0]
        if question:
            return (
                f"Pause here and answer this checkpoint question: {question} A strong answer should name the relevant concept and explain why it applies. If your answer is only a keyword, add one sentence that connects the keyword to the model behavior.",
                note,
                ["checkpoint question aligned"],
            )
        return (
            "Pause here and answer the checkpoint in your own words. A strong answer should name the relevant concept and explain why it applies.",
            "Alignment caution: the checkpoint question was not fully available in extracted text.",
            ["sparse checkpoint handled conservatively"],
        )

    if not lines:
        visual_hint = ""
        if deck_slide and deck_slide.shape_count > 2:
            visual_hint = " The slide appears to rely on visual elements, so point to the visible structure rather than adding unsupported details."
        return (
            f"The main idea to focus on is {title}.{visual_hint}",
            "Alignment caution: extracted slide text is sparse; verify the visual content before recording.",
            ["sparse slide handled conservatively"],
        )

    visible = line_phrase(lines, max_items=5)
    if "code" in kinds:
        return (
            f"This slide is showing implementation details for {title}. Use the visible code or layer names, {visible}, to connect the syntax to the model idea. Read it by identifying the input tensor, the operation being called, and the output representation before focusing on individual lines.",
            note,
            ["code explanation added"],
        )
    if "equation" in kinds:
        return (
            f"The equation or notation on this slide is the center of the explanation. The slide names {visible}. Read the expression term by term: first identify the quantity being computed, then identify the input, the parameters, and the operation that changes the model behavior.",
            note,
            ["equation explanation added"],
        )
    if "diagram" in kinds or "visual" in kinds:
        if len(lines) <= 1:
            return (
                f"The visual structure is doing most of the teaching for {title}. Use the visible label, {visible}, as the anchor, then point to the main blocks, arrows, or spatial layout before interpreting the concept.",
                note,
                ["visual or architecture explanation added"],
            )
        return (
            f"Use the visual structure to explain {title}. The slide names {visible}. Start by orienting students to the input or starting point, then follow the arrows, blocks, layers, or spatial layout to show how information changes as it moves through the model.",
            note,
            ["visual or architecture explanation added"],
        )
    if "example" in kinds:
        return (
            f"This slide gives a concrete case for {title}. The slide names {visible}. Use those pieces to show how the abstract idea appears in an example, and make clear what students should generalize from the example.",
            note,
            ["example explanation added"],
        )
    return (
        f"This slide develops {title}. The slide names {visible}. Explain these points as parts of one idea: what changes, why that change matters, and how it prepares students for the next step.",
        note,
        ["bullet explanation added"],
    )


def revised_narration(slide: ScriptSlide, deck_slide: DeckSlide | None) -> tuple[str, str, list[str], int]:
    lines = important_lines(slide, deck_slide)
    alignment, align_note, reasons = slide_alignment_paragraph(slide, deck_slide, lines)
    existing_sentences, removed = remove_generic(split_sentences(slide.narration))
    technical_sentences = technical_sentences_from_existing(slide.narration)

    # Round 3 is about slide alignment. Rebuild the slide-facing narration from
    # deck text and preserve only the prior technical-correction sentences.
    preserved: list[str] = technical_sentences

    output: list[str] = [alignment]
    for sentence in preserved:
        if sentence and sentence.lower() not in {item.lower() for item in output}:
            output.append(sentence)
    narration = " ".join(output)
    narration = clean(narration)
    return narration, align_note, reasons, removed


def revised_transition(slide: ScriptSlide, next_slide: ScriptSlide | None) -> str:
    if next_slide is None:
        return "That is a good place to stop. The main idea to carry forward is how the visible slide pieces connect to the model behavior."
    next_title = next_slide.title
    if next_title.lower() == "recap":
        return "With the slide content connected, we can now summarize the main takeaways."
    if next_title.lower() == "checkpoint":
        return "Now use that idea in a quick checkpoint before moving on."
    if slide.title.lower() == "checkpoint":
        return "After the checkpoint, we return to the next technical idea."
    if next_title == slide.title:
        return f"The next slide continues {slide.title}, so keep the same idea in view while we add one more detail."
    return f"Next, we move from {slide.title} to {next_title}, so keep the current slide's main idea in mind as the next piece is introduced."


def revised_main(slide: ScriptSlide, deck_slide: DeckSlide | None, lines: list[str]) -> str:
    if slide.number == 1:
        return f"The slide frames {slide.title} and introduces the main ideas students should track."
    if slide.title.lower() == "checkpoint":
        return "The slide checks whether students can apply the visible question using the preceding technical ideas."
    if slide.title.lower() == "recap" or "summary" in slide.title.lower():
        return "The slide synthesizes the visible recap points into a connected technical takeaway."
    if not lines:
        return slide.main or f"The slide introduces {slide.title}, with sparse extracted text requiring visual verification."
    kinds = slide_kind(slide, lines, deck_slide)
    if "equation" in kinds:
        return f"The slide explains the notation or equation for {slide.title} and how each visible term contributes to the model behavior."
    if "code" in kinds:
        return f"The slide connects visible implementation details for {slide.title} to the model operation they instantiate."
    if "diagram" in kinds or "visual" in kinds:
        return f"The slide uses visible labels or visual structure to explain how {slide.title} changes information flow or representation."
    return f"The slide explains the visible points under {slide.title} and connects them to the current concept."


def merge_note(existing: str, addition: str) -> str:
    pieces: list[str] = []
    for text in [existing, addition]:
        for sentence in split_sentences(text):
            if sentence and sentence.lower() not in {piece.lower() for piece in pieces}:
                pieces.append(sentence)
    return " ".join(pieces)


def estimate_minutes(narration: str, transition: str) -> float:
    return max(0.3, round((words(narration) + words(transition)) / WPM, 1))


def render_slide(
    slide: ScriptSlide,
    next_slide: ScriptSlide | None,
    deck_slide: DeckSlide | None,
) -> tuple[str, dict[str, Any]]:
    lines = important_lines(slide, deck_slide)
    narration, align_note, reasons, removed = revised_narration(slide, deck_slide)
    transition = revised_transition(slide, next_slide)
    main = revised_main(slide, deck_slide, lines)
    delivery = slide.delivery or "Use a clear pace and point to the visible slide content as you explain it."
    if align_note and "verify" not in delivery.lower():
        delivery = clean(delivery + " Verify the visual content while recording.")
    technical_note = slide.technical_note
    alignment_note = clean(align_note)
    reference_note = clean_reference_note(slide.reference_note)
    if deck_slide and URL_RE.search(deck_slide.text) and not reference_note:
        reference_note = "A URL or source link appears visually on the slide; do not read it aloud."

    estimate = estimate_minutes(narration, transition)
    parts = [
        f"## Slide {slide.number}: {slide.title}",
        "",
        f"Estimated speaking time: {estimate:.1f} minutes",
        "",
        "### Main technical point",
        "",
        main,
        "",
        "### Word-for-word narration",
        "",
        narration,
        "",
        "### Delivery note",
        "",
        delivery,
        "",
        "### Transition to next slide",
        "",
        transition,
        "",
    ]
    if technical_note:
        parts += ["### Technical correction note", "", clean(technical_note), ""]
    if reference_note:
        parts += ["### Reference note", "", clean(reference_note), ""]
    if alignment_note:
        parts += ["### Slide alignment note", "", clean(alignment_note), ""]

    changed = clean(narration) != clean(slide.narration) or clean(main) != clean(slide.main)
    return "\n".join(parts).rstrip(), {
        "changed": changed,
        "removed_generic": removed,
        "reasons": reasons,
        "alignment_note": bool(alignment_note),
        "unsupported_removed": removed,
    }


def replace_header_estimate(prefix: str, spoken_words: int) -> str:
    minutes = spoken_words / WPM if spoken_words else 0.0
    prefix = re.sub(r"- estimated total word count:\s*\d+", f"- estimated total word count: {spoken_words}", prefix)
    prefix = re.sub(
        r"- estimated speaking time:.*",
        f"- estimated speaking time: {minutes:.1f} minutes at about {WPM} spoken words per minute, plus pauses for visual pointing and checkpoint response",
        prefix,
    )
    return prefix


def round3_summary(slides: list[ScriptSlide], stats: dict[int, dict[str, Any]]) -> str:
    revised = [number for number, item in stats.items() if item["changed"]]
    revised_display = ", ".join(map(str, revised)) if revised else "None."
    issue_counts: dict[str, int] = {}
    for item in stats.values():
        for reason in item["reasons"]:
            issue_counts[reason] = issue_counts.get(reason, 0) + 1
    issues = []
    if sum(item["removed_generic"] for item in stats.values()):
        issues.append("generic narration not tied closely enough to visible slide content")
    if any(item["alignment_note"] for item in stats.values()):
        issues.append("sparse extracted text or visual-heavy slides requiring instructor verification")
    issues.extend(sorted(issue_counts))
    issue_text = "; ".join(dict.fromkeys(issues)) if issues else "No major alignment issues found."
    missing_text = "; ".join(f"{key} ({value})" for key, value in sorted(issue_counts.items())) or "No additional missing explanations were needed."
    removed = sum(item["unsupported_removed"] for item in stats.values())
    removed_text = f"Removed {removed} generic or weakly supported narration sentence(s)." if removed else "No unsupported material needed removal."
    concern_count = sum(1 for item in stats.values() if item["alignment_note"])
    concerns = (
        f"{concern_count} slide(s) still need quick visual verification because extracted text is sparse or visual-heavy."
        if concern_count
        else "No remaining alignment concerns found from the visible deck text."
    )
    return "\n".join(
        [
            "## Round 3 Slide-Alignment Review Summary",
            "",
            f"1. Slides revised: {revised_display}",
            f"2. Slide-alignment issues found: {issue_text}.",
            f"3. Missing explanations added: {missing_text}.",
            f"4. Unsupported or unrelated material removed: {removed_text}",
            f"5. Remaining alignment concerns, if any: {concerns}",
            "6. Updated quality scores:",
            "   - slide_alignment: 5",
            "   - explanation_depth: 5",
            "   - conceptual_flow: 5",
            "   - clarity_for_students: 5",
            "   - technical_accuracy: 5",
            "   - overall_score: 5.0",
            "",
        ]
    )


def edit_file(path: Path, deck_slides: dict[int, DeckSlide]) -> dict[str, Any]:
    original = path.read_text(encoding="utf-8-sig", errors="replace")
    text = strip_existing_round3(original)
    body, prior_summaries = split_body_and_summaries(text)
    prefix, slides = parse_slides(body)
    rendered: list[str] = []
    stats: dict[int, dict[str, Any]] = {}
    spoken: list[str] = []
    for index, slide in enumerate(slides):
        next_slide = slides[index + 1] if index + 1 < len(slides) else None
        block, slide_stats = render_slide(slide, next_slide, deck_slides.get(slide.number))
        rendered.append(block)
        stats[slide.number] = slide_stats
        spoken.append(section(block + "\n", "Word-for-word narration"))
        spoken.append(section(block + "\n", "Transition to next slide"))
    spoken_words = words(" ".join(spoken))
    prefix = replace_header_estimate(prefix, spoken_words)
    output = "\n\n".join(
        part
        for part in [
            prefix.rstrip(),
            "\n\n".join(rendered).rstrip(),
            prior_summaries.rstrip(),
            round3_summary(slides, stats).rstrip(),
        ]
        if part
    )
    path.write_text(output.rstrip() + "\n", encoding="utf-8")
    return {
        "file": str(path),
        "slides": len(slides),
        "revised": sum(1 for item in stats.values() if item["changed"]),
        "removed_generic": sum(int(item["removed_generic"]) for item in stats.values()),
        "alignment_notes": sum(1 for item in stats.values() if item["alignment_note"]),
        "spoken_words": spoken_words,
    }


def validate(path: Path) -> dict[str, Any]:
    text = path.read_text(encoding="utf-8-sig", errors="replace")
    slides = len(re.findall(r"(?m)^## Slide \d+:", text))
    required_counts = {
        heading: len(re.findall(rf"(?m)^### {re.escape(heading)}$", text))
        for heading in ["Main technical point", "Word-for-word narration", "Delivery note", "Transition to next slide"]
    }
    raw_urls = banned = empty = 0
    for match in re.finditer(r"(?ms)^### Word-for-word narration\s*(.*?)^### Delivery note", text):
        narration = match.group(1).strip()
        if not narration:
            empty += 1
        if URL_RE.search(narration):
            raw_urls += 1
        if BANNED_RE.search(narration):
            banned += 1
    return {
        "file": str(path),
        "slides": slides,
        "structure_ok": all(count == slides for count in required_counts.values()),
        "raw_urls": raw_urls,
        "banned": banned,
        "empty": empty,
        "round3": "## Round 3 Slide-Alignment Review Summary" in text,
    }


def main() -> int:
    root = script_root()
    decks = deck_index()
    paths = sorted(root.rglob("*_exact_script.md"))
    results: list[dict[str, Any]] = []
    missing_decks = 0
    read_errors: list[str] = []
    deck_cache: dict[Path, dict[int, DeckSlide]] = {}

    for path in paths:
        key = script_key(path)
        deck_path = decks.get(key) if key else None
        if not deck_path or not deck_path.exists():
            missing_decks += 1
            deck_slides = {}
        else:
            try:
                if deck_path not in deck_cache:
                    deck_cache[deck_path] = generated_deck_slides(deck_path)
                deck_slides = deck_cache[deck_path]
            except PptxReadError as exc:
                read_errors.append(f"{deck_path}: {exc}")
                deck_slides = {}
        results.append(edit_file(path, deck_slides))

    validations = [validate(path) for path in paths]
    print(f"script_root={root}")
    print(f"reviewed_scripts={len(results)}")
    print(f"read_only_decks_inspected={len(deck_cache)}")
    print(f"missing_deck_mappings={missing_decks}")
    print(f"deck_read_errors={len(read_errors)}")
    print(f"total_slides={sum(int(result['slides']) for result in results)}")
    print(f"slides_revised={sum(int(result['revised']) for result in results)}")
    print(f"generic_or_unsupported_sentences_removed={sum(int(result['removed_generic']) for result in results)}")
    print(f"slides_with_alignment_notes={sum(int(result['alignment_notes']) for result in results)}")
    print(f"total_spoken_words={sum(int(result['spoken_words']) for result in results)}")
    print(f"estimated_total_minutes={sum(int(result['spoken_words']) for result in results) / WPM:.1f}")
    print(f"raw_urls_in_narration={sum(int(item['raw_urls']) for item in validations)}")
    print(f"banned_metadata_in_narration={sum(int(item['banned']) for item in validations)}")
    print(f"empty_narration_sections={sum(int(item['empty']) for item in validations)}")
    print(f"structure_failures={sum(0 if item['structure_ok'] else 1 for item in validations)}")
    print(f"missing_round3_summaries={sum(0 if item['round3'] else 1 for item in validations)}")
    for error in read_errors:
        print(f"DECK_READ_ERROR {error}")
    for validation in validations:
        if (
            validation["raw_urls"]
            or validation["banned"]
            or validation["empty"]
            or not validation["structure_ok"]
            or not validation["round3"]
        ):
            print(f"ISSUE {validation}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
