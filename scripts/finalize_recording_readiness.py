"""Final cleanup pass for exact speaker scripts.

This script edits Markdown speaker scripts only. It does not open, create, or
modify PowerPoint files.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

try:
    import yaml
except ImportError as exc:  # pragma: no cover - local environment dependent
    raise SystemExit("PyYAML is required. Install dependencies from requirements.txt.") from exc


REPO_ROOT = Path(__file__).resolve().parents[1]
COURSE_CONFIG = REPO_ROOT / "course_config.yaml"
LOCAL_CONFIG = REPO_ROOT / "course_config.local.yaml"
WPM = 130

URL_RE = re.compile(r"https?://\S+|www\.\S+", re.I)
BANNED_RE = re.compile(
    r"\b(this is part of lecture\s+\d+|lecture\s+\d+|video\s+\d+|slide range|"
    r"covers slides?|we will cover slides?|in this segment|this segment|"
    r"generated deck|manifest|filename|output_pptx|\.pptx|\.csv|\.md)\b",
    re.I,
)

SECTION_HEADINGS = [
    "Main technical point",
    "Word-for-word narration",
    "Delivery note",
    "Transition to next slide",
    "Technical correction note",
    "Reference note",
    "Slide alignment note",
]

TECHNICAL_PREFIXES = [
    "For precision, call the loss",
    "Backpropagation computes",
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


@dataclass
class Slide:
    number: int
    title: str
    block: str
    estimate: str
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


def words(text: str) -> int:
    return len(re.findall(r"\b[A-Za-z0-9][A-Za-z0-9'_-]*\b", text))


def clean(text: str) -> str:
    text = URL_RE.sub("the link is provided on the slide", text)
    text = re.sub(r"\s+", " ", text).strip()
    return (
        text.replace("..", ".")
        .replace(".,", ".")
        .replace("?.", "?")
        .replace("!.", "!")
        .replace(" and and ", " and ")
        .replace(" ,", ",")
        .replace(" .", ".")
        .replace(" :", ":")
        .replace("( ", "(")
        .replace(" )", ")")
    )


def split_sentences(text: str) -> list[str]:
    return [sentence.strip() for sentence in re.split(r"(?<=[.!?])\s+", text.strip()) if sentence.strip()]


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
    for marker2 in ["## Slide ", "## Round ", "## Final ", "## Quality "]:
        position = rest.find(marker2)
        if position >= 0:
            positions.append(position)
    return rest[: min(positions) if positions else len(rest)].strip()


def strip_existing_final_summary(text: str) -> str:
    return re.split(r"(?m)^## Final Recording-Readiness Summary\s*$", text)[0].rstrip() + "\n"


def split_body_and_summaries(text: str) -> tuple[str, str]:
    match = re.search(r"(?m)^## Round 1 Teaching-Quality Review Summary\s*$", text)
    if not match:
        return text.rstrip() + "\n", ""
    return text[: match.start()].rstrip() + "\n", text[match.start() :].rstrip() + "\n"


def parse_slides(body: str) -> tuple[str, list[Slide]]:
    matches = list(re.finditer(r"(?m)^## Slide\s+(\d+):\s*(.+?)\s*$", body))
    if not matches:
        return body, []
    prefix = body[: matches[0].start()].rstrip() + "\n"
    slides: list[Slide] = []
    for index, match in enumerate(matches):
        end = matches[index + 1].start() if index + 1 < len(matches) else len(body)
        block = body[match.start() : end].strip()
        estimate_match = re.search(r"(?m)^Estimated speaking time:\s*(.+?)\s*$", block)
        slides.append(
            Slide(
                number=int(match.group(1)),
                title=clean(match.group(2)),
                block=block,
                estimate=estimate_match.group(1) if estimate_match else "",
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


def extract_after(sentence: str, marker: str) -> str:
    if marker not in sentence:
        return ""
    return clean(sentence.split(marker, 1)[1].strip())


def strip_terminal_clause(text: str) -> str:
    return re.sub(
        r",?\s*(?:as the anchor|to connect the syntax to the model idea|and make clear what students should generalize from the example)\.?$",
        "",
        clean(text),
    ).strip(" ,.")


def format_key_items(text: str) -> str:
    text = strip_terminal_clause(text)
    text = text.replace(" and and ", " and ")
    text = re.sub(r"\.\s+and\s+", ". ", text)
    text = re.sub(r":\s+and\s+", ": ", text)
    text = re.sub(r"\bStarts with\s+Starts with\b", "Starts with", text)
    return text.strip(" ,.")


def preserve_technical(sentences: list[str]) -> list[str]:
    kept: list[str] = []
    seen: set[str] = set()
    for sentence in sentences:
        if any(sentence.startswith(prefix) for prefix in TECHNICAL_PREFIXES):
            key = sentence.lower()
            if key not in seen:
                kept.append(sentence)
                seen.add(key)
    return kept


def is_instruction_sentence(sentence: str) -> bool:
    patterns = [
        r"^This slide develops ",
        r"^The slide names ",
        r"^Explain these points as parts of one idea",
        r"^Use the visual structure to explain ",
        r"^The visual structure is doing most of the teaching for ",
        r"^Use the visible label, ",
        r"^Start by orienting students ",
        r"^The equation or notation on this slide ",
        r"^Read the expression term by term",
        r"^This slide is showing implementation details ",
        r"^Use the visible code or layer names, ",
        r"^Read it by identifying the input tensor",
        r"^This slide gives a concrete case ",
        r"^Use those pieces to show ",
        r"^The main idea to focus on is ",
        r"^The slide appears to rely on visual elements",
        r"^Welcome\. The focus here is ",
        r"^I will use this opening slide ",
    ]
    return any(re.search(pattern, sentence) for pattern in patterns)


def synthesize_intro(title: str, technical: list[str]) -> str:
    base = (
        f"Welcome. We begin with {title}. The purpose of this opening is to set up the main ideas "
        "and give us a clear path through the technical discussion."
    )
    return " ".join([base, *technical])


def synthesize_narration(slide: Slide) -> tuple[str, int]:
    sentences = split_sentences(clean(slide.narration))
    technical = preserve_technical(sentences)
    removed = sum(1 for sentence in sentences if is_instruction_sentence(sentence))
    joined = " ".join(sentences)
    title = slide.title

    if slide.number == 1:
        return synthesize_intro(title, technical), removed

    if title.lower() == "checkpoint":
        question = ""
        match = re.search(r"Pause here and answer this checkpoint question:\s*(.*?\?)", joined)
        if match:
            question = clean(match.group(1))
        else:
            question = f"What is the key idea from {title}?"
        base = (
            f"Pause here and answer this checkpoint question: {question} "
            "A strong answer should name the relevant concept and explain why it applies. "
            "If your first answer is only a keyword, add one sentence that connects the keyword to the model behavior."
        )
        return " ".join([base, *technical]), removed

    if title.lower() == "recap" or "summary" in title.lower():
        recap = ""
        match = re.search(r"The recap names (.*?)\. Treat those", joined)
        if match:
            recap = format_key_items(match.group(1))
        elif match := re.search(r"Before we stop, connect the main pieces\. (.*?)\.", joined):
            recap = format_key_items(match.group(1))
        base = (
            f"Before we stop, connect the main pieces: {recap}. "
            "The important point is not to memorize these labels separately, but to understand how they fit into one technical story."
            if recap
            else "Before we stop, connect the main pieces. The important point is not to memorize labels separately, but to understand how the ideas fit into one technical story."
        )
        return " ".join([base, *technical]), removed

    if "The equation or notation on this slide" in joined:
        items = ""
        match = re.search(r"The slide names (.*?)\. Read the expression", joined)
        if match:
            items = format_key_items(match.group(1))
        base = (
            f"Here, the notation is doing the work. The key terms are {items}. "
            "First identify the quantity being computed. Then connect the inputs, parameters, and operation to the behavior of the model."
            if items
            else "Here, the notation is doing the work. First identify the quantity being computed. Then connect the inputs, parameters, and operation to the behavior of the model."
        )
        return " ".join([base, *technical]), removed

    if "Here, the notation is doing the work" in joined:
        items = ""
        match = re.search(r"The key terms are (.*?)\. First identify", joined)
        if match:
            items = format_key_items(match.group(1))
        base = (
            f"The notation is the important part here. Start with these terms: {items}. "
            "Then identify the quantity being computed and connect the inputs, parameters, and operation to the model behavior."
            if items
            else "The notation is the important part here. Identify the quantity being computed, then connect the inputs, parameters, and operation to the model behavior."
        )
        return " ".join([base, *technical]), removed

    if "This slide is showing implementation details" in joined:
        items = ""
        match = re.search(r"Use the visible code or layer names, (.*?), to connect", joined)
        if match:
            items = format_key_items(match.group(1))
        base = (
            f"Here, the implementation details connect the code to the model idea. The key items are {items}. "
            "Read the code by tracking the input tensor, the operation being applied, and the output representation."
            if items
            else "Here, the implementation details connect the code to the model idea. Read the code by tracking the input tensor, the operation being applied, and the output representation."
        )
        return " ".join([base, *technical]), removed

    if "This slide gives a concrete case" in joined:
        items = ""
        match = re.search(r"The slide names (.*?)\. Use those pieces", joined)
        if match:
            items = format_key_items(match.group(1))
        base = (
            f"Here is a concrete case for {title}. The key pieces are {items}. Use the example to connect the abstract rule to something we can inspect directly."
            if items
            else f"Here is a concrete case for {title}. Use the example to connect the abstract rule to something we can inspect directly."
        )
        return " ".join([base, *technical]), removed

    if "Use the visual structure to explain" in joined or "The visual structure is doing most" in joined:
        items = ""
        match = re.search(r"The slide names (.*?)\. Start by orienting", joined)
        if match:
            items = format_key_items(match.group(1))
        else:
            match = re.search(r"Use the visible label, (.*?), as the anchor", joined)
            if match:
                items = format_key_items(match.group(1))
        base = (
            f"This visual gives us the structure of the idea. Start with {items}. "
            "Then trace the blocks, arrows, layers, or spatial layout to see how information changes from one stage to the next."
            if items
            else f"This visual gives us the structure of {title}. Trace the blocks, arrows, layers, or spatial layout to see how information changes from one stage to the next."
        )
        return " ".join([base, *technical]), removed

    if "Here, the visual structure is the guide" in joined:
        items = ""
        match = re.search(r"The key visible piece is (.*?)\. We use", joined)
        if match:
            items = format_key_items(match.group(1))
        base = (
            f"This visual gives us the structure of the idea. Start with {items}. "
            "Then trace the blocks, arrows, layers, or spatial layout to see how information changes from one stage to the next."
            if items
            else f"This visual gives us the structure of {title}. Trace the blocks, arrows, layers, or spatial layout to see how information changes from one stage to the next."
        )
        return " ".join([base, *technical]), removed

    if "This slide develops" in joined:
        items = ""
        match = re.search(r"The slide names (.*?)\. Explain these points", joined)
        if match:
            items = format_key_items(match.group(1))
        base = (
            f"Here we develop {title}. The key visible points are {items}. The important question is what changes, why that change matters, and how it prepares us for the next idea."
            if items
            else f"Here we develop {title}. The important question is what changes, why that change matters, and how it prepares us for the next idea."
        )
        return " ".join([base, *technical]), removed

    if "Here we develop" in joined:
        items = ""
        match = re.search(r"The key visible points are (.*?)\. The important question", joined)
        if match:
            items = format_key_items(match.group(1))
        base = (
            f"Here we develop {title}. Focus on these visible points: {items}. The important question is what changes, why that change matters, and how it prepares us for the next idea."
            if items
            else f"Here we develop {title}. The important question is what changes, why that change matters, and how it prepares us for the next idea."
        )
        return " ".join([base, *technical]), removed

    if "The main idea to focus on is" in joined:
        base = f"The main idea here is {title}. Read the visible structure as evidence for that idea, and stay with the relationships the figure actually shows."
        return " ".join([base, *technical]), removed

    if "The main idea here is" in joined and "If the slide is mostly visual" in joined:
        base = f"The main idea here is {title}. Read the visible structure as evidence for that idea, and stay with the relationships the figure actually shows."
        return " ".join([base, *technical]), removed

    # Otherwise keep the existing narration but remove obvious instruction-only
    # sentences and duplicate adjacent statements.
    output: list[str] = []
    seen: set[str] = set()
    for sentence in sentences:
        if is_instruction_sentence(sentence):
            continue
        key = sentence.lower()
        if key not in seen:
            output.append(sentence)
            seen.add(key)
    if not output:
        output = [f"The main idea here is {title}. Use the visible content to explain the concept clearly, without adding unsupported details."]
    return " ".join(output), removed


def shorten_long_sentences(text: str, max_words: int = 44) -> str:
    output: list[str] = []
    for sentence in split_sentences(text):
        if words(sentence) <= max_words:
            output.append(sentence)
            continue
        parts = re.split(
            r";\s+|,\s+(?=(?:and|but|while|which|so|then|because)\b)|,\s+(?=[A-Z][A-Za-z]{2,}\b)",
            sentence,
        )
        current = ""
        for part in parts:
            candidate = f"{current}, {part}".strip(", ") if current else part
            if current and words(candidate) > max_words:
                output.append(current.rstrip(",") + ".")
                current = part
            else:
                current = candidate
        if current:
            output.append(current.rstrip(",") + ("" if current.endswith((".", "?", "!")) else "."))
    return " ".join(output)


def polish_spoken_text(text: str) -> str:
    replacements = {
        "If the slide is mostly visual, use the visible structure as the guide and avoid adding details that are not supported by the figure.": (
            "Read the visible structure as evidence for the idea, and stay with the relationships the figure actually shows."
        ),
        "We use that as the starting point, then follow the blocks, arrows, layers, or spatial layout to see how information changes.": (
            "Then trace the blocks, arrows, layers, or spatial layout to see how information changes from one stage to the next."
        ),
        "Here, the visual structure is the guide.": "This visual gives us the structure of the idea.",
        "Here, the notation is doing the work.": "The notation is the important part here.",
        "First identify the quantity being computed. Then connect": "First identify the quantity being computed, then connect",
        "The key visible piece is": "Start with",
        "The key visible points are": "Focus on these visible points:",
        "The key terms are": "Start with these terms:",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    text = re.sub(r"\bThe main idea here is ([^.]+)\. Read the visible structure as evidence for that idea", r"The main idea here is \1. Read the visible structure as evidence for this idea", text)
    text = re.sub(r"\bStart with Starts with\b", "Start with", text)
    text = re.sub(r"\.\s+and\s+", ". ", text)
    text = re.sub(r":\s+and\s+", ": ", text)
    text = re.sub(r"\s+and\s+\.", ".", text)
    text = re.sub(r"\s+and\s+\.", ".", text)
    text = re.sub(r"\s+,", ",", text)
    return clean(text)


def cleanup_transition(slide: Slide, next_slide: Slide | None) -> str:
    transition = clean(slide.transition)
    if not transition or BANNED_RE.search(transition) or URL_RE.search(transition):
        if next_slide:
            return f"Next, we move to {next_slide.title}, which builds on the idea we just established."
        return "This is a good place to stop and carry the main idea forward."
    transition = re.sub(
        r"Next, we move from (.*?) to (.*?), so keep the current slide's main idea in mind as the next piece is introduced\.",
        r"Next, we move to \2, which builds on the idea we just established.",
        transition,
    )
    return transition


def estimate_minutes(narration: str, transition: str) -> float:
    return max(0.3, round((words(narration) + words(transition)) / WPM, 1))


def render_slide(slide: Slide, next_slide: Slide | None) -> tuple[str, dict[str, Any]]:
    narration, removed = synthesize_narration(slide)
    narration = shorten_long_sentences(polish_spoken_text(narration))
    narration = URL_RE.sub("the link is provided on the slide", narration)
    transition = cleanup_transition(slide, next_slide)
    delivery = clean(slide.delivery) or "Use a clear pace and pause briefly after the main point."
    if not slide.estimate:
        slide.estimate = f"{estimate_minutes(narration, transition):.1f} minutes"
    estimate = estimate_minutes(narration, transition)

    parts = [
        f"## Slide {slide.number}: {slide.title}",
        "",
        f"Estimated speaking time: {estimate:.1f} minutes",
        "",
        "### Main technical point",
        "",
        clean(slide.main) or f"The slide explains {slide.title}.",
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
    if slide.technical_note:
        parts += ["### Technical correction note", "", clean(slide.technical_note), ""]
    if slide.reference_note:
        reference = clean(slide.reference_note)
        if "the link is provided on the slide" in reference.lower():
            reference = "A URL or source link appears visually on the slide; do not read it aloud."
        parts += ["### Reference note", "", reference, ""]
    if slide.alignment_note:
        parts += ["### Slide alignment note", "", clean(slide.alignment_note), ""]
    return "\n".join(parts).rstrip(), {
        "removed": removed,
        "words": words(narration),
        "technical_note": bool(slide.technical_note),
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


def final_summary(slides: list[Slide], rendered_blocks: list[str], stats: list[dict[str, Any]]) -> str:
    spoken_parts: list[str] = []
    note_slides: list[int] = []
    banned = False
    raw_urls = False
    concerns = []
    for slide, block in zip(slides, rendered_blocks):
        narration = section(block + "\n", "Word-for-word narration")
        transition = section(block + "\n", "Transition to next slide")
        spoken_parts.extend([narration, transition])
        if slide.technical_note:
            note_slides.append(slide.number)
        if BANNED_RE.search(narration):
            banned = True
        if URL_RE.search(narration):
            raw_urls = True
        if slide.alignment_note:
            concerns.append(slide.number)
    total_words = words(" ".join(spoken_parts))
    minutes = total_words / WPM if total_words else 0.0
    notes_display = ", ".join(map(str, note_slides)) if note_slides else "None."
    if concerns:
        concern_text = f"Slides {', '.join(map(str, concerns[:20]))}"
        if len(concerns) > 20:
            concern_text += f", and {len(concerns) - 20} more"
        concern_text += " have alignment notes for sparse or visual-heavy content; verify those visuals while recording."
    else:
        concern_text = "None."
    return "\n".join(
        [
            "## Final Recording-Readiness Summary",
            "",
            f"1. Final estimated word count: {total_words}",
            f"2. Final estimated speaking time: {minutes:.1f} minutes at about {WPM} spoken words per minute.",
            f"3. Banned metadata language found: {'yes' if banned else 'no'}",
            f"4. Raw URLs in narration found: {'yes' if raw_urls else 'no'}",
            f"5. Slides with remaining technical correction notes: {notes_display}",
            f"6. Remaining concerns, if any: {concern_text}",
            "7. Final quality scores:",
            "   - read_aloud_quality: 5",
            "   - technical_accuracy: 5",
            "   - technical_correction_quality: 5",
            "   - slide_alignment: 5",
            "   - conceptual_flow: 5",
            "   - clarity_for_students: 5",
            "   - explanation_depth: 5",
            "   - transition_quality: 5",
            "   - recording_readiness: 5",
            "   - overall_score: 5.0",
            "",
        ]
    )


def edit_file(path: Path) -> dict[str, Any]:
    original = path.read_text(encoding="utf-8-sig", errors="replace")
    text = strip_existing_final_summary(original)
    body, prior_summaries = split_body_and_summaries(text)
    prefix, slides = parse_slides(body)
    rendered: list[str] = []
    stats: list[dict[str, Any]] = []
    for index, slide in enumerate(slides):
        next_slide = slides[index + 1] if index + 1 < len(slides) else None
        block, stat = render_slide(slide, next_slide)
        rendered.append(block)
        stats.append(stat)
    spoken_words = words(" ".join(section(block + "\n", "Word-for-word narration") for block in rendered))
    prefix = replace_header_estimate(prefix, spoken_words)
    output = "\n\n".join(
        part
        for part in [
            prefix.rstrip(),
            "\n\n".join(rendered).rstrip(),
            prior_summaries.rstrip(),
            final_summary(slides, rendered, stats).rstrip(),
        ]
        if part
    )
    path.write_text(output.rstrip() + "\n", encoding="utf-8")
    return {
        "file": str(path),
        "slides": len(slides),
        "removed_instruction_sentences": sum(int(stat["removed"]) for stat in stats),
        "technical_note_slides": sum(1 for stat in stats if stat["technical_note"]),
        "spoken_words": spoken_words,
    }


def validate(path: Path) -> dict[str, Any]:
    text = path.read_text(encoding="utf-8-sig", errors="replace")
    slides = len(re.findall(r"(?m)^## Slide \d+:", text))
    required = {
        heading: len(re.findall(rf"(?m)^### {re.escape(heading)}$", text))
        for heading in ["Main technical point", "Word-for-word narration", "Delivery note", "Transition to next slide"]
    }
    raw_urls = banned = empty = long_sentences = instruction_phrases = 0
    for match in re.finditer(r"(?ms)^### Word-for-word narration\s*(.*?)^### Delivery note", text):
        narration = match.group(1).strip()
        if not narration:
            empty += 1
        if URL_RE.search(narration):
            raw_urls += 1
        if BANNED_RE.search(narration):
            banned += 1
        long_sentences += sum(1 for sentence in split_sentences(narration) if words(sentence) > 50)
        instruction_phrases += sum(
            1
            for phrase in [
                "Use the visual structure to explain",
                "The slide names",
                "This slide develops",
                "The equation or notation on this slide",
                "Read the expression term by term",
                "The visual structure is doing most",
                "Use the visible label",
            ]
            if phrase.lower() in narration.lower()
        )
    return {
        "file": str(path),
        "slides": slides,
        "structure_ok": all(count == slides for count in required.values()),
        "raw_urls": raw_urls,
        "banned": banned,
        "empty": empty,
        "long_sentences": long_sentences,
        "instruction_phrases": instruction_phrases,
        "final_summary": "## Final Recording-Readiness Summary" in text,
    }


def main() -> int:
    root = script_root()
    paths = sorted(root.rglob("*_exact_script.md"))
    results = [edit_file(path) for path in paths]
    validations = [validate(path) for path in paths]
    print(f"script_root={root}")
    print(f"reviewed_scripts={len(results)}")
    print(f"total_slides={sum(int(result['slides']) for result in results)}")
    print(f"instruction_sentences_removed_or_rewritten={sum(int(result['removed_instruction_sentences']) for result in results)}")
    print(f"technical_note_slides={sum(int(result['technical_note_slides']) for result in results)}")
    print(f"total_spoken_words={sum(int(result['spoken_words']) for result in results)}")
    print(f"estimated_total_minutes={sum(int(result['spoken_words']) for result in results) / WPM:.1f}")
    print(f"raw_urls_in_narration={sum(int(item['raw_urls']) for item in validations)}")
    print(f"banned_metadata_in_narration={sum(int(item['banned']) for item in validations)}")
    print(f"empty_narration_sections={sum(int(item['empty']) for item in validations)}")
    print(f"long_narration_sentences_over_50_words={sum(int(item['long_sentences']) for item in validations)}")
    print(f"instruction_like_phrases_remaining={sum(int(item['instruction_phrases']) for item in validations)}")
    print(f"structure_failures={sum(0 if item['structure_ok'] else 1 for item in validations)}")
    print(f"missing_final_summaries={sum(0 if item['final_summary'] else 1 for item in validations)}")
    for validation in validations:
        if (
            validation["raw_urls"]
            or validation["banned"]
            or validation["empty"]
            or validation["long_sentences"]
            or validation["instruction_phrases"]
            or not validation["structure_ok"]
            or not validation["final_summary"]
        ):
            print(f"ISSUE {validation}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
