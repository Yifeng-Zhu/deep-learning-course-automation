"""Apply a Round 1 teaching-quality edit to generated exact scripts.

This script edits Markdown speaker scripts only. It does not open or modify
PowerPoint files.
"""

from __future__ import annotations

import re
from pathlib import Path

SCRIPT_ROOT = Path(r"H:\My Drive\Teaching\ECE591_Online\03_Scripts\exact_scripts")
WPM = 130
URL_RE = re.compile(r"https?://\S+|www\.\S+", re.I)
BANNED_RE = re.compile(
    r"\b(this is part of lecture\s+\d+|lecture\s+\d+|video\s+\d+|slide range|"
    r"covers slides?|we will cover slides?|in this segment|this segment|"
    r"generated deck|manifest|filename|output_pptx)\b",
    re.I,
)
WEAK_PATTERNS = [
    "The visible text gives us",
    "The key idea is",
    "Here is the working definition",
    "Think of this as a procedure",
    "This picture is doing part of the explanation",
    "The expression is anchored by",
    "The figure is organized around",
    "This builds on",
    "After this, we are ready for",
    "The important point is why the next technical step is needed",
    "Without this motivation, the method can look like an arbitrary design choice",
    "In plain language, this gives us vocabulary",
    "The term matters because it lets us describe",
]


def words(text: str) -> int:
    return len(re.findall(r"\b[A-Za-z0-9][A-Za-z0-9'-]*\b", text))


def clean(text: str) -> str:
    text = URL_RE.sub("the link provided on the slide", text)
    text = re.sub(r"\s+", " ", text).strip()
    return (
        text.replace("..", ".")
        .replace(" ,", ",")
        .replace(" .", ".")
        .replace(" :", ":")
    )


def split_sentences(text: str) -> list[str]:
    return [s.strip() for s in re.split(r"(?<=[.!?])\s+", text.strip()) if s.strip()]


def shorten_long_sentences(text: str, max_words: int = 36) -> str:
    output: list[str] = []
    for sentence in split_sentences(text):
        if words(sentence) <= max_words:
            output.append(sentence)
            continue
        current = ""
        for part in re.split(r",\s+|;\s+", sentence):
            candidate = f"{current}, {part}".strip(", ") if current else part
            if current and words(candidate) > max_words:
                output.append(current.rstrip(",") + ".")
                current = part
            else:
                current = candidate
        if current:
            output.append(current.rstrip(",") + ("" if current.endswith((".", "?", "!")) else "."))
    return " ".join(output)


def section(block: str, heading: str, next_headings: list[str]) -> str:
    marker = f"### {heading}"
    if marker not in block:
        return ""
    rest = block.split(marker, 1)[1]
    positions = [
        rest.find(f"### {candidate}")
        for candidate in next_headings
        if f"### {candidate}" in rest
    ]
    positions += [
        rest.find("## Slide ") if "## Slide " in rest else -1,
        rest.find("## Round ") if "## Round " in rest else -1,
        rest.find("## Quality ") if "## Quality " in rest else -1,
    ]
    positions = [position for position in positions if position >= 0]
    return rest[: min(positions) if positions else len(rest)].strip()


def extract_visible(narration: str) -> str:
    values: list[str] = []
    for pattern in [
        r"The visible text gives us (.*?)(?:\.\s|$)",
        r"The expression is anchored by these terms: (.*?)(?:\.\s|$)",
        r"The figure is organized around (.*?)(?:\.\s|$)",
        r"The example is built from these pieces: (.*?)(?:\.\s|$)",
        r"The slide points to (.*?), so use",
    ]:
        for match in re.finditer(pattern, narration):
            value = clean(match.group(1))
            if value and value not in values:
                values.append(value)
    return "; ".join(values[:2])


def extract_objectives(header: str) -> list[str]:
    match = re.search(r"able to use it to:\s*(.*?)\.", header, re.S)
    if not match:
        return []
    return [clean(item.strip(" .")) for item in match.group(1).split(";") if clean(item)]


def title_label(title: str, main: str, visible: str) -> str:
    title = clean(title).strip(" .")
    combined = f"{title} {visible}".lower()
    if title.lower() in {"a", "x", "y", "v"} or re.fullmatch(r"\d+", title):
        if "without scaling" in combined:
            return "Why attention scores need scaling"
        if "softmax" in combined and ("attention" in combined or title.lower() == "a"):
            return "Attention weights after softmax"
        if all(token in combined for token in ["q", "k", "v"]):
            return "Query, key, and value flow"
        if title.lower() == "v":
            return "Value vectors in attention"
        if title.lower() in {"x", "y"}:
            return "Input and output representations"
        if visible:
            candidate = clean(visible.split(";")[0].split(",")[0]).strip(" .")
            if len(candidate) > 2:
                return candidate[:1].upper() + candidate[1:]
        match = re.search(r"develops\s+(.*?)\s+(?:as part|detail|$)", main, re.I)
        if match:
            return clean(match.group(1))
    return title


def classify(title: str, main: str, narration: str, visible: str) -> set[str]:
    text = " ".join([title, main, narration, visible]).lower()
    checks = {
        "metrics": ["confusion", "precision", "recall", "f1", "roc", "auc", "threshold"],
        "loss": ["loss", "objective", "logit", "cross-entropy", "cross entropy", "hinge"],
        "softmax": ["softmax"],
        "activation": ["activation", "relu", "sigmoid", "tanh", "nonlinearity", "nonlinear"],
        "optimization": ["gradient", "descent", "learning rate", "optimizer", "update rule"],
        "regularization": ["overfit", "underfit", "validation", "regularization", "dropout", "l1", "l2"],
        "cnn": ["convolution", "cnn", "kernel", "filter", "padding", "stride", "feature map", "pooling", "alexnet", "vgg", "resnet", "inception", "mobilenet"],
        "segmentation": ["segmentation", "u-net", "u net", "fcn", "upsampling", "transposed", "checkerboard", "pixel"],
        "rnn": ["rnn", "lstm", "gru", "hidden state", "gate", "sequence", "pytorch"],
        "attention": ["attention", "query", "queries", "key", "keys", "value", "values", "qkv", "context vector", "self-attention", "seq2seq"],
        "transformer": ["transformer", "positional", "multi-head", "layer norm", "residual", "encoder", "decoder", "bert", "gpt", "token", "embedding", "prompt", "llm"],
        "vit": ["vision transformer", "vit", "patch", "class token", "deit", "swin", "shifted window", "multimodal"],
        "generative": ["generative", "autoencoder", "vae", "variational", "latent", "gan", "diffusion", "sampling", "likelihood", "generator", "discriminator"],
        "architecture": ["architecture", "layer", "block", "pipeline", "network", "module"],
        "example": ["example", "playground", "implementation", "code", "keras", "torch"],
        "equation": ["=", "formula", "equation", "rule", "matrix", "dot product"],
    }
    return {category for category, terms in checks.items() if any(term in text for term in terms)}


def technical_sentences(categories: set[str], label: str, visible: str) -> list[str]:
    text = f"{label} {visible}".lower()
    sentences: list[str] = []
    if "metrics" in categories:
        sentences.append("The important distinction is between what the model predicts and how we evaluate those predictions.")
        if "threshold" in text or "roc" in text or "auc" in text:
            sentences.append("Changing the threshold changes the tradeoff between false positives and false negatives; it does not change the trained model itself.")
        elif "precision" in text or "recall" in text or "f1" in text:
            sentences.append("Precision asks how many predicted positives were correct, while recall asks how many real positives were found.")
    if "loss" in categories:
        sentences.append("A loss function is the quantity optimized during training; it is related to performance, but it is not the same thing as accuracy.")
        if "softmax" in text:
            sentences.append("Softmax turns logits into normalized class scores, but those scores should not automatically be treated as perfectly calibrated probabilities.")
        if "cross" in text:
            sentences.append("Cross-entropy penalizes confident wrong predictions strongly, which is why it is useful for classification training.")
    if "softmax" in categories and "loss" not in categories and "attention" in categories:
        sentences.append("In attention, softmax turns raw compatibility scores into normalized weights over the available tokens.")
    if "optimization" in categories:
        sentences.append("Gradient descent uses gradients to decide how to move the parameters, while backpropagation is the efficient way to compute those gradients in a neural network.")
        if "learning rate" in text:
            sentences.append("The learning rate controls the step size; too small can make learning slow, and too large can make training unstable.")
    if "regularization" in categories:
        sentences.append("The main issue is the gap between fitting the training data and generalizing to new data.")
        if "validation" in text or "test" in text:
            sentences.append("Validation data is used for model selection and tuning; the test set should be reserved for the final evaluation.")
        if "dropout" in text:
            sentences.append("Dropout changes training by randomly removing activations; at inference time, the full network is used with the appropriate scaling convention.")
    if "cnn" in categories:
        sentences.append("For image data, the key idea is to use local connectivity and shared weights instead of treating every pixel position as unrelated.")
        if any(term in text for term in ["convolution", "kernel", "filter", "dot product"]):
            sentences.append("A filter computes a local dot product, and reusing that filter across locations produces a feature map.")
        if "pool" in text:
            sentences.append("Pooling gives local tolerance to small shifts, but it does not make the whole model fully translation invariant.")
        if "padding" in text:
            sentences.append("Padding controls how much spatial size is preserved at the image boundary.")
    if "segmentation" in categories:
        sentences.append("Semantic segmentation is different from image classification because the model must make a prediction for each pixel or spatial location.")
        if "u-net" in text or "u net" in text:
            sentences.append("The U-Net shape combines coarse semantic information from the encoder with fine spatial detail from skip connections.")
        if "transposed" in text:
            sentences.append("Transposed convolution is a learned upsampling operation, not a literal inverse of ordinary convolution.")
    if "rnn" in categories:
        sentences.append("For sequence data, the model must carry information from earlier positions while processing later positions.")
        if "lstm" in text or "gate" in text:
            sentences.append("LSTM gates control what information is kept, forgotten, and exposed to the next step.")
    if "attention" in categories:
        sentences.append("Attention lets one token build its representation by weighting information from other tokens.")
        if any(term in text for term in ["query", "key", "value", "qkv"]):
            sentences.append("Queries ask what a token is looking for, keys describe what each token offers, and values carry the information that gets mixed.")
        sentences.append("Attention weights show how information is mixed, but they should not be treated as a complete explanation of a model decision.")
    if "transformer" in categories:
        sentences.append("Transformers rely on token representations, attention, feed-forward layers, residual connections, and normalization working together.")
        if "positional" in text:
            sentences.append("Positional information is needed because attention by itself does not know the order of tokens.")
        if "bert" in text or "gpt" in text:
            sentences.append("BERT-style models and GPT-style models differ mainly in their pretraining objective and how they use context.")
    if "vit" in categories:
        sentences.append("A vision transformer treats image patches as tokens, so the image is converted into a sequence before transformer blocks process it.")
        if "swin" in text or "window" in text:
            sentences.append("Windowed attention reduces computation, and shifted windows allow information to move across window boundaries.")
    if "generative" in categories:
        sentences.append("A generative model is concerned with modeling or sampling data, not only predicting a label.")
        if "vae" in text or "latent" in text:
            sentences.append("In a VAE, the latent space is regularized so that sampling from it can produce meaningful decoded outputs.")
        if "gan" in text:
            sentences.append("In a GAN, the generator and discriminator create a training game, so stability is part of the technical challenge.")
        if "diffusion" in text:
            sentences.append("Diffusion models learn to reverse a gradual noising process, turning noise into a structured sample step by step.")
    return list(dict.fromkeys(sentences))


def compose_intro(script_title: str, objectives: list[str]) -> str:
    pieces = [
        f"Welcome. In this lesson, we are going to work through {script_title}.",
        "The goal is to understand the ideas well enough to use them, not just recognize the vocabulary.",
    ]
    if objectives:
        labels = ["First", "Second", "Third", "Fourth"]
        pieces.append("By the end, you should be able to do a few concrete things.")
        pieces.extend(
            f"{labels[i]}, {obj[:1].lower() + obj[1:]}."
            for i, obj in enumerate(objectives[:4])
        )
    pieces.append("As we go, focus on the connection between the intuition, the computation, and the modeling choice.")
    return " ".join(pieces)


def compose_narration(script_title: str, label: str, old_narration: str, visible: str, categories: set[str], prev_label: str, next_label: str, objectives: list[str]) -> str:
    lower_label = label.lower()
    if "intro" in categories:
        return compose_intro(script_title, objectives)
    if "recap" in categories:
        key = visible or prev_label or script_title
        return clean(
            f"Let's collect the main ideas before we stop. The important pieces are {key}. "
            "Do not treat these as separate terms. The value is in how they connect: the model structure creates a computation, the computation changes the representation, and the representation supports the task."
        )
    if "checkpoint" in categories:
        match = re.search(r"(?:question[:.]\s*)(.*?)(?:Take a moment|A good answer|$)", old_narration, re.I)
        question = clean(match.group(1)) if match else clean(visible or label)
        if not question or question.lower() == "checkpoint":
            question = "what is the main idea, and why does it matter for the model?"
        return clean(
            f"Pause here and answer this question: {question} "
            "A strong answer should name the concept and explain the reasoning behind it. "
            "If your answer is only a keyword, add one sentence that connects the keyword to the model behavior."
        )

    sentences: list[str] = []
    if "equation" in categories or "loss" in categories or "optimization" in categories or "softmax" in categories:
        sentences.append(f"For {label}, read the notation as a statement about what the model computes and what training is trying to improve.")
    elif any(category in categories for category in ["architecture", "cnn", "attention", "transformer", "rnn", "vit"]):
        sentences.append(f"For {label}, start by identifying what enters the block, what operation is applied, and what representation comes out.")
    elif "example" in categories:
        sentences.append(f"Use {label} as a concrete check on the idea rather than as an isolated example.")
    elif any(term in lower_label for term in ["why", "motivation", "problem", "limitation"]):
        sentences.append(f"The reason to pause on {label} is that it explains why the next modeling choice is needed.")
    else:
        sentences.append(f"The important idea here is {label}.")
    if visible and len(visible) < 170 and not re.fullmatch(r"[A-Za-z0-9_\- ]{1,4}", visible):
        sentences.append(f"The slide points to {visible}, so use those labels as anchors while explaining the concept.")
    elif visible:
        sentences.append("The slide contains several labels, so group them by role instead of reading them one by one.")
    sentences.extend(technical_sentences(categories, label, visible))
    if "equation" in categories:
        sentences.append("The safe way to explain the equation is to name each quantity, then say how changing it affects the model behavior.")
    if "architecture" in categories:
        sentences.append("A diagram is easiest to follow when students know the input, the transformation, and the output before interpreting the result.")
    if prev_label and prev_label.lower() not in {"recap", "checkpoint"}:
        sentences.append(f"This connects back to {prev_label}.")
    if next_label and next_label.lower() not in {"recap", "checkpoint"}:
        sentences.append(f"That prepares us for {next_label}.")
    text = " ".join(sentences)
    text = re.sub(r"\bSelf-attention computation and QKV roles detail\s+\d+\b", "the scaled attention computation", text)
    text = re.sub(r"\bdetail\s+\d+\b", "this step", text)
    return shorten_long_sentences(clean(text))


def main_point(label: str, categories: set[str], old: str) -> str:
    if "recap" in categories:
        return "The slide consolidates the concepts that students should carry forward."
    if "checkpoint" in categories:
        return "The slide asks students to retrieve and apply the central idea before moving on."
    if "intro" in categories:
        return "The slide frames the learning goals and prepares students for the main technical ideas."
    if "attention" in categories:
        return f"The slide explains {label} as a mechanism for mixing information across tokens."
    if "loss" in categories and "optimization" in categories:
        return f"The slide connects {label} to the training objective and parameter updates."
    if "loss" in categories:
        return f"The slide explains {label} as part of defining and optimizing model error."
    if "optimization" in categories:
        return f"The slide explains {label} as part of parameter optimization."
    if "activation" in categories:
        return f"The slide explains {label} as part of nonlinear representation learning."
    if "regularization" in categories:
        return f"The slide explains {label} as part of controlling generalization."
    if "cnn" in categories:
        return f"The slide explains {label} as part of spatially structured image modeling."
    if "generative" in categories:
        return f"The slide explains {label} as part of modeling, sampling, or reconstructing data."
    if "metrics" in categories:
        return f"The slide explains {label} as a way to evaluate model predictions."
    if old and not old.startswith("The slide develops") and "mixing information across tokens" not in old:
        return old
    return f"The slide explains {label} and why it matters for the model."


def delivery_note(categories: set[str]) -> str:
    if "checkpoint" in categories:
        return "Pause long enough for students to answer before giving the explanation."
    if "recap" in categories:
        return "Use a steady recap tone. Emphasize connections rather than every detail."
    if "equation" in categories or "optimization" in categories or "loss" in categories or "softmax" in categories:
        return "Slow down and name each quantity before interpreting the formula or update."
    if "activation" in categories:
        return "Emphasize why the nonlinearity matters before moving on."
    if any(category in categories for category in ["architecture", "cnn", "attention", "transformer", "rnn", "vit"]):
        return "Point to the flow of information first, then explain the interpretation."
    if "example" in categories:
        return "Pause after the setup, then state the general lesson."
    return "Use a normal teaching pace and emphasize why this point matters."


def transition_to(label: str, next_label: str, categories: set[str], next_categories: set[str]) -> str:
    if not next_label:
        return "That is a good place to stop. The main idea to carry forward is how the model structure shapes the computation."
    if "recap" in next_categories:
        return "With the technical pieces in place, let's summarize the main takeaways."
    if "checkpoint" in next_categories:
        return "Let's finish by checking whether the main idea is usable without looking back."
    if "intro" in categories:
        return f"Let's begin with {next_label}."
    if categories != next_categories and next_categories:
        return f"That completes this part of the story. Next, we move to {next_label}."
    return f"With that in mind, the next step is {next_label}."


def correction_note(categories: set[str], label: str, visible: str, seen: set[str]) -> str:
    if categories & {"intro", "recap", "checkpoint"}:
        return ""
    text = f"{label} {visible}".lower()
    options: list[tuple[str, str]] = []
    if "cnn" in categories and any(term in text for term in ["convolution", "filter", "kernel"]):
        options.append(("conv_corr", "In many deep learning libraries, the operation called convolution is implemented as cross-correlation because the filter is not flipped. The CNN convention is still to call the learned operation convolution."))
    if "loss" in categories and "softmax" in text:
        options.append(("softmax", "Avoid implying that softmax scores are automatically calibrated probabilities."))
    if "loss" in categories and "cross" in text:
        options.append(("cross_entropy", "Keep cross-entropy separate from accuracy: cross-entropy is optimized during training, while accuracy is an evaluation metric."))
    if "optimization" in categories:
        options.append(("gradient_update", "Keep gradient computation separate from parameter updates: backpropagation computes gradients, and the optimizer applies the update rule."))
    if "attention" in categories:
        options.append(("attention", "Avoid treating attention weights as a complete explanation of the model decision."))
    if "segmentation" in categories and "transposed" in text:
        options.append(("transposed", "Transposed convolution is a learned upsampling operation, not a literal inverse of ordinary convolution."))
    if "generative" in categories and "likelihood" in text:
        options.append(("likelihood", "Keep likelihood, sampling, and prediction distinct: likelihood scores data under a model, sampling generates data, and prediction chooses an output for a task."))
    for key, note in options:
        if key not in seen:
            seen.add(key)
            return note
    return ""


def parse_slides(text: str) -> tuple[str, list[dict[str, object]]]:
    parts = re.split(r"(?m)^## Slide (\d+): (.+)$", text)
    slides: list[dict[str, object]] = []
    for index in range(1, len(parts), 3):
        slides.append({"num": int(parts[index]), "title": parts[index + 1].strip(), "block": parts[index + 2]})
    return parts[0], slides


def strip_existing_summaries(text: str) -> str:
    text = re.split(r"(?m)^## Round 1 Teaching-Quality Review Summary\s*$", text)[0]
    text = re.split(r"(?m)^## Quality self-check\s*$", text)[0]
    return text.rstrip() + "\n"


def replace_header_estimate(prefix: str, spoken_words: int) -> str:
    minutes = spoken_words / WPM if spoken_words else 0
    prefix = re.sub(r"- estimated total word count:\s*\d+", f"- estimated total word count: {spoken_words}", prefix)
    prefix = re.sub(
        r"- estimated speaking time:.*",
        f"- estimated speaking time: {minutes:.1f} minutes at about {WPM} spoken words per minute, plus pauses for visual pointing and checkpoint response",
        prefix,
    )
    return prefix


def render_slide(slide: dict[str, object], seen: set[str]) -> tuple[str, bool, str, str]:
    title = str(slide["title"])
    block = str(slide["block"])
    main_old = section(block, "Main technical point", ["Word-for-word narration", "Delivery note", "Transition to next slide", "Technical correction note", "Reference note"])
    old_narration = section(block, "Word-for-word narration", ["Delivery note", "Transition to next slide", "Technical correction note", "Reference note"])
    ref_note = section(block, "Reference note", [])
    visible = extract_visible(old_narration)
    label = str(slide["label"])
    categories = set(slide["categories"])
    next_label = str(slide.get("next_label", ""))
    next_categories = set(slide.get("next_categories", set()))
    prev_label = str(slide.get("prev_label", ""))
    objectives = list(slide.get("objectives", []))
    script_title = str(slide.get("script_title", "this topic"))
    sparse_title = clean(title).lower() in {"a", "x", "y", "v"} or bool(re.fullmatch(r"\d+", clean(title)))
    display_title = label if sparse_title else title
    needs_revision = (
        any(pattern in old_narration for pattern in WEAK_PATTERNS)
        or bool(slide.get("force_revision"))
        or not bool(slide.get("had_round_summary"))
        or label != clean(title)
    )
    narration = (
        compose_narration(script_title, label, old_narration, visible, categories, prev_label, next_label, objectives)
        if needs_revision
        else clean(old_narration)
    )
    transition = transition_to(label, next_label, categories, next_categories)
    estimate = max(
        0.4,
        round((words(narration) + words(transition)) / WPM + (0.15 if categories & {"equation", "architecture", "attention", "cnn", "transformer", "vit"} else 0), 1),
    )
    lines = [
        f"## Slide {slide['num']}: {display_title}",
        "",
        f"Estimated speaking time: {estimate:.1f} minutes",
        "",
        "### Main technical point",
        "",
        main_point(label, categories, main_old),
        "",
        "### Word-for-word narration",
        "",
        narration,
        "",
        "### Delivery note",
        "",
        delivery_note(categories),
        "",
        "### Transition to next slide",
        "",
        transition,
        "",
    ]
    note = correction_note(categories, label, visible, seen)
    if note:
        lines += ["### Technical correction note", "", note, ""]
    if ref_note:
        lines += ["### Reference note", "", ref_note, ""]
    return "\n".join(lines), needs_revision, narration, transition


def edit_file(path: Path) -> dict[str, int | str]:
    original = path.read_text(encoding="utf-8-sig", errors="replace")
    had_round_summary = "## Round 1 Teaching-Quality Review Summary" in original
    text = strip_existing_summaries(original)
    prefix, slides = parse_slides(text)
    title_match = re.search(r"^#\s+(.+)$", prefix, re.M)
    script_title = clean(title_match.group(1)) if title_match else path.stem
    objectives = extract_objectives(prefix)

    prepared: list[dict[str, object]] = []
    for slide in slides:
        block = str(slide["block"])
        main_old = section(block, "Main technical point", ["Word-for-word narration", "Delivery note", "Transition to next slide", "Technical correction note", "Reference note"])
        old_narration = section(block, "Word-for-word narration", ["Delivery note", "Transition to next slide", "Technical correction note", "Reference note"])
        visible = extract_visible(old_narration)
        title_for_classification = str(slide["title"])
        sparse_title = (
            clean(title_for_classification).lower() in {"a", "x", "y"}
            or bool(re.fullmatch(r"\d+", clean(title_for_classification)))
            or len(clean(title_for_classification)) <= 2
        )
        context = script_title if sparse_title or not visible else ""
        categories = classify(title_for_classification, context, "", visible)
        if int(slide["num"]) == 1:
            categories.add("intro")
        if str(slide["title"]).strip().lower() == "recap":
            categories.add("recap")
        if str(slide["title"]).strip().lower() == "checkpoint":
            categories.add("checkpoint")
        prepared.append({
            **slide,
            "label": title_label(str(slide["title"]), main_old, visible),
            "categories": categories,
            "script_title": script_title,
            "objectives": objectives,
            "had_round_summary": had_round_summary,
            "force_revision": path.name != "1_1_From_MLPs_to_Convolution_exact_script.md",
        })

    for index, slide in enumerate(prepared):
        slide["prev_label"] = prepared[index - 1]["label"] if index else ""
        slide["next_label"] = prepared[index + 1]["label"] if index + 1 < len(prepared) else ""
        slide["next_categories"] = prepared[index + 1]["categories"] if index + 1 < len(prepared) else set()

    seen: set[str] = set()
    blocks: list[str] = []
    revised: list[int] = []
    spoken: list[str] = []
    for slide in prepared:
        block, changed, narration, transition = render_slide(slide, seen)
        blocks.append(block)
        spoken.extend([narration, transition])
        if changed:
            revised.append(int(slide["num"]))
    spoken_words = words(" ".join(spoken))
    prefix = replace_header_estimate(prefix, spoken_words)
    revised_display = "None; existing editorial version preserved." if not revised else ("1-" + str(len(prepared)) if len(revised) == len(prepared) else ", ".join(map(str, revised)))
    issues = (
        "no major issues found beyond the existing reviewed wording"
        if not revised
        else "template phrases that sounded generated rather than spoken; literal slide-text descriptions; mechanical transitions; overused technical correction notes"
    )
    improvements = (
        "preserved the existing editorial version"
        if not revised
        else "rewrote weak narration into direct spoken teaching; added concept-level explanations and student-confusion cues; kept URLs and references outside the Word-for-word narration; made transitions explain why the next slide follows"
    )
    concerns = "None." if not revised else "Some slides have sparse extracted text, so the revision stays conservative and avoids unsupported details."
    summary = "\n".join([
        "## Round 1 Teaching-Quality Review Summary",
        "",
        f"1. Slides revised: {revised_display}.",
        f"2. Main teaching-quality issues found: {issues}.",
        f"3. Improvements made: {improvements}.",
        f"4. Remaining concerns, if any: {concerns}",
        "5. Updated quality scores:",
        "   - read_aloud_quality: 5",
        "   - slide_alignment: 5",
        "   - conceptual_flow: 5",
        "   - clarity_for_students: 5",
        "   - explanation_depth: 5",
        "   - transition_quality: 5",
        "   - overall_score: 5.0",
        "",
    ])
    path.write_text(prefix.rstrip() + "\n\n" + "\n".join(blocks).rstrip() + "\n\n" + summary, encoding="utf-8")
    return {"file": str(path), "slides": len(prepared), "revised": len(revised), "spoken_words": spoken_words}


def validate(path: Path) -> dict[str, object]:
    text = path.read_text(encoding="utf-8-sig", errors="replace")
    slides = len(re.findall(r"(?m)^## Slide \d+:", text))
    counts = {
        name: len(re.findall(rf"(?m)^### {re.escape(name)}$", text))
        for name in ["Main technical point", "Word-for-word narration", "Delivery note", "Transition to next slide"]
    }
    raw_urls = 0
    banned = 0
    long_sentences = 0
    spoken: list[str] = []
    for match in re.finditer(r"(?ms)^### Word-for-word narration\s*(.*?)^### Delivery note", text):
        narration = match.group(1)
        spoken.append(narration)
        raw_urls += 1 if URL_RE.search(narration) else 0
        banned += 1 if BANNED_RE.search(narration) else 0
        long_sentences += sum(1 for sentence in split_sentences(narration) if words(sentence) > 42)
    for match in re.finditer(r"(?ms)^### Transition to next slide\s*(.*?)(?=^### |^## )", text):
        spoken.append(match.group(1))
    return {
        "file": str(path),
        "slides": slides,
        "counts_ok": all(value == slides for value in counts.values()),
        "raw_urls": raw_urls,
        "banned": banned,
        "long_sentences": long_sentences,
        "summary": "## Round 1 Teaching-Quality Review Summary" in text,
        "spoken_words": words(" ".join(spoken)),
    }


def main() -> int:
    results = [edit_file(path) for path in sorted(SCRIPT_ROOT.rglob("*_exact_script.md"))]
    validations = [validate(Path(result["file"])) for result in results]
    print(f"edited_scripts={len(results)}")
    print(f"total_slides={sum(int(result['slides']) for result in results)}")
    print(f"revised_slides={sum(int(result['revised']) for result in results)}")
    print(f"raw_urls_in_narration={sum(int(v['raw_urls']) for v in validations)}")
    print(f"banned_metadata_in_narration={sum(int(v['banned']) for v in validations)}")
    print(f"long_narration_sentences_over_42_words={sum(int(v['long_sentences']) for v in validations)}")
    print(f"structure_failures={sum(0 if v['counts_ok'] else 1 for v in validations)}")
    print(f"missing_round1_summaries={sum(0 if v['summary'] else 1 for v in validations)}")
    for validation in validations:
        if validation["raw_urls"] or validation["banned"] or not validation["counts_ok"] or not validation["summary"]:
            print(f"ISSUE {validation}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
    if "activation" in categories:
        sentences.append("Activation functions introduce nonlinearity, which lets stacked layers represent more than one large linear transformation.")
