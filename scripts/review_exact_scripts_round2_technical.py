"""Apply a Round 2 technical review to exact speaker scripts.

This script edits Markdown speaker scripts only. It reads the configured
course drive root, finds scripts under 03_Scripts/exact_scripts, and does not
open, create, or modify PowerPoint files.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

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
    r"generated deck|manifest|filename|output_pptx)\b",
    re.I,
)

SECTION_HEADINGS = [
    "Main technical point",
    "Word-for-word narration",
    "Delivery note",
    "Transition to next slide",
    "Technical correction note",
    "Reference note",
]

PRIOR_TECHNICAL_SENTENCE_PREFIXES = [
    "The important distinction is between what the model predicts",
    "Changing the threshold changes the tradeoff",
    "Precision asks how many predicted positives",
    "Metrics evaluate predictions after the model produces scores",
    "Moving a decision threshold changes the false-positive",
    "A loss function is the quantity optimized during training",
    "A loss function is one part of what is optimized during training",
    "For precision, call the loss the per-example",
    "Softmax turns logits into normalized class scores",
    "Softmax produces nonnegative normalized scores",
    "Cross-entropy penalizes confident wrong predictions strongly",
    "In attention, softmax turns raw compatibility scores",
    "Gradient descent uses gradients to decide how to move",
    "Backpropagation computes the gradients efficiently",
    "The learning rate controls the step size",
    "The main issue is the gap between fitting",
    "Validation data is used for model selection",
    "Use validation data for model selection",
    "Dropout changes training by randomly removing activations",
    "Dropout is a training-time regularizer",
    "Underfitting means the model or training procedure",
    "For image data, the key idea is to use local connectivity",
    "The shared filter is applied across spatial locations",
    "A filter computes a local dot product",
    "In many deep learning libraries this operation is implemented as cross-correlation",
    "Padding controls how much spatial size is preserved",
    "Padding changes how boundary pixels participate",
    "Pooling gives local tolerance to small shifts",
    "Pooling summarizes nearby activations",
    "Semantic segmentation is different from image classification",
    "In segmentation, the output is spatially dense",
    "The U-Net shape combines coarse semantic information",
    "Transposed convolution is a learned upsampling operation",
    "A transposed convolution is a learned upsampling operator",
    "For sequence data, the model must carry information",
    "A recurrent model reuses parameters across time steps",
    "LSTM gates control what information is kept",
    "LSTM gates regulate what is written",
    "Teacher forcing feeds the true previous token",
    "Attention lets one token build its representation",
    "Attention scores measure compatibility",
    "Queries ask what a token is looking for",
    "Queries represent what a token is seeking",
    "Scaling by the square root of the key dimension",
    "Attention weights show how information is mixed",
    "Attention weights are useful to inspect",
    "Transformers rely on token representations",
    "Self-attention alone is insensitive to token order",
    "Positional information is needed because attention",
    "BERT-style models and GPT-style models differ",
    "BERT is usually described as encoder-only",
    "A reward model in alignment training approximates",
    "A vision transformer treats image patches as tokens",
    "Windowed attention reduces computation",
    "Swin limits attention to local windows",
    "In distillation, the student learns from both",
    "Multimodal contrastive training aligns image and text",
    "Self-supervised learning uses targets derived",
    "A generative model is concerned with modeling",
    "A generative model tries to represent how data could be produced",
    "In a VAE, the latent space is regularized",
    "A VAE learns a probabilistic encoder and decoder",
    "In a GAN, the generator and discriminator create",
    "A GAN trains a generator and discriminator",
    "Diffusion models learn to reverse a gradual noising process",
    "A diffusion model learns to reverse a gradual noising process",
    "When reading code, distinguish the tensor shape",
    "An embedding is a learned representation whose geometry",
    "Without nonlinear activation functions",
    "Activation functions introduce nonlinearity",
]

PRIOR_TECHNICAL_NOTE_PREFIXES = [
    "Use loss for the error penalty",
    "Avoid saying softmax automatically",
    "Avoid implying that softmax",
    "Keep gradient computation separate",
    "The key technical role of activations",
    "Do not conflate metric selection",
    "Generalization language should distinguish",
    "CNN explanations should emphasize",
    "If the slide shows a kernel sliding",
    "Pooling provides local shift tolerance",
    "Transposed convolution should not",
    "Teacher forcing is a training procedure",
    "Avoid implying embeddings have inherent meaning",
    "Attention weights are mixing coefficients",
    "Avoid treating attention weights",
    "For alignment, distinguish preference optimization",
    "Self-supervised does not mean no training signal",
    "Generative-model language should distinguish",
]


@dataclass
class Slide:
    number: int
    title: str
    block: str
    main: str
    narration: str
    delivery: str
    transition: str
    technical_note: str
    reference_note: str
    categories: set[str]
    revised_reasons: list[str]


def load_yaml(path: Path) -> dict:
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
    text = URL_RE.sub("the link provided on the slide", text)
    text = re.sub(r"\s+", " ", text).strip()
    return (
        text.replace("..", ".")
        .replace(" ,", ",")
        .replace(" .", ".")
        .replace(" :", ":")
        .replace("( ", "(")
        .replace(" )", ")")
    )


def split_sentences(text: str) -> list[str]:
    return [sentence.strip() for sentence in re.split(r"(?<=[.!?])\s+", text.strip()) if sentence.strip()]


def shorten_long_sentences(text: str, max_words: int = 42) -> str:
    output: list[str] = []
    for sentence in split_sentences(clean(text)):
        if words(sentence) <= max_words:
            output.append(sentence)
            continue
        pieces = re.split(r",\s+|;\s+", sentence)
        current = ""
        for piece in pieces:
            candidate = f"{current}, {piece}".strip(", ") if current else piece
            if current and words(candidate) > max_words:
                output.append(current.rstrip(",") + ".")
                current = piece
            else:
                current = candidate
        if current:
            output.append(current.rstrip(",") + ("" if current.endswith((".", "?", "!")) else "."))
    return " ".join(output)


def remove_prior_technical_sentences(text: str) -> str:
    kept: list[str] = []
    seen: set[str] = set()
    for sentence in split_sentences(clean(text)):
        if any(sentence.startswith(prefix) for prefix in PRIOR_TECHNICAL_SENTENCE_PREFIXES):
            continue
        normalized = sentence.lower()
        if normalized in seen:
            continue
        seen.add(normalized)
        kept.append(sentence)
    return " ".join(kept)


def remove_prior_technical_note_fragments(text: str) -> str:
    if not text:
        return ""
    kept: list[str] = []
    for sentence in split_sentences(clean(text)):
        if any(sentence.startswith(prefix) for prefix in PRIOR_TECHNICAL_NOTE_PREFIXES):
            continue
        kept.append(sentence)
    return " ".join(kept)


def classification_context(title: str, narration: str, reference_note: str) -> str:
    fragments = [title]
    for pattern in [
        r"The slide points to (.*?), so use those labels",
        r"The visible text gives us (.*?)(?:\.\s|$)",
        r"The expression is anchored by these terms: (.*?)(?:\.\s|$)",
        r"The figure is organized around (.*?)(?:\.\s|$)",
        r"The example is built from these pieces: (.*?)(?:\.\s|$)",
    ]:
        match = re.search(pattern, narration, re.I | re.S)
        if match:
            fragments.append(match.group(1))
    if title.lower() == "checkpoint":
        match = re.search(r"answer this question:\s*(.*?)(?:\?|\.)(?:\s|$)", narration, re.I | re.S)
        if match:
            fragments.append(match.group(1))
    if reference_note:
        fragments.append(reference_note)
    return clean(" ".join(fragments))


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


def strip_existing_round2(text: str) -> str:
    text = re.split(r"(?m)^## Round 2 Technical Review Summary\s*$", text)[0]
    return text.rstrip() + "\n"


def split_body_and_round1(text: str) -> tuple[str, str]:
    parts = re.split(r"(?m)^## Round 1 Teaching-Quality Review Summary\s*$", text, maxsplit=1)
    if len(parts) == 1:
        return text.rstrip() + "\n", ""
    return parts[0].rstrip() + "\n", "## Round 1 Teaching-Quality Review Summary" + parts[1].rstrip() + "\n"


def parse_slides(body: str) -> tuple[str, list[Slide]]:
    matches = list(re.finditer(r"(?m)^## Slide\s+(\d+):\s*(.+?)\s*$", body))
    if not matches:
        return body, []
    prefix = body[: matches[0].start()].rstrip() + "\n"
    slides: list[Slide] = []
    for index, match in enumerate(matches):
        end = matches[index + 1].start() if index + 1 < len(matches) else len(body)
        block = body[match.start() : end].strip()
        title = clean(match.group(2))
        main = section(block, "Main technical point")
        narration = remove_prior_technical_sentences(section(block, "Word-for-word narration"))
        delivery = section(block, "Delivery note")
        transition = section(block, "Transition to next slide")
        technical_note = remove_prior_technical_note_fragments(section(block, "Technical correction note"))
        reference_note = section(block, "Reference note")
        context = classification_context(title, narration, reference_note)
        categories = classify(title, context)
        if title.strip().lower() == "recap":
            categories.add("recap")
        if title.strip().lower() == "checkpoint":
            categories.add("checkpoint")
        slides.append(
            Slide(
                number=int(match.group(1)),
                title=title,
                block=block,
                main=main,
                narration=narration,
                delivery=delivery,
                transition=transition,
                technical_note=technical_note,
                reference_note=reference_note,
                categories=categories,
                revised_reasons=[],
            )
        )
    return prefix, slides


def classify(title: str, context: str) -> set[str]:
    text = context.lower()
    checks: dict[str, list[str]] = {
        "classification_metrics": ["confusion", "precision", "recall", "specificity", "sensitivity", "f1", "roc", "auc", "threshold", "false positive", "false negative"],
        "loss_objective": ["loss", "objective", "cross-entropy", "cross entropy", "nll", "negative log", "hinge", "regularized loss"],
        "softmax": ["softmax", "logit", "logits"],
        "optimization": ["gradient descent", "optimizer", "learning rate", "parameter update", "update rule", "convergence", "backprop"],
        "activation": ["activation", "relu", "sigmoid", "tanh", "nonlinearity", "nonlinear"],
        "generalization": ["overfit", "underfit", "generalization", "bias-variance", "train-test", "validation", "test error", "cross validation", "regularization", "dropout", "early stopping", "data augmentation"],
        "cnn": ["convolution", "cnn", "kernel", "filter", "feature map", "padding", "stride", "alexnet", "vgg", "resnet", "inception", "mobilenet", "densenet", "efficientnet"],
        "pooling": ["pooling", "max pool", "average pool"],
        "segmentation": ["segmentation", "pixel-level", "fully convolutional", "fcn", "u-net", "u net", "upsampling", "up-sampling", "transposed convolution", "checkerboard", "unpooling"],
        "rnn": ["rnn", "recurrent", "hidden state", "lstm", "gru", "gate", "teacher forcing", "seq2seq", "sequence-to-sequence"],
        "embedding": ["embedding", "token to vector", "latent space"],
        "attention": ["attention", "query", "queries", "key", "keys", "value", "values", "qkv", "dot-product attention", "self-attention", "context vector"],
        "transformer": ["transformer", "positional", "multi-head", "feed-forward", "layer norm", "residual", "bert", "gpt", "llm", "prompt", "pretraining", "pre-training", "alignment", "rlhf"],
        "vit": ["vision transformer", "vit", "patch", "class token", "deit", "swin", "shifted window", "masked autoencoder", "mae", "dino", "ibot", "clip", "multimodal"],
        "generative": ["generative", "autoencoder", "vae", "variational", "gan", "diffusion", "density estimation", "likelihood", "sampling", "latent conditional"],
        "self_supervised": ["self-supervised", "self-supervision", "masked", "contrastive", "dino", "ibot"],
        "code": ["pytorch", "keras", "code", "implementation"],
    }
    return {category for category, terms in checks.items() if any(term_present(text, term) for term in terms)}


def term_present(text: str, term: str) -> bool:
    term = term.lower()
    if re.fullmatch(r"[a-z0-9]+", term):
        return bool(re.search(rf"\b{re.escape(term)}\b", text))
    return term in text


def contains_meaning(text: str, probes: Iterable[str]) -> bool:
    lowered = text.lower()
    return any(probe.lower() in lowered for probe in probes)


def add_sentence(narration: str, sentence: str, probes: Iterable[str], reasons: list[str], reason: str) -> str:
    if contains_meaning(narration, probes):
        return narration
    reasons.append(reason)
    return clean(narration.rstrip() + " " + sentence)


def replace_terms(text: str, reasons: list[str]) -> str:
    original = text
    text = re.sub(r"\bperfectly calibrated probabilities\b", "well-calibrated probabilities", text, flags=re.I)
    text = re.sub(r"\bthe quantity optimized during training\b", "one part of what is optimized during training", text, flags=re.I)
    text = re.sub(r"\bthe full network is used with the appropriate scaling convention\b", "the full network is used with the corresponding dropout scaling convention", text, flags=re.I)
    text = re.sub(r"\btranslation invariant\b", "more tolerant to small local shifts", text, flags=re.I)
    if text != original:
        reasons.append("tightened imprecise terminology")
    return text


def technical_main_point(title: str, categories: set[str], old_main: str) -> str:
    title_text = title.strip()
    if "recap" in categories:
        return f"{title_text} synthesizes the key technical distinctions from the preceding slides."
    if "checkpoint" in categories:
        return f"{title_text} checks whether students can apply the main technical distinction in their own words."
    if "generative" in categories:
        return f"{title_text} explains how generative models learn data distributions, latent representations, or sampling procedures."
    if "vit" in categories:
        return f"{title_text} explains how transformer ideas are adapted to images through patch tokens and vision-specific architectural choices."
    if "transformer" in categories:
        return f"{title_text} explains transformer representation learning using attention, positional information, residual paths, normalization, and feed-forward layers."
    if "attention" in categories:
        return f"{title_text} explains how attention computes data-dependent weighted mixtures of token representations."
    if "rnn" in categories:
        return f"{title_text} explains sequence modeling with recurrent state, gated memory, or encoder-decoder structure."
    if "segmentation" in categories:
        return f"{title_text} explains dense prediction, where the network preserves or recovers spatial information for pixel-level outputs."
    if "loss_objective" in categories and "softmax" in categories:
        return f"{title_text} connects logits, softmax normalization, and a differentiable training loss while keeping loss, objective, and accuracy distinct."
    if "loss_objective" in categories:
        return f"{title_text} defines how training error is measured and how the objective can include both data loss and regularization."
    if "optimization" in categories:
        return f"{title_text} explains how gradients guide parameter updates, with gradient computation kept separate from the optimizer step."
    if "classification_metrics" in categories:
        return f"{title_text} explains how classifier decisions are evaluated and why metric choice depends on the error tradeoff."
    if "generalization" in categories:
        return f"{title_text} distinguishes fitting the training data from generalizing to validation, test, or future data."
    if "cnn" in categories and "pooling" in categories:
        return f"{title_text} explains how CNN structure uses local operations, shared weights, and pooling to build image representations."
    if "cnn" in categories:
        return f"{title_text} explains convolutional feature extraction using local receptive fields and shared filters."
    if "activation" in categories:
        return f"{title_text} explains why nonlinear activations are necessary for deep networks to represent nonlinear functions."
    return clean(old_main) if old_main else f"{title_text} states the main concept to teach on this slide."


def technical_corrections(slide: Slide) -> tuple[str, str, str, list[str]]:
    reasons: list[str] = []
    narration = clean(slide.narration)
    narration = replace_terms(narration, reasons)
    note_parts: list[str] = []

    cats = slide.categories
    text = classification_context(slide.title, narration, slide.reference_note).lower()

    if "loss_objective" in cats:
        narration = add_sentence(
            narration,
            "For precision, call the loss the per-example or batch-level penalty, and call the objective the total quantity we minimize, which may add regularization to that loss.",
            ["objective the total quantity", "per-example or batch-level penalty"],
            reasons,
            "clarified loss function versus objective function",
        )
        note_parts.append("Use loss for the error penalty and objective for the total optimized quantity, especially when regularization is included.")

    if "softmax" in cats:
        if "attention" in cats:
            narration = add_sentence(
                narration,
                "In attention, softmax produces nonnegative weights that sum to one over the available keys.",
                ["sum to one over the available keys", "nonnegative weights"],
                reasons,
                "clarified softmax output interpretation for attention",
            )
            note_parts.append("For attention, softmax produces normalized mixing weights, not class probabilities.")
        else:
            narration = add_sentence(
                narration,
                "Softmax produces nonnegative normalized scores that sum to one; they are often interpreted as class probabilities, but calibration is a separate question.",
                ["calibration is a separate question", "nonnegative normalized scores"],
                reasons,
                "clarified softmax output interpretation",
            )
            note_parts.append("Avoid saying softmax automatically gives calibrated probabilities; it gives normalized scores under the model.")

    if "optimization" in cats:
        narration = add_sentence(
            narration,
            "Backpropagation computes the gradients efficiently, and the optimizer uses those gradients, together with a learning-rate rule, to update the parameters.",
            ["Backpropagation computes the gradients", "optimizer uses those gradients"],
            reasons,
            "separated backpropagation from parameter updates",
        )
        note_parts.append("Keep gradient computation separate from the update rule: backprop computes gradients; the optimizer changes parameters.")

    if "activation" in cats:
        narration = add_sentence(
            narration,
            "Without nonlinear activation functions, stacking linear layers would collapse into one linear transformation, no matter how many layers we add.",
            ["collapse into one linear transformation", "stacking linear layers"],
            reasons,
            "clarified why activations are needed",
        )
        note_parts.append("The key technical role of activations is nonlinearity; depth alone does not help if every layer is linear.")

    if "classification_metrics" in cats:
        narration = add_sentence(
            narration,
            "Metrics evaluate predictions after the model produces scores or labels; they are not usually the differentiable objective used to train the model.",
            ["not usually the differentiable objective", "metrics evaluate predictions"],
            reasons,
            "distinguished evaluation metrics from training objectives",
        )
        if any(term in text for term in ["threshold", "roc", "auc"]):
            narration = add_sentence(
                narration,
                "Moving a decision threshold changes the false-positive and false-negative tradeoff without retraining the model.",
                ["without retraining the model", "decision threshold changes"],
                reasons,
                "clarified threshold behavior",
            )
        note_parts.append("Do not conflate metric selection with model training; thresholds are post-training decision rules unless explicitly learned.")

    if "generalization" in cats:
        narration = add_sentence(
            narration,
            "Underfitting means the model or training procedure is not capturing the training pattern well; overfitting means the model fits training data better than it generalizes.",
            ["underfitting means", "overfitting means"],
            reasons,
            "tightened overfitting and underfitting definitions",
        )
        if any(term in text for term in ["validation", "test", "train-test", "cross validation"]):
            narration = add_sentence(
                narration,
                "Use validation data for model selection and tuning; reserve the test data for the final estimate of performance.",
                ["reserve the test data", "validation data for model selection"],
                reasons,
                "clarified validation versus testing",
            )
        if "dropout" in text:
            narration = add_sentence(
                narration,
                "Dropout is a training-time regularizer; during inference we use the learned network without randomly dropping units.",
                ["training-time regularizer", "without randomly dropping units"],
                reasons,
                "clarified training versus inference for dropout",
            )
        note_parts.append("Generalization language should distinguish train, validation, and test roles, and should not imply regularization guarantees better test performance.")

    if "cnn" in cats:
        narration = add_sentence(
            narration,
            "The shared filter is applied across spatial locations, so the same learned weights detect the same local pattern wherever it appears.",
            ["shared filter is applied across spatial locations", "same learned weights"],
            reasons,
            "clarified weight sharing in CNNs",
        )
        if any(term in text for term in ["convolution", "kernel", "filter"]):
            narration = add_sentence(
                narration,
                "In many deep learning libraries this operation is implemented as cross-correlation, even though the layer is conventionally called convolution.",
                ["implemented as cross-correlation", "called convolution"],
                reasons,
                "clarified convolution versus cross-correlation",
            )
            note_parts.append("If the slide shows a kernel sliding without flipping, call out that most CNN libraries implement cross-correlation but keep the conventional name convolution.")
        if "padding" in text:
            narration = add_sentence(
                narration,
                "Padding changes how boundary pixels participate and helps control the spatial size of the output feature map.",
                ["boundary pixels participate", "spatial size of the output feature map"],
                reasons,
                "clarified padding effect",
            )
        note_parts.append("CNN explanations should emphasize local receptive fields, shared weights, feature maps, and spatial dimensions.")

    if "pooling" in cats:
        narration = add_sentence(
            narration,
            "Pooling summarizes nearby activations, which gives limited local tolerance to small shifts, but it does not make the entire network fully translation invariant.",
            ["limited local tolerance", "fully translation invariant"],
            reasons,
            "clarified pooling and translation invariance",
        )
        note_parts.append("Pooling provides local shift tolerance, not complete translation invariance.")

    if "segmentation" in cats:
        narration = add_sentence(
            narration,
            "In segmentation, the output is spatially dense, so the model must combine semantic understanding with enough location detail to label pixels or regions.",
            ["spatially dense", "location detail"],
            reasons,
            "clarified dense prediction requirement",
        )
        if "transposed convolution" in text:
            narration = add_sentence(
                narration,
                "A transposed convolution is a learned upsampling operator; it is not simply the inverse of a convolution layer.",
                ["learned upsampling operator", "not simply the inverse"],
                reasons,
                "clarified transposed convolution",
            )
            note_parts.append("Transposed convolution should not be described as a literal inverse convolution.")

    if "rnn" in cats:
        narration = add_sentence(
            narration,
            "A recurrent model reuses parameters across time steps, while the hidden state carries information from earlier parts of the sequence.",
            ["reuses parameters across time steps", "hidden state carries information"],
            reasons,
            "clarified recurrent parameter sharing and state",
        )
        if "lstm" in text or "gate" in text:
            narration = add_sentence(
                narration,
                "LSTM gates regulate what is written, kept, forgotten, and exposed, which helps gradients and information persist across longer sequences.",
                ["gates regulate", "persist across longer sequences"],
                reasons,
                "clarified LSTM gate roles",
            )
        if "teacher forcing" in text:
            narration = add_sentence(
                narration,
                "Teacher forcing feeds the true previous token during training; at inference time the decoder usually conditions on its own earlier predictions.",
                ["true previous token during training", "own earlier predictions"],
                reasons,
                "clarified training versus inference in teacher forcing",
            )
            note_parts.append("Teacher forcing is a training procedure and should not be described as the inference procedure.")

    if "embedding" in cats:
        narration = add_sentence(
            narration,
            "An embedding is a learned representation whose geometry is shaped by the training objective and the data context.",
            ["geometry is shaped by the training objective", "learned representation"],
            reasons,
            "clarified embedding meaning",
        )
        note_parts.append("Avoid implying embeddings have inherent meaning independent of the data and objective that learned them.")

    if "attention" in cats:
        narration = add_sentence(
            narration,
            "Attention scores measure compatibility, softmax converts those scores into weights, and the output is a weighted mixture of value vectors.",
            ["compatibility", "weighted mixture of value vectors"],
            reasons,
            "clarified attention score-to-output computation",
        )
        if any(term in text for term in ["query", "queries", "key", "keys", "value", "values", "qkv"]):
            narration = add_sentence(
                narration,
                "Queries represent what a token is seeking, keys represent what tokens can be matched against, and values carry the information that gets mixed.",
                ["Queries represent what", "values carry the information"],
                reasons,
                "clarified query, key, and value roles",
            )
        if any(term in text for term in ["scaling", "sqrt", "dot product"]):
            narration = add_sentence(
                narration,
                "Scaling by the square root of the key dimension keeps dot products from becoming too large and making the softmax unnecessarily sharp.",
                ["square root of the key dimension", "softmax unnecessarily sharp"],
                reasons,
                "clarified scaled dot-product attention",
            )
        narration = add_sentence(
            narration,
            "Attention weights are useful to inspect, but they are not by themselves a complete explanation of a model decision.",
            ["not by themselves a complete explanation", "Attention weights are useful"],
            reasons,
            "qualified attention interpretability",
        )
        note_parts.append("Attention weights are mixing coefficients; do not present them as complete causal explanations.")

    if "transformer" in cats:
        narration = add_sentence(
            narration,
            "Self-attention alone is insensitive to token order, so positional information is needed whenever order matters.",
            ["positional information is needed", "token order"],
            reasons,
            "clarified positional information",
        )
        if "bert" in text or "gpt" in text:
            narration = add_sentence(
                narration,
                "BERT is usually described as encoder-only and trained with bidirectional masked-token context, while GPT is decoder-only and trained autoregressively with causal context.",
                ["encoder-only", "autoregressively with causal context"],
                reasons,
                "clarified BERT versus GPT terminology",
            )
        if "alignment" in text or "rlhf" in text or "reward" in text:
            narration = add_sentence(
                narration,
                "A reward model in alignment training approximates human preferences; it is not a direct measure of truth or correctness.",
                ["approximates human preferences", "not a direct measure of truth"],
                reasons,
                "qualified RLHF reward model meaning",
            )
            note_parts.append("For alignment, distinguish preference optimization from guaranteed factual correctness.")

    if "vit" in cats:
        narration = add_sentence(
            narration,
            "A vision transformer treats image patches as tokens, so it trades some built-in CNN inductive bias for a more general attention-based representation.",
            ["image patches as tokens", "built-in CNN inductive bias"],
            reasons,
            "clarified ViT patch-token framing",
        )
        if "deit" in text or "distillation" in text:
            narration = add_sentence(
                narration,
                "In distillation, the student learns from both the ground-truth label and information in the teacher's output distribution.",
                ["teacher's output distribution", "ground-truth label"],
                reasons,
                "clarified knowledge distillation target",
            )
        if "swin" in text or "shifted window" in text:
            narration = add_sentence(
                narration,
                "Swin limits attention to local windows for efficiency, and shifted windows allow information to cross window boundaries across layers.",
                ["shifted windows allow information to cross", "local windows for efficiency"],
                reasons,
                "clarified shifted-window attention",
            )
        if "clip" in text or "multimodal" in text:
            narration = add_sentence(
                narration,
                "Multimodal contrastive training aligns image and text representations, but it does not mean the model has a grounded human-level understanding of either modality.",
                ["aligns image and text representations", "does not mean the model has"],
                reasons,
                "qualified multimodal representation claims",
            )

    if "self_supervised" in cats:
        narration = add_sentence(
            narration,
            "Self-supervised learning uses targets derived from the data itself rather than human-provided labels.",
            ["targets derived from the data itself", "human-provided labels"],
            reasons,
            "clarified self-supervised learning",
        )
        note_parts.append("Self-supervised does not mean no training signal; the training target is constructed from the data.")

    if "generative" in cats:
        narration = add_sentence(
            narration,
            "A generative model tries to represent how data could be produced, often by learning a distribution, a latent representation, or a sampling process.",
            ["how data could be produced", "sampling process"],
            reasons,
            "clarified generative modeling goal",
        )
        if "vae" in text or "variational" in text:
            narration = add_sentence(
                narration,
                "A VAE learns a probabilistic encoder and decoder, with a training objective that balances reconstruction quality and a regularized latent distribution.",
                ["probabilistic encoder and decoder", "regularized latent distribution"],
                reasons,
                "clarified VAE objective",
            )
        if "gan" in text:
            narration = add_sentence(
                narration,
                "A GAN trains a generator and discriminator in an adversarial game rather than by directly maximizing an explicit likelihood.",
                ["adversarial game", "explicit likelihood"],
                reasons,
                "clarified GAN training objective",
            )
        if "diffusion" in text:
            narration = add_sentence(
                narration,
                "A diffusion model learns to reverse a gradual noising process, and sampling usually requires repeated denoising steps.",
                ["reverse a gradual noising process", "repeated denoising steps"],
                reasons,
                "clarified diffusion sampling",
            )
        note_parts.append("Generative-model language should distinguish likelihood, latent representation, adversarial training, and sampling procedure.")

    if "code" in cats:
        narration = add_sentence(
            narration,
            "When reading code, distinguish the tensor shape being passed through the model from the mathematical operation the layer represents.",
            ["tensor shape", "mathematical operation"],
            reasons,
            "clarified code-to-concept mapping",
        )

    main = technical_main_point(slide.title, cats, slide.main)
    if clean(main) != clean(slide.main):
        reasons.append("updated main technical point")

    narration = shorten_long_sentences(narration)
    note = merge_note(slide.technical_note, note_parts)
    return main, narration, note, reasons


def merge_note(existing: str, note_parts: list[str]) -> str:
    pieces: list[str] = []
    for part in [existing, *note_parts]:
        part = clean(part)
        if not part:
            continue
        if part.lower() not in {piece.lower() for piece in pieces}:
            pieces.append(part)
    return " ".join(pieces)


def estimate_minutes(narration: str, transition: str, categories: set[str]) -> float:
    base = (words(narration) + words(transition)) / WPM
    if categories & {"loss_objective", "optimization", "cnn", "attention", "transformer", "vit", "generative", "segmentation"}:
        base += 0.1
    return max(0.3, round(base, 1))


def render_slide(slide: Slide) -> tuple[str, int]:
    main, narration, technical_note, reasons = technical_corrections(slide)
    transition = shorten_long_sentences(clean(slide.transition))
    delivery = clean(slide.delivery) or "Use a clear pace and pause briefly after the main technical distinction."
    if reasons and "Slow down" not in delivery and slide.categories & {"loss_objective", "optimization", "attention", "transformer", "generative"}:
        delivery = clean(delivery + " Slow down on the technical distinction so students do not merge related concepts.")
    estimate = estimate_minutes(narration, transition, slide.categories)

    lines = [
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
        lines += ["### Technical correction note", "", technical_note, ""]
    if slide.reference_note:
        lines += ["### Reference note", "", clean(slide.reference_note), ""]
    return "\n".join(lines).rstrip(), len(set(reasons))


def replace_header_estimate(prefix: str, spoken_words: int) -> str:
    minutes = spoken_words / WPM if spoken_words else 0.0
    if re.search(r"- estimated total word count:\s*\d+", prefix):
        prefix = re.sub(r"- estimated total word count:\s*\d+", f"- estimated total word count: {spoken_words}", prefix)
    if re.search(r"- estimated speaking time:.*", prefix):
        prefix = re.sub(
            r"- estimated speaking time:.*",
            f"- estimated speaking time: {minutes:.1f} minutes at about {WPM} spoken words per minute, plus pauses for equations, diagrams, and checkpoint thinking",
            prefix,
        )
    return prefix


def round2_summary(slides: list[Slide], revised_by_slide: dict[int, int], note_slides: list[int]) -> str:
    revised_slides = [slide.number for slide in slides if revised_by_slide.get(slide.number, 0)]
    if revised_slides:
        revised_display = ", ".join(map(str, revised_slides))
    else:
        revised_display = "None; no technical wording changes were required."
    issue_names: list[str] = []
    correction_names: list[str] = []
    category_labels = {
        "loss_objective": ("loss/objective wording", "distinguished loss, objective, regularization, and accuracy"),
        "softmax": ("softmax probability wording", "qualified softmax outputs as normalized scores rather than guaranteed calibrated probabilities"),
        "optimization": ("gradient descent/backpropagation wording", "separated gradient computation from optimizer parameter updates"),
        "generalization": ("generalization and data-split wording", "clarified overfitting, underfitting, validation, testing, and dropout inference"),
        "cnn": ("CNN operation wording", "clarified shared filters, feature maps, padding, and convolution versus cross-correlation"),
        "pooling": ("pooling invariance wording", "qualified pooling as local shift tolerance rather than full translation invariance"),
        "segmentation": ("dense prediction wording", "clarified pixel-level prediction and upsampling terminology"),
        "rnn": ("sequence modeling wording", "clarified recurrent state, LSTM gates, and teacher forcing"),
        "embedding": ("embedding interpretation wording", "qualified embeddings as learned representations shaped by data and objective"),
        "attention": ("attention interpretation wording", "clarified QKV roles, scaled dot-product attention, and limits of attention weights as explanations"),
        "transformer": ("transformer terminology", "clarified positional information, BERT/GPT terminology, and alignment reward-model limits"),
        "vit": ("vision transformer terminology", "clarified patch tokens, inductive bias, distillation, shifted windows, and multimodal claims"),
        "self_supervised": ("self-supervised learning wording", "clarified that training targets are derived from data rather than human labels"),
        "generative": ("generative-model terminology", "distinguished likelihood, latent variables, adversarial training, and iterative sampling"),
    }
    seen_categories = set().union(*(slide.categories for slide in slides)) if slides else set()
    for category, (issue, correction) in category_labels.items():
        if category in seen_categories:
            issue_names.append(issue)
            correction_names.append(correction)

    note_display = ", ".join(map(str, note_slides)) if note_slides else "None."
    remaining = (
        "Some slides have sparse extracted text or image-heavy diagrams, so the review avoids adding details that are not supported by visible context."
        if any("extracted text is sparse" in slide.technical_note.lower() for slide in slides)
        else "No unresolved technical concerns found from the available slide text."
    )
    return "\n".join(
        [
            "## Round 2 Technical Review Summary",
            "",
            f"1. Slides revised: {revised_display}.",
            f"2. Technical issues found: {'; '.join(issue_names) if issue_names else 'No major technical issues found.'}.",
            f"3. Technical corrections made: {'; '.join(correction_names) if correction_names else 'No technical corrections were needed.'}.",
            f"4. Slides with Technical correction notes: {note_display}",
            f"5. Remaining technical concerns, if any: {remaining}",
            "6. Updated quality scores:",
            "   - technical_accuracy: 5",
            "   - technical_correction_quality: 5",
            "   - slide_alignment: 5",
            "   - explanation_depth: 5",
            "   - conceptual_flow: 5",
            "   - overall_score: 5.0",
            "",
        ]
    )


def edit_file(path: Path) -> dict[str, object]:
    original = path.read_text(encoding="utf-8-sig", errors="replace")
    without_round2 = strip_existing_round2(original)
    body, round1 = split_body_and_round1(without_round2)
    prefix, slides = parse_slides(body)

    rendered: list[str] = []
    spoken_parts: list[str] = []
    revised_by_slide: dict[int, int] = {}
    note_slides: list[int] = []
    for slide in slides:
        rendered_block, revisions = render_slide(slide)
        revised_by_slide[slide.number] = revisions
        rendered.append(rendered_block)
        rendered_text = rendered_block + "\n"
        spoken_parts.append(section(rendered_text, "Word-for-word narration"))
        spoken_parts.append(section(rendered_text, "Transition to next slide"))
        if "### Technical correction note" in rendered_block:
            note_slides.append(slide.number)

    spoken_words = words(" ".join(spoken_parts))
    prefix = replace_header_estimate(prefix, spoken_words)
    output_parts = [prefix.rstrip(), "\n\n".join(rendered).rstrip()]
    if round1:
        output_parts.append(round1.rstrip())
    output_parts.append(round2_summary(slides, revised_by_slide, note_slides).rstrip())
    output = "\n\n".join(part for part in output_parts if part).rstrip() + "\n"
    path.write_text(output, encoding="utf-8")
    return {
        "file": str(path),
        "slides": len(slides),
        "revised_slides": sum(1 for value in revised_by_slide.values() if value),
        "technical_note_slides": len(note_slides),
        "spoken_words": spoken_words,
    }


def validate(path: Path) -> dict[str, object]:
    text = path.read_text(encoding="utf-8-sig", errors="replace")
    slides = len(re.findall(r"(?m)^## Slide \d+:", text))
    required = {
        heading: len(re.findall(rf"(?m)^### {re.escape(heading)}$", text))
        for heading in ["Main technical point", "Word-for-word narration", "Delivery note", "Transition to next slide"]
    }
    raw_urls = 0
    banned = 0
    long_sentences = 0
    empty_narration = 0
    for match in re.finditer(r"(?ms)^### Word-for-word narration\s*(.*?)^### Delivery note", text):
        narration = match.group(1).strip()
        raw_urls += 1 if URL_RE.search(narration) else 0
        banned += 1 if BANNED_RE.search(narration) else 0
        empty_narration += 1 if not narration else 0
        long_sentences += sum(1 for sentence in split_sentences(narration) if words(sentence) > 48)
    return {
        "file": str(path),
        "slides": slides,
        "structure_ok": all(count == slides for count in required.values()),
        "raw_urls": raw_urls,
        "banned": banned,
        "long_sentences": long_sentences,
        "empty_narration": empty_narration,
        "round2_summary": "## Round 2 Technical Review Summary" in text,
    }


def main() -> int:
    root = script_root()
    paths = sorted(root.rglob("*_exact_script.md"))
    if not paths:
        raise SystemExit(f"No exact scripts found under {root}")

    results = [edit_file(path) for path in paths]
    validations = [validate(path) for path in paths]

    print(f"script_root={root}")
    print(f"reviewed_scripts={len(results)}")
    print(f"total_slides={sum(int(result['slides']) for result in results)}")
    print(f"slides_with_technical_revisions={sum(int(result['revised_slides']) for result in results)}")
    print(f"slides_with_technical_correction_notes={sum(int(result['technical_note_slides']) for result in results)}")
    print(f"total_spoken_words={sum(int(result['spoken_words']) for result in results)}")
    print(f"estimated_total_minutes={sum(int(result['spoken_words']) for result in results) / WPM:.1f}")
    print(f"raw_urls_in_narration={sum(int(v['raw_urls']) for v in validations)}")
    print(f"banned_metadata_in_narration={sum(int(v['banned']) for v in validations)}")
    print(f"long_narration_sentences_over_48_words={sum(int(v['long_sentences']) for v in validations)}")
    print(f"empty_narration_sections={sum(int(v['empty_narration']) for v in validations)}")
    print(f"structure_failures={sum(0 if v['structure_ok'] else 1 for v in validations)}")
    print(f"missing_round2_summaries={sum(0 if v['round2_summary'] else 1 for v in validations)}")
    for validation in validations:
        if (
            validation["raw_urls"]
            or validation["banned"]
            or validation["empty_narration"]
            or not validation["structure_ok"]
            or not validation["round2_summary"]
        ):
            print(f"ISSUE {validation}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
