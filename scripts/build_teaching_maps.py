"""Build concept-aware teaching maps from extracted slide text CSVs."""

from __future__ import annotations

import csv
import re
import sys
from dataclasses import dataclass
from pathlib import Path


INPUT_DIR = Path("manifests") / "slide_text"
OUTPUT_DIR = Path("manifests") / "teaching_maps"
CSV_FIELDS = [
    "lecture_number",
    "lecture_title",
    "slide_number",
    "slide_title",
    "core_concept",
    "slide_role",
    "concept_group",
    "prerequisite_concept",
    "next_concept",
    "estimated_teaching_minutes",
    "logical_boundary_after_slide",
    "boundary_strength",
    "boundary_reason",
    "notes_for_future_script",
]


@dataclass(frozen=True)
class ConceptGroup:
    start: int
    end: int
    name: str
    prerequisite: str
    next_concept: str
    boundary_strength: str = "medium"


LECTURE_GROUPS: dict[int, list[ConceptGroup]] = {
    1: [
        ConceptGroup(1, 1, "CNN lecture framing", "Deep learning basics", "MLP refresher"),
        ConceptGroup(2, 7, "MLP and neural network refresher", "Neural network function approximation", "Image data motivation", "strong"),
        ConceptGroup(8, 13, "Image datasets and CNN motivation", "MLP limitations for images", "Convolution operation", "strong"),
        ConceptGroup(14, 21, "Convolution fundamentals", "Image tensors and local receptive fields", "Classical image filters", "strong"),
        ConceptGroup(22, 33, "Classical filters padding and kernel mechanics", "Convolution local dot product", "Multichannel convolution", "strong"),
        ConceptGroup(34, 41, "Multichannel and multilayer convolution", "Single-channel convolution", "CNN inductive bias", "strong"),
        ConceptGroup(42, 45, "CNN inductive bias and hierarchy", "Stacked convolutional feature maps", "Pooling", "medium"),
        ConceptGroup(46, 48, "Pooling and spatial reduction", "Convolutional feature maps", "AlexNet architecture", "strong"),
        ConceptGroup(49, 51, "AlexNet architecture and implementation", "Convolution and pooling layers", "CNN summary", "strong"),
        ConceptGroup(52, 52, "CNN summary", "CNN architecture components", "Next lecture on evaluation metrics", "strong"),
    ],
    2: [
        ConceptGroup(1, 2, "Evaluation motivation", "Classifier predictions", "Binary decision framing", "strong"),
        ConceptGroup(3, 6, "Binary decisions and confusion matrix", "Classification scores", "Basic metrics", "strong"),
        ConceptGroup(7, 14, "Accuracy precision recall specificity and F1", "Confusion matrix counts", "Metric tradeoff discussions", "strong"),
        ConceptGroup(15, 21, "Threshold and metric tradeoff discussions", "Precision recall specificity and F1", "ROC and AUC", "strong"),
        ConceptGroup(22, 27, "ROC curves and AUC", "Threshold-dependent metrics", "AUC interpretation exercises", "strong"),
        ConceptGroup(28, 30, "AUC interpretation exercises", "ROC curve geometry", "Metric selection summary", "medium"),
        ConceptGroup(31, 32, "Metric selection summary", "Evaluation metric tradeoffs", "Next lecture on loss functions", "strong"),
    ],
    3: [
        ConceptGroup(1, 4, "Score functions loss functions and activations", "Classifier evaluation", "Gradient descent optimization", "strong"),
        ConceptGroup(5, 16, "Gradient descent optimization", "Loss surface and differentiability", "Linear classifier worked example", "strong"),
        ConceptGroup(17, 23, "Linear classifier worked example", "Gradient-based training", "Softmax conversion", "strong"),
        ConceptGroup(24, 29, "Softmax derivation and numerical stability", "Linear classifier scores", "Cross entropy loss", "strong"),
        ConceptGroup(30, 37, "Entropy cross entropy and loss comparison", "Softmax probabilities", "Lecture summary", "strong"),
        ConceptGroup(38, 38, "Loss function summary", "Softmax and cross entropy", "Next lecture on generalization", "strong"),
    ],
    4: [
        ConceptGroup(1, 7, "Underfitting overfitting and bias variance", "Training and test error", "Generalization splits", "strong"),
        ConceptGroup(8, 12, "Generalization splits and cross validation", "Bias variance tradeoff", "Regularization objective", "strong"),
        ConceptGroup(13, 18, "Regularized loss norms and least squares", "Validation-based model selection", "L1 and L2 geometry", "strong"),
        ConceptGroup(19, 27, "L1 L2 geometry shrinkage and sparsity", "Norm definitions and regularized loss", "Dropout regularization", "strong"),
        ConceptGroup(28, 31, "Dropout regularization", "Overfitting prevention", "Data augmentation", "medium"),
        ConceptGroup(32, 33, "Data augmentation", "Training data variability", "Optimization controls", "medium"),
        ConceptGroup(34, 37, "Early stopping and learning-rate control", "Validation monitoring", "Training recipe", "strong"),
        ConceptGroup(38, 39, "Generalization recipe and summary", "Regularization and optimization controls", "Next lecture on modern CNNs", "strong"),
    ],
    5: [
        ConceptGroup(1, 4, "Modern CNN evolution overview", "Convolutional network basics", "Foundation CNN architectures", "strong"),
        ConceptGroup(5, 14, "LeNet AlexNet GAP and early CNN practice", "Convolution pooling and classification heads", "Batch normalization and depth", "strong"),
        ConceptGroup(15, 16, "Batch normalization transition", "Early CNN design", "Depth and VGG", "medium"),
        ConceptGroup(17, 26, "Depth width hierarchy and VGG", "CNN feature hierarchy", "Deep optimization problems", "strong"),
        ConceptGroup(27, 33, "Degradation vanishing gradients and optimization limits", "Very deep CNNs", "Residual learning", "strong"),
        ConceptGroup(34, 40, "ResNet and residual learning", "Deep optimization problem", "Multi-scale efficient CNN design", "strong"),
        ConceptGroup(41, 50, "Inception and multi-scale efficient convolution", "Residual design principles", "Mobile and dense connectivity", "strong"),
        ConceptGroup(51, 58, "MobileNet and DenseNet efficiency patterns", "Efficient convolution and feature reuse", "Compound scaling", "strong"),
        ConceptGroup(59, 67, "EfficientNet scaling MBConv SE and modern principles", "Efficient architecture blocks", "CNN evolution summary", "strong"),
        ConceptGroup(68, 69, "CNN evolution and objective-function wrap-up", "Modern CNN design principles", "Next lecture on segmentation", "strong"),
    ],
    6: [
        ConceptGroup(1, 2, "Semantic segmentation task definition", "Image classification", "Fully convolutional motivation", "strong"),
        ConceptGroup(3, 6, "Why sliding windows fail and FCN motivation", "Pixel-level prediction goal", "Fully convolutional networks", "strong"),
        ConceptGroup(7, 12, "Fully convolutional network conversion", "Classification CNNs", "Encoder-decoder segmentation", "strong"),
        ConceptGroup(13, 16, "Downsampling U-Net and encoder-decoder structure", "FCN dense prediction", "Upsampling methods", "strong"),
        ConceptGroup(17, 24, "Upsampling methods and interpolation", "Downsampled feature maps", "Transposed convolution", "strong"),
        ConceptGroup(25, 31, "Transposed convolution spatial examples", "Upsampling by interpolation", "Matrix view of transposed convolution", "strong"),
        ConceptGroup(32, 39, "Matrix view and stride-2 transposed convolution derivation", "Convolution as a matrix", "Artifacts and summary", "strong"),
        ConceptGroup(40, 41, "Transposed convolution limitations and summary", "Transpose convolution mechanics", "Next lecture on sequences", "strong"),
    ],
    7: [
        ConceptGroup(1, 5, "Sequence modeling task taxonomy", "Neural network mappings", "RNN recurrence", "strong"),
        ConceptGroup(6, 8, "RNN recurrence and long-term dependency problem", "Sequence inputs and outputs", "LSTM gates", "strong"),
        ConceptGroup(9, 17, "LSTM cell gates state update and parameter count", "RNN hidden state", "LSTM implementation", "strong"),
        ConceptGroup(18, 23, "LSTM affine form and PyTorch workflow", "LSTM gate equations", "Sequence model variants", "strong"),
        ConceptGroup(24, 28, "Sequence model variants and bidirectionality", "LSTM implementation", "Seq2Seq response generation", "strong"),
        ConceptGroup(29, 38, "Seq2Seq response generation worked example", "Encoder-decoder sequence models", "Image captioning example", "strong"),
        ConceptGroup(39, 46, "Image captioning worked example", "Seq2Seq generation", "LSTM summary", "strong"),
        ConceptGroup(47, 47, "LSTM summary", "RNN LSTM and Seq2Seq models", "Next lecture on transformers", "strong"),
    ],
    8: [
        ConceptGroup(1, 2, "Transformer lecture roadmap", "Sequence modeling with LSTMs", "Seq2Seq bottleneck", "strong"),
        ConceptGroup(3, 10, "Seq2Seq mechanics and bottleneck", "Encoder-decoder RNNs", "Attention over encoder states", "strong"),
        ConceptGroup(11, 18, "Attention for Seq2Seq alignment", "Seq2Seq bottleneck", "Self-attention", "strong"),
        ConceptGroup(19, 31, "Self-attention computation and QKV roles", "Attention weights and context vectors", "Positional encoding", "strong"),
        ConceptGroup(32, 40, "Attention visualization and positional encoding", "Self-attention without recurrence", "Multi-head attention", "strong"),
        ConceptGroup(41, 45, "Multi-head and cross-attention", "Single-head self-attention", "Transformer summary", "strong"),
        ConceptGroup(46, 47, "Transformer summary and references", "Attention and positional encoding", "Next lecture on LLMs", "strong"),
    ],
    9: [
        ConceptGroup(1, 4, "LLM overview tokenization and embeddings", "Transformer fundamentals", "Long-distance dependencies", "strong"),
        ConceptGroup(5, 8, "Long-distance dependency and transformer advantage", "Token embeddings", "Pretraining and fine-tuning", "strong"),
        ConceptGroup(9, 16, "BERT pretraining and fine-tuning examples", "Transformer encoder representations", "GPT autoregressive modeling", "strong"),
        ConceptGroup(17, 21, "GPT BERT comparison and autoregressive next-token prediction", "BERT encoder-only modeling", "Few-shot and in-context learning", "strong"),
        ConceptGroup(22, 24, "Few-shot in-context learning and scaling laws", "Autoregressive LMs", "Alignment and RLHF", "strong"),
        ConceptGroup(25, 31, "Instruction tuning reward modeling and RLHF", "Pretrained LLM behavior", "LLM summary", "strong"),
        ConceptGroup(32, 32, "LLM summary", "Pretraining fine-tuning and alignment", "Next lecture on vision transformers", "strong"),
    ],
    10: [
        ConceptGroup(1, 5, "Vision Transformer motivation and overview", "Transformers and CNNs", "ViT patch pipeline", "strong"),
        ConceptGroup(6, 15, "ViT patch embeddings encoder and classification token", "Transformer encoder basics", "ViT limitations and training", "strong"),
        ConceptGroup(16, 23, "ViT inductive bias cost pretraining and evaluation", "ViT core architecture", "DeiT and distillation", "strong"),
        ConceptGroup(24, 32, "DeiT and knowledge distillation", "ViT data efficiency limits", "Tokenization and hierarchical variants", "strong"),
        ConceptGroup(33, 38, "T2T PVT and spatial-reduction attention", "Patch tokenization limits", "Swin hierarchy", "strong"),
        ConceptGroup(39, 58, "Swin hierarchy window attention and shifted windows", "Hierarchical vision features", "Vision transformer variants", "strong"),
        ConceptGroup(59, 63, "Vision transformer variants and segmentation", "Swin attention design", "Self-supervised vision transformers", "strong"),
        ConceptGroup(64, 67, "MAE DINO iBOT and self-supervised vision", "Transformer vision backbones", "Multimodal transformers", "strong"),
        ConceptGroup(68, 70, "CLIP multimodal transformers and recent evolution", "Self-supervised visual representations", "Vision transformer summary", "strong"),
        ConceptGroup(71, 71, "Vision transformer summary", "Vision transformer evolution", "References and backup material", "strong"),
        ConceptGroup(72, 75, "References and backup slides", "Completed main lecture", "Next lecture on generative models", "medium"),
    ],
    11: [
        ConceptGroup(1, 2, "Generative modeling overview and learning types", "Supervised deep learning", "Unsupervised learning bridge", "strong"),
        ConceptGroup(3, 5, "Unsupervised learning bridge to generative models", "Learning without labels", "Generative model goals", "strong"),
        ConceptGroup(6, 8, "Goals and families of generative models", "Density estimation and sampling", "Autoencoders and VAEs", "strong"),
        ConceptGroup(9, 13, "Autoencoders VAEs and latent spaces", "Representation learning", "GANs", "strong"),
        ConceptGroup(14, 16, "GAN objectives strengths and weaknesses", "Sampleable latent models", "Diffusion models", "strong"),
        ConceptGroup(17, 18, "Diffusion and latent conditional generation", "Adversarial generation", "Generative model summary", "strong"),
        ConceptGroup(19, 19, "Generative model summary", "Autoencoders GANs and diffusion", "Course wrap-up or next module", "strong"),
    ],
}


def clean_text(value: str) -> str:
    return " ".join(str(value or "").replace("\r", "\n").split())


def project_root() -> Path:
    return Path(__file__).resolve().parents[1]


def input_files(input_dir: Path) -> list[Path]:
    files = sorted(
        input_dir.glob("*_slide_text.csv"),
        key=lambda path: int(path.name.split("_", 1)[0]),
    )
    return [path for path in files if int(path.name.split("_", 1)[0]) >= 1]


def read_rows(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def group_for_slide(lecture_number: int, slide_number: int) -> ConceptGroup:
    for group in LECTURE_GROUPS[lecture_number]:
        if group.start <= slide_number <= group.end:
            return group
    raise ValueError(f"No concept group configured for Lecture {lecture_number}, slide {slide_number}")


def slide_role(title: str, group: ConceptGroup) -> str:
    text = f"{title} {group.name}".lower()
    if "summary" in text or "reference" in text:
        return "summary"
    if "discussion" in text or "question" in text or "which cases" in text:
        return "exercise"
    if "worked example" in text or "example" in text or "demo" in text or "playground" in text:
        return "example"
    if "derivation" in text or "gradient" in text or "formula" in text or "entropy" in text:
        return "derivation"
    if "algorithm" in text or "workflow" in text or "pytorch" in text or "keras" in text or "training" in text or "step " in text:
        return "algorithm"
    if "architecture" in text or "visualization" in text or "curve" in text or "animation" in text or "contour" in text or "layout" in text:
        return "visualization"
    if "what is" in text or "definition" in text or "notation" in text or "core idea" in text:
        return "definition"
    if "why" in text or "problem" in text or "challenge" in text or "motivation" in text or "limitation" in text:
        return "motivation"
    if "overview" in text or "roadmap" in text or "from " in text or "vs" in text:
        return "transition"
    return "definition"


def core_concept(title: str, group: ConceptGroup, slide_number: int) -> str:
    cleaned_title = clean_text(title)
    if not cleaned_title or re.fullmatch(r"\d+|[a-zA-Z]", cleaned_title):
        return f"{group.name} detail {slide_number}"
    if len(cleaned_title) > 110:
        return cleaned_title[:107].rstrip() + "..."
    return cleaned_title


def estimated_minutes(role: str, title: str, group: ConceptGroup) -> str:
    text = f"{title} {group.name}".lower()
    minutes = {
        "logistics": 0.5,
        "transition": 0.7,
        "motivation": 0.9,
        "definition": 1.0,
        "visualization": 1.0,
        "example": 1.2,
        "algorithm": 1.3,
        "derivation": 1.5,
        "exercise": 1.5,
        "summary": 0.9,
        "other": 0.8,
    }[role]
    if any(token in text for token in ["worked example", "step-by-step", "derivation", "matrix view"]):
        minutes += 0.3
    if any(token in text for token in ["summary", "reference", "backup"]):
        minutes = min(minutes, 0.8)
    return f"{minutes:.1f}"


def boundary_for_slide(slide_number: int, group: ConceptGroup, is_last_slide: bool) -> tuple[str, str, str]:
    if slide_number == group.end:
        strength = "strong" if is_last_slide else group.boundary_strength
        reason = (
            f"Completes the concept group '{group.name}' and the next slide begins "
            f"'{group.next_concept}'. This is a natural stopping point without cutting a derivation, algorithm, example, architecture explanation, or comparison."
        )
        if is_last_slide:
            reason = f"Ends the lecture after completing '{group.name}'."
        return "yes", strength, reason

    return (
        "no",
        "weak",
        f"Continue through the current concept group '{group.name}'; stopping here would interrupt the conceptual sequence.",
    )


def notes_for_script(role: str, group: ConceptGroup, title: str, boundary: str) -> str:
    base = {
        "motivation": "Use this slide to establish why the topic matters before adding formal detail.",
        "definition": "State the key definition carefully and connect it to the previous slide.",
        "example": "Walk through the example step by step; avoid summarizing before the learner sees the complete case.",
        "derivation": "Keep the mathematical chain continuous and reserve recap until the derivation boundary.",
        "algorithm": "Narrate the procedure in order and emphasize inputs, outputs, and update rules.",
        "visualization": "Describe the visual structure explicitly for online viewers who may be watching on a small screen.",
        "summary": "Use this slide to consolidate the preceding concept group and prepare a checkpoint.",
        "exercise": "Pause for learner reasoning before revealing or discussing the answer.",
        "transition": "Use this slide as a bridge; preview the next concept without overloading the slide.",
        "logistics": "Keep this short and orient the learner to the lecture goal.",
        "other": "Use concise narration and tie the slide back to the current concept group.",
    }[role]
    if boundary == "yes":
        return f"{base} Add a brief recap or checkpoint here before moving on."
    return base


def build_map_rows(rows: list[dict[str, str]]) -> list[dict[str, str]]:
    output_rows: list[dict[str, str]] = []
    max_slide_number = max(int(row["slide_number"]) for row in rows)

    concepts: dict[int, str] = {}
    groups: dict[int, ConceptGroup] = {}
    for row in rows:
        lecture_number = int(row["lecture_number"])
        slide_number = int(row["slide_number"])
        group = group_for_slide(lecture_number, slide_number)
        groups[slide_number] = group
        concepts[slide_number] = core_concept(row.get("possible_slide_title", ""), group, slide_number)

    for row in rows:
        lecture_number = int(row["lecture_number"])
        slide_number = int(row["slide_number"])
        group = groups[slide_number]
        slide_title = clean_text(row.get("possible_slide_title", ""))
        role = slide_role(slide_title, group)
        boundary, strength, reason = boundary_for_slide(
            slide_number, group, slide_number == max_slide_number
        )

        previous_concept = (
            group.prerequisite if slide_number == group.start else concepts.get(slide_number - 1, "")
        )
        next_concept = (
            group.next_concept if slide_number == group.end else concepts.get(slide_number + 1, "")
        )

        output_rows.append(
            {
                "lecture_number": str(lecture_number),
                "lecture_title": row.get("lecture_title", ""),
                "slide_number": str(slide_number),
                "slide_title": slide_title,
                "core_concept": concepts[slide_number],
                "slide_role": role,
                "concept_group": group.name,
                "prerequisite_concept": previous_concept,
                "next_concept": next_concept,
                "estimated_teaching_minutes": estimated_minutes(role, slide_title, group),
                "logical_boundary_after_slide": boundary,
                "boundary_strength": strength,
                "boundary_reason": reason,
                "notes_for_future_script": notes_for_script(role, group, slide_title, boundary),
            }
        )

    return output_rows


def write_csv(path: Path, rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=CSV_FIELDS)
        writer.writeheader()
        writer.writerows(rows)


def main() -> int:
    root = project_root()
    input_dir = root / INPUT_DIR
    output_dir = root / OUTPUT_DIR

    try:
        files = input_files(input_dir)
        total_slides = 0
        for file_path in files:
            source_rows = read_rows(file_path)
            if not source_rows:
                continue
            lecture_number = int(source_rows[0]["lecture_number"])
            map_rows = build_map_rows(source_rows)
            write_csv(output_dir / f"{lecture_number}_teaching_map.csv", map_rows)
            total_slides += len(map_rows)
    except (OSError, RuntimeError, ValueError, KeyError) as exc:
        print("Teaching map build failed.")
        print(f"Error: {exc}")
        return 1

    print("Teaching map build complete")
    print(f"Input folder:       {input_dir}")
    print(f"Output folder:      {output_dir}")
    print(f"Lectures processed: {len(files)}")
    print(f"Slides mapped:      {total_slides}")
    print("PowerPoint safety: used extracted CSVs only; no PowerPoint files were read or modified.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
