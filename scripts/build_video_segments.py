"""Create video segmentation plans from teaching maps."""

from __future__ import annotations

import csv
import sys
from dataclasses import dataclass
from pathlib import Path


INPUT_DIR = Path("manifests") / "teaching_maps"
OUTPUT_DIR = Path("manifests") / "video_segments"
REVIEW_STATUS = "needs_review"
CSV_FIELDS = [
    "lecture_number",
    "lecture_title",
    "video_number",
    "video_title",
    "original_slide_start",
    "original_slide_end",
    "estimated_minutes",
    "main_concept_group",
    "learning_objectives",
    "why_this_is_a_logical_segment",
    "why_the_ending_slide_is_a_good_stopping_point",
    "suggested_recap_slide",
    "suggested_quiz_question",
    "review_status",
]


@dataclass(frozen=True)
class SegmentSpec:
    video_number: int
    start: int
    end: int
    video_title: str
    learning_objectives: str
    why_segment: str
    recap: str
    quiz: str


SEGMENTS: dict[int, list[SegmentSpec]] = {
    1: [
        SegmentSpec(
            1,
            1,
            21,
            "From MLPs to Convolution",
            "Review MLP components; explain why image data needs spatial structure; define convolution as a local sliding dot product; compute convolution output size.",
            "This segment moves from neural-network refresher material into the motivation for CNNs and completes the first full mathematical explanation of convolution before switching to filter mechanics and multichannel CNNs.",
            "Recap MLP limitations for images, local receptive fields, sliding filters, and output-size reasoning.",
            "Why does convolution use fewer parameters than a fully connected layer on the same image?",
        ),
        SegmentSpec(
            2,
            22,
            33,
            "Filters, Padding, and Kernel Mechanics",
            "Explain image filters and padding; compute kernel behavior at image borders; connect classical filters to learned convolution kernels.",
            "This shorter segment isolates filter and padding mechanics after learners have already seen the basic convolution operation.",
            "Recap filters, padding, border effects, and how fixed image-processing kernels relate to learned CNN filters.",
            "Why can padding preserve spatial dimensions while still allowing a convolution filter to inspect border pixels?",
        ),
        SegmentSpec(
            3,
            34,
            41,
            "Multichannel Convolution and Feature Maps",
            "Describe multichannel filters; explain feature maps; connect stacked convolution layers to learned hierarchical representations.",
            "This segment isolates tensor and feature-map mechanics before moving to pooling and complete CNN architecture.",
            "Recap filter depth, multichannel convolution, feature maps, and multilayer feature hierarchies.",
            "How do multichannel filters combine input channels to create one feature map?",
        ),
        SegmentSpec(
            4,
            42,
            52,
            "Pooling, Inductive Bias, and AlexNet",
            "Explain pooling and spatial reduction; connect CNN inductive bias to hierarchical representations; summarize AlexNet's architecture.",
            "This segment completes the CNN-building-block story by moving from spatial reduction into a concrete architecture example and final summary.",
            "Recap pooling, spatial hierarchy, translation tolerance, AlexNet layers, and the main CNN takeaways.",
            "How does pooling change the spatial representation before a CNN reaches its classifier layers?",
        ),
    ],
    2: [
        SegmentSpec(
            1,
            1,
            14,
            "Confusion Matrices and Core Classification Metrics",
            "Explain binary classifier decisions; read a confusion matrix; compute accuracy, precision, recall, specificity, and F1.",
            "This segment builds the vocabulary and formulas needed before learners reason about threshold movement and ROC curves.",
            "Recap true and false positives and negatives, accuracy, precision, recall, specificity, and F1.",
            "For a rare-disease classifier, why can high accuracy still be misleading, and which metric would you inspect next?",
        ),
        SegmentSpec(
            2,
            15,
            32,
            "Threshold Tradeoffs, ROC Curves, and Metric Selection",
            "Explain threshold-dependent tradeoffs; interpret ROC curves and AUC; choose metrics based on task consequences.",
            "This segment keeps threshold movement, ROC/AUC, and final metric choice together because the interpretation depends on comparing operating points.",
            "Recap threshold changes, precision-recall tradeoffs, ROC curves, AUC interpretation, and task-driven metric selection.",
            "If false negatives are much more costly than false positives, how should that affect your threshold and metric choice?",
        ),
    ],
    3: [
        SegmentSpec(
            1,
            1,
            16,
            "Loss Functions and Gradient Descent",
            "Connect score functions to loss functions; describe gradient descent; explain learning-rate effects; compare optimization behavior.",
            "This segment completes the optimization foundation before the lecture moves into the linear-classifier worked example and classification-specific loss construction.",
            "Recap model scores, loss, gradient direction, update steps, convergence, and learning-rate behavior.",
            "What can go wrong if the learning rate is too large, and how does the training curve reveal it?",
        ),
        SegmentSpec(
            2,
            17,
            23,
            "Linear Classifier Scores",
            "Interpret linear classifier scores; connect class weights to logits; explain how a worked example produces one score per class.",
            "This segment separates the linear scoring example from the probability and loss derivation that follows.",
            "Recap logits, class score vectors, linear weights, and how a classifier ranks classes before normalization.",
            "What information does a raw class score contain before it has been converted into a probability?",
        ),
        SegmentSpec(
            3,
            24,
            38,
            "Softmax and Cross-Entropy",
            "Convert logits to probabilities with softmax; explain numerical stability; derive cross-entropy for one-hot labels; compare cross-entropy with squared error.",
            "This segment keeps softmax normalization and cross-entropy together because the loss derivation depends directly on the probability model.",
            "Recap softmax normalization, stable computation, entropy, cross-entropy, and why cross-entropy fits classification.",
            "Why does subtracting the maximum logit improve softmax stability without changing the predicted probabilities?",
        ),
    ],
    4: [
        SegmentSpec(
            1,
            1,
            18,
            "Generalization, Validation, and Regularized Loss",
            "Distinguish overfitting and underfitting; explain bias-variance behavior; use train, validation, and test splits; describe regularized loss minimization.",
            "This segment completes the diagnosis and measurement side of generalization before moving into the geometry of L1/L2 penalties and practical control methods.",
            "Recap underfit, good fit, overfit, bias-variance tradeoff, validation splits, cross validation, and regularized objectives.",
            "Why is a validation set used for model selection instead of tuning directly on the test set?",
        ),
        SegmentSpec(
            2,
            19,
            39,
            "Regularization Geometry and Training Controls",
            "Explain L1 and L2 geometry; connect sparsity and shrinkage; describe dropout, augmentation, early stopping, and learning-rate scheduling; assemble a training recipe.",
            "The segment keeps the full regularization toolbox together: L1/L2 geometry leads naturally into dropout, augmentation, early stopping, and the final learning recipe.",
            "Recap L1 sparsity, L2 shrinkage, dropout, data augmentation, early stopping, learning-rate scheduling, and the final recipe.",
            "When would L1 regularization be preferred over L2 regularization, and what behavior would you expect in the learned weights?",
        ),
    ],
    5: [
        SegmentSpec(
            1,
            1,
            14,
            "Early CNN Milestones",
            "Trace the CNN evolution from LeNet to AlexNet; explain global average pooling; connect early design choices to practical CNN training.",
            "This segment covers the historical CNN foundation without mixing in the later depth and optimization problem.",
            "Recap LeNet, AlexNet, global average pooling, and why early CNN design choices mattered.",
            "Which AlexNet design choices helped make CNNs practical for large-scale image recognition?",
        ),
        SegmentSpec(
            2,
            15,
            33,
            "Depth, VGG, and the Degradation Problem",
            "Explain batch normalization; describe why depth and hierarchy help; identify degradation, vanishing gradients, and optimization limits.",
            "This segment isolates the build-up to the depth problem so the next video can begin cleanly with residual learning as the solution.",
            "Recap batch normalization, depth, VGG-style stacking, vanishing gradients, and the degradation problem.",
            "Why was the degradation problem surprising if deeper networks should be able to represent shallower ones?",
        ),
        SegmentSpec(
            3,
            34,
            50,
            "ResNet and Inception Design Patterns",
            "Explain residual learning; describe skip connections; interpret Inception-style multi-scale branches and efficient convolution choices.",
            "This segment starts with the ResNet answer to degradation and then covers Inception as a separate modern CNN design pattern.",
            "Recap residual shortcuts, identity mappings, parallel branches, 1x1 bottlenecks, and multi-scale CNN design.",
            "How do residual connections and dense connections solve related but different information-flow problems?",
        ),
        SegmentSpec(
            4,
            51,
            69,
            "Efficient CNNs and Modern Design Principles",
            "Compare MobileNet and DenseNet efficiency ideas; explain compound scaling, MBConv, and channel attention; summarize modern CNN design principles.",
            "This segment groups the efficiency-focused architectures and final design principles after ResNet and Inception are complete.",
            "Recap depthwise separable convolution, dense concatenation, compound scaling, MBConv, SE, and modern CNN principles.",
            "How do depthwise separable convolution and compound scaling reduce cost while preserving model capacity?",
        ),
    ],
    6: [
        SegmentSpec(
            1,
            1,
            12,
            "Semantic Segmentation and Fully Convolutional Networks",
            "Define semantic segmentation; explain why sliding windows fail; convert classifiers into fully convolutional networks.",
            "This segment keeps the task definition and FCN conversion together before introducing encoder-decoder details.",
            "Recap dense prediction, sliding-window limitations, and how fully convolutional networks preserve spatial outputs.",
            "Why is a fully convolutional network better suited than a sliding-window classifier for segmentation?",
        ),
        SegmentSpec(
            2,
            13,
            24,
            "U-Net Structure and Upsampling Choices",
            "Describe U-Net and encoder-decoder structure; compare interpolation-based upsampling methods; explain why resolution recovery matters.",
            "This segment focuses on the architecture and upsampling bridge before the lecture moves into transposed-convolution mechanics.",
            "Recap downsampling, skip connections, encoder-decoder structure, nearest-neighbor upsampling, and bilinear interpolation.",
            "Why do segmentation networks need a mechanism to recover spatial resolution after downsampling?",
        ),
        SegmentSpec(
            3,
            25,
            41,
            "Transposed Convolution Mechanics and Artifacts",
            "Explain transposed convolution spatially; connect convolution to matrix multiplication; derive transposed convolution from the matrix view; identify checkerboard artifacts.",
            "This segment keeps the transposed-convolution worked examples and matrix derivation intact, then ends after discussing limitations and summarizing the operation.",
            "Recap overlapping contributions, stride effects, matrix transposition, stride-2 examples, and checkerboard artifacts.",
            "In the matrix view, why is transposed convolution not simply the inverse of convolution?",
        ),
    ],
    7: [
        SegmentSpec(
            1,
            1,
            23,
            "RNNs, LSTM Gates, and PyTorch Workflow",
            "Classify sequence modeling tasks; explain RNN recurrence; identify why long-term dependencies are difficult; describe LSTM gates and cell state updates; use the PyTorch LSTM workflow.",
            "This segment completes the LSTM mechanism and implementation foundation before moving to higher-level sequence model patterns and worked generation examples.",
            "Recap sequence task types, recurrence, long-term dependency, forget/input/output gates, cell state updates, parameter count, and PyTorch inputs.",
            "Which LSTM gate decides how much old cell-state information should be retained, and why is that useful?",
        ),
        SegmentSpec(
            2,
            24,
            38,
            "Sequence Variants and Seq2Seq Generation",
            "Compare one-to-many, many-to-one, and many-to-many models; explain bidirectional LSTMs; walk through Seq2Seq response generation.",
            "This segment keeps sequence model variants with the Seq2Seq worked example, then stops before the separate image-captioning application.",
            "Recap sequence mapping patterns, bidirectionality, encoder-decoder structure, and Seq2Seq response generation.",
            "Why does an encoder-decoder sequence model need a decoder state rather than only the final predicted token?",
        ),
        SegmentSpec(
            3,
            39,
            47,
            "Image Captioning with LSTMs",
            "Connect CNN image features to LSTM decoding; walk through the image-captioning example; summarize where LSTMs work well and where they struggle.",
            "This segment isolates image captioning as its own application instead of folding it into the Seq2Seq module.",
            "Recap visual feature extraction, caption decoding, start and end tokens, and the strengths and limits of LSTMs.",
            "Why does image captioning combine a visual encoder with a language decoder instead of using only one model type?",
        ),
    ],
    8: [
        SegmentSpec(
            1,
            1,
            18,
            "Seq2Seq Bottlenecks and Attention",
            "Explain the Seq2Seq encoder-decoder bottleneck; describe teacher forcing; compute attention context from encoder states; interpret alignment weights.",
            "This segment completes the motivation for attention and the decoder-looking-back mechanism before starting self-attention, which is a different architecture idea.",
            "Recap encoder-decoder flow, final-state bottleneck, teacher forcing, attention scores, normalized weights, and context vectors.",
            "How does attention reduce the information bottleneck in a basic Seq2Seq model?",
        ),
        SegmentSpec(
            2,
            19,
            31,
            "Self-Attention and QKV Computation",
            "Compute self-attention from queries, keys, and values; explain attention weights; connect token-to-token comparison to contextual representations.",
            "This segment isolates the core self-attention calculation before adding position information and multi-head structure.",
            "Recap Q/K/V roles, dot-product attention, normalized weights, values, and contextual token representations.",
            "What does the query-key dot product measure in self-attention?",
        ),
        SegmentSpec(
            3,
            32,
            47,
            "Positional Encoding and Transformer Blocks",
            "Explain positional encoding; interpret attention visualizations; describe multi-head attention; distinguish self-attention from cross-attention.",
            "This segment completes the Transformer block after learners understand self-attention mechanics.",
            "Recap positional encoding, sinusoidal patterns, attention visualization, multi-head attention, cross-attention, and the transformer summary.",
            "Why does a transformer need positional encoding if self-attention can compare every token with every other token?",
        ),
    ],
    9: [
        SegmentSpec(
            1,
            1,
            8,
            "LLM Inputs and Transformer Advantage",
            "Explain tokenization and embeddings; describe why transformers help with long-distance dependencies.",
            "This segment introduces LLM representations and the architecture motivation before moving into pretraining objectives.",
            "Recap tokens, embeddings, context, long-distance dependencies, and the transformer advantage.",
            "Why do token embeddings need context before they become useful language representations?",
        ),
        SegmentSpec(
            2,
            9,
            21,
            "BERT, GPT, and Pretraining Objectives",
            "Explain BERT-style pretraining and fine-tuning; describe GPT-style autoregressive next-token prediction; compare encoder and decoder language modeling objectives.",
            "This segment keeps the pretraining and fine-tuning examples together and ends after the BERT/GPT objective comparison.",
            "Recap masked language modeling, BERT fine-tuning, GPT autoregressive prediction, and encoder-decoder objective differences.",
            "What is the key training objective difference between BERT-style masked modeling and GPT-style autoregressive modeling?",
        ),
        SegmentSpec(
            3,
            22,
            32,
            "Prompting, Scaling, and Alignment",
            "Explain few-shot and in-context learning; outline scaling-law intuition; describe instruction tuning, reward modeling, and RLHF.",
            "This segment focuses on how pretrained LLMs are adapted, prompted, scaled, and aligned after the core objectives are established.",
            "Recap few-shot prompting, in-context learning, scaling laws, supervised fine-tuning, reward models, RLHF, and the LLM summary.",
            "Why does RLHF add a reward model after supervised fine-tuning instead of relying only on next-token prediction?",
        ),
    ],
    10: [
        SegmentSpec(
            1,
            1,
            23,
            "Vision Transformer Core Architecture",
            "Explain why transformers can be used for images; convert images into patch tokens; describe patch embeddings, positional embeddings, the CLS token, and ViT training limitations.",
            "This segment completes the core ViT pipeline and its main practical constraints before moving to data-efficient improvements and hierarchical variants.",
            "Recap image patches, linear patch embedding, positional embedding, transformer encoder, CLS token, inductive bias, attention cost, and pretraining.",
            "Why does ViT usually need more data than a CNN when trained from scratch?",
        ),
        SegmentSpec(
            2,
            24,
            38,
            "Data-Efficient Vision Transformers",
            "Explain DeiT distillation; describe progressive tokenization; explain PVT and spatial-reduction attention.",
            "This segment focuses on data-efficient and early hierarchical improvements before the Swin architecture begins.",
            "Recap DeiT, teacher-student distillation, distillation tokens, T2T, PVT, and spatial-reduction attention.",
            "How does a distillation token help DeiT learn from a teacher model?",
        ),
        SegmentSpec(
            3,
            39,
            58,
            "Swin Transformer Hierarchy and Shifted Windows",
            "Explain Swin patch merging; describe window attention; show how shifted windows enable cross-window communication.",
            "This segment keeps the Swin architecture explanation intact from hierarchy through shifted-window attention.",
            "Recap Swin hierarchy, patch merging, local windows, shifted windows, and cross-window information flow.",
            "How does shifted-window attention let Swin communicate across local windows without using full global attention?",
        ),
        SegmentSpec(
            4,
            59,
            70,
            "Vision Transformer Extensions and Multimodal Models",
            "Compare early vision transformer variants; describe transformer segmentation; explain MAE, DINO, and CLIP at a high level; summarize how vision transformers evolved.",
            "This segment focuses on ViT extensions, self-supervised vision, and multimodal transformers before the closing summary and backup material.",
            "Recap Twins, Segmenter, MAE, DINO, CLIP, and recent multimodal evolution.",
            "What changes when a vision transformer is trained with image-text contrastive learning rather than only image labels?",
        ),
        SegmentSpec(
            5,
            71,
            75,
            "Vision Transformer Summary and Backup",
            "Summarize the main ViT architecture trends; identify which reference and backup slides should support later script writing.",
            "This short closing segment keeps summary, references, and backup slides separate from the main technical extensions.",
            "Recap the core ViT trend line, what belongs in the final summary, and which backup slides are reference-only.",
            "Which ViT ideas from this lecture are core teaching points, and which should remain as backup context?",
        ),
    ],
    11: [
        SegmentSpec(
            1,
            1,
            8,
            "Generative Modeling Foundations",
            "Distinguish supervised, unsupervised, and generative modeling; connect clustering and density estimation to generative goals.",
            "This segment establishes what generative modeling is before comparing specific model families.",
            "Recap learning types, clustering, density estimation, and the goals of generative models.",
            "What makes a model generative rather than only discriminative?",
        ),
        SegmentSpec(
            2,
            9,
            13,
            "Autoencoders, VAEs, and Latent Spaces",
            "Explain reconstruction; describe latent spaces; compare autoencoders and VAEs as generative modeling tools.",
            "This segment separates reconstruction-based generative models from adversarial and diffusion models.",
            "Recap reconstruction, bottlenecks, latent spaces, autoencoders, and VAEs.",
            "Why does a latent space make autoencoders useful for generative modeling?",
        ),
        SegmentSpec(
            3,
            14,
            19,
            "GANs, Diffusion, and Generative Model Summary",
            "Explain GAN objectives; describe diffusion at a high level; summarize how major generative model families differ.",
            "This segment keeps adversarial and diffusion generation together and closes with the lecture summary.",
            "Recap generator-discriminator training, GAN strengths and weaknesses, diffusion, latent conditional generation, and the course summary.",
            "How do GANs and diffusion models differ in the way they learn to generate samples?",
        ),
    ],
}


def project_root() -> Path:
    return Path(__file__).resolve().parents[1]


def read_teaching_map(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def rows_for_segment(
    rows: list[dict[str, str]],
    start: int,
    end: int,
) -> list[dict[str, str]]:
    segment_rows = [
        row for row in rows if start <= int(row["slide_number"]) <= end
    ]
    if not segment_rows:
        raise ValueError(f"No teaching-map rows found for slides {start}-{end}")
    return segment_rows


def main_concept_group(segment_rows: list[dict[str, str]]) -> str:
    groups: list[str] = []
    for row in segment_rows:
        group = row["concept_group"]
        if group not in groups:
            groups.append(group)
    return "; ".join(groups)


def estimated_minutes(segment_rows: list[dict[str, str]]) -> str:
    total = sum(float(row["estimated_teaching_minutes"]) for row in segment_rows)
    return f"{total:.1f}"


def validate_boundary(segment_rows: list[dict[str, str]], spec: SegmentSpec) -> None:
    last = segment_rows[-1]
    if int(last["slide_number"]) != spec.end:
        raise ValueError(f"Segment {spec.video_number} does not end at slide {spec.end}")
    if last["logical_boundary_after_slide"] != "yes":
        raise ValueError(
            f"Lecture segment {spec.video_number} ends at slide {spec.end}, "
            "which is not marked as a logical boundary."
        )
    if last["boundary_strength"] not in {"medium", "strong"}:
        raise ValueError(
            f"Lecture segment {spec.video_number} ends at slide {spec.end}, "
            f"but boundary strength is {last['boundary_strength']}."
        )


def build_segment_rows(
    lecture_number: int,
    teaching_rows: list[dict[str, str]],
) -> list[dict[str, str]]:
    segment_rows: list[dict[str, str]] = []
    lecture_title = teaching_rows[0]["lecture_title"]

    for spec in SEGMENTS[lecture_number]:
        source_rows = rows_for_segment(teaching_rows, spec.start, spec.end)
        validate_boundary(source_rows, spec)
        last = source_rows[-1]
        segment_rows.append(
            {
                "lecture_number": str(lecture_number),
                "lecture_title": lecture_title,
                "video_number": str(spec.video_number),
                "video_title": spec.video_title,
                "original_slide_start": str(spec.start),
                "original_slide_end": str(spec.end),
                "estimated_minutes": estimated_minutes(source_rows),
                "main_concept_group": main_concept_group(source_rows),
                "learning_objectives": spec.learning_objectives,
                "why_this_is_a_logical_segment": spec.why_segment,
                "why_the_ending_slide_is_a_good_stopping_point": last["boundary_reason"],
                "suggested_recap_slide": spec.recap,
                "suggested_quiz_question": spec.quiz,
                "review_status": REVIEW_STATUS,
            }
        )

    covered = {
        slide
        for spec in SEGMENTS[lecture_number]
        for slide in range(spec.start, spec.end + 1)
    }
    all_slides = {int(row["slide_number"]) for row in teaching_rows}
    missing = sorted(all_slides - covered)
    extra = sorted(covered - all_slides)
    if missing or extra:
        raise ValueError(
            f"Lecture {lecture_number} segment coverage mismatch. "
            f"Missing slides: {missing}; extra slides: {extra}"
        )

    return segment_rows


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
        total_segments = 0
        for lecture_number in sorted(SEGMENTS):
            teaching_map_path = input_dir / f"{lecture_number}_teaching_map.csv"
            teaching_rows = read_teaching_map(teaching_map_path)
            segment_rows = build_segment_rows(lecture_number, teaching_rows)
            write_csv(output_dir / f"{lecture_number}_video_segments.csv", segment_rows)
            total_segments += len(segment_rows)
    except (OSError, RuntimeError, ValueError, KeyError) as exc:
        print("Video segmentation build failed.")
        print(f"Error: {exc}")
        return 1

    print("Video segmentation build complete")
    print(f"Input folder:       {input_dir}")
    print(f"Output folder:      {output_dir}")
    print(f"Lectures processed: {len(SEGMENTS)}")
    print(f"Segments created:   {total_segments}")
    print("PowerPoint safety: used teaching-map CSVs only; no PowerPoint files were read or modified.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
