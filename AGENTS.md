# Codex Instructions for ECE 591 Online Deep Learning

## Goal

Help convert in-person deep learning lecture slides into online 30-minute video modules.

## Course folder

The Google Drive course folder is machine-specific.

Read the local Google Drive root from:

- `course_config.local.yaml`

Read shared folder names and policies from:

- `course_config.yaml`

Do not hardcode drive letters such as `H:\` or `G:\` in scripts.

## PowerPoint filename convention

Original lecture PowerPoint files use this pattern:

- `0 - Title.pptx`
- `1 - Title.pptx`
- `2 - Title.pptx`

The number before the hyphen is the lecture number.

When processing files:

- Sort PowerPoint files numerically by the leading lecture number.
- Treat `0 - *.pptx` as Lecture 0.
- Lecture 0 is a special case and is already approximately one 30-minute video.
- Batch segmentation should process files starting from `1 - *.pptx` unless explicitly instructed otherwise.
- Do not assume filenames such as `Lecture_00_Intro.pptx`.

## Rules

- Do not overwrite original PowerPoint files.
- Original slides are stored in `01_Original_Slides`.
- Revised slides must be written only to `02_Revised_Slides`.
- Preserve the original PowerPoint theme unless instructed otherwise.
- Preserve technical correctness.
- Keep slide text concise.
- Put detailed explanations in speaker notes.
- Each 30-minute video should include:
  - clear title
  - 2–4 learning objectives
  - coherent slide sequence
  - speaker notes
  - one recap slide
  - one quiz/checkpoint slide

## Workflow

1. Analyze the original lecture deck.
2. Extract slide text.
3. Create a slide inventory.
4. Estimate teaching time.
5. Segment the lecture into 30-minute videos.
6. Suggest revisions.
7. Only after review, generate revised PowerPoint files.

## Speaker script quality

The exact speaker script is the primary teaching product. When creating or revising scripts:

- Write exact word-for-word read-aloud narration, not outlines, bullet notes, or metadata summaries.
- Keep spoken narration natural, clear, and instructor-like. Avoid dense textbook prose, generic filler, and directions such as "explain this" or "mention that".
- Do not mention lecture numbers, video numbers, slide ranges, filenames, manifests, generated decks, or segmentation details in spoken narration.
- Do not put raw URLs in spoken narration. If a slide includes a URL, describe the resource briefly and put the URL in a separate reference note.
- Teach the meaning behind each slide, not just the visible bullets. Explain why the concept matters and connect it to the previous and next ideas.
- Keep narration slide-aligned. Explain equations, diagrams, arrows, architectures, examples, comparisons, and code only when they are supported by the slide context.
- Do not force a fixed 25-30 minute duration. Estimate speaking time after narration is generated using about 125-135 spoken words per minute.
- Allocate more time to equations, algorithms, model architectures, code, diagrams, and worked examples; use less time for title, agenda, transition, recap, and simple checkpoint slides.

## Technical correction rules

For each slide script:

- Identify the main technical claim before writing narration.
- Check whether the claim is accurate for a deep learning course.
- Correct misleading, incomplete, outdated, ambiguous, or imprecise explanations in the spoken narration while preserving the intended meaning.
- Add a separate technical correction note when the slide itself appears technically wrong, ambiguous, outdated, or potentially misleading.
- Do not silently introduce unsupported technical claims.
- Be especially careful with loss functions vs. objectives, gradient descent vs. backpropagation, gradients vs. parameter updates, training vs. inference, validation vs. testing, overfitting vs. underfitting, capacity vs. generalization, activations and nonlinear representations, softmax scores vs. calibrated probabilities, cross-entropy vs. accuracy, convolution vs. correlation, pooling and translation invariance, embeddings and representations, attention weights and interpretability, encoder/decoder/transformer terminology, generative vs. discriminative models, likelihood vs. sampling vs. prediction, and supervised vs. unsupervised vs. self-supervised learning.

## Safety

Never edit the same PowerPoint file from two machines at the same time.
Never overwrite original PowerPoint files.
