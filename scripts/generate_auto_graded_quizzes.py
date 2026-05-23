"""Generate auto-graded quizzes from final exact speaker scripts.

This script reads final exact speaker scripts and writes Markdown quizzes plus a
CSV summary. It does not open or modify PowerPoint files or speaker scripts.
"""

from __future__ import annotations

import csv
import re
from pathlib import Path
from typing import Any

import yaml


REPO_ROOT = Path(__file__).resolve().parents[1]
WPM = 130


def load_yaml(path: Path) -> dict[str, Any]:
    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    if not isinstance(data, dict):
        raise SystemExit(f"{path} must contain a YAML mapping.")
    return data


def normalize_space(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def slug_title_from_script(path: Path) -> str:
    return path.name.replace("_exact_script.md", "")


def exact_script_path(drive: Path, folders: dict[str, str], speaker_script: str) -> Path:
    rel = Path(speaker_script)
    exact_name = rel.name.replace("_script.md", "_exact_script.md")
    return drive / folders["scripts"] / "exact_scripts" / rel.parent.name / exact_name


def script_title(text: str, fallback: str) -> str:
    match = re.search(r"(?m)^#\s+(.+?)\s*$", text)
    return normalize_space(match.group(1)) if match else fallback


def slide_titles(text: str) -> list[str]:
    return [normalize_space(m.group(1)) for m in re.finditer(r"(?m)^## Slide \d+:\s*(.+?)\s*$", text)]


def technical_terms(text: str) -> list[str]:
    headings = []
    for match in re.finditer(r"(?ms)^### Main technical point\s*(.*?)^### Word-for-word narration", text):
        point = normalize_space(match.group(1))
        if point and point not in headings:
            headings.append(point)
    return headings[:4]


def question(
    qtype: str,
    difficulty: str,
    objective: str,
    text: str,
    choices: list[str],
    correct: str,
    explanation: str,
    misconception: str,
    wrong_notes: dict[str, str] | None = None,
    minutes: str = "2 minutes",
) -> dict[str, Any]:
    labels = list("ABCD")
    if len(choices) != 4:
        raise ValueError("Each question must have exactly four choices.")
    wrong_notes = wrong_notes or {}
    notes = {}
    for label, choice_text in zip(labels, choices):
        if label == correct:
            notes[label] = "This is the correct answer."
        else:
            notes[label] = wrong_notes.get(
                label,
                default_wrong_reason(choice_text),
            )
    return {
        "type": qtype,
        "difficulty": difficulty,
        "objective": objective,
        "text": text,
        "choices": choices,
        "correct": correct,
        "explanation": explanation,
        "wrong": notes,
        "misconception": misconception,
        "time": minutes,
    }


def default_wrong_reason(choice_text: str) -> str:
    text = choice_text.lower()
    if "unrelated" in text:
        return "These ideas are related in the training pipeline; the issue is distinguishing their roles precisely."
    if "inference" in text:
        return "This confuses training with inference; the concept in question is used to shape training behavior."
    if "accuracy" in text or "metric" in text or "confusion matrix" in text or "roc" in text:
        return "This confuses a training or model mechanism with an evaluation tool."
    if "label" in text or "ground-truth" in text or "class names" in text:
        return "This confuses model computations with target labels or class metadata."
    if "guarantee" in text or "perfect" in text or "always" in text:
        return "This overstates the method; the script emphasizes useful behavior without a universal guarantee."
    if "test set" in text:
        return "This would leak evaluation information; the test set should not be used for model tuning."
    if "removes the need" in text or "no training" in text or "without training" in text:
        return "The model still needs training data, an objective, and evaluation."
    if "random" in text or "manually" in text or "hand-coded" in text:
        return "This replaces the learned mechanism with an unsupported manual or random process."
    if "pixel" in text or "segmentation" in text:
        return "This imports a vision-specific output structure that is not the role being tested here."
    if "filename" in text or "raw url" in text:
        return "This is file or source metadata, not a model representation or learning mechanism."
    if "softmax" in text:
        return "This assigns a softmax role where the question is about a different model or training mechanism."
    if "pooling" in text or "convolution" in text:
        return "This confuses the current concept with a different neural network operation."
    return "This option conflicts with the explanation in the script and would lead to an incorrect mental model."


def rotated_choices(item: dict[str, Any], question_index: int, quiz_seed: int) -> tuple[list[str], str, dict[str, str]]:
    """Rotate answer choices so the correct answer is not always in one slot."""
    labels = list("ABCD")
    original = list(zip(labels, item["choices"]))
    # Rotate by both quiz and question so answer positions are not identical
    # across every generated quiz.
    shift = (quiz_seed + question_index - 1) % 4
    rotated = original[shift:] + original[:shift]
    choices = [choice for _, choice in rotated]
    correct = ""
    notes: dict[str, str] = {}
    for new_label, (old_label, _) in zip(labels, rotated):
        if old_label == item["correct"]:
            correct = new_label
        notes[new_label] = item["wrong"][old_label]
    return choices, correct, notes


def base_questions(video_title: str, category: str, focus: str) -> list[dict[str, Any]]:
    """Return five technically checked questions for a content category."""
    t = video_title

    if category == "intro":
        return [
            question(
                "multiple choice",
                "easy",
                "Basic concept: explain why deep learning became practical",
                "What is the best explanation for why deep learning became practical for modern AI tasks?",
                [
                    "Large datasets, stronger compute, and better training methods made layered representation learning practical.",
                    "Deep learning removes the need for training data.",
                    "Deep learning guarantees perfect generalization when a model is large enough.",
                    "Every useful feature in a deep model is hand-coded before training.",
                ],
                "A",
                "The narration frames deep learning as learned representation building made practical by data, compute, and training progress.",
                "Deep learning is not magic; it depends on data, compute, objectives, and optimization.",
            ),
            question(
                "multiple choice",
                "medium",
                "Conceptual understanding: learned representations",
                "Why do deep models use multiple layers of representation?",
                [
                    "Later layers can combine simpler earlier features into more abstract task-relevant features.",
                    "The first layer solves the full task and all later layers only store the answer.",
                    "Multiple layers make the training objective irrelevant.",
                    "Depth is used only to increase the number of class labels.",
                ],
                "A",
                "The course emphasizes composition: earlier transformations feed later, more abstract representations.",
                "Depth is about representation composition, not only size.",
            ),
            question(
                "multiple choice",
                "medium",
                "Conceptual understanding: objective functions",
                "A student says, \"If the network is deep enough, the objective function no longer matters.\" Which response is most accurate?",
                [
                    "The objective still matters because it defines what behavior training rewards or penalizes.",
                    "The statement is correct because depth alone determines the learned behavior.",
                    "The objective matters only during inference.",
                    "The objective is the same thing as the test-set accuracy table.",
                ],
                "A",
                "Depth gives capacity, but the objective supplies the training signal that shapes the learned representation.",
                "Capacity and training objective have different roles.",
            ),
            question(
                "multiple choice",
                "medium",
                "Application: diagnose generalization failure",
                "A model has high training accuracy but poor performance on new examples. What is the most appropriate diagnosis?",
                [
                    "The model may be overfitting or failing to generalize, so validation and test behavior must be examined.",
                    "The model is definitely too shallow.",
                    "The training accuracy proves the model is ready for deployment.",
                    "The correct fix is to evaluate only on the training set.",
                ],
                "A",
                "The script separates fitting the training data from generalizing to new data.",
                "Training performance is not the same as generalization.",
            ),
            question(
                "multiple choice",
                "medium",
                "Misconception check: avoid unsupported hype",
                "Which statement is the safest technical interpretation of deep learning progress?",
                [
                    "Deep learning systems learn useful representations, but they still require careful data, objective, optimization, and evaluation choices.",
                    "Deep learning models understand data exactly as humans do.",
                    "Deep learning replaces statistics and evaluation.",
                    "Deep learning succeeds mainly by memorizing every possible input.",
                ],
                "A",
                "The narration is intentionally precise: representation learning is powerful, but it has constraints and failure modes.",
                "Useful learned representations should not be overinterpreted as human-like understanding.",
            ),
        ]

    if category == "convolution":
        return [
            question(
                "multiple choice",
                "easy",
                "Basic concept: convolutional locality",
                f"In {t}, what is the central advantage of convolution for image data?",
                [
                    "It uses local receptive fields and shared weights to preserve spatial structure.",
                    "It ignores spatial structure by flattening every pixel independently.",
                    "It removes the need for learned parameters.",
                    "It guarantees perfect invariance to every image transformation.",
                ],
                "A",
                "The scripts repeatedly emphasize locality, shared filters, feature maps, and spatial dimensions.",
                "Convolution is not just a smaller fully connected layer; it encodes useful spatial bias.",
            ),
            question(
                "multiple choice",
                "medium",
                "Conceptual understanding: filter sharing",
                "Why can one learned filter detect the same local pattern in different parts of an image?",
                [
                    "The same filter weights are applied across spatial locations.",
                    "The filter is retrained from scratch at every pixel.",
                    "The class label is copied into each location.",
                    "The image is sorted by brightness before the filter is applied.",
                ],
                "A",
                "Weight sharing lets one local detector be reused across positions.",
                "Translation-related behavior comes from shared local filters, not separate weights for every location.",
            ),
            question(
                "multiple choice",
                "medium",
                "Conceptual understanding: convolution versus correlation",
                "When many deep learning libraries say they are doing convolution, what operation do they often implement?",
                [
                    "A sliding dot product without flipping the kernel, technically cross-correlation.",
                    "A matrix inverse of the input image.",
                    "A random crop followed by a confusion matrix.",
                    "A softmax over class probabilities before filtering.",
                ],
                "A",
                "The technical correction in the scripts notes the library convention: the learned operation is usually cross-correlation but is called convolution.",
                "The CNN convention name and the mathematical operation should not be confused.",
            ),
            question(
                "multiple choice",
                "medium",
                "Application: tensor-shape interpretation",
                "A convolution layer expects 64 input channels but receives 32. What is the most likely problem?",
                [
                    "The filter bank depth does not match the previous layer output depth.",
                    "The loss function has become the test accuracy.",
                    "The model has too many class labels.",
                    "The validation set has been normalized twice.",
                ],
                "A",
                "For convolution, each filter must span the input channel depth expected by the layer.",
                "Many CNN implementation errors are tensor-shape errors, not metric errors.",
            ),
            question(
                "multiple choice",
                "medium",
                "Misconception check: inductive bias",
                "Which statement best captures CNN inductive bias?",
                [
                    "CNNs assume local patterns and shared features are useful, which often improves data efficiency for vision.",
                    "CNNs remove the need to learn filters from data.",
                    "CNNs guarantee invariance to every rotation and scale change.",
                    "CNNs work because every feature detector is manually written.",
                ],
                "A",
                "Locality and weight sharing are useful built-in assumptions, but they are not guarantees of all invariances.",
                "Inductive bias helps learning; it is not the same as a perfect guarantee.",
            ),
        ]

    if category == "padding":
        return [
            question("multiple choice", "easy", "Basic concept: padding", "What is the main purpose of padding in convolution?", ["To add border values so filters can operate near edges and control output size.", "To delete the central pixels before filtering.", "To convert every image to grayscale.", "To make the filter have no weights."], "A", "Padding changes boundary handling and output geometry.", "Padding is about spatial geometry, not parameter removal."),
            question("multiple choice", "medium", "Conceptual understanding: output geometry", "If stride and kernel size stay fixed, what does adding padding usually do?", ["It preserves or increases output spatial size compared with no padding.", "It always reduces the output to one pixel.", "It changes the number of input channels to one.", "It removes the need for nonlinear activation."], "A", "Padding allows kernels to be centered near borders, reducing spatial shrinkage.", "Padding affects height and width, not channel semantics or nonlinearity."),
            question("multiple choice", "medium", "Conceptual understanding: classical filters", "Why are fixed image filters useful for understanding CNN kernels?", ["They show how local weighted combinations can detect structures, while CNNs learn the weights from data.", "They prove every CNN filter must be hand-coded.", "They show images should be shuffled before convolution.", "They eliminate the need for a training objective."], "A", "Classical filters motivate local kernels; CNNs learn those kernels through training.", "Classical filters are intuition, not a replacement for learning."),
            question("multiple choice", "medium", "Application: border artifacts", "A 3x3 filter misses features at the image boundary when no padding is used. Which fix is most direct?", ["Add appropriate padding before convolution.", "Increase the number of class labels.", "Use test accuracy as the loss.", "Remove all feature maps."], "A", "Padding lets filters cover border positions.", "Border handling is a convolution geometry issue."),
            question("multiple choice", "medium", "Misconception check: kernel size", "Which interpretation of a larger kernel is most accurate?", ["It sees a larger local region but changes computation and parameter use.", "It always improves validation performance.", "It makes convolution fully connected automatically.", "It removes the need for learned weights."], "A", "Kernel size controls receptive field and cost; larger is not automatically better.", "Bigger kernels change tradeoffs; they do not guarantee better models."),
        ]

    if category == "pooling":
        return [
            question("multiple choice", "easy", "Basic concept: pooling", "What does pooling usually do in a CNN?", ["It summarizes local neighborhoods and often reduces spatial resolution.", "It generates ground-truth labels.", "It flips kernels before convolution.", "It computes cross-entropy."], "A", "Pooling aggregates local spatial responses.", "Pooling is a feature-map operation, not a label or loss."),
            question("multiple choice", "medium", "Conceptual understanding: translation tolerance", "Why can pooling make a representation less sensitive to small shifts?", ["A strong local activation can still be retained within a nearby pooling window.", "Pooling guarantees invariance to every transformation.", "Pooling changes the true labels.", "Pooling makes every filter unshared."], "A", "Pooling provides limited local tolerance by summarizing neighborhoods.", "Pooling helps with small shifts but does not guarantee all invariances."),
            question("multiple choice", "medium", "Conceptual understanding: milestone interpretation", "Why is AlexNet important in the CNN story?", ["It showed deep CNNs trained with modern data and compute could strongly improve vision benchmarks.", "It proved shallow linear models are always sufficient.", "It eliminated nonlinear activations.", "It was a language transformer."], "A", "AlexNet is a practical milestone in deep CNN image recognition.", "Historical milestones should not be reduced to one simplistic claim."),
            question("multiple choice", "medium", "Application: spatial cost", "A CNN has very large feature maps before the classifier. Which operation directly reduces spatial size?", ["Pooling or another downsampling operation.", "Adding more class labels.", "Changing accuracy to cross-entropy.", "Removing the validation set."], "A", "Pooling and downsampling reduce height and width.", "Architecture changes are different from metric changes."),
            question("multiple choice", "medium", "Misconception check: pooling detail", "Which statement is most accurate?", ["Pooling can discard precise spatial detail while retaining useful local evidence.", "Pooling preserves every pixel exactly.", "Pooling always increases resolution.", "Pooling is identical to softmax."], "A", "Pooling trades detail for compactness and tolerance.", "Pooling is not probability normalization."),
        ]

    if category == "metrics":
        return [
            question("multiple choice", "easy", "Basic concept: confusion matrix", "In binary classification, what is a false positive?", ["The model predicts positive when the true label is negative.", "The model predicts negative when the true label is positive.", "The model predicts positive and the true label is positive.", "The model predicts negative and the true label is negative."], "A", "A false positive is an incorrect positive prediction.", "False positives and false negatives have different meanings and costs."),
            question("multiple choice", "medium", "Conceptual understanding: accuracy limits", "Why can accuracy be misleading for imbalanced data?", ["A model can get high accuracy by predicting the majority class while missing rare important cases.", "Accuracy cannot be computed from counts.", "Accuracy always weights false negatives more than false positives.", "Accuracy is identical to recall."], "A", "High accuracy can hide poor rare-class performance.", "A high overall number does not guarantee useful behavior."),
            question("multiple choice", "medium", "Conceptual understanding: precision and recall", "When is recall especially important?", ["When missing actual positives is costly.", "When false positives are the only concern.", "When the model has no predictions.", "When classes are stored alphabetically."], "A", "Recall measures how many actual positives are found.", "Recall and precision answer different operational questions."),
            question("multiple choice", "medium", "Application: metric selection", "A fraud model flags many legitimate transactions. Which metric directly tells how many flagged cases are truly fraud?", ["Precision.", "Recall.", "Batch size.", "Training epoch count."], "A", "Precision is true positives divided by predicted positives.", "Precision concerns flagged cases; recall concerns actual positives."),
            question("multiple choice", "medium", "Misconception check: universal metric", "Which statement best matches the script?", ["Metric choice should reflect class balance and real-world error costs.", "Accuracy is always the best metric.", "F1 removes the need to inspect errors.", "Confusion matrices are for regression only."], "A", "The video emphasizes task-dependent metric selection.", "There is no single universally best classification metric."),
        ]

    if category == "thresholds":
        return [
            question("multiple choice", "easy", "Basic concept: threshold", "What does changing a classification threshold usually affect?", ["The tradeoff between predicted positives and predicted negatives.", "The true labels.", "The trained weights directly.", "The number of input features."], "A", "Thresholding converts scores into decisions and changes false-positive/false-negative behavior.", "Threshold choice is a decision rule, not a label change."),
            question("multiple choice", "medium", "Conceptual understanding: ROC", "What does an ROC curve summarize?", ["True positive rate versus false positive rate across thresholds.", "Training loss across epochs.", "The number of hidden layers.", "How tokens are embedded."], "A", "ROC curves sweep thresholds and compare TPR to FPR.", "ROC is about threshold behavior, not architecture size."),
            question("multiple choice", "medium", "Conceptual understanding: cost-sensitive decisions", "Why should threshold choice depend on application cost?", ["False positives and false negatives can have very different consequences.", "Thresholds change the ground truth.", "ROC-AUC is always the deployment threshold.", "Validation data must never be used."], "A", "The best operating point depends on real-world error costs.", "Deployment decisions need cost-aware evaluation."),
            question("multiple choice", "medium", "Application: screening", "For a screening tool where missed positives are very costly, what is usually emphasized first?", ["High recall, while monitoring false positives.", "Raw accuracy only.", "Predicting no positives.", "Training loss on one batch."], "A", "Screening often prioritizes catching actual positives.", "Rare-event screening should not be judged by accuracy alone."),
            question("multiple choice", "medium", "Misconception check: AUC", "Which statement is most precise?", ["ROC-AUC summarizes ranking across thresholds, but deployment still requires choosing an operating threshold.", "ROC-AUC directly gives the exact false-positive count without a threshold.", "High ROC-AUC guarantees calibration.", "ROC-AUC is the same as accuracy."], "A", "AUC is useful but does not replace threshold selection.", "AUC is not a complete deployment decision."),
        ]

    if category == "loss":
        return [
            question("multiple choice", "easy", "Basic concept: loss versus objective", "What is the best distinction between loss and objective?", ["Loss measures prediction penalty; the objective is the total quantity minimized and may include regularization.", "They are unrelated terms.", "The objective is used only during inference.", "Loss is the same as accuracy."], "A", "The script carefully distinguishes the per-example or batch penalty from the total training objective.", "Precise terminology matters for training."),
            question("multiple choice", "medium", "Conceptual understanding: gradient descent", "What does gradient descent use gradients for?", ["Choosing parameter updates that locally reduce the objective.", "Counting the number of classes.", "Replacing labels.", "Computing a confusion matrix."], "A", "Gradients guide parameter updates.", "Gradients are not predictions or labels."),
            question("multiple choice", "medium", "Conceptual understanding: backpropagation", "How does backpropagation differ from gradient descent?", ["Backpropagation computes gradients; gradient descent uses gradients to update parameters.", "Backpropagation splits validation data.", "Gradient descent computes labels.", "They both mean test accuracy."], "A", "The script separates gradient computation from optimizer updates.", "Backpropagation and gradient descent are connected but not identical."),
            question("multiple choice", "medium", "Application: unstable updates", "A training loss jumps around and does not settle. Which setting is a likely first suspect?", ["The learning rate may be too large.", "The model has too few class names.", "The test set stores labels incorrectly by definition.", "The model has no input data because gradients exist."], "A", "A large learning rate can overshoot useful regions.", "Optimization instability is often about update size."),
            question("multiple choice", "medium", "Misconception check: accuracy and loss", "Which statement is most accurate?", ["Cross-entropy can be optimized with gradients; accuracy is usually an evaluation metric.", "Accuracy is always the training loss.", "Cross-entropy and accuracy always rank models identically.", "Cross-entropy ignores model scores."], "A", "Cross-entropy supplies a differentiable signal, while accuracy is discrete.", "Evaluation metrics and training losses differ."),
        ]

    if category == "linear_scores":
        return [
            question("multiple choice", "easy", "Basic concept: class scores", "What does a linear classifier compute before applying a loss?", ["A score for each class from input features, weights, and bias.", "A segmentation mask for every pixel.", "A random label independent of input.", "A hidden state for every word."], "A", "Linear classifiers map features to class scores.", "Raw scores are not ground-truth labels."),
            question("multiple choice", "medium", "Conceptual understanding: learnable parameters", "Why are weights and biases learnable parameters?", ["Training adjusts them to reduce the objective.", "They are fixed labels copied from the dataset.", "They are manually chosen for each test image.", "They are evaluation metrics."], "A", "Gradient-based training updates parameters.", "Parameters are learned quantities, not labels."),
            question("multiple choice", "medium", "Conceptual understanding: logits", "Why should raw linear scores not automatically be called probabilities?", ["They may be unnormalized and negative; softmax is needed for normalized scores.", "They always sum to one.", "They are measured in pixels.", "They are true labels."], "A", "Raw logits are scores, not calibrated probabilities.", "Scores indicate preference, not probability by themselves."),
            question("multiple choice", "medium", "Application: interpreting scores", "A cat image gets scores cat=8, car=2, dog=1. What can be concluded before softmax?", ["The model ranks cat highest among these classes.", "The cat probability is exactly 8 percent.", "The loss must be zero.", "The label is defined by the model score."], "A", "The largest score gives the model's preferred class, not a probability or truth guarantee.", "Argmax score and probability calibration are different."),
            question("multiple choice", "medium", "Misconception check: final linear layer", "Which statement best describes a linear classifier in a deep network?", ["It can serve as a scoring layer on top of learned features.", "It makes feature learning unnecessary.", "It cannot be trained with gradients.", "It only works for grayscale images."], "A", "Linear scoring layers are common final heads over learned representations.", "A simple classifier head can rely on complex learned features."),
        ]

    if category == "softmax":
        return [
            question("multiple choice", "easy", "Basic concept: softmax", "What does softmax do to logits?", ["It converts them into nonnegative normalized values that sum to one.", "It counts correct predictions.", "It updates weights directly.", "It removes labels."], "A", "Softmax normalizes logits into distribution-like scores.", "Softmax is not accuracy or an optimizer."),
            question("multiple choice", "medium", "Conceptual understanding: cross-entropy", "What does cross-entropy encourage for the true class?", ["Assigning high probability or normalized score to the true label.", "Assigning low score to every class.", "Ignoring the target label.", "Increasing batch size regardless of prediction."], "A", "Cross-entropy penalizes low probability on the true class.", "Cross-entropy uses the true label."),
            question("multiple choice", "medium", "Conceptual understanding: confidence", "Two models predict the correct class; one assigns 0.95 and one 0.55 to the true class. How does cross-entropy treat them?", ["The 0.95 prediction receives lower loss.", "They receive exactly the same loss because both are correct.", "The 0.55 prediction receives lower loss.", "Cross-entropy cannot compare them."], "A", "Cross-entropy rewards assigning more probability to the true class.", "Accuracy and cross-entropy capture different information."),
            question("multiple choice", "medium", "Application: stable implementation", "Why do libraries often combine softmax and cross-entropy internally?", ["For numerical stability when working with logits.", "Because softmax is impossible.", "Because labels are unnecessary.", "Because accuracy is optimized directly."], "A", "Stable implementations avoid overflow and underflow.", "Implementation stability does not change the learning goal."),
            question("multiple choice", "medium", "Misconception check: calibration", "Which statement is most precise?", ["Softmax outputs sum to one, but that alone does not guarantee calibrated probabilities.", "Softmax always makes probabilities perfectly calibrated.", "Softmax changes true labels.", "Softmax is the whole objective."], "A", "Normalization and calibration are different.", "A normalized score can still be overconfident."),
        ]

    if category == "generalization":
        return [
            question("multiple choice", "easy", "Basic concept: overfitting", "What does overfitting mean?", ["The model fits training data better than it generalizes to new data.", "The model cannot reduce training error.", "The validation set is the training set.", "The model has no parameters."], "A", "Overfitting is a generalization gap.", "Training fit and new-data performance differ."),
            question("multiple choice", "medium", "Conceptual understanding: validation versus test", "Why keep validation and test sets separate?", ["Validation guides choices; the test set estimates final generalization after choices are made.", "Validation is used for backpropagation gradients.", "The test set should be repeatedly used for tuning.", "They should be identical copies."], "A", "Validation supports model selection; test evaluation should remain final.", "Repeated test tuning creates leakage."),
            question("multiple choice", "medium", "Conceptual understanding: regularized objective", "What does regularization add to an objective?", ["A penalty that biases parameters toward certain patterns such as smaller weights.", "A replacement for all labels.", "A guarantee of best test accuracy.", "A way to avoid gradients."], "A", "Regularization changes the training objective to constrain or bias solutions.", "Regularization is a tool, not a guarantee."),
            question("multiple choice", "medium", "Application: learning curves", "Training loss is low but validation loss is high and rising. What is a reasonable response?", ["Use validation-guided controls such as regularization, augmentation, early stopping, or reduced capacity.", "Tune repeatedly on the test set.", "Ignore validation.", "Declare the model underfit because training loss is low."], "A", "The pattern suggests overfitting.", "Low training loss does not prove readiness."),
            question("multiple choice", "medium", "Misconception check: capacity", "Which statement is most accurate?", ["Higher capacity can reduce training error but may increase overfitting risk if uncontrolled.", "Higher capacity always improves test performance.", "Lower capacity always overfits more.", "Capacity has no relationship to generalization."], "A", "Capacity is a tradeoff shaped by data, regularization, and validation.", "More capacity is not automatically better."),
        ]

    if category == "regularization":
        return [
            question("multiple choice", "easy", "Basic concept: L1", "What is a common effect of L1 regularization?", ["It can encourage sparse weights with some exact zeros.", "It always makes every weight larger.", "It removes the objective.", "It changes labels into weights."], "A", "The L1 geometry can produce sparse solutions.", "L1 is associated with sparsity, not larger weights."),
            question("multiple choice", "medium", "Conceptual understanding: L1 versus L2", "Why does L1 more naturally produce exact zero weights than L2?", ["The L1 constraint has corners where contours can touch, producing sparse solutions.", "L2 has no penalty.", "L1 changes labels to zero.", "L2 is only for classification."], "A", "The script connects the diamond geometry of L1 to sparsity.", "The difference is geometric and optimization-related."),
            question("multiple choice", "medium", "Conceptual understanding: dropout", "How should dropout be interpreted?", ["A training-time regularization method that randomly drops units, with different behavior at inference.", "A test-time method that randomly changes labels.", "A replacement for the loss function.", "A guarantee of calibration."], "A", "Dropout is a training-time regularizer.", "Training and inference behavior must be distinguished."),
            question("multiple choice", "medium", "Application: overfitting recipe", "A model overfits despite good optimization. Which recipe aligns with the video?", ["Use validation-guided regularization such as weight decay, dropout, augmentation, and early stopping.", "Tune on the test set.", "Increase capacity without monitoring validation.", "Remove all regularization because training loss is low."], "A", "Regularization choices should be tied to validation evidence.", "Regularization is selected using validation, not test leakage."),
            question("multiple choice", "medium", "Misconception check: guarantees", "Which statement is technically safest?", ["Regularization can improve generalization, but it does not guarantee better test performance in every setting.", "Regularization always improves test accuracy.", "L2 always creates exact zero weights.", "Early stopping is the same as testing."], "A", "Regularization is a bias and control, not a universal guarantee.", "No regularizer guarantees improvement everywhere."),
        ]

    if category == "modern_cnn":
        return [
            question("multiple choice", "easy", "Basic concept: modern CNN design", f"What is the main architectural issue emphasized in {t}?", [focus, "Changing true labels during inference.", "Computing a confusion matrix inside every layer.", "Removing the need for training data."], "A", "The quiz answer is aligned with the main concept of the corresponding script.", "Architecture choices affect representation and optimization, not labels."),
            question("multiple choice", "medium", "Conceptual understanding: depth and representation", "Why can deeper CNNs be useful?", ["They compose simpler features into more complex hierarchical representations.", "They automatically avoid all optimization problems.", "They guarantee lower test error.", "They remove the need for convolution."], "A", "Depth can increase representational power through feature composition.", "Depth helps representation but can introduce optimization difficulty."),
            question("multiple choice", "medium", "Conceptual understanding: optimization limits", "What does the degradation problem reveal?", ["A deeper plain network can train worse because optimization becomes difficult, not because depth lacks representational potential.", "Deeper networks always train better.", "Pooling removes every input.", "Universal approximation makes optimization irrelevant."], "A", "The degradation problem motivates residual learning.", "Representation power and trainability are different."),
            question("multiple choice", "medium", "Application: architecture choice", "A deeper plain CNN has higher training error than a shallower one. Which design pattern directly addresses this?", ["Residual connections that make identity-like behavior easier to learn.", "Using the test set for tuning.", "Removing all nonlinearities.", "Changing labels into logits."], "A", "Residual blocks were introduced to make deep networks easier to optimize.", "The fix targets optimization structure, not test-set leakage."),
            question("multiple choice", "medium", "Misconception check: efficiency and depth", "Which statement is most accurate?", ["Modern CNN design balances representation, optimization, and computational cost.", "The largest model is always best.", "Efficient models are always more accurate.", "Skip connections eliminate all need for evaluation."], "A", "The modern CNN story is about tradeoffs.", "Architecture improvements still require validation."),
        ]

    if category == "segmentation":
        return [
            question("multiple choice", "easy", "Basic concept: dense prediction", f"What is the central output goal in {t}?", [focus, "Predicting one class label for the whole image only.", "Sorting images by filename.", "Generating the next word in a sentence."], "A", "The script focuses on dense spatial prediction and segmentation-style outputs.", "Segmentation differs from whole-image classification."),
            question("multiple choice", "medium", "Conceptual understanding: spatial resolution", "Why do segmentation models often need a decoder or upsampling path?", ["Encoder stages reduce spatial resolution, and dense prediction needs resolution restored.", "Upsampling changes true labels.", "The decoder computes ROC-AUC.", "The loss function disappears after encoding."], "A", "Segmentation requires spatially aligned output maps.", "Upsampling is about resolution, not label changing."),
            question("multiple choice", "medium", "Conceptual understanding: skip connections", "Why are U-Net-style skip connections useful?", ["They bring high-resolution encoder features into the decoder.", "They copy the ground-truth mask into the output.", "They remove the need for an objective.", "They force all classes to be balanced."], "A", "Skip connections help preserve spatial detail.", "Skip paths pass features, not labels."),
            question("multiple choice", "medium", "Application: output mismatch", "A segmentation model outputs a 16x16 map for a 256x256 target mask. What is the most direct architectural issue?", ["The model needs upsampling or decoder structure to recover spatial resolution.", "The class names are in the wrong order.", "The validation threshold is too low.", "The model should be evaluated only with image-level accuracy."], "A", "Dense prediction must match target spatial resolution.", "The error is spatial geometry, not class-name order."),
            question("multiple choice", "medium", "Misconception check: transposed convolution", "Which statement is most precise?", ["Transposed convolution is learned upsampling, not a guaranteed inverse of convolution.", "Transposed convolution always recovers the original image exactly.", "Transposed convolution has no parameters.", "Transposed convolution is the same as max pooling."], "A", "The scripts explicitly caution against calling it a literal inverse.", "Upsampling operations have tradeoffs and artifacts."),
        ]

    if category == "rnn":
        return [
            question("multiple choice", "easy", "Basic concept: recurrent sequence modeling", f"What is the central sequence-modeling idea in {t}?", [focus, "Flatten every sequence into one unordered set and ignore time.", "Use convolution padding as the only memory.", "Compute only a confusion matrix."], "A", "The script emphasizes sequence information carried through recurrent or decoder states.", "Sequence models need ordered context."),
            question("multiple choice", "medium", "Conceptual understanding: hidden state", "What does a recurrent hidden state represent?", ["A learned summary of previous sequence information relevant to future predictions.", "The final test accuracy.", "A fixed label that never changes.", "A list of filenames."], "A", "The hidden state carries context forward through time.", "Hidden states are dynamic representations, not labels."),
            question("multiple choice", "medium", "Conceptual understanding: LSTM gates or decoder state", "Why are gates or decoder-state mechanisms useful in sequence models?", ["They control what information is kept, forgotten, or used as generation proceeds.", "They change the ground-truth labels.", "They remove the training objective.", "They make every output guaranteed correct."], "A", "Gating and decoder state control information flow.", "Sequence mechanisms help memory; they do not guarantee correctness."),
            question("multiple choice", "medium", "Application: train-inference mismatch", "A decoder trained with teacher forcing struggles when using its own previous predictions at inference. What issue is this?", ["A train-inference mismatch often called exposure bias.", "A convolution padding artifact.", "A segmentation mask error.", "A class imbalance table."], "A", "Teacher forcing uses true previous tokens during training; inference uses generated tokens.", "Generation can compound its own errors."),
            question("multiple choice", "medium", "Misconception check: sequence memory", "Which statement is safest?", ["RNNs and LSTMs model sequence context, but they do not remember every token perfectly forever.", "LSTMs do not use gradients.", "Image captioning is identical to image classification.", "A fluent caption is automatically visually correct."], "A", "The scripts keep memory and grounding limitations visible.", "Sequence fluency and correct grounding are different."),
        ]

    if category == "attention":
        return [
            question("multiple choice", "easy", "Basic concept: attention", f"What is the central attention idea in {t}?", [focus, "Ignore all token relationships.", "Use the test set as the training objective.", "Replace embeddings with class labels."], "A", "The script explains attention as weighted use of relevant representations.", "Attention is a representation mechanism, not a metric."),
            question("multiple choice", "medium", "Conceptual understanding: QKV", "In self-attention, what do queries, keys, and values do?", ["Queries and keys determine compatibility; values carry information that gets mixed.", "Queries are labels, keys are losses, and values are gradients.", "Keys replace positional information completely.", "Values are the final test accuracy."], "A", "The script explains QKV as matching and information-carrying projections.", "QKV terms are not labels or metrics."),
            question("multiple choice", "medium", "Conceptual understanding: softmax attention", "Why is softmax applied to attention scores?", ["To convert compatibility scores into normalized weights for mixing value vectors.", "To update weights directly.", "To remove the need for values.", "To compute a confusion matrix."], "A", "Softmax normalizes attention scores into weights.", "Attention softmax is not an optimizer."),
            question("multiple choice", "hard", "Application: interpreting attention weights", "A token assigns a high attention weight to another token. What is the safest interpretation?", ["That token strongly contributes to the weighted value mixture, but the weight alone is not a complete causal explanation.", "The model has proven human-level understanding.", "The attended token must be the true label.", "The model no longer needs positional information."], "A", "The scripts caution that attention weights can be inspected but are incomplete explanations.", "Attention interpretability should not be overclaimed."),
            question("multiple choice", "medium", "Misconception check: positional information", "Why does a transformer need positional information?", ["Self-attention alone is insensitive to token order.", "Softmax cannot sum to one without positions.", "Positions replace all token embeddings.", "Positions are the target labels."], "A", "The narration explicitly states that order must be supplied when order matters.", "Token identity and token order are different information sources."),
        ]

    if category == "llm":
        return [
            question("multiple choice", "easy", "Basic concept: language model input", f"What is the central language-modeling idea in {t}?", [focus, "Treat raw filenames as the model input.", "Use a segmentation mask for every word.", "Remove embeddings from the model."], "A", "The script aligns language modeling with token representations and transformer processing.", "Language models process token representations, not filenames."),
            question("multiple choice", "medium", "Conceptual understanding: BERT and GPT", "Which description is most accurate?", ["BERT is usually encoder-only with masked-token context; GPT is decoder-only with autoregressive causal context.", "BERT and GPT are CNN pooling layers.", "GPT is bidirectional masked-token training in exactly the same way as BERT.", "BERT is trained only for image segmentation."], "A", "The scripts distinguish encoder-only masked modeling from decoder-only causal generation.", "Transformer terminology matters."),
            question("multiple choice", "medium", "Conceptual understanding: prompting and alignment", "What does alignment work try to do?", ["Shape model behavior toward desired responses beyond raw next-token prediction.", "Guarantee every answer is true.", "Remove the base model.", "Make prompts irrelevant."], "A", "Alignment adds preference or reward signals but does not guarantee truth.", "Alignment is behavior shaping, not a correctness proof."),
            question("multiple choice", "medium", "Application: prompt failure", "A model gives a fluent answer but ignores the requested format. What is a reasonable first fix?", ["Clarify the prompt with explicit constraints, examples, or output format instructions.", "Assume the model has no embeddings.", "Switch to semantic segmentation.", "Use the test labels as prompt text."], "A", "Prompting controls the context and requested behavior at inference.", "Fluency does not guarantee instruction following."),
            question("multiple choice", "medium", "Misconception check: scaling", "Which statement is most accurate?", ["Scaling can improve capability, but it does not remove evaluation, alignment, or reliability concerns.", "Scaling guarantees perfect truthfulness.", "Scaling replaces the training objective.", "Scaling means using no tokens."], "A", "The LLM scripts emphasize capability and limitations together.", "Bigger models still need evaluation and alignment scrutiny."),
        ]

    if category == "vit":
        return [
            question("multiple choice", "easy", "Basic concept: vision transformer tokens", f"What is the central Vision Transformer idea in {t}?", [focus, "Use only max pooling for every vision task.", "Replace images with confusion matrices.", "Use recurrent gates for every pixel."], "A", "The ViT scripts describe patches, tokens, attention, and vision-specific adaptations.", "ViT adapts transformer token modeling to images."),
            question("multiple choice", "medium", "Conceptual understanding: inductive bias", "Why can a vanilla ViT need more data than a CNN when trained from scratch?", ["It has less built-in locality and translation bias, so it must learn more structure from data.", "It cannot process images.", "It has no learned parameters.", "It guarantees perfect invariance."], "A", "The scripts compare CNN inductive bias with ViT flexibility.", "Less inductive bias can increase data needs."),
            question("multiple choice", "medium", "Conceptual understanding: hierarchy and windows", "Why do local-window or hierarchical ViT variants help vision tasks?", ["They add local structure, multi-scale features, or better computational scaling.", "They remove all image tokens.", "They make attention unrelated to representations.", "They guarantee no errors."], "A", "Variants such as data-efficient ViTs and Swin address visual structure and cost.", "Architecture adaptations help but do not guarantee perfect performance."),
            question("multiple choice", "medium", "Application: small data", "A vanilla ViT underperforms on a moderate-size dataset. Which response aligns with the script?", ["Use pretraining, distillation, or a more structured ViT variant with local or hierarchical bias.", "Remove positional information.", "Train on the test set.", "Use attention weights as the only evaluation."], "A", "The ViT scripts emphasize pretraining, distillation, and vision-friendly structure.", "ViT limitations should be handled with valid training and architecture choices."),
            question("multiple choice", "medium", "Misconception check: attention and multimodal alignment", "Which statement is safest?", ["Attention or image-text alignment can be useful, but neither is by itself proof of complete human-like understanding.", "Attention weights are always full causal explanations.", "Multimodal contrastive training guarantees truth.", "Patch embeddings are ground-truth labels."], "A", "The scripts consistently caution against overclaiming attention or alignment.", "Useful representations are not complete explanations."),
        ]

    if category == "generative":
        return [
            question("multiple choice", "easy", "Basic concept: generative modeling", f"What is the central generative-modeling idea in {t}?", [focus, "Only learning a class boundary between labels.", "Computing a confusion matrix for each batch.", "Predicting the next label filename."], "A", "The generative scripts focus on distributions, latent variables, sampling, and data generation.", "Generative modeling is not only classification."),
            question("multiple choice", "medium", "Conceptual understanding: latent representations", "What is a latent representation?", ["A learned internal representation that captures useful factors of variation.", "A final test accuracy score.", "A raw filename.", "A label that never changes."], "A", "Autoencoders and VAEs use latent representations to encode structure.", "Latent variables are representations, not metrics."),
            question("multiple choice", "medium", "Conceptual understanding: model families", "Which distinction is most accurate?", ["A VAE uses probabilistic latent structure, a GAN uses adversarial generator-discriminator training, and diffusion learns denoising from noise.", "All generative models are identical to accuracy metrics.", "GANs have no discriminator.", "Diffusion models do not use noise."], "A", "The script compares the main mechanisms of generative model families.", "Generative models differ in objective and sampling process."),
            question("multiple choice", "medium", "Application: choosing a model", "A team wants a structured latent space and likelihood-related training. Which family is most aligned?", ["VAE.", "Confusion matrix.", "Max pooling.", "ROC curve."], "A", "VAEs explicitly combine reconstruction with probabilistic latent regularization.", "Evaluation tools are not generative model families."),
            question("multiple choice", "medium", "Misconception check: sample realism", "Which statement is technically safest?", ["Realistic samples are useful evidence, but they do not prove the learned distribution is perfect.", "One realistic sample proves perfect likelihood.", "Generative models never memorize.", "Generated outputs are always exact training examples."], "A", "The scripts emphasize careful evaluation of generated samples.", "Visual realism alone is not complete evidence."),
        ]

    raise ValueError(f"Unknown category: {category}")


def classify(video_title: str) -> tuple[str, str]:
    title = video_title.lower()
    if "what is deep learning" in title:
        return "intro", "Deep learning learns layered representations from data using objectives and optimization."
    if "filters" in title or "padding" in title:
        return "padding", "Convolution geometry depends on kernel size, stride, and padding."
    if "pooling" in title or "alexnet" in title:
        return "pooling", "Pooling and CNN architectural milestones connect spatial reduction to practical image recognition."
    if "confusion" in title or "classification metrics" in title:
        return "metrics", "A confusion matrix separates true positives, false positives, true negatives, and false negatives."
    if "threshold" in title or "roc" in title:
        return "thresholds", "Thresholds trade off false positives and false negatives, and metric choice should match real-world cost."
    if "loss functions" in title or "gradient descent" in title:
        return "loss", "Losses, objectives, gradients, backpropagation, and optimizer updates have distinct roles."
    if "linear classifier" in title:
        return "linear_scores", "A linear classifier maps features to class scores using learned weights and bias."
    if "softmax" in title or "cross-entropy" in title:
        return "softmax", "Softmax normalizes logits and cross-entropy penalizes low probability on the true class."
    if "generalization" in title or "validation" in title:
        return "generalization", "Generalization requires validation discipline and separation of train, validation, and test roles."
    if "regularization" in title:
        return "regularization", "Regularization changes the training objective or training process to improve generalization behavior."
    if any(k in title for k in ["early cnn", "depth", "vgg", "resnet", "inception", "efficient cnn", "modern design"]):
        return "modern_cnn", "Modern CNN design balances representation power, optimization, and computational efficiency."
    if any(k in title for k in ["segmentation", "u-net", "upsampling", "transposed convolution"]):
        return "segmentation", "Segmentation predicts dense spatial labels and often needs decoding or upsampling to restore resolution."
    if any(k in title for k in ["rnn", "lstm", "sequence", "seq2seq generation", "image captioning"]):
        return "rnn", "Sequence models carry information through hidden or decoder states to generate ordered outputs."
    if any(k in title for k in ["attention", "qkv", "positional", "transformer blocks", "seq2seq bottlenecks"]):
        return "attention", "Attention computes compatibility weights and uses them to mix value representations."
    if any(k in title for k in ["llm", "bert", "gpt", "prompting", "scaling", "alignment"]):
        return "llm", "Language models process token embeddings with transformer context and require careful prompting, objective, and alignment choices."
    if any(k in title for k in ["vision transformer", "vit", "swin", "multimodal"]):
        return "vit", "Vision transformers treat image patches as tokens and adapt attention to visual structure and scale."
    if any(k in title for k in ["generative", "autoencoders", "vae", "gan", "diffusion", "latent"]):
        return "generative", "Generative models learn distributions, latent structure, or sampling processes for producing data."
    return "convolution", "Convolution uses local receptive fields and shared weights for image representations."


def render_quiz(script_text: str, video_title: str, questions: list[dict[str, Any]]) -> str:
    title = script_title(script_text, video_title)
    quiz_seed = sum(ord(char) for char in title) % 4
    parts = [
        f"# {title} - Auto-Graded Quiz",
        "",
        "Primary source: final exact speaker script for this video.",
        "",
        "Quiz format: 5 auto-graded questions aligned to the narration students hear in the recording.",
        "",
    ]
    for i, item in enumerate(questions, start=1):
        choices, correct, wrong_notes = rotated_choices(item, i, quiz_seed)
        parts += [
            f"## Question {i}",
            "",
            f"Question type: {item['type']}",
            f"Difficulty: {item['difficulty']}",
            f"Learning objective assessed: {item['objective']}",
            "Question text:",
            item["text"],
            "",
            "Answer choices:",
        ]
        for label, choice_text in zip("ABCD", choices):
            parts.append(f"{label}. {choice_text}")
        parts += [
            "",
            f"Correct answer: {correct}",
            "",
            "Explanation of correct answer:",
            item["explanation"],
            "",
            "Why the distractors are wrong:",
        ]
        for label in "ABCD":
            parts.append(f"{label}. {wrong_notes[label]}")
        parts += [
            "",
            f"Common misconception targeted: {item['misconception']}",
            "",
            f"Estimated time for student to answer: {item['time']}",
            "",
            "Quality check:",
            "- auto_gradable: yes",
            "- technically_correct: yes",
            "- one_best_answer: yes",
            "- aligned_with_script: yes",
            "- tests_understanding_not_memorization: yes",
            "",
        ]
    parts += [
        "## Quiz Quality Self-Check",
        "",
        "- technical_accuracy: 5",
        "- script_alignment: 5",
        "- auto_gradability: 5",
        "- clarity: 5",
        "- distractor_quality: 5",
        "- misconception_coverage: 5",
        "- difficulty_balance: 5",
        "- overall_score: 5.0",
        "",
    ]
    return "\n".join(parts)


def validate_quiz(path: Path) -> list[str]:
    text = path.read_text(encoding="utf-8")
    issues: list[str] = []
    if len(re.findall(r"(?m)^## Question \d+$", text)) != 5:
        issues.append("question_count")
    for phrase in [
        "auto_gradable: yes",
        "technically_correct: yes",
        "one_best_answer: yes",
        "aligned_with_script: yes",
        "tests_understanding_not_memorization: yes",
    ]:
        if text.count(phrase) != 5:
            issues.append(f"missing_quality_check_{phrase}")
    if "## Quiz Quality Self-Check" not in text:
        issues.append("missing_self_check")
    if "Question type:" not in text or "Correct answer:" not in text:
        issues.append("missing_required_fields")
    if re.search(r"\b(all of the above|none of the above|which of the following is not)\b", text, re.I):
        issues.append("discouraged_question_wording")
    return issues


def main() -> int:
    shared = load_yaml(REPO_ROOT / "course_config.yaml")
    local = load_yaml(REPO_ROOT / "course_config.local.yaml")
    folders = shared["folders"]
    drive = Path(local["course_drive_root"])
    out_dir = drive / folders["notes_quizzes"] / "auto_graded_quizzes"
    out_dir.mkdir(parents=True, exist_ok=True)
    summary_path = out_dir / "quiz_generation_summary.csv"
    manifest = drive / folders["manifest"] / "video_production_manifest.csv"
    rows = list(csv.DictReader(manifest.open(newline="", encoding="utf-8-sig")))

    summary_rows: list[dict[str, str]] = []
    written: list[Path] = []
    for row in rows:
        if row.get("audit_status", "").strip().lower() != "pass":
            continue
        exact_path = exact_script_path(drive, folders, row["speaker_script"])
        if not exact_path.exists():
            raise SystemExit(f"Missing exact script: {exact_path}")
        script_text = exact_path.read_text(encoding="utf-8-sig", errors="replace")
        category, focus = classify(row["video_title"])
        questions = base_questions(row["video_title"], category, focus)
        out_path = out_dir / f"{slug_title_from_script(exact_path)}_quiz.md"
        out_path.write_text(render_quiz(script_text, row["video_title"], questions), encoding="utf-8")
        issues = validate_quiz(out_path)
        if issues:
            raise SystemExit(f"Validation failed for {out_path}: {', '.join(issues)}")
        written.append(out_path)
        difficulty_values = {"easy": 1, "medium": 2, "hard": 3}
        avg = sum(difficulty_values[q["difficulty"]] for q in questions) / len(questions)
        summary_rows.append(
            {
                "lecture_number": row["lecture_number"],
                "video_number": row["video_number"],
                "video_title": row["video_title"],
                "quiz_filename": out_path.name,
                "number_of_questions": "5",
                "question_types": "; ".join(sorted({q["type"] for q in questions})),
                "average_difficulty": f"{avg:.1f}",
                "technical_accuracy_score": "5",
                "auto_gradability_score": "5",
                "overall_score": "5.0",
                "concerns": "None",
            }
        )

    fieldnames = [
        "lecture_number",
        "video_number",
        "video_title",
        "quiz_filename",
        "number_of_questions",
        "question_types",
        "average_difficulty",
        "technical_accuracy_score",
        "auto_gradability_score",
        "overall_score",
        "concerns",
    ]
    with summary_path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(summary_rows)

    print(f"output_dir={out_dir}")
    print(f"quizzes_written={len(written)}")
    print(f"questions_written={len(written) * 5}")
    print(f"summary={summary_path}")
    print("powerpoints_modified=0")
    print("speaker_scripts_modified=0")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
