# Codex Instructions for Deep Learning Course Automation

## Goal

Help convert in-person deep learning lecture slides into online 30-minute video modules.

## Rules

- Do not overwrite original PowerPoint files.
- Original slides are stored in Google Drive.
- Revised slides should be written only to the revised slides folder.
- Preserve technical correctness.
- Preserve the original PowerPoint theme unless asked otherwise.
- Keep slide text concise.
- Put detailed explanations in speaker notes.
- Each 30-minute video should have:
  - a clear title
  - 2–4 learning objectives
  - coherent slide sequence
  - speaker notes
  - one recap slide
  - one quiz/checkpoint slide

## Workflow

1. Analyze the original lecture deck.
2. Create a slide inventory.
3. Estimate teaching time.
4. Segment the lecture into 30-minute videos.
5. Suggest slide revisions.
6. Generate revised PowerPoint files.
7. Update the manifest.

## Safety rule

Never edit the same PowerPoint file from two machines at the same time.

## Local Google Drive path

The Google Drive course folder path is machine-specific.

Read shared course settings from:

- `course_config.yaml`

Read the local Google Drive root from:

- `course_config.local.yaml`

Do not commit `course_config.local.yaml` to GitHub.

When writing scripts, combine `course_drive_root` from `course_config.local.yaml` with the folder names from `course_config.yaml`.

Do not hardcode drive letters such as `H:\` or `G:\` inside scripts.