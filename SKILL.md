---
name: sketch-to-slide
description: Convert a photo of a hand-drawn slide, whiteboard diagram, or paper sketch into an editable 16:9 PowerPoint slide rebuilt from native PPT shapes (text boxes, rectangles, diamonds, ellipses, lines, arrows). Faithful conversion only — no redesign, no invented content. Use when the user shares a sketch, whiteboard, or diagram photo and wants a .pptx they can edit.
license: MIT
compatibility: Requires vision, shell access, and python3. Works in Claude Code, Claude Cowork, and GitHub Copilot CLI.
metadata:
  author: Alberto Jiménez Bákit (albertojb)
---

You convert a photo of a hand-drawn slide or diagram into an editable PowerPoint slide. You do the seeing; the bundled scripts do the rendering. Never build the .pptx yourself and never edit it by hand.

NORTH STAR: faithful conversion, not reinterpretation. You decide only shape mapping and placement — never content or layout redesign.

SKILL_DIR below means the directory containing this SKILL.md.

# Workflow

## 1. Get the photo

The user provides a photo path (whiteboard, paper, flipchart). If no photo was provided, ask for one. Look at the image before doing anything else.

## 2. Extract the layout

Read SKILL_DIR/references/contract.md (once per session) — it defines the JSON format exactly.

Transcribe every drawn element into `layout.json`, written next to the photo:

- One drawn element = one JSON element. Never add, merge, split, or rearrange.
- Estimate positions on a normalized 0–1 grid (origin top-left). Preserve relative positions and sizes.
- Snap elements that are clearly meant to align (within ~3% of each other) to the same coordinate, and give clearly same-sized boxes identical w/h. That is mapping, not redesign.
- Arrow direction follows the drawn arrowhead: `from` is the tail, `to` is the head. Arrowheads on both ends means `double_arrow`; no arrowhead means `line`.
- A title written across the top of the sketch goes in the top-level `title` field, not in `elements`.
- Text you cannot read becomes "[illegible]" — never guess or invent words.
- Ignore smudges, eraser ghosts, and camera artifacts. Do not ignore faint but deliberate marks.

## 3. Render

```bash
python3 -c "import pptx" 2>/dev/null || python3 -m pip install --quiet python-pptx
python3 SKILL_DIR/scripts/render.py layout.json -o <photo-stem>.pptx
```

## 4. Verify — machine-readable gate

```bash
python3 SKILL_DIR/scripts/verify.py layout.json <photo-stem>.pptx
```

The JSON output has a `passed` boolean. `passed: true` is the only acceptable success state. If false, fix layout.json (or report the failure honestly) and re-run steps 3–4. Never claim success verbally when the gate failed.

## 5. Report

Tell the user:

- The output .pptx path.
- The verifier verdict: `passed` plus element counts.
- Any "[illegible]" placeholders they need to fill in.
- Nothing else. Do not offer to redesign the slide.

# Failure modes

- Photo too blurry or diagram unreadable → ask for a better photo; do not guess a layout.
- render.py or verify.py errors → stderr JSON says why; fix layout.json accordingly.
- A drawn shape has no contract type (e.g. a cloud) → use the closest contract type (`rounded_rect`) and tell the user about the substitution.
