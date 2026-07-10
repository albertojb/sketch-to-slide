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

## 2. Normalize orientation — before reading any coordinates

Photos are often taken sideways or upside down. Use the handwriting as the compass: find the orientation in which the text reads normally (left to right, letters upright). ALL coordinates you transcribe must be in that corrected orientation, never in photo orientation. A sketch photographed sideways must NOT produce a sideways slide.

- Mentally rotate the whole sketch first, then transcribe. Positions like "top", "left", and arrow directions all refer to the corrected orientation.
- Cross-check with reading logic: labels like Context → Solution → Impact or a numbered sequence should flow left-to-right or top-to-bottom. If your transcription has a narrative running bottom-to-top or right-to-left, you got the rotation wrong — fix it before rendering.
- State in your report which rotation you applied (e.g. "photo was rotated 90° clockwise; corrected").

## 3. Extract the layout

Read SKILL_DIR/references/contract.md (once per session) — it defines the JSON format exactly. Also read SKILL_DIR/experiences/transcription.md (once per session) — transcription pitfalls this install has learned from earlier runs; apply every rule that fits.

Transcribe every drawn element into `layout.json`, written next to the photo:

- One drawn element = one JSON element. Never add, merge, split, or rearrange.
- The sketch's outer border or page frame is the slide canvas, NOT an element. Do not transcribe it. Ignore paper edges too.
- Estimate positions on a normalized 0–1 grid (origin top-left, corrected orientation). Rough is fine — the renderer deterministically aligns rows/columns, unifies near-identical sizes, and fits the content to the canvas. Your job is structure: what connects to what, what sits in which row and column.
- Bullet rows drawn as a dot plus a squiggle/line are placeholder bullets, not illegible text: put them in the box's `bullets` list as `"…"`. Legible bullet words are transcribed as written. Never silently drop bullet rows.
- Anything drawn as a circle uses type `circle` (renders perfectly round). Chevrons and big open arrows between sections use `chevron_*` / `triangle_*` types.
- For every line or arrow, identify the two shapes its endpoints actually touch in the sketch and connect exactly those (`from`/`to`). Never reroute a connector to a different shape because it "makes more sense".
- Arrow direction follows the drawn arrowhead: `from` is the tail, `to` is the head. Arrowheads on both ends means `double_arrow`; no arrowhead means `line`.
- A title written across the top of the sketch goes in the top-level `title` field, not in `elements`.
- A small label box attached to or on top of a bigger box (a header chip) is its own `rect` with `bold: true`. Transcribe it touching (or nearly touching) its body box; the renderer snaps it flush.
- A drawn table (grid with rows and columns) is ONE `table` element — never transcribe its grid lines, header cells, or body cells as separate elements. Put column headers in `columns`, cell text in `rows` (row-major, `"[illegible]"` for unreadable cells), and set `row_headers: true` when the first column is clearly a header column. The renderer builds the grid from native shapes; consulting slides never use PowerPoint table objects.
- Text you cannot read becomes "[illegible]" — never guess or invent words.
- Ignore smudges, eraser ghosts, and camera artifacts. Do not ignore faint but deliberate marks.

Consulting layout conventions (the renderer enforces these deterministically — your job is only to transcribe structure honestly):

- 3+ similar boxes in a row or column are almost always meant to be equal-size and equidistant, even with arrows or a chevron between them. Transcribe rough positions in the right ORDER with roughly similar gaps; do not fuss over exact spacing. Only spread gaps unevenly when the sketch is clearly deliberate about it (e.g. a timeline with clustered milestones) — the renderer preserves gaps that differ by more than 2×.
- 2–3 tall side-by-side frames dividing the page follow canonical splits: 2/3 + 1/3, equal thirds, or 1/4 + 1/2 + 1/4. Transcribe their rough widths; if they are near one of these splits the renderer snaps them exactly, so don't force it — a genuine 50/50 stays 50/50.
- A small shape between two row members (a triangle or chevron marking flow) belongs to the gap; the renderer re-centers it there.

## 4. Render

```bash
python3 -c "import pptx" 2>/dev/null || python3 -m pip install --quiet python-pptx
python3 SKILL_DIR/scripts/render.py layout.json -o <photo-stem>.pptx
```

## 5. Verify — machine-readable gate

```bash
python3 SKILL_DIR/scripts/verify.py layout.json <photo-stem>.pptx
```

The JSON output has a `passed` boolean. `passed: true` is the only acceptable success state. If false, fix layout.json (or report the failure honestly) and re-run steps 4–5. Never claim success verbally when the gate failed.

Then do a visual check (requires Pillow; skip if unavailable):

```bash
python3 SKILL_DIR/scripts/preview.py <photo-stem>.pptx
```

Look at the generated PNG next to the original photo. Check orientation (nothing sideways), that every connector joins the same two shapes as in the sketch, and that no element is missing. If something is off, fix layout.json and re-run steps 4–5.

## 6. Report

Tell the user:

- The output .pptx path.
- The verifier verdict: `passed` plus element counts.
- Any "[illegible]" placeholders they need to fill in.
- Nothing else. Do not offer to redesign the slide.

# Failure modes

- Photo too blurry or diagram unreadable → ask for a better photo; do not guess a layout.
- render.py or verify.py errors → first check SKILL_DIR/experiences/rendering.md for a known pitfall, then the stderr JSON says why; fix layout.json accordingly.
- A drawn shape has no contract type (e.g. a cloud or a star) → use the closest contract type (clouds/blobs → `rounded_rect`) and tell the user about the substitution.

# Self-learning (per install)

At the end of every run, classify each mistake you had to correct — a transcription error caught in the visual check, a gate failure, a user correction:

- **ONE-TIME** — specific to this photo → no persistence.
- **PATTERN** — could recur on future sketches → append an entry to the matching file:
  - seeing/transcription mistakes → SKILL_DIR/experiences/transcription.md
  - render/verify/script pitfalls → SKILL_DIR/experiences/rendering.md

Entry format (files are append-only; never delete an entry — mark outdated ones `Superseded by NNN`):

```markdown
## Experience NNN — <short title>

**Date**: YYYY-MM-DD
**Problem**: <one line>
**Root Cause**: <why it happened>
**Fix**: <what was done this run>
**Rule**: <one sentence a future run can apply>
```

Experiences stay local to this install. Describe the pitfall generically — do not copy the user's sketch content into an entry beyond what the pitfall needs.
