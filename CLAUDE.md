# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

An Agent Skill (SKILL.md per the Agent Skills spec) that converts a photo of a hand-drawn slide into an editable 16:9 .pptx. One repo serves three hosts: Claude Code, Claude Cowork, GitHub Copilot CLI. Only runtime dependency is `python-pptx` (Pillow optional, for preview.py).

NORTH STAR (from GOAL.md, the drift anchor): **faithful conversion, not reinterpretation.** The skill maps drawn elements to PPT shapes and positions — never redesigns layout, never invents content. Every change must serve this. Illegible text becomes `"[illegible]"`, never a guess.

## Commands

```bash
# Full self-check: tidy-pass assertions + end-to-end render+verify on every example fixture
python3 tests/run.py

# Manual single run
python3 scripts/render.py examples/sample-layout.json -o /tmp/sample.pptx
python3 scripts/verify.py examples/sample-layout.json /tmp/sample.pptx   # JSON with "passed": true/false

# Optional visual check (needs Pillow)
python3 scripts/preview.py /tmp/sample.pptx   # writes a PNG wireframe
```

`verify.py`'s `passed` boolean is the only success gate — never claim success verbally when it failed. Scripts fail with JSON on stderr and exit code 2.

## Architecture

Division of labor is fixed (GOAL.md constraint): **the host model does the vision, the scripts do the rendering.** The model reads the photo and writes `layout.json`; it never builds or edits the .pptx directly.

Pipeline: photo → (agent transcribes per SKILL.md rules) → `layout.json` → `render.py` → .pptx → `verify.py` gate → `preview.py` visual check.

- **references/contract.md** — the layout.json contract (v1): element types, normalized 0–1 coordinates, connector `from`/`to` semantics, fixed black-on-white styling with no font family. Renderer, verifier, and SKILL.md transcription rules must all stay in sync with it.
- **scripts/render.py** — contract → .pptx via python-pptx. Contains the deterministic **tidy pass** (canvas fit, row/column alignment within tolerances, size unification) applied before rendering; tolerances are module constants at the top.
- **scripts/verify.py** — re-opens the .pptx and checks every layout element landed. It re-applies the same tidy pass to know the expected geometry, so tidy-pass changes in render.py must be mirrored here.
- **scripts/preview.py** — .pptx → PNG wireframe for the agent's visual comparison.
- **experiences/** — per-install, append-only self-learning journals (transcription.md, rendering.md). Entries are never deleted, only marked `Superseded by NNN`. Don't commit install-specific entries as if they were product content.
- **specs/**, **ROADMAP.md**, **GOAL.md** — epic specs and status. Currently pre-v1: Epic 4 (cross-host validation) in progress.

## Constraints to preserve

- Neutral styling only: black on white, 1pt outlines, no colors, no shadows, **no font family set** (so slides inherit any template's theme font).
- Connectors glue to shape sides and route as orthogonal elbows so they survive editing in PowerPoint.
- Self-contained repo, MIT, no dependencies beyond python-pptx.
- Verification stays machine-readable (`passed` boolean), never a verbal claim.
