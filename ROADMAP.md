# ROADMAP

Ordered epics. Every epic serves the NORTH STAR line in GOAL.md.

## 1. Conversion contract & skill scaffold — DONE (2026-07-09)

Layout JSON contract in references/contract.md; SKILL.md with the photo → layout.json → render → verify → report workflow; repo scaffold (README, LICENSE, requirements, examples).

## 2. Renderer — DONE (2026-07-09)

`scripts/render.py`: layout JSON → 16:9 .pptx via python-pptx. Native text boxes, autoshapes, and connectors with arrowheads. Neutral black-on-white, no template, no font family set.

## 3. Fidelity rules & machine-readable verification — DONE (2026-07-09)

Extraction rules in SKILL.md (spatial mapping, snapping exception, arrow direction, the [illegible] policy) plus `scripts/verify.py`, which checks every layout element landed in the .pptx and prints `{"passed": true/false, ...}`. Smoke-tested end to end on examples/sample-layout.json: passed.

## 4. Cross-host install & validation — IN PROGRESS

Install docs written for Claude Code, Claude Cowork, and GitHub Copilot CLI. Awaiting hands-on testing with real photos on each host — checklist in specs/epic-4-cross-host-validation.md. Sign-off deferred until Epic 5 lands: validating output with known layout deficiencies would burn test photos on results we already know are wrong.

## 5. Consulting layout conventions — IN PROGRESS (2026-07-10)

Bounded equidistant distribution, flush snap for header chips, canonical frame ratios (2/3+1/3, thirds, 25/50/25), contract-level `table` type expanded to lines + header boxes + borderless text (never the native PPT table object), consulting transcription rules in SKILL.md, tracked photo fixtures. Spec: specs/epic-5-consulting-layout-conventions.md. Issues #1–#4.

## 6. Publish v1 — PENDING

Blocked on Epics 4–5. Repo is public at github.com/albertojb/sketch-to-slide; v1 tag waits on validation.
