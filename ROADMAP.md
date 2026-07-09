# ROADMAP

Ordered epics. Every epic serves the NORTH STAR line in GOAL.md.

## 1. Conversion contract & skill scaffold — DONE (2026-07-09)

Layout JSON contract in references/contract.md; SKILL.md with the photo → layout.json → render → verify → report workflow; repo scaffold (README, LICENSE, requirements, examples).

## 2. Renderer — DONE (2026-07-09)

`scripts/render.py`: layout JSON → 16:9 .pptx via python-pptx. Native text boxes, autoshapes, and connectors with arrowheads. Neutral black-on-white, no template, no font family set.

## 3. Fidelity rules & machine-readable verification — DONE (2026-07-09)

Extraction rules in SKILL.md (spatial mapping, snapping exception, arrow direction, the [illegible] policy) plus `scripts/verify.py`, which checks every layout element landed in the .pptx and prints `{"passed": true/false, ...}`. Smoke-tested end to end on examples/sample-layout.json: passed.

## 4. Cross-host install & validation — IN PROGRESS

Install docs written for Claude Code, Claude Cowork, and GitHub Copilot CLI. Awaiting hands-on testing with real photos on each host — checklist in specs/epic-4-cross-host-validation.md.

## 5. Publish v1 — PENDING

Blocked on Epic 4. Public repo github.com/albertojb/sketch-to-slide not yet created; local git history is ready to push.
