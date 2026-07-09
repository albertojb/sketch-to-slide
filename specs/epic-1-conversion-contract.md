# Epic 1 — Conversion contract & skill scaffold

Acceptance criteria. Done means every box checks.

## Layout contract

- [x] A documented JSON schema covering at minimum: rectangle, rounded rectangle, ellipse, diamond, freestanding text, line, single-head arrow, double-head arrow.
- [x] Every element carries: id, type, position and size in normalized 0–1 coordinates (origin top-left), optional text. Arrows and lines reference endpoints as coordinates or element ids.
- [x] Text styling is limited to: size tier (title / body / small), bold flag, alignment. No colors beyond black, white, gray.
- [x] The contract states the fidelity rule verbatim: elements are transcribed, never added, merged, or rearranged; illegible text becomes "[illegible]".

## SKILL.md

- [x] Frontmatter passes the Agent Skills spec: name `sketch-to-slide` (matches the directory), description states what it does and when to trigger, license MIT, author Alberto Jiménez Bákit (albertojb).
- [x] Body instructs the host model step by step: read the photo → emit layout.json per the contract → run `scripts/render.py` → run `scripts/verify.py` → report the output path and the verifier verdict.
- [x] Instructions assume nothing Zo-specific; only python3 and pip are available on the host.

## Repo scaffold

- [x] Folder layout exists: SKILL.md, README.md (one-paragraph pitch plus install placeholder), LICENSE (MIT), scripts/ (stubs only), examples/ (empty), specs/, GOAL.md, ROADMAP.md.
- [x] README states the renderer lands in Epic 2 — no false "works now" claim.

Out of scope for this epic: rendering code beyond stubs, host testing, publishing.
