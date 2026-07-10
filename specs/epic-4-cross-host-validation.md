# Epic 4 — Cross-host install & validation

Acceptance criteria. Done means every box checks. This is Alberto's hands-on test checklist.

Reference photos from the 2026-07-09 field test live in examples/field-tests/ (situation-conflict-resolution, calls-flow-desk-to-field) — use these (or fresh photos) on each host and compare against examples/situation-conflict-resolution.json, the frozen transcription of the first one.

- [ ] Claude Code: skill installed at `~/.claude/skills/sketch-to-slide`; a real whiteboard photo produces a .pptx; verify.py reports `passed: true`; the slide visually matches the sketch.
- [ ] Claude Cowork: same photo via ZIP-uploaded skill → same outcome.
- [ ] GitHub Copilot CLI: same photo via `~/.copilot/skills/sketch-to-slide` (or repo `.github/skills/`) → same outcome.
- [ ] A wonky hand-drawn sketch (uneven boxes, crooked rows) comes out with rows/columns aligned where clearly intended — and nothing else changed.
- [ ] An illegible scribble becomes "[illegible]" in the slide; no invented text anywhere.
- [ ] The slide pasted into a branded corporate template inherits the theme font and keeps its black-on-white look.
- [ ] README install instructions confirmed accurate on each host; corrections applied.
