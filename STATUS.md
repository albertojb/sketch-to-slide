# STATUS

## Next run starts here (2026-07-10, end of day)

- Epic 5 merged (PRs #5–#9) plus a same-day field-fix round from two real Claude Code runs, all independently reviewed to APPROVE and merged: PR #12 (flush snap repairs overlaps), PR #15 (MCK-style black header bands via `header: true`, overlap-aware bottom→top connector routing, fill-aware preview), PR #18 (bands drawn as a frame's top compartment seat flush and render on top; frame detection gates on tallest single box; detected frames align top/bottom edges; SKILL.md defaults: rect unless clearly rounded, no arrowhead = line).
- **Next bucket: Epic 4 — cross-host validation.** Alberto hands-on (next session, different PC): install fresh from GitHub, run the skill in Claude Code, Claude Cowork, and GitHub Copilot CLI against the photos in `examples/field-tests/`; checklist in `specs/epic-4-cross-host-validation.md`.
- Test command: `python3 tests/run.py` (18 checks, green on main).

## What landed (Epic 5)

- Equidistant distribution for 3+ aligned boxes/column-groups (2× gap bound; gaps occupied by a drawn shape or bridged by a connector exempt; in-gap shapes re-centered; deliberately uneven layouts untouched).
- Flush snap for near-touching boxes (header chips), tolerance measured in drawn space.
- Canonical frame ratios: 2/3+1/3, 1/3+2/3, equal thirds, 25/50/25 (±8 points; 55% height / 70% coverage gates; 50/50 untouched).
- Contract-level `table` element expanded to native shapes (bold header boxes, borderless text cells, row separator lines) — never a PPT table object; connectors to table ids rejected.
- SKILL.md consulting transcription rules; contract amended with three named bounded exceptions; Epic 5 spec written in Allium (`specs/epic-5-consulting-layout-conventions.md`).
- Fixtures: `examples/situation-conflict-resolution.json`, `examples/effort-table.json`, downscaled field-test photos in `examples/field-tests/`.

## Degunk backlog (minors, deliberately not filed as issues)

- Bare `except Exception` guards in `render.py` `_no_shadow` and `verify.py` auto-shape probe — add one-line justification comments.
- `_snap_flush` iterates ordered pairs, examining each unordered pair twice per axis.
- `_distribute` keys its decoration map by `id()`; an index would be clearer.
- `tests/run.py`: fixed temp filenames can collide across users; `check=True` hides render stderr on failure.
- SKILL.md gap-decoration bullet reads slightly broader than the code (re-centering requires a 3+ member group).
- `DECOR_FRAC` heuristic ceiling: a genuinely narrow content box (<50% of median group extent) is silently classed as an in-gap decoration.
