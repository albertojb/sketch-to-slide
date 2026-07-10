# DECISIONS

One line each: decision — why — which GOAL line it serves.

## Epic 5 (2026-07-10)

- Bounded equidistance (2× gap ratio; occupied/bridged gaps exempt) instead of always-equalize — a clustered timeline is deliberate and must survive — "never redesigns the layout".
- Distribution and frame snapping move column-groups as units (chip + body together) — internal alignment must survive every snap — "same relative positions".
- Flush tolerance measured in drawn space (scaled by canvas-fit factors) — the contract bounds what the human drew, not the rescaled result; independent-review finding — "mapping, not redesign".
- Contract-level `table` element expanded by the renderer, never a native PPT table, connectors to tables rejected — a deterministic grid cannot come from freehand cell-by-cell transcription — "editable native shapes, no manual rebuilding".
- Frame-ratio snapping (±8 points) added as a third *named, bounded* contract exception rather than silently widening the 3% snapping rule — fidelity exceptions must be auditable — NORTH STAR.
- Epic 5 spec written in Allium (epics 1/4 stay markdown) — precise behavioral rules with invariants and testable scenarios — "verification is machine-readable".
- MBB reference slides (SAMPLES/) not committed — third-party images in a public MIT repo; their conventions are encoded in contract.md and SKILL.md — "public open-source repo, installable by anyone".
- Consulting knowledge lives in versioned SKILL.md/contract, not experiences/ — the journal is for run-time-discovered pitfalls only; keeping that meaning keeps self-learning trustworthy — self-contained-skill constraint.
- Epic 4 sign-off deferred until Epic 5 landed — validating output with known layout defects burns test photos on foreseen failures — "slide visually matches the sketch".
- Stack recovered via PR #9 after gh silently merged #6–#8 into stacked bases — process lesson recorded in CONTEXT.md: verify state after every gh write — (process, all GOAL lines).
