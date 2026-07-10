# CONTEXT — what future runs need to know

- Canon: GOAL.md NORTH STAR → ROADMAP.md → specs/ → STATUS.md checkpoint. Epic 5's spec is Allium; epics 1 and 4 are markdown checklists — both forms are fine, pick per epic.
- `normalize_layout` in `scripts/render.py` is the single tidy pass; `verify.py` imports it, so any change lands in renderer and verifier simultaneously — never fork it. Pass order is fixed and documented in the Epic 5 spec's Resolved Questions: circles → canvas fit → align → unify → flush (drawn-space, scaled by fit factors) → frame snap → distribute horizontal → distribute vertical → circles → table expansion.
- Every tolerance is a named constant at the top of `render.py` and documented as a bounded exception in `references/contract.md`. New snaps must follow the same pattern: explicit numeric bound + contract entry, or they violate the NORTH STAR.
- Distribution and frame snapping operate on column-groups (boxes clustered by center), so attached elements (header chips) move/scale with their group — preserve that property in any new pass.
- Table expansion ids follow `{id}.hN` / `{id}.rNcM` / `{id}.lnN`; connectors referencing a table id are a validation error.
- `Test_files/` (scratch) and `SAMPLES/` (third-party MBB slide images — copyright, MIT repo) are gitignored on purpose. Curated fixtures go in `examples/`.
- Tooling caveat on this machine: gh 2.46 hits the GraphQL projectCards deprecation — `gh pr edit --base` fails silently and merge commands emit noise. ALWAYS confirm with `gh pr view N --json state` after any gh write, and never suppress stderr on gh commands. If a stacked base ever breaks: PR head refs (`refs/pull/N/head`) preserve everything; recover by branching from the top ref (that is how PR #9 recovered #6–#8).
- Stacked PRs: squash-merge top-down and retarget each PR to main (verify the retarget!) before merging.
