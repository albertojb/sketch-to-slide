# Experiences — Rendering & verify pitfalls

> **Append-only.** Self-learning persistence layer, local to this install. Read this file whenever render.py or verify.py fails, before debugging from scratch. When a pattern-level render/verify pitfall is diagnosed and fixed during a run, append an `Experience NNN` entry below. Never delete entries; mark outdated ones `Superseded by NNN`.
>
> Mechanism adapted from mbb-ppt-generator's Self-Refinement protocol.

---

## Experience 001 — Vertical unglued elbows render as diagonals

**Date**: 2026-07-09
**Problem**: A diamond→ellipse arrow between vertically stacked shapes rendered as an ugly diagonal instead of a clean elbow.
**Root Cause**: PowerPoint's `bentConnector3` routes horizontal-first (H-V-H); with no glued connection sites, a vertical-dominant elbow has nothing to re-route it.
**Fix**: render.py uses an elbow only when the connector is horizontal-dominant or both ends are glued to shapes; otherwise a straight connector.
**Rule**: Fixed in render.py — if diagonal connectors reappear, check the glue/dominance logic there before touching layout.json.
