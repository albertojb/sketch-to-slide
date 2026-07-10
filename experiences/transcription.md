# Experiences — Transcription pitfalls

> **Append-only.** Self-learning persistence layer, local to this install. Read this file once per session before extracting a layout (workflow step 3) and apply every rule that fits. When a pattern-level transcription mistake is caught and fixed during a run, append an `Experience NNN` entry below. Never delete entries; mark outdated ones `Superseded by NNN`.
>
> Mechanism adapted from mbb-ppt-generator's Self-Refinement protocol.

---

## Experience 001 — Photo orientation transcribed as-is

**Date**: 2026-07-09
**Problem**: A sketch photographed rotated 90° produced a sideways slide.
**Root Cause**: Coordinates were read in photo orientation instead of the orientation in which the handwriting reads normally.
**Fix**: Orientation-normalization step added to SKILL.md (step 2).
**Rule**: Decide the rotation before transcribing a single coordinate; cross-check that the narrative flows left-to-right / top-to-bottom.

## Experience 002 — Placeholder bullet rows silently dropped

**Date**: 2026-07-09
**Problem**: Bullet rows drawn as a dot plus a squiggle were treated as noise and omitted, so boxes rendered emptier than the sketch.
**Root Cause**: Squiggles were classified as illegible smudges instead of deliberate placeholder content.
**Fix**: `bullets` field added to the contract; squiggle rows transcribed as `"…"`.
**Rule**: A dot + horizontal squiggle inside a box is a placeholder bullet, never a smudge — count the rows and transcribe each one.

## Experience 003 — Connector rerouted to a "more sensible" shape

**Date**: 2026-07-09
**Problem**: A line drawn between Truck and Review was transcribed as Truck→AI Assess because that flow seemed more logical.
**Root Cause**: Inferred intent overrode what the pen actually drew.
**Fix**: Endpoint-fidelity rule added to SKILL.md (step 3).
**Rule**: Connect exactly the two shapes the drawn endpoints touch, even when a different connection would "make more sense".
