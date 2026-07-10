# Layout contract — v1

The intermediate representation between the agent's eyes and the renderer. The agent writes this JSON; `scripts/render.py` consumes it; `scripts/verify.py` checks the result against it.

## Coordinate system

The slide is 16:9 (13.333in × 7.5in). All positions and sizes are normalized 0.0–1.0: `x` and `w` are fractions of slide width, `y` and `h` are fractions of slide height. Origin is the top-left corner.

## Top level

```json
{
  "version": 1,
  "title": "Optional action title",
  "elements": []
}
```

| field | required | meaning |
|---|---|---|
| `version` | yes | must be `1` |
| `title` | no | rendered top-left, bold, title size |
| `elements` | yes | the drawn elements, one JSON object each |

## Element types

| type | geometry fields | notes |
|---|---|---|
| `rect` | `x, y, w, h` | plain box |
| `rounded_rect` | `x, y, w, h` | rounded box; also the fallback for clouds and blobs |
| `ellipse` | `x, y, w, h` | ovals |
| `circle` | `x, y, w` | true circle; `h` is derived so it renders round — use for anything drawn as a circle |
| `diamond` | `x, y, w, h` | decision diamonds |
| `chevron_right` | `x, y, w, h` | chevron / open arrow pointing right |
| `chevron_down` | `x, y, w, h` | chevron pointing down |
| `triangle_up` | `x, y, w, h` | solid-outline triangle pointing up |
| `triangle_right` | `x, y, w, h` | triangle pointing right |
| `triangle_down` | `x, y, w, h` | triangle pointing down |
| `triangle_left` | `x, y, w, h` | triangle pointing left |
| `text` | `x, y, w, h` | free-standing text, no outline, no fill |
| `table` | `x, y, w, h` + `columns`, `rows` | expanded by the renderer into header boxes, borderless text cells, and separator lines — never a native PPT table object (see below) |
| `line` | `from`+`to` or `x1, y1, x2, y2` | no arrowheads |
| `arrow` | `from`+`to` or `x1, y1, x2, y2` | arrowhead at the `to` / `(x2, y2)` end |
| `double_arrow` | `from`+`to` or `x1, y1, x2, y2` | arrowheads at both ends |

## Common fields

| field | applies to | default | meaning |
|---|---|---|---|
| `id` | all | — | unique string; required on any element a connector references |
| `text` | all | `""` | content; `\n` for line breaks; on connectors it renders as a small label at the midpoint |
| `bullets` | boxes, text | — | list of strings rendered as `• ` lines below `text`, top-anchored, left-aligned; use `"…"` for bullet rows drawn as placeholder squiggles |
| `size` | boxes, text | `body` | `title` (20pt) / `body` (12pt) / `small` (9pt) |
| `bold` | boxes, text | `false` | |
| `header` | boxes | `false` | section-header band: fixed slim height (~5.5% of slide, anchored at its drawn bottom edge), black fill, white bold left-aligned text; sits flush on its body box; width unification includes it (a band tracks its frame's width) but height unification excludes it |
| `align` | boxes, text | `center` for shapes, `left` for `text` | `left` / `center` / `right` |

## Tables

A drawn table (grid of rows and columns) is one `table` element — never transcribe its individual lines and cells. `columns` is an optional list of header strings (rendered as bold header boxes across the top); `rows` is a required non-empty list of rows, each a list of cell strings (rendered as borderless left-aligned text); every row must have the same cell count. Optional `row_headers: true` renders each row's first cell as a bold header box. The renderer expands the table deterministically: equal column widths, equal row heights, a thin horizontal separator under each body row — native shapes only, never a PowerPoint table object, so the result stays editable anywhere. Connectors cannot reference a table's id.

Connectors with `from`/`to` reference element ids; the renderer attaches them at side midpoints (top/left/bottom/right, picked by direction), glues them to the shapes so they survive editing, and routes misaligned connections as clean orthogonal elbows instead of diagonals. Use raw `x1, y1, x2, y2` only for lines that don't connect two boxes (dividers, axes, underlines).

## Tidy pass (deterministic — applied by both renderer and verifier)

Hand-estimated coordinates are cleaned up mechanically before rendering; structure is never changed:

1. **Canvas fit** — the content bounding box is scaled and centered into the slide's content area (margins, below the title band).
2. **Row/column alignment** — shapes whose centers are within ~5% vertically (or ~3.5% horizontally) are aligned exactly.
3. **Size unification** — shapes whose widths/heights are within a few percent are given identical dimensions.
4. Circles stay circular through all of the above.

Coordinates only need to be roughly right; alignment and consistent sizing come out clean automatically. Getting structure right (what connects to what, what sits in which row/column) is what matters.

## Styling (fixed — not expressible per element)

Black text and outlines on white fill. 1pt shape outlines, 1.5pt connector lines, no shadows, no colors. Font family is never set, so the slide inherits the theme font of any template it is pasted into. Single sanctioned inversion: `header: true` boxes render as black bands with white bold text (consulting section-header style).

Connectors pick attachment sides by geometry: vertically disjoint boxes (a row change) attach bottom→top so the elbow routes through the inter-row gap; same-row boxes attach left/right; overlapping boxes fall back to whichever direction dominates.

## Fidelity rule (verbatim)

Elements are transcribed, never added, merged, or rearranged; illegible text becomes "[illegible]".

Snapping exception: elements clearly meant to align (within ~3%) may be given identical coordinates or sizes. That is mapping, not redesign.

Consulting-format exceptions (bounded, deterministic — mapping, not redesign):

1. **Equidistant distribution** — 3+ boxes (or column-groups such as a header chip plus its body box) aligned in a row or column, whose edge-to-edge gaps are already similar (largest ≤ 2× smallest; a gap occupied by a small drawn shape or bridged by a connector is exempt from the bound), are redistributed to exactly equal gaps with the group's outer span kept fixed. Shapes drawn inside a gap are re-centered within it. Gaps outside the bound are deliberate and never touched.
2. **Flush snap** — boxes drawn touching or nearly touching (within ~1.5%, whether slightly apart or slightly overlapping, with ≥50% overlap on the perpendicular axis) snap flush; e.g. a header chip sits exactly on its body box, never doubling or crossing its border.
3. **Canonical frame ratios** — 2–3 tall frames (column-groups whose height spans ≥55% of the content height, jointly covering ≥70% of the content width, non-overlapping) whose width fractions are within 8 points of a canonical consulting split — 1/3+2/3, 2/3+1/3, equal thirds, or 1/4+1/2+1/4 — snap to the exact ratio; gutters are then equalized by the distribution rule when its bound allows. A split matching no canon (e.g. 50/50) stays untouched.

## Example

```json
{
  "version": 1,
  "title": "Order-to-cash: current state",
  "elements": [
    {"id": "b1", "type": "rect", "x": 0.06, "y": 0.22, "w": 0.20, "h": 0.14, "text": "Customer places order"},
    {"id": "b2", "type": "rounded_rect", "x": 0.40, "y": 0.22, "w": 0.20, "h": 0.14, "text": "Order validation"},
    {"id": "d1", "type": "diamond", "x": 0.72, "y": 0.18, "w": 0.22, "h": 0.22, "text": "In stock?"},
    {"id": "e1", "type": "ellipse", "x": 0.40, "y": 0.62, "w": 0.20, "h": 0.16, "text": "Back-order queue"},
    {"id": "t1", "type": "text", "x": 0.06, "y": 0.86, "w": 0.40, "h": 0.06, "text": "Source: ops team sketch", "size": "small"},
    {"id": "a1", "type": "arrow", "from": "b1", "to": "b2"},
    {"id": "a2", "type": "arrow", "from": "b2", "to": "d1", "text": "SLA 24h"},
    {"id": "a3", "type": "arrow", "from": "d1", "to": "e1", "text": "no"},
    {"id": "a4", "type": "double_arrow", "x1": 0.06, "y1": 0.50, "x2": 0.26, "y2": 0.50},
    {"id": "l1", "type": "line", "x1": 0.06, "y1": 0.18, "x2": 0.94, "y2": 0.18}
  ]
}
```
