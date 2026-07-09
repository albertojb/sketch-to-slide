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
| `align` | boxes, text | `center` for shapes, `left` for `text` | `left` / `center` / `right` |

Connectors with `from`/`to` reference element ids; the renderer attaches them at side midpoints (top/left/bottom/right, picked by direction), glues them to the shapes so they survive editing, and routes misaligned connections as clean orthogonal elbows instead of diagonals. Use raw `x1, y1, x2, y2` only for lines that don't connect two boxes (dividers, axes, underlines).

## Tidy pass (deterministic — applied by both renderer and verifier)

Hand-estimated coordinates are cleaned up mechanically before rendering; structure is never changed:

1. **Canvas fit** — the content bounding box is scaled and centered into the slide's content area (margins, below the title band).
2. **Row/column alignment** — shapes whose centers are within ~5% vertically (or ~3.5% horizontally) are aligned exactly.
3. **Size unification** — shapes whose widths/heights are within a few percent are given identical dimensions.
4. Circles stay circular through all of the above.

Coordinates only need to be roughly right; alignment and consistent sizing come out clean automatically. Getting structure right (what connects to what, what sits in which row/column) is what matters.

## Styling (fixed — not expressible per element)

Black text and outlines on white fill. 1pt shape outlines, 1.5pt connector lines, no shadows, no colors. Font family is never set, so the slide inherits the theme font of any template it is pasted into.

## Fidelity rule (verbatim)

Elements are transcribed, never added, merged, or rearranged; illegible text becomes "[illegible]".

Snapping exception: elements clearly meant to align (within ~3%) may be given identical coordinates or sizes. That is mapping, not redesign.

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
