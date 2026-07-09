#!/usr/bin/env python3
"""sketch-to-slide renderer: layout JSON -> 16:9 .pptx of native PowerPoint shapes."""

import argparse
import json
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

SLIDE_W_EMU = 12192000
SLIDE_H_EMU = 6858000
BLACK = RGBColor(0, 0, 0)
WHITE = RGBColor(255, 255, 255)
SIZE_PT = {"title": 20, "body": 12, "small": 9}
ALIGN_MAP = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
SHAPE_MAP = {
    "rect": MSO_SHAPE.RECTANGLE,
    "rounded_rect": MSO_SHAPE.ROUNDED_RECTANGLE,
    "ellipse": MSO_SHAPE.OVAL,
    "diamond": MSO_SHAPE.DIAMOND,
}
CONNECTOR_TYPES = ("line", "arrow", "double_arrow")
TITLE_BOX = (0.04, 0.035, 0.92, 0.10)


def fail(msg):
    print(json.dumps({"error": msg}), file=sys.stderr)
    sys.exit(2)


def load_layout(path):
    try:
        with open(path) as f:
            layout = json.load(f)
    except (OSError, json.JSONDecodeError) as e:
        fail(f"cannot read layout: {e}")
    if layout.get("version") != 1:
        fail("layout version must be 1")
    elements = layout.get("elements")
    if not isinstance(elements, list) or not elements:
        fail("layout needs a non-empty 'elements' list")
    seen = set()
    for i, el in enumerate(elements):
        t = el.get("type")
        if t not in SHAPE_MAP and t != "text" and t not in CONNECTOR_TYPES:
            fail(f"element {i}: unknown type '{t}'")
        eid = el.get("id")
        if eid is not None:
            if eid in seen:
                fail(f"duplicate id '{eid}'")
            seen.add(eid)
        if t in SHAPE_MAP or t == "text":
            for k in ("x", "y", "w", "h"):
                if not isinstance(el.get(k), (int, float)):
                    fail(f"element {i} ('{t}'): missing numeric '{k}'")
        elif "from" not in el and "to" not in el:
            for k in ("x1", "y1", "x2", "y2"):
                if not isinstance(el.get(k), (int, float)):
                    fail(f"element {i} ('{t}'): needs from/to ids or x1,y1,x2,y2")
    return layout


def boxes_by_id(layout):
    return {
        el["id"]: (el["x"], el["y"], el["w"], el["h"])
        for el in layout["elements"]
        if el.get("id") and el["type"] not in CONNECTOR_TYPES
    }


def _edge_point(box, toward):
    x, y, w, h = box
    cx, cy = x + w / 2, y + h / 2
    dx, dy = toward[0] - cx, toward[1] - cy
    if dx == 0 and dy == 0:
        return (cx, cy)
    tx = (w / 2) / abs(dx) if dx else float("inf")
    ty = (h / 2) / abs(dy) if dy else float("inf")
    t = min(tx, ty)
    return (cx + dx * t, cy + dy * t)


def resolve_connector(el, boxes):
    """Return ((x1, y1), (x2, y2)) in normalized coordinates."""
    if "from" in el or "to" in el:
        for key in ("from", "to"):
            if el.get(key) not in boxes:
                fail(f"connector '{el.get('id', '?')}' references unknown id '{el.get(key)}'")
        b1, b2 = boxes[el["from"]], boxes[el["to"]]
        c1 = (b1[0] + b1[2] / 2, b1[1] + b1[3] / 2)
        c2 = (b2[0] + b2[2] / 2, b2[1] + b2[3] / 2)
        return _edge_point(b1, c2), _edge_point(b2, c1)
    return (el["x1"], el["y1"]), (el["x2"], el["y2"])


def label_box(el, p1, p2):
    text = str(el["text"])
    w = min(0.30, max(0.10, 0.008 * len(text)))
    h = 0.05
    mx, my = (p1[0] + p2[0]) / 2, (p1[1] + p2[1]) / 2
    return mx - w / 2, my - h / 2, w, h


def _ex(v):
    return Emu(int(round(v * SLIDE_W_EMU)))


def _ey(v):
    return Emu(int(round(v * SLIDE_H_EMU)))


def set_text(shape, text, size="body", bold=False, align="center", anchor=MSO_ANCHOR.MIDDLE):
    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    pt = Pt(SIZE_PT.get(size, SIZE_PT["body"]))
    for i, line in enumerate(str(text).split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.alignment = ALIGN_MAP.get(align, PP_ALIGN.CENTER)
        p.font.size = pt
        p.font.bold = bool(bold)
        p.font.color.rgb = BLACK
        for r in p.runs:
            r.font.size = pt
            r.font.bold = bool(bold)
            r.font.color.rgb = BLACK


def _no_shadow(shape):
    try:
        shape.shadow.inherit = False
    except Exception:
        pass


def add_arrowheads(conn, head, tail):
    # python-pptx has no arrowhead API; append DrawingML line-end elements directly
    ln = conn.line._get_or_add_ln()
    if head:
        ln.append(ln.makeelement(qn("a:headEnd"), {"type": "triangle", "w": "med", "len": "med"}))
    if tail:
        ln.append(ln.makeelement(qn("a:tailEnd"), {"type": "triangle", "w": "med", "len": "med"}))


def render(layout, out_path):
    prs = Presentation()
    prs.slide_width = Emu(SLIDE_W_EMU)
    prs.slide_height = Emu(SLIDE_H_EMU)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    boxes = boxes_by_id(layout)
    counts = {"shapes": 0, "textboxes": 0, "connectors": 0, "labels": 0}

    if layout.get("title"):
        x, y, w, h = TITLE_BOX
        tb = slide.shapes.add_textbox(_ex(x), _ey(y), _ex(w), _ey(h))
        set_text(tb, layout["title"], size="title", bold=True, align="left")
        counts["textboxes"] += 1

    connectors = []
    for el in layout["elements"]:
        t = el["type"]
        if t in CONNECTOR_TYPES:
            connectors.append(el)
            continue
        if t == "text":
            tb = slide.shapes.add_textbox(_ex(el["x"]), _ey(el["y"]), _ex(el["w"]), _ey(el["h"]))
            set_text(tb, el.get("text", ""), el.get("size", "body"), el.get("bold", False),
                     el.get("align", "left"), anchor=MSO_ANCHOR.TOP)
            counts["textboxes"] += 1
            continue
        sp = slide.shapes.add_shape(SHAPE_MAP[t], _ex(el["x"]), _ey(el["y"]), _ex(el["w"]), _ey(el["h"]))
        sp.fill.solid()
        sp.fill.fore_color.rgb = WHITE
        sp.line.color.rgb = BLACK
        sp.line.width = Pt(1)
        _no_shadow(sp)
        if el.get("text"):
            set_text(sp, el["text"], el.get("size", "body"), el.get("bold", False), el.get("align", "center"))
        counts["shapes"] += 1

    for el in connectors:
        p1, p2 = resolve_connector(el, boxes)
        conn = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT, _ex(p1[0]), _ey(p1[1]), _ex(p2[0]), _ey(p2[1]))
        conn.line.color.rgb = BLACK
        conn.line.width = Pt(1.5)
        _no_shadow(conn)
        t = el["type"]
        add_arrowheads(conn, head=(t == "double_arrow"), tail=(t in ("arrow", "double_arrow")))
        counts["connectors"] += 1
        if el.get("text"):
            lx, ly, lw, lh = label_box(el, p1, p2)
            tb = slide.shapes.add_textbox(_ex(lx), _ey(ly), _ex(lw), _ey(lh))
            tb.fill.solid()
            tb.fill.fore_color.rgb = WHITE
            tb.line.fill.background()
            set_text(tb, el["text"], "small", False, "center")
            counts["labels"] += 1

    prs.save(out_path)
    return counts


def main():
    ap = argparse.ArgumentParser(description="Render layout JSON into a 16:9 .pptx of native shapes")
    ap.add_argument("layout")
    ap.add_argument("-o", "--output")
    args = ap.parse_args()
    out = args.output or (args.layout.rsplit(".", 1)[0] + ".pptx")
    layout = load_layout(args.layout)
    counts = render(layout, out)
    print(json.dumps({"output": out, **counts}))


if __name__ == "__main__":
    main()
