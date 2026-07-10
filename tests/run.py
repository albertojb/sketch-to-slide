#!/usr/bin/env python3
"""Self-checks: tidy-pass behavior plus end-to-end render+verify on every example fixture.

Run: python3 tests/run.py — exits non-zero on the first failed assertion.
"""

import json
import os
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "scripts"))
import render as R  # noqa: E402


def load_example(name):
    with open(os.path.join(ROOT, "examples", name)) as f:
        return json.load(f)


def gaps_of(els, ids):
    byid = {e["id"]: e for e in els}
    ms = sorted((byid[i] for i in ids), key=lambda b: b["x"])
    return [b["x"] - (a["x"] + a["w"]) for a, b in zip(ms, ms[1:])]


def test_scr_distribution():
    els = R.normalize_layout(load_example("situation-conflict-resolution.json"))["elements"]
    byid = {e["id"]: e for e in els}
    g = gaps_of(els, ["c1", "c2", "c3"])
    assert abs(g[0] - g[1]) < 1e-6, f"column gaps not equal: {g}"
    tr, c2, c3 = byid["tr"], byid["c2"], byid["c3"]
    mid = ((c2["x"] + c2["w"]) + c3["x"]) / 2
    assert abs((tr["x"] + tr["w"] / 2) - mid) < 1e-6, "triangle not centered in its gap"
    for h, c in (("h1", "c1"), ("h2", "c2"), ("h3", "c3")):
        assert abs((byid[h]["y"] + byid[h]["h"]) - byid[c]["y"]) < 1e-9, f"{h} not flush on {c}"
        assert abs(byid[h]["x"] - byid[c]["x"]) < 1e-9, f"{h} not left-aligned with {c}"
    assert len({round(byid[c]["w"], 6) for c in ("c1", "c2", "c3")}) == 1, "column widths differ"


def test_uneven_untouched():
    layout = {"version": 1, "elements": [
        {"id": "a", "type": "rect", "x": 0.05, "y": 0.40, "w": 0.10, "h": 0.20, "text": "a"},
        {"id": "b", "type": "rect", "x": 0.20, "y": 0.40, "w": 0.10, "h": 0.20, "text": "b"},
        {"id": "c", "type": "rect", "x": 0.80, "y": 0.40, "w": 0.10, "h": 0.20, "text": "c"},
    ]}
    g = gaps_of(R.normalize_layout(layout)["elements"], ["a", "b", "c"])
    assert g[1] / g[0] > 2.5, f"deliberately uneven gaps were equalized: {g}"


def test_bridged_gaps_equalize():
    # arrows between boxes exempt their gaps from the similarity bound
    layout = {"version": 1, "elements": [
        {"id": "a", "type": "rect", "x": 0.05, "y": 0.40, "w": 0.15, "h": 0.20, "text": "a"},
        {"id": "b", "type": "rect", "x": 0.26, "y": 0.40, "w": 0.15, "h": 0.20, "text": "b"},
        {"id": "c", "type": "rect", "x": 0.70, "y": 0.40, "w": 0.15, "h": 0.20, "text": "c"},
        {"id": "ab", "type": "arrow", "from": "a", "to": "b"},
        {"id": "bc", "type": "arrow", "from": "b", "to": "c"},
    ]}
    g = gaps_of(R.normalize_layout(layout)["elements"], ["a", "b", "c"])
    assert abs(g[0] - g[1]) < 1e-6, f"bridged gaps not equalized: {g}"


def test_table_expansion():
    els = R.normalize_layout(load_example("effort-table.json"))["elements"]
    parts = [e for e in els if e.get("id", "").startswith("t1.")]
    kinds = {"rect": 0, "text": 0, "line": 0}
    for e in parts:
        kinds[e["type"]] += 1
    assert kinds == {"rect": 3, "text": 9, "line": 3}, f"unexpected expansion: {kinds}"
    headers = sorted(e["x"] for e in parts if e["type"] == "rect")
    assert abs((headers[1] - headers[0]) - (headers[2] - headers[1])) < 1e-9, "unequal column widths"
    assert not any(e["type"] == "table" for e in els), "table not expanded"


def test_no_native_table_object():
    from pptx import Presentation
    src = os.path.join(ROOT, "examples", "effort-table.json")
    out = os.path.join(tempfile.gettempdir(), "effort-table-check.pptx")
    subprocess.run(
        [sys.executable, os.path.join(ROOT, "scripts", "render.py"), src, "-o", out],
        check=True, capture_output=True,
    )
    prs = Presentation(out)
    assert not any(getattr(sp, "has_table", False) for sp in prs.slides[0].shapes), \
        "native PPT table object found in output"


def test_connector_to_table_rejected():
    layout = {"version": 1, "elements": [
        {"id": "t", "type": "table", "x": 0.1, "y": 0.1, "w": 0.5, "h": 0.4, "rows": [["a"]]},
        {"id": "b", "type": "rect", "x": 0.7, "y": 0.1, "w": 0.2, "h": 0.2, "text": "b"},
        {"id": "arr", "type": "arrow", "from": "t", "to": "b"},
    ]}
    path = os.path.join(tempfile.gettempdir(), "bad-table-conn.json")
    with open(path, "w") as f:
        json.dump(layout, f)
    r = subprocess.run(
        [sys.executable, os.path.join(ROOT, "scripts", "render.py"), path,
         "-o", os.path.join(tempfile.gettempdir(), "bad.pptx")],
        capture_output=True, text=True,
    )
    assert r.returncode == 2 and "cannot attach to tables" in r.stderr, \
        f"connector-to-table not rejected: rc={r.returncode} stderr={r.stderr}"


def test_flush_in_drawn_space():
    # a small sketch upscaled 2.2x: chip drawn 1.2% above its body must still snap
    layout = {"version": 1, "elements": [
        {"id": "chip", "type": "rect", "x": 0.30, "y": 0.300, "w": 0.10, "h": 0.04, "text": "chip"},
        {"id": "body", "type": "rect", "x": 0.30, "y": 0.352, "w": 0.10, "h": 0.15, "bullets": ["…"]},
    ]}
    byid = {e["id"]: e for e in R.normalize_layout(layout)["elements"]}
    gap = byid["body"]["y"] - (byid["chip"]["y"] + byid["chip"]["h"])
    assert abs(gap) < 1e-9, f"chip not flush after upscale: gap {gap}"


def test_end_to_end():
    for name in sorted(os.listdir(os.path.join(ROOT, "examples"))):
        if not name.endswith(".json"):
            continue
        src = os.path.join(ROOT, "examples", name)
        out = os.path.join(tempfile.gettempdir(), name.replace(".json", ".pptx"))
        subprocess.run(
            [sys.executable, os.path.join(ROOT, "scripts", "render.py"), src, "-o", out],
            check=True, capture_output=True,
        )
        r = subprocess.run(
            [sys.executable, os.path.join(ROOT, "scripts", "verify.py"), src, out],
            capture_output=True, text=True,
        )
        v = json.loads(r.stdout)
        assert v["passed"], f"{name}: verify failed: {v['failures']}"


if __name__ == "__main__":
    for check in (test_scr_distribution, test_uneven_untouched,
                  test_bridged_gaps_equalize, test_flush_in_drawn_space,
                  test_table_expansion, test_no_native_table_object,
                  test_connector_to_table_rejected, test_end_to_end):
        check()
        print(f"ok {check.__name__}")
    print("all checks passed")
