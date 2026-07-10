# sketch-to-slide

Draw a slide on paper or a whiteboard. Photograph it. Get back an editable 16:9 PowerPoint slide rebuilt from native PPT objects — text boxes, rectangles, circles, ellipses, diamonds, chevrons, triangles, bullets, lines, and arrows — laid out the way you drew it.

Faithful conversion, not reinterpretation: nothing is added, merged, redesigned, or "improved". A deterministic tidy pass aligns rows and columns, unifies near-identical sizes, and fits the drawing to the slide — mapping, not redesign. Connectors glue to shape sides and route as clean orthogonal elbows, so they survive editing in PowerPoint. Neutral black-on-white styling with no font family set, so the slide adapts to any corporate template you paste it into.

## How it works

1. Your AI agent (Claude or Copilot) looks at the photo, corrects for camera rotation, and transcribes every drawn element into `layout.json` per [references/contract.md](references/contract.md).
2. `scripts/render.py` deterministically tidies the layout and builds the .pptx with python-pptx.
3. `scripts/verify.py` re-opens the .pptx and checks that every element landed — it prints machine-readable JSON with a `passed` boolean. No verbal gate-passes.
4. `scripts/preview.py` draws a PNG wireframe of the .pptx so the agent can visually compare the result against your photo.

## Self-learning

The skill keeps a per-install journal in `experiences/` (transcription pitfalls, render/verify pitfalls). When a run hits a pattern-level mistake, the agent appends a short append-only `Experience` entry; future runs read the journal before transcribing — or when a gate fails — so the same mistake isn't repeated. Entries stay local to your install.

## Requirements

Any agent host with vision, shell access, and python3. The only dependency is `python-pptx` (the skill installs it automatically if missing). `Pillow` is optional, used only for the PNG preview.

## Install

**Claude Code** — clone into your personal skills folder:

```bash
git clone https://github.com/albertojb/sketch-to-slide ~/.claude/skills/sketch-to-slide
```

(or into a project's `.claude/skills/sketch-to-slide` for project-level use)

**Claude Cowork** — download this repo as a ZIP with the skill folder as the ZIP root, then in Cowork: Customize → "+" → Skills tab → upload the ZIP.

**GitHub Copilot CLI** — clone into your personal skills folder:

```bash
git clone https://github.com/albertojb/sketch-to-slide ~/.copilot/skills/sketch-to-slide
```

(or into a repo's `.github/skills/sketch-to-slide` for project-level use)

## Use

> "Turn this whiteboard photo into a slide: ~/Desktop/IMG_2041.jpg"

You get back `IMG_2041.pptx` plus the verifier's verdict, with any unreadable handwriting marked `[illegible]`.

## Manual run (no agent)

```bash
python3 scripts/render.py examples/sample-layout.json -o sample.pptx
python3 scripts/verify.py examples/sample-layout.json sample.pptx
```

## Status

Pre-v1. Renderer and verifier are smoke-tested; cross-host validation (Claude Code, Claude Cowork, GitHub Copilot CLI) is in progress — see [specs/epic-4-cross-host-validation.md](specs/epic-4-cross-host-validation.md).

## License

MIT — Alberto Jiménez Bákit ([albertojb](https://github.com/albertojb))
