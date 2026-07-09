# sketch-to-slide

Draw a slide on paper or a whiteboard. Photograph it. Get back an editable 16:9 PowerPoint slide rebuilt from native PPT objects — text boxes, rectangles, diamonds, ellipses, lines, and arrows — in the same positions you drew them.

Faithful conversion, not reinterpretation: nothing is added, merged, redesigned, or "improved". Neutral black-on-white styling with no font family set, so the slide adapts to any corporate template you paste it into.

## How it works

1. Your AI agent (Claude or Copilot) looks at the photo and transcribes every drawn element into `layout.json` per [references/contract.md](references/contract.md).
2. `scripts/render.py` deterministically builds the .pptx with python-pptx.
3. `scripts/verify.py` re-opens the .pptx and checks that every element landed — it prints machine-readable JSON with a `passed` boolean. No verbal gate-passes.

## Requirements

Any agent host with vision, shell access, and python3. The only dependency is `python-pptx` (the skill installs it automatically if missing).

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
