NORTH STAR: Faithful conversion, not reinterpretation — the skill only maps drawn elements to PPT shapes and positions; it never redesigns the layout or invents content. Sketch in, client-editable slide out, no manual rebuilding.

# Goal

A photo of a hand-drawn slide or diagram (whiteboard or paper) becomes an editable .pptx slide that recreates the same layout with native PowerPoint objects — text boxes, shapes, lines, and arrows — in the same relative positions.

# Core decision

Faithful conversion, not reinterpretation. The skill decides only shape mapping and placement, never content or layout redesign. The user goes from sketch to client-editable slide without manually rebuilding it.

# Definition of done

- One photo in, one 16:9 slide out, in a .pptx that opens cleanly in PowerPoint.
- Every box, arrow, connector, and label in the sketch exists as an editable native shape in the same relative position.
- Style is neutral black-on-white with no template or branding, so the slide adapts when pasted into any user's corporate template.
- The skill runs in Claude Code, Claude Cowork, and GitHub Copilot CLI.
- It lives as a public open-source repo under github.com/albertojb, installable and runnable by anyone, with no Zo dependency.

# Out of scope

- Creative redesign, "improving" the layout, or generating content that was not drawn.
- Embedding the source photo in the slide.
- Multi-slide deck generation (mbb-ppt-generator covers that).
- Branded templates, colors, logos, or footers.
- Perfect OCR: illegible handwriting becomes an "[illegible]" placeholder, never guessed content.

# Fixed constraints

- Self-contained: the repo bundles all code it needs; the only external dependency is python-pptx.
- Skill format follows the Agent Skills spec (SKILL.md with frontmatter) so one repo serves all three hosts.
- Division of labor: the host model does the vision (reads the photo and emits a structured layout file); bundled scripts do the rendering deterministically. The model never builds the .pptx directly.
- Verification is machine-readable: a script asserts the rendered slide matches the extracted layout and outputs a passed boolean, not a verbal claim.
- License MIT. Author: Alberto Jiménez Bákit (albertojb).
