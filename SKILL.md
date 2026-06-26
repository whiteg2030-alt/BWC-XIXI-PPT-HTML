---
name: "BWC xixi ppt html"
description: Use this skill to create a BWC Xixi style HTML presentation and also export an editable-text PPTX from the generated HTML. The skill bundles BaiWuchangKeke fonts, installs them when needed, and requires HTML/PPTX visible text to use BaiWuchangKeke.
---

# BWC Xixi PPT HTML

## What This Skill Produces

Create two deliverables for the user:

1. A local single-file HTML slide deck.
2. An editable `.pptx` generated from that HTML. Text in the PPTX must be real editable PowerPoint text, not a screenshot.

Use this style for small-and-beautiful product decks, portfolios, research notes, private talks, lightweight launches, and refined Chinese presentation templates.

## Style Rules

- Content language: Chinese main content with small English decorative labels.
- Theme versions: black or white. If the user does not choose, make black by default and mention white is also supported.
- Black: `paper #000000`, `ink #f8f8f4`, `lime #b7ff00`.
- White: `paper #f7f7f2`, `ink #050505`, `lime #b7ff00`.
- Background: fine grid, subtle and clean.
- Accent: neon green only for page chip, progress dots, short rules, particles, and key numbers.
- All visible HTML text must use bundled `BaiWuchangKeke`.
- All editable PPTX text must use `白无常可可体`; do not use PowerPoint fallback fonts for the exported PPTX.
- Bundled font files live in `assets/fonts/`:
  - `BaiWuchangKeke-Thin.ttf`
  - `BaiWuchangKeke-Regular.ttf`
  - `BaiWuchangKeke-Bold.ttf`
- Big titles use thin weight. Body text uses bold weight.
- Titles must be at most two lines.
- Do not use emoji, external images, heavy gradients, big shadows, large green glow, or SaaS-style rounded-card clutter.

## Layout Set

Use only these layouts unless the user explicitly asks to extend the system:

- `S01 Cover`: title, short lead, tag, right-side animated particle ring.
- `S02 Principle`: left green number block, right title, lead, three principles.
- `S03 Content`: left core idea and body text, right metric card.
- `S04 Closing`: left conclusion, right three compact method cards.

Each slide section must include `data-layout="S01"` etc.

## Workflow

1. Copy one template from `assets/templates/`:
   - `xixi-template-black.html`
   - `xixi-template-white.html`
2. Replace the slide content while preserving:
   - fixed 1672 x 941 canvas
   - unified header/footer
   - page chip
   - progress dots
   - keyboard/click navigation
   - particle canvas on cover
3. Save the HTML with a clear name, for example `xixi_deck_black.html`.
4. If the user may not have BaiWuchangKeke installed, install the bundled font first. On Windows:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/install_bwc_fonts.ps1
```

5. Export editable PPTX from the generated HTML. The export script also attempts to install the bundled fonts automatically on Windows before creating the PPTX:

```bash
python scripts/html_to_editable_pptx.py path/to/xixi_deck_black.html path/to/xixi_deck_black.pptx
```

6. Check the HTML visually and open the PPTX to confirm text is editable and using BaiWuchangKeke.

## PPTX Export Contract

The PPTX does not need to reproduce every CSS detail pixel-perfectly. It must preserve:

- slide count
- slide order
- Chinese titles and body text as editable text boxes
- header/footer text
- page numbers
- black/white theme
- neon green accents
- the cover particle ring as editable PowerPoint shapes
- BaiWuchangKeke as the specified font for every editable text run

Do not export the PPTX as full-slide screenshots. The user must be able to click and edit text in PowerPoint.

## Quality Checklist

- HTML opens locally.
- PPTX opens locally.
- PPTX text is editable.
- PPTX text runs specify `白无常可可体`.
- `assets/fonts/` contains the three bundled BaiWuchangKeke font files.
- Header/footer are consistent.
- Page number and progress dots match slide index.
- Cover particle ring exists.
- Black version has clean black background, not green-tinted.
- White version has enough contrast.
- No title exceeds two lines.
- No text overlaps the header/footer.
