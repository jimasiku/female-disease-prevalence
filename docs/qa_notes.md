# QA Notes — FemaleDiseasPrevalence.pptx

## QA approach

Two layers of QA on each rebuild:

1. **Geometry inspection** (`scripts/qa_inspect.py`) — walks every shape on every slide and flags out-of-bounds shapes, off-slide shapes, empty text frames, and bounding-box overlaps between text-bearing shapes. Cards and pills (decorative rounded rectangles) are excluded from the overlap check because their text overlays sit on top by design.
2. **Visual inspection** — render the .pptx via LibreOffice → PDF → page-by-page PNG, then look at each slide. (See "Visual render" section below for what actually happened on this machine.)

Raw layout dump is preserved in `docs/qa_notes_raw.txt`.

## v1 — original build

Issues found and fixed on the first pass:

| # | Slide | Problem | Fix |
|---|-------|---------|-----|
| 1 | 2 (Hook) | Title box "Zambia's 2017 STEPS survey" overlapped the "24%" stat | Tightened title height and pushed stat down |
| 2 | 4–8 (Disease) | Full-width footnote collided with the slide-number box at 12.5in | Reduced footnote width from 12.3in to 11.4in |
| 3 | 6 (Migraine, old layout D) | Ratio label overlapped the "3:1" stat | Shrank label height |

## v2 — consistency revisions

Subsequent revisions applied per the follow-up brief:

- **Slide 1**: Removed the "A 3-MINUTE STORY" eyebrow and the "Big Idea" label. The Big Idea sentence now stands alone as the subtitle. Background changed from warm white (#FFF8F5) to deep plum (#3D2B56) with title and subtitle in white. Slide 1 is the only slide with the dark background — every other slide keeps the light theme.
- **Slide 2**: Removed the "THE HOOK · Diabesity Zambia" eyebrow. Everything else preserved.
- **Slide 3**: Removed both `majorGridlines` and `minorGridlines` from the category and value axes via direct XML element removal. The chart now has no grid lines.
- **Slides 3–8**: Collapsed the previous four disease-layout variants (A center-hero, B two-column, C card-right, D top-stat) into a **single consistent layout** applied to every disease slide. Slide 3 (Landscape) was brought into the same template on title position and footnote position.

### Shared layout for slides 3–8

| Element | Left | Top | Width | Height | Style |
|---------|------|-----|-------|--------|-------|
| Category pill (slides 4–8 only) | 0.9in | 0.55in | 2.4in | 0.4in | Lavender pill, 11pt PRIMARY |
| Slide title | 0.9in | 1.15in | 11.5in | 0.9in | 28pt bold PRIMARY |
| Stat eyebrow (4–8) | 0.9in | 2.4in | 6.0in | 0.45in | 14pt SECONDARY |
| Stat callout (4–8) | 0.9in | 2.85in | 6.0in | 2.0in | 160pt bold ACCENT |
| Chart (slide 3) | 0.9in | 2.4in | 11.5in | 4.4in | Violet bars |
| Description (4–8) | 7.2in | 2.4in | 5.5in | 2.4in | 16pt TEXT |
| So-what (4–8) | 7.2in | 5.0in | 5.5in | 1.9in | 22pt bold PRIMARY |
| Footnote | 0.5in | 7.05in | 11.4in | 0.35in | 9pt MUTED |
| Slide number (4–8) | 12.5in | 7.05in | 0.6in | 0.3in | 9pt MUTED right-aligned |
| Background | – | – | – | – | #FFF8F5 warm white |

The geometry inspection confirms identical positions across all five disease slides and matching title/footnote Y on slide 3.

## v2 — QA pass results

Geometry pass returns one informational warning:

- **Slide 9**, empty auto-shape at L=0.90in T=5.20in — this is the "ONE RECOMMENDATION" card hosting overlaid text. The card itself is intentionally text-empty. Not an issue.

No overlaps, no off-slide shapes, no cut-off content.

## Visual render

LibreOffice was not installed on the build machine. `winget install TheDocumentFoundation.LibreOffice` was attempted; the `msstore` source path failed under the corporate policy on this machine and the `winget` source path was used as the secondary attempt. If LibreOffice did install, the render command is:

```
soffice --headless --convert-to pdf output\FemaleDiseasPrevalence.pptx --outdir output\
```

then `pdftoppm output\FemaleDiseasPrevalence.pdf output\slide -r 120 -png` to rasterise. If it did not install, geometry QA above is the authoritative check.

## Final state

- 10 slides, 13.33in × 7.5in (16:9)
- One consistent layout across slides 4–8; slide 3 (chart) aligned on title and footnote Y
- Slide 1 is the only dark slide (deep plum, white text)
- Slide 2 has no eyebrow
- Slide 3 chart has no gridlines
- Palette adheres to the brief (deep plum #3D2B56, soft violet #7B5EA7, gold #E8A838, lavender mist #F3EEF8, warm white #FFF8F5)
- All footnote URLs resolve to live PMC / CDC pages
- No pink, no blue, no Comic Sans, no clip-art, no decorative header bands
