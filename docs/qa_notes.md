# QA Notes — FemaleDiseasPrevalence.pptx

## QA approach (deviation from spec)

The brief specified LibreOffice + pdftoppm for visual QA. LibreOffice is not installed in this environment and the user opted for the Python-only fallback path, so QA was performed by **geometry inspection of the underlying OOXML** instead of pixel rendering.

`scripts/qa_inspect.py` walks every shape on every slide and checks for:

1. Any shape extending past the slide bounds (cut-off content)
2. Any shape positioned off-slide (negative left/top)
3. Empty text frames (placeholders the build forgot to populate)
4. Bounding-box overlap between text-bearing shapes (excluding decorative auto-shapes such as pills and cards, which are intentionally text-hosting containers)

The raw layout dump is preserved in `docs/qa_notes_raw.txt`.

## Issues found and fixes

| # | Slide | Problem (1st pass) | Fix |
|---|-------|--------------------|-----|
| 1 | 2 (Hook) | Title box "Zambia's 2017 STEPS survey" (top=1.6in, h=1.2in) overlapped the "24%" stat (top=2.5in) | Tightened title height to 0.55in and moved stat to top=2.25in |
| 2 | 4–8 (Disease) | Full-width footnote (ends at 12.8in) collided with the slide-number box at 12.5in | Reduced footnote width from 12.3in to 11.4in (ends at 11.9in) |
| 3 | 6 (Migraine, Layout D) | "female-to-male ratio" label (h=0.6in, ends 1.75in) overlapped the "3:1" stat starting at top=1.6in | Shrank label height to 0.35in (ends 1.5in) |

After the fixes, the second QA pass returned only two informational warnings:

- **Slide 7**, empty auto-shape at L=7.60in T=1.30in — this is the lavender card hosting the "LIFETIME RISK GAP" stat. Text is layered on top in separate boxes for typographic precision; the card itself is intentionally text-empty. Not an issue.
- **Slide 9**, empty auto-shape at L=0.90in T=5.20in — same pattern: the "ONE RECOMMENDATION" card. Not an issue.

## Spot-checks on text fit

Hand-estimated character budgets vs. box areas (avg Calibri char width ≈ 0.45× the point size in pixels):

| Slide | Longest text | Chars | Box | Verdict |
|-------|--------------|-------|-----|---------|
| 1 (Title) | Big Idea | 278 | 11.5in × 2.2in @ 20pt | ≈3.5 lines, room for 6 → fits |
| 2 (Hook) | "more women than men…" | 154 | 11.5in × 1.8in @ 22pt | ≈2 lines, room for 4 → fits |
| 7 (Alzheimer) | "1 in 5 women vs 1 in 10 men…" | 96 | 5.1in × 1.4in @ 14pt | 2 lines, fits |
| 9 (So-What) | Big Idea + recommendation | 278 + 165 | separate boxes | fits |
| 10 (Sources) | 6 entries | ≈700 total | 11.5in × 5.5in @ 11–12pt | fits without scrolling |

## Limitations of this QA

Without rendering to pixels we cannot catch:

- Font substitution surprises if Calibri is unavailable at presentation time
- Sub-pixel kerning artefacts
- Chart data-label collisions at unusual zoom levels

If the file is ever rendered with LibreOffice or PowerPoint and a visual issue surfaces, edit `scripts/build_presentation.py` and re-run; the geometry pass will catch most layout regressions automatically.

## Final state

- 10 slides, 13.33in × 7.5in (16:9)
- All 5 disease slides use one of four layouts (A center-hero, B two-column, C card-right, D top-stat) — exceeds the ≥3 distinct arrangements rule
- Palette adheres to the brief (deep plum #3D2B56, soft violet #7B5EA7, gold #E8A838, lavender mist #F3EEF8, warm white #FFF8F5)
- All footnote URLs resolve to live PMC / CDC pages (verified at research time)
- No pink, no blue, no Comic Sans, no clip-art, no decorative header bands
