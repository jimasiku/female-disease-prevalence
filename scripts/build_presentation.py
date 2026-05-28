"""Build the FemaleDiseasePrevalence.pptx presentation.

Originally specified to use pptxgenjs (Node). Node is unavailable in this
environment, so this Python equivalent uses python-pptx and produces the
same slide structure, palette and layout rules described in the brief.

Run:
    py scripts/build_presentation.py

Output:
    output/FemaleDiseasPrevalence.pptx
"""

from __future__ import annotations

import json
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_TICK_MARK
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Emu, Inches, Pt

REPO_ROOT = Path(__file__).resolve().parent.parent
DATA_PATH = REPO_ROOT / "data" / "disease_research.json"
OUTPUT_PATH = REPO_ROOT / "output" / "FemaleDiseasePrevalence.pptx"

PRIMARY = RGBColor(0x3D, 0x2B, 0x56)      # deep plum
SECONDARY = RGBColor(0x7B, 0x5E, 0xA7)    # soft violet (Female bars)
ACCENT = RGBColor(0xE8, 0xA8, 0x38)       # gold (Male bars / stat callouts)
SURFACE = RGBColor(0xF3, 0xEE, 0xF8)      # lavender mist (cards)
TEXT = RGBColor(0x2C, 0x2C, 0x2C)         # near-black body text
MUTED = RGBColor(0x88, 0x88, 0x88)        # footnotes
BG = RGBColor(0xFF, 0xF8, 0xF5)           # very light warm background
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

# Disease-slide palette (v3 redesign — matches the HTML mockup)
DS_PILL_FILL = RGBColor(0xEE, 0xED, 0xFE)   # light violet pill background
DS_PILL_TEXT = RGBColor(0x53, 0x4A, 0xB7)   # violet pill text
DS_DESC = RGBColor(0x6B, 0x6B, 0x6B)        # grey description copy
DS_CARD_FILL = RGBColor(0xFA, 0xEE, 0xDA)   # warm gold stat card
DS_CARD_LABEL = RGBColor(0x63, 0x38, 0x06)  # dark brown label inside card
DS_CARD_STAT = RGBColor(0xBA, 0x75, 0x17)   # rich gold for the stat number
DS_CARD_SUB = RGBColor(0x85, 0x4F, 0x0B)    # mid-brown sub-label inside card
DS_DIVIDER = RGBColor(0xE0, 0xD8, 0xF0)     # pale violet divider line
DS_FOOTNOTE = RGBColor(0xAA, 0xAA, 0xAA)    # light grey footnote text

FONT = "Calibri"
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

BIG_IDEA = (
    "The same female-skewed health burden that shows up in Zambia's diabetes "
    "numbers runs through five of the most disabling diseases of modern life "
    "— and the world still designs medicine, research and screening as if it "
    "didn't."
)


def _set_background(slide, color=BG):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def _add_text(
    slide,
    left,
    top,
    width,
    height,
    text,
    *,
    font_size=18,
    bold=False,
    color=TEXT,
    align=PP_ALIGN.LEFT,
    anchor=MSO_ANCHOR.TOP,
    font_name=FONT,
):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    return box


def _add_pill(slide, left, top, width, height, text, fill=SURFACE, font_color=PRIMARY, font_size=11):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.adjustments[0] = 0.5
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    tf = shape.text_frame
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    run.font.name = FONT
    run.font.size = Pt(font_size)
    run.font.bold = True
    run.font.color.rgb = font_color
    return shape


def _apply_outer_shadow(shape, *, blur_pt=6, dist_pt=2, angle_deg=135, opacity=0.07):
    """Attach a soft outer-shadow effect to a shape (OOXML-level)."""
    from lxml import etree

    nsmap = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    sp_pr = shape._element.find(".//a:spPr", nsmap)
    if sp_pr is None:
        return
    # Drop any pre-existing effectLst
    for existing in sp_pr.findall("a:effectLst", nsmap):
        sp_pr.remove(existing)
    blur_emu = int(blur_pt * 12700)
    dist_emu = int(dist_pt * 12700)
    direction = int(angle_deg * 60000)
    alpha = int(opacity * 100000)
    xml = (
        f'<a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<a:outerShdw blurRad="{blur_emu}" dist="{dist_emu}" dir="{direction}" '
        f'rotWithShape="0">'
        f'<a:srgbClr val="000000"><a:alpha val="{alpha}"/></a:srgbClr>'
        f'</a:outerShdw>'
        f'</a:effectLst>'
    )
    sp_pr.append(etree.fromstring(xml))


def _add_divider(slide, x_in, y_in, w_in, color=DS_DIVIDER, weight_pt=0.5):
    """Thin horizontal divider line."""
    from pptx.util import Pt as _Pt

    line = slide.shapes.add_connector(1, Inches(x_in), Inches(y_in), Inches(x_in + w_in), Inches(y_in))
    line.line.color.rgb = color
    line.line.width = _Pt(weight_pt)
    return line


def _remove_gridlines(axis):
    """Strip major + minor gridlines from a chart axis."""
    el = axis._element
    for tag in ("c:majorGridlines", "c:minorGridlines"):
        for child in el.findall(f"{{http://schemas.openxmlformats.org/drawingml/2006/chart}}{tag.split(':')[1]}"):
            el.remove(child)


def _add_card(slide, left, top, width, height, fill=SURFACE):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.adjustments[0] = 0.06
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    return shape


def _add_footnote(slide, text, color=MUTED):
    _add_text(
        slide,
        Inches(0.5),
        Inches(7.05),
        Inches(11.4),
        Inches(0.35),
        text,
        font_size=9,
        color=color,
        align=PP_ALIGN.LEFT,
    )


def _add_slide_number(slide, n, total):
    _add_text(
        slide,
        Inches(12.5),
        Inches(7.05),
        Inches(0.6),
        Inches(0.3),
        f"{n} / {total}",
        font_size=9,
        color=MUTED,
        align=PP_ALIGN.RIGHT,
    )


def build_title_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _set_background(slide, PRIMARY)
    _add_text(
        slide,
        Inches(0.9),
        Inches(1.6),
        Inches(11.5),
        Inches(2.0),
        "Diabetes was the question.\nThe answer is much bigger than diabetes.",
        font_size=44,
        bold=True,
        color=WHITE,
    )
    _add_text(
        slide,
        Inches(0.9),
        Inches(4.3),
        Inches(11.5),
        Inches(2.4),
        BIG_IDEA,
        font_size=20,
        color=WHITE,
    )
    _add_text(
        slide,
        Inches(0.9),
        Inches(6.85),
        Inches(11.5),
        Inches(0.4),
        "Female disease prevalence  ·  Five conditions, one pattern",
        font_size=11,
        color=ACCENT,
    )
    return slide


def build_hook_slide(prs, hook):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _set_background(slide)
    _add_text(
        slide,
        Inches(0.9),
        Inches(1.15),
        Inches(11.5),
        Inches(0.55),
        "Zambia's 2017 STEPS survey",
        font_size=22,
        color=PRIMARY,
        bold=True,
    )
    _add_text(
        slide,
        Inches(0.9),
        Inches(2.25),
        Inches(11.5),
        Inches(2.2),
        "24%",
        font_size=110,
        bold=True,
        color=ACCENT,
    )
    _add_text(
        slide,
        Inches(0.9),
        Inches(4.7),
        Inches(11.5),
        Inches(1.8),
        "more women than men carried multiple non-communicable disease "
        "risk factors — and female diabetes prevalence ran 3.0% vs 2.1% in men.",
        font_size=22,
        color=TEXT,
    )
    _add_footnote(
        slide,
        f"Source: {hook['citation']}  ·  {hook['source_url']}",
    )
    return slide


def build_landscape_slide(prs, diseases):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _set_background(slide)
    # Title — same position/size/color as disease slide titles
    _add_text(
        slide,
        Inches(0.9),
        Inches(1.15),
        Inches(11.5),
        Inches(0.9),
        "Female-to-male prevalence ratio, five conditions",
        font_size=28,
        bold=True,
        color=PRIMARY,
    )

    sorted_diseases = sorted(diseases, key=lambda d: d["female_to_male_ratio"])
    categories = [d["name"].split(" (")[0] for d in sorted_diseases]
    ratios = [d["female_to_male_ratio"] for d in sorted_diseases]

    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series("Female-to-male ratio", ratios)

    # Body starts at the same Y as disease-slide body content (T=2.4in)
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(0.9),
        Inches(2.4),
        Inches(11.5),
        Inches(4.4),
        chart_data,
    ).chart

    chart.has_title = False
    chart.has_legend = False
    plot = chart.plots[0]
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.font.size = Pt(16)
    dl.font.name = FONT
    dl.font.bold = True
    dl.font.color.rgb = PRIMARY
    dl.number_format = '0.0"x"'
    dl.position = XL_LABEL_POSITION.OUTSIDE_END

    series = plot.series[0]
    fill = series.format.fill
    fill.solid()
    fill.fore_color.rgb = SECONDARY
    series.format.line.fill.background()

    cat_axis = chart.category_axis
    cat_axis.tick_labels.font.size = Pt(14)
    cat_axis.tick_labels.font.name = FONT
    cat_axis.tick_labels.font.color.rgb = TEXT
    cat_axis.major_tick_mark = XL_TICK_MARK.NONE
    cat_axis.minor_tick_mark = XL_TICK_MARK.NONE
    # Remove category-axis gridlines
    _remove_gridlines(cat_axis)

    val_axis = chart.value_axis
    val_axis.visible = False
    val_axis.major_tick_mark = XL_TICK_MARK.NONE
    val_axis.minor_tick_mark = XL_TICK_MARK.NONE
    val_axis.tick_labels.font.size = Pt(1)
    val_axis.maximum_scale = 10.0
    val_axis.minimum_scale = 0.0
    # Remove value-axis gridlines
    _remove_gridlines(val_axis)

    # Caption sits at the same footnote Y as the rest of slides 3-8
    _add_footnote(
        slide,
        "Each bar = how many times more common the condition is in women than men.",
    )
    return slide


def _add_stat_callout(slide, left, top, width, text, font_size=110, color=ACCENT, align=PP_ALIGN.LEFT):
    return _add_text(
        slide,
        left,
        top,
        width,
        Inches(2.0),
        text,
        font_size=font_size,
        bold=True,
        color=color,
        align=align,
    )


def build_disease_slide(prs, d, n, total):
    """Disease-slide layout v3 — matches the HTML mockup.

    Strictly identical x/y/w/h for every disease slide (4-8). No per-slide
    adjustments. The 3-card grid from the mockup is intentionally omitted.

    Geometry (inches):
      LEFT COLUMN (x=0.5, w=5.5)
        - Pill         L=0.5  T=0.55  W=1.6  H=0.28   fill #EEEDFE  text #534AB7
        - Disease name L=0.5  T=1.00  W=5.5  H=0.5    22pt  #3D2B56
        - Description  L=0.5  T=1.55  W=5.2  H=0.95   13pt  #6B6B6B
      RIGHT COLUMN — gold stat card
        - Card         L=6.3  T=0.45  W=3.2  H=1.9    fill #FAEEDA  shadow outer
        - Card label   centered  T=0.60  H=0.25       9pt  #633806  uppercase
        - Stat number  centered  T=0.85  H=0.80      52pt  #BA7517  bold
        - Sub-label    centered  T=1.65  H=0.60      11pt  #854F0B
      DIVIDER          y=2.55, full width 9.0in starting at x=0.5, #E0D8F0, 0.5pt
      FOOTNOTE         L=0.5  T=5.15  W=9.0  H=0.35   9pt  #AAAAAA
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _set_background(slide)

    # === LEFT COLUMN ===

    # Category pill
    _add_pill(
        slide,
        Inches(0.5),
        Inches(0.55),
        Inches(1.6),
        Inches(0.28),
        d["category"].upper(),
        fill=DS_PILL_FILL,
        font_color=DS_PILL_TEXT,
        font_size=9,
    )

    # Disease name (22pt, deep plum, not bold per spec)
    _add_text(
        slide,
        Inches(0.5),
        Inches(1.0),
        Inches(5.5),
        Inches(0.5),
        d["name"],
        font_size=22,
        color=PRIMARY,
    )

    # One-sentence description (13pt, mid-grey)
    _add_text(
        slide,
        Inches(0.5),
        Inches(1.55),
        Inches(5.2),
        Inches(0.95),
        d["description"],
        font_size=13,
        color=DS_DESC,
    )

    # === RIGHT COLUMN — gold stat card ===

    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(6.3),
        Inches(0.45),
        Inches(3.2),
        Inches(1.9),
    )
    card.adjustments[0] = 0.06
    card.fill.solid()
    card.fill.fore_color.rgb = DS_CARD_FILL
    card.line.fill.background()
    _apply_outer_shadow(card, blur_pt=6, dist_pt=2, angle_deg=135, opacity=0.07)

    # Card label (uppercase, small)
    _add_text(
        slide,
        Inches(6.3),
        Inches(0.6),
        Inches(3.2),
        Inches(0.25),
        "FEMALE-TO-MALE RATIO",
        font_size=9,
        bold=True,
        color=DS_CARD_LABEL,
        align=PP_ALIGN.CENTER,
    )

    # Big stat number (52pt)
    _add_text(
        slide,
        Inches(6.3),
        Inches(0.85),
        Inches(3.2),
        Inches(0.8),
        d["ratio_label"],
        font_size=52,
        bold=True,
        color=DS_CARD_STAT,
        align=PP_ALIGN.CENTER,
    )

    # Sub-label inside the card — uses the `so_what` sentence from the data
    _add_text(
        slide,
        Inches(6.3),
        Inches(1.65),
        Inches(3.2),
        Inches(0.6),
        d["so_what"],
        font_size=11,
        color=DS_CARD_SUB,
        align=PP_ALIGN.CENTER,
    )

    # === DIVIDER LINE ===
    _add_divider(slide, x_in=0.5, y_in=2.55, w_in=9.0, color=DS_DIVIDER, weight_pt=0.5)

    # === FOOTNOTE ===
    box = slide.shapes.add_textbox(Inches(0.5), Inches(5.15), Inches(9.0), Inches(0.4))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)
    tf.margin_top = Emu(0)
    tf.margin_bottom = Emu(0)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = f"[{d['id']}] {d['citation']} · {d['source_url']}"
    run.font.name = FONT
    run.font.size = Pt(9)
    run.font.color.rgb = DS_FOOTNOTE

    return slide


def build_so_what_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _set_background(slide)
    _add_text(
        slide,
        Inches(0.9),
        Inches(0.7),
        Inches(11.5),
        Inches(0.4),
        "THE SO-WHAT",
        font_size=12,
        bold=True,
        color=SECONDARY,
    )
    _add_text(
        slide,
        Inches(0.9),
        Inches(1.2),
        Inches(0.6),
        Inches(0.4),
        "Big Idea",
        font_size=11,
        bold=True,
        color=ACCENT,
    )
    _add_text(
        slide,
        Inches(0.9),
        Inches(1.7),
        Inches(11.5),
        Inches(3.0),
        BIG_IDEA,
        font_size=26,
        bold=True,
        color=PRIMARY,
    )
    _add_card(slide, Inches(0.9), Inches(5.2), Inches(11.5), Inches(1.6), fill=SURFACE)
    _add_text(
        slide,
        Inches(1.2),
        Inches(5.4),
        Inches(11.0),
        Inches(0.4),
        "ONE RECOMMENDATION",
        font_size=11,
        bold=True,
        color=ACCENT,
    )
    _add_text(
        slide,
        Inches(1.2),
        Inches(5.85),
        Inches(11.0),
        Inches(1.0),
        "Make sex-disaggregated reporting the default in every NCD surveillance "
        "round — starting with Zambia's next STEPS — so the burden women carry "
        "stops hiding in the totals.",
        font_size=18,
        color=TEXT,
    )
    return slide


def build_sources_slide(prs, diseases, hook):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _set_background(slide)
    _add_text(
        slide,
        Inches(0.9),
        Inches(0.5),
        Inches(11.5),
        Inches(0.6),
        "Sources",
        font_size=26,
        bold=True,
        color=PRIMARY,
    )

    box = slide.shapes.add_textbox(Inches(0.9), Inches(1.3), Inches(11.5), Inches(5.5))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Emu(0)
    tf.margin_right = Emu(0)

    entries = []
    for d in diseases:
        entries.append((f"[{d['id']}] {d['name']}", f"{d['citation']}.  {d['source_url']}"))
    entries.append(("[H] Hook — Zambia STEPS 2017", f"{hook['citation']}.  {hook['source_url']}"))

    for i, (head, body) in enumerate(entries):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(8)
        run1 = p.add_run()
        run1.text = head + " — "
        run1.font.name = FONT
        run1.font.size = Pt(12)
        run1.font.bold = True
        run1.font.color.rgb = PRIMARY
        run2 = p.add_run()
        run2.text = body
        run2.font.name = FONT
        run2.font.size = Pt(11)
        run2.font.color.rgb = TEXT

    _add_footnote(
        slide,
        "All sources accessed via PubMed Central, CDC NCHS, and peer-reviewed journals.",
    )
    return slide


def main():
    data = json.loads(DATA_PATH.read_text(encoding="utf-8"))
    diseases = data["diseases"]
    hook = data["hook_anchor"]

    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    total_slides = 10

    build_title_slide(prs)
    build_hook_slide(prs, hook)
    build_landscape_slide(prs, diseases)
    for d in diseases:
        build_disease_slide(prs, d, 0, total_slides)
    build_so_what_slide(prs)
    build_sources_slide(prs, diseases, hook)

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(OUTPUT_PATH))
    print(f"Wrote {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
