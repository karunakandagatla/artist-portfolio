#!/usr/bin/env python3
"""Generate a traditional cream/maroon/saffron program PPT from program.html content."""

from __future__ import annotations

import shutil
from pathlib import Path

from PIL import Image, ImageEnhance
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

ROOT = Path(__file__).resolve().parent
IMG = ROOT / "assets" / "images"
PROG = IMG / "program_images"
OUT = ROOT / "Shakthi-Program.pptx"
TMP = ROOT / ".pptx_build"

BLACK = RGBColor(0x00, 0x00, 0x00)
GOLD = RGBColor(0xF2, 0xD1, 0x5C)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
RED = GOLD
CREAM = GOLD
MAROON = GOLD
ORANGE = GOLD
BROWN_DARK = GOLD
BROWN_MID = GOLD
BROWN_LIGHT = GOLD
CARD = BLACK
NAME_RED = GOLD
MAROON_SOFT = GOLD

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

FONT_HEAD = "Palatino"
FONT_BODY = "Georgia"
FONT_SANS = "Helvetica Neue"


ITEMS = [
    {
        "num": "1.a",
        "title": "Pushpanjali",
        "ragam": "Vagdheshwari",
        "thalam": "Adi",
        "choreography": "Guru Smt. Chandna Talluri",
        "accent": RGBColor(0xF2, 0xD1, 0x5C),
        "image": PROG / "pushpanjali-tanvika.png",
        "fit": "contain",
        "bg": (12, 8, 6),
    },
    {
        "num": "1.b",
        "title": "Saraswathi Kouthvam",
        "ragam": "Vagdheshwari",
        "thalam": "Chaturashra Eka (Traditional Kouthvam rhythm structure)",
        "choreography": "Guru Smt. Chandna Talluri",
        "accent": RGBColor(0xFF, 0x9F, 0x1C),
        "image": PROG / "TA-78.jpg",
        "fit": "cover",
        "focus": "top",
    },
    {
        "num": "02",
        "title": "Mahalakshmi Ashtaka Alarippu",
        "ragam": "Ragamalika",
        "thalam": "Misra Chapu",
        "choreography": "Dr Sridhar Vasudevan",
        "accent": RGBColor(0xF6, 0xDC, 0x6B),
        "image": PROG / "TA-92.jpg",
        "fit": "cover",
        "focus": "top",
    },
    {
        "num": "03",
        "title": "Kirtana — Sharade Shyamala, Shyamale Sharada",
        "ragam": "Shyamala (Malayamarutham / Derived Classicals)",
        "thalam": "Adi",
        "choreography": "Guru Smt. Chandna Talluri",
        "accent": RGBColor(0x3D, 0xFF, 0x9A),
        "image": PROG / "shyamala-tanvika.png",
        "fit": "contain",
        "bg": (12, 8, 6),
    },
    {
        "num": "04",
        "title": "Varnam Drishya Bharatham",
        "ragam": "Ragamalika",
        "thalam": "Adi",
        "choreography": "Guru Smt. Chandna Talluri",
        "accent": RGBColor(0xFF, 0x7A, 0x00),
        "image": PROG / "TA-80.jpg",
        "fit": "cover",
        "focus": "top",
    },
    {
        "num": "05",
        "title": "Bathukamma",
        "ragam": "Folk Melody",
        "thalam": "Traditional Folk Beat (Dappu rhythm alignment)",
        "choreography": "Guru Smt. Chandna Talluri",
        "accent": RGBColor(0xFF, 0x4F, 0xA3),
        "image": PROG / "TA-108.jpg",
        "fit": "cover",
        "focus": "top",
    },
    {
        "num": "06",
        "title": "Kuruvanji",
        "ragam": "Ragamalika (Traditional Folk-Classical Blend)",
        "thalam": "Adi / Misra Chapu",
        "choreography": "Guru Smt. Chandna Talluri",
        "accent": RGBColor(0x4E, 0xCB, 0xFF),
        "image": PROG / "TA-93.jpg",
        "fit": "cover",
        "focus": "top",
    },
    {
        "num": "07",
        "title": "Mahishasura Mardini",
        "ragam": "Revathi",
        "thalam": "Adi",
        "choreography": "Guru Smt. Chandna Talluri",
        "accent": RGBColor(0xFF, 0x3B, 0x3B),
        "image": PROG / "TA-35.jpg",
        "fit": "cover",
        "focus": "top",
    },
    {
        "num": "08",
        "title": "Thillana — Hamsanandi",
        "ragam": "Hamsanandi",
        "thalam": "Adi",
        "choreography": "Guru Smt. Chandna Talluri",
        "accent": RGBColor(0xFF, 0xE1, 0x4A),
        "image": PROG / "thillana-tanvika.png",
        "fit": "contain",
        "bg": (12, 8, 6),
    },
    {
        "num": "09",
        "title": "Mangalam",
        "ragam": "Madhyamavathi / Saurashtram",
        "thalam": "Adi",
        "choreography": "Guru Smt. Chandna Talluri",
        "accent": RGBColor(0xFF, 0xF3, 0xC4),
        "image": PROG / "TA-132.jpg",
        "fit": "cover",
        "focus": "top",
    },
]


def set_run_font(run, name, size, bold=False, italic=False, color=None, spacing=None):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    latin = rPr.find(qn("a:latin"))
    if latin is None:
        latin = etree.SubElement(rPr, qn("a:latin"))
    latin.set("typeface", name)
    if spacing is not None:
        rPr.set("spc", str(int(spacing)))


def set_shape_alpha(shape, alpha_pct):
    """alpha_pct: 0 transparent, 100 opaque."""
    spPr = shape._element.spPr
    solid = spPr.find(qn("a:solidFill"))
    if solid is None:
        return
    srgb = solid.find(qn("a:srgbClr"))
    if srgb is None:
        return
    for child in list(srgb):
        if child.tag == qn("a:alpha"):
            srgb.remove(child)
    alpha = etree.SubElement(srgb, qn("a:alpha"))
    alpha.set("val", str(int(alpha_pct * 1000)))


def no_line(shape):
    shape.line.fill.background()


def add_rect(slide, l, t, w, h, color, alpha=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    no_line(shape)
    if alpha is not None:
        set_shape_alpha(shape, alpha)
    return shape


def add_textbox(slide, l, t, w, h, anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.auto_size = None
    tf.margin_left = Inches(0.02)
    tf.margin_right = Inches(0.02)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    try:
        tf._txBody.bodyPr.set("anchor", {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}[anchor])
    except Exception:
        pass
    return tf


def write_lines(tf, lines, align=PP_ALIGN.LEFT):
    """lines: list of dicts with text and font kwargs. Use text='\\n' or empty to skip."""
    first = True
    for line in lines:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()
        p.alignment = line.get("align", align)
        p.space_before = Pt(line.get("before", 0))
        p.space_after = Pt(line.get("after", 0))
        run = p.add_run()
        run.text = line.get("text", "")
        set_run_font(
            run,
            line.get("font", FONT_BODY),
            line.get("size", 14),
            bold=line.get("bold", False),
            italic=line.get("italic", False),
            color=line.get("color", RED),
            spacing=line.get("spacing"),
        )
    return tf


def set_slide_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def prepare_image(src: Path, dest: Path, tw: int, th: int, fit="cover", focus="top", bg=(28, 13, 10)):
    img = Image.open(src)
    if img.mode == "P":
        img = img.convert("RGBA")
    if img.mode == "RGBA":
        canvas = Image.new("RGB", img.size, bg)
        canvas.paste(img, mask=img.split()[-1])
        img = canvas
    else:
        img = img.convert("RGB")

    src_w, src_h = img.size
    if fit == "contain":
        scale = min(tw / src_w, th / src_h)
        new_w = max(1, int(src_w * scale))
        new_h = max(1, int(src_h * scale))
        img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
        canvas = Image.new("RGB", (tw, th), bg)
        canvas.paste(img, ((tw - new_w) // 2, (th - new_h) // 2))
        img = canvas
    else:
        scale = max(tw / src_w, th / src_h)
        new_w = max(1, int(src_w * scale))
        new_h = max(1, int(src_h * scale))
        img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
        left = max(0, (new_w - tw) // 2)
        if focus == "top":
            top = 0
        elif focus == "bottom":
            top = max(0, new_h - th)
        else:
            top = max(0, (new_h - th) // 2)
        img = img.crop((left, top, left + tw, top + th))

    dest.parent.mkdir(parents=True, exist_ok=True)
    img.save(dest, "JPEG", quality=88, optimize=True)
    return dest


def add_full_picture(slide, src, box_l, box_t, box_w, box_h, dest_name):
    """Place the complete photo inside a box, with no cropping."""
    img = Image.open(src)
    if img.mode == "P":
        img = img.convert("RGBA")
    if img.mode == "RGBA":
        canvas = Image.new("RGB", img.size, (0, 0, 0))
        canvas.paste(img, mask=img.split()[-1])
        img = canvas
    else:
        img = img.convert("RGB")

    iw, ih = img.size
    max_w = box_w.inches
    max_h = box_h.inches
    aspect = iw / float(ih)
    if aspect >= (max_w / max_h):
        w_in = max_w
        h_in = max_w / aspect
    else:
        h_in = max_h
        w_in = max_h * aspect

    px_w = max(1, int(w_in * 160))
    px_h = max(1, int(h_in * 160))
    img = img.resize((px_w, px_h), Image.Resampling.LANCZOS)
    dest = TMP / dest_name
    dest.parent.mkdir(parents=True, exist_ok=True)
    img.save(dest, "JPEG", quality=88, optimize=True)

    left = box_l + Inches((max_w - w_in) / 2)
    top = box_t + Inches((max_h - h_in) / 2)
    slide.shapes.add_picture(str(dest), left, top, Inches(w_in), Inches(h_in))


def add_outer_border(slide):
    add_rect(slide, Inches(0), Inches(0), SLIDE_W, Inches(0.03), RED)
    add_rect(slide, Inches(0), SLIDE_H - Inches(0.03), SLIDE_W, Inches(0.03), RED)
    add_rect(slide, Inches(0), Inches(0), Inches(0.03), SLIDE_H, RED)
    add_rect(slide, SLIDE_W - Inches(0.03), Inches(0), Inches(0.03), SLIDE_H, RED)


def add_header_bar(slide, right_label="Bharathanatyam Arangetram"):
    tf = add_textbox(slide, Inches(0.35), Inches(0.18), Inches(7.2), Inches(0.46), MSO_ANCHOR.MIDDLE)
    write_lines(
        tf,
        [
            {
                "text": "SHAKTHI   ·   THE POWER WITHIN",
                "font": FONT_SANS,
                "size": 14,
                "bold": True,
                "color": GOLD,
                "spacing": 280,
            }
        ],
        PP_ALIGN.LEFT,
    )
    tf2 = add_textbox(slide, Inches(7.6), Inches(0.18), Inches(5.3), Inches(0.46), MSO_ANCHOR.MIDDLE)
    write_lines(
        tf2,
        [
            {
                "text": right_label.upper(),
                "font": FONT_SANS,
                "size": 13,
                "bold": True,
                "color": GOLD,
                "spacing": 180,
            }
        ],
        PP_ALIGN.RIGHT,
    )


def add_footer(slide, page, total):
    tf = add_textbox(slide, Inches(0.35), SLIDE_H - Inches(0.42), Inches(9.5), Inches(0.30), MSO_ANCHOR.MIDDLE)
    write_lines(
        tf,
        [
            {
                "text": "Tanvika Karupakula  ·  Saturday, September 5, 2026  ·  Dublin, California",
                "font": FONT_SANS,
                "size": 12,
                "color": GOLD,
                "spacing": 80,
            }
        ],
    )
    tf2 = add_textbox(slide, Inches(10.4), SLIDE_H - Inches(0.42), Inches(2.5), Inches(0.30), MSO_ANCHOR.MIDDLE)
    write_lines(
        tf2,
        [
            {
                "text": f"{page:02d}  /  {total:02d}",
                "font": FONT_SANS,
                "size": 12,
                "bold": True,
                "color": GOLD,
                "spacing": 80,
            }
        ],
        PP_ALIGN.RIGHT,
    )


def add_photo_frame(slide, img_path, l, t, w, h):
    add_rect(slide, l - Inches(0.06), t - Inches(0.06), w + Inches(0.12), h + Inches(0.12), MAROON)
    add_rect(slide, l - Inches(0.03), t - Inches(0.03), w + Inches(0.06), h + Inches(0.06), GOLD)
    slide.shapes.add_picture(str(img_path), l, t, w, h)


def build_title(prs, total):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, BLACK)

    photo = TMP / "title-hero.jpg"
    prepare_image(IMG / "AG_Hero_img.png", photo, 2100, 2250, fit="cover", focus="center")
    slide.shapes.add_picture(str(photo), Inches(6.55), Inches(0), Inches(6.783), SLIDE_H)
    add_rect(slide, Inches(6.55), Inches(0), Inches(0.03), SLIDE_H, RED)

    tf = add_textbox(slide, Inches(0.55), Inches(0.85), Inches(5.7), Inches(0.4))
    write_lines(
        tf,
        [
            {
                "text": "BHARATHANATYAM   ·   ARANGETRAM",
                "font": FONT_SANS,
                "size": 14,
                "bold": True,
                "color": GOLD,
                "spacing": 220,
            }
        ],
    )

    tf = add_textbox(slide, Inches(0.55), Inches(1.35), Inches(5.7), Inches(0.95))
    write_lines(
        tf,
        [{"text": "Shakthi", "font": FONT_HEAD, "size": 58, "bold": True, "color": GOLD}],
    )
    tf = add_textbox(slide, Inches(0.55), Inches(2.2), Inches(5.7), Inches(0.4))
    write_lines(
        tf,
        [
            {
                "text": "THE POWER WITHIN",
                "font": FONT_SANS,
                "size": 14,
                "bold": True,
                "color": ORANGE,
                "spacing": 360,
            }
        ],
    )

    add_rect(slide, Inches(0.55), Inches(2.72), Inches(2.1), Inches(0.045), MAROON)

    tf = add_textbox(slide, Inches(0.55), Inches(2.95), Inches(5.7), Inches(0.55))
    write_lines(
        tf,
        [{"text": "Program", "font": FONT_HEAD, "size": 40, "bold": True, "color": GOLD}],
    )
    tf = add_textbox(slide, Inches(0.55), Inches(3.5), Inches(5.7), Inches(0.55))
    write_lines(
        tf,
        [{"text": "Tanvika Karupakula", "font": FONT_HEAD, "size": 28, "bold": True, "color": WHITE}],
    )

    add_rect(slide, Inches(0.55), Inches(4.25), Inches(2.1), Inches(0.06), GOLD)

    tf = add_textbox(slide, Inches(0.55), Inches(4.55), Inches(5.7), Inches(2.0))
    write_lines(
        tf,
        [
            {
                "text": "Saturday, September 5, 2026",
                "font": FONT_SANS,
                "size": 20,
                "bold": True,
                "color": GOLD,
                "after": 8,
            },
            {
                "text": "2:30 PM PDT",
                "font": FONT_SANS,
                "size": 20,
                "bold": True,
                "color": WHITE,
                "after": 12,
            },
            {
                "text": "Dublin High School",
                "font": FONT_HEAD,
                "size": 22,
                "bold": True,
                "color": GOLD,
                "after": 4,
            },
            {
                "text": "8151 Village Parkway, Dublin, CA 94568",
                "font": FONT_SANS,
                "size": 16,
                "bold": True,
                "color": WHITE,
            },
        ],
    )


def build_overview(prs, page, total):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, BLACK)
    add_outer_border(slide)
    add_header_bar(slide, "Program Sequence")
    add_footer(slide, page, total)

    tf = add_textbox(slide, Inches(0.55), Inches(0.82), Inches(12.2), Inches(0.55))
    write_lines(
        tf,
        [{"text": "The Margam", "font": FONT_HEAD, "size": 26, "bold": True, "color": MAROON}],
        PP_ALIGN.CENTER,
    )
    add_rect(slide, Inches(5.9), Inches(1.38), Inches(1.5), Inches(0.03), GOLD)

    def draw_column(items, x, heading):
        tf = add_textbox(slide, Inches(x), Inches(1.48), Inches(5.55), Inches(0.32))
        write_lines(
            tf,
            [
                {
                    "text": heading,
                    "font": FONT_SANS,
                    "size": 11,
                    "bold": True,
                    "color": ORANGE,
                    "spacing": 180,
                }
            ],
        )
        y = 1.85
        for item in items:
            tf = add_textbox(slide, Inches(x), Inches(y), Inches(5.55), Inches(0.38))
            write_lines(
                tf,
                [{"text": item["title"], "font": FONT_HEAD, "size": 18, "bold": True, "color": BROWN_DARK}],
            )
            tf = add_textbox(slide, Inches(x), Inches(y + 0.36), Inches(5.55), Inches(0.32))
            write_lines(
                tf,
                [
                    {
                        "text": f"Ragam  {item['ragam']}   ·   Thalam  {item['thalam']}",
                        "font": FONT_SANS,
                        "size": 12,
                        "bold": True,
                        "color": MAROON,
                    }
                ],
            )
            y += 0.82

    draw_column(ITEMS[:5], 0.55, "BEFORE INTERVAL")
    draw_column(ITEMS[5:], 7.2, "AFTER INTERVAL")


def build_item(prs, item, page, total, photo_left=True):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, BLACK)

    accent = item.get("accent", GOLD)
    if photo_left:
        photo_box = (Inches(0.15), Inches(0.15), Inches(6.35), Inches(7.20))
        text_x = 6.70
    else:
        photo_box = (Inches(6.85), Inches(0.15), Inches(6.35), Inches(7.20))
        text_x = 0.40

    add_full_picture(
        slide,
        item["image"],
        *photo_box,
        dest_name=f"item-{item['num'].replace('.', '')}.jpg",
    )

    title = item["title"].upper()
    title_len = len(title)
    title_size = 26 if title_len > 42 else 32 if title_len > 28 else 36 if title_len > 18 else 40
    title_h = 1.45 if title_len > 28 else 0.95

    tf = add_textbox(slide, Inches(text_x), Inches(1.85), Inches(6.15), Inches(title_h), MSO_ANCHOR.BOTTOM)
    write_lines(
        tf,
        [{"text": title, "font": FONT_SANS, "size": title_size, "bold": True, "color": accent}],
    )

    y = 1.85 + title_h + 0.28
    rows = [
        ("RAGAM", item["ragam"]),
        ("THALAM", item["thalam"]),
        ("CHOREOGRAPHY", item["choreography"]),
    ]
    for label, value in rows:
        long_value = len(value) > 28
        row_h = 0.95 if long_value else 0.72
        tf = add_textbox(slide, Inches(text_x), Inches(y), Inches(2.45), Inches(row_h), MSO_ANCHOR.TOP)
        write_lines(
            tf,
            [
                {
                    "text": label,
                    "font": FONT_SANS,
                    "size": 16,
                    "bold": True,
                    "color": GOLD,
                    "spacing": 160,
                }
            ],
        )
        value_size = 18 if long_value else 22
        tf = add_textbox(
            slide,
            Inches(text_x + 2.45),
            Inches(y),
            Inches(3.70),
            Inches(row_h),
            MSO_ANCHOR.TOP,
        )
        write_lines(
            tf,
            [{"text": value.upper(), "font": FONT_SANS, "size": value_size, "bold": True, "color": WHITE}],
        )
        y += row_h + 0.08


def build_interval(prs, page, total):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, BLACK)
    add_outer_border(slide)
    add_header_bar(slide)
    add_footer(slide, page, total)

    add_rect(slide, Inches(3.1), Inches(3.48), Inches(2.2), Inches(0.06), GOLD)
    add_rect(slide, Inches(8.05), Inches(3.48), Inches(2.2), Inches(0.06), GOLD)

    tf = add_textbox(slide, Inches(1.8), Inches(2.55), Inches(9.7), Inches(1.5), MSO_ANCHOR.MIDDLE)
    write_lines(
        tf,
        [
            {
                "text": "Interval",
                "font": FONT_HEAD,
                "size": 72,
                "bold": True,
                "color": GOLD,
                "align": PP_ALIGN.CENTER,
            }
        ],
        PP_ALIGN.CENTER,
    )

    tf = add_textbox(slide, Inches(2.2), Inches(4.2), Inches(8.9), Inches(0.7))
    write_lines(
        tf,
        [
            {
                "text": "Bathukamma",
                "font": FONT_SANS,
                "size": 28,
                "bold": True,
                "color": RGBColor(0xFF, 0x4F, 0xA3),
                "align": PP_ALIGN.CENTER,
                "spacing": 120,
            }
        ],
        PP_ALIGN.CENTER,
    )


def build_closing(prs, page, total):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, BLACK)
    add_outer_border(slide)
    add_header_bar(slide)
    add_footer(slide, page, total)

    add_full_picture(
        slide,
        PROG / "TA-132.jpg",
        Inches(6.85),
        Inches(0.15),
        Inches(6.35),
        Inches(7.20),
        dest_name="close-photo.jpg",
    )

    tf = add_textbox(slide, Inches(0.50), Inches(1.8), Inches(6.3), Inches(0.5))
    write_lines(
        tf,
        [
            {
                "text": "MANGALAM",
                "font": FONT_SANS,
                "size": 18,
                "bold": True,
                "color": RGBColor(0xFF, 0xF3, 0xC4),
                "spacing": 320,
            }
        ],
    )
    tf = add_textbox(slide, Inches(0.50), Inches(2.35), Inches(6.3), Inches(1.3))
    write_lines(
        tf,
        [{"text": "Thank you", "font": FONT_HEAD, "size": 56, "bold": True, "color": GOLD}],
    )
    add_rect(slide, Inches(0.50), Inches(3.8), Inches(2.4), Inches(0.07), GOLD)
    tf = add_textbox(slide, Inches(0.50), Inches(4.1), Inches(6.3), Inches(1.4))
    write_lines(
        tf,
        [
            {"text": "Shakthi — The Power Within", "font": FONT_HEAD, "size": 24, "bold": True, "color": GOLD, "after": 10},
            {
                "text": "Tanvika Karupakula",
                "font": FONT_SANS,
                "size": 20,
                "bold": True,
                "color": WHITE,
                "after": 6,
            },
            {
                "text": "September 5, 2026",
                "font": FONT_SANS,
                "size": 18,
                "color": GOLD,
            },
        ],
    )


def main():
    if TMP.exists():
        shutil.rmtree(TMP)
    TMP.mkdir(parents=True)

    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    prs.core_properties.title = "Shakthi — Program"
    prs.core_properties.author = "Tanvika Karupakula"
    prs.core_properties.subject = "Bharathanatyam Arangetram Program"

    # title, 5 items, interval, 5 items, closing
    total = 1 + len(ITEMS) + 2

    build_title(prs, total)
    page = 2
    for i, item in enumerate(ITEMS[:5]):
        build_item(prs, item, page, total, photo_left=(i % 2 == 0))
        page += 1
    build_interval(prs, page, total)
    page += 1
    for i, item in enumerate(ITEMS[5:], start=5):
        build_item(prs, item, page, total, photo_left=(i % 2 == 0))
        page += 1
    build_closing(prs, page, total)

    prs.save(OUT)
    shutil.rmtree(TMP, ignore_errors=True)
    print(f"Wrote {OUT} ({OUT.stat().st_size / 1_000_000:.1f} MB, {total} slides)")


if __name__ == "__main__":
    main()
