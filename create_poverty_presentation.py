from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import MSO_AUTO_SIZE, MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


GOOGLE_SLIDES_FILE = "poverty_presentation_google_slides.pptx"
LEGACY_FILE = "poverty_presentation_bg.pptx"
FONT = "Arial"

NAVY = RGBColor(18, 34, 58)
DEEP_NAVY = RGBColor(9, 24, 43)
BLUE = RGBColor(43, 83, 128)
SLATE = RGBColor(76, 86, 100)
LIGHT_BG = RGBColor(246, 248, 251)
WHITE = RGBColor(255, 255, 255)
ORANGE = RGBColor(238, 137, 41)
YELLOW = RGBColor(249, 196, 73)
GREEN = RGBColor(70, 142, 101)
RED = RGBColor(184, 71, 68)
GREY = RGBColor(135, 144, 157)
PALE_BLUE = RGBColor(227, 235, 246)
PALE_ORANGE = RGBColor(255, 239, 217)


def inches(value):
    return Inches(value)


def shape(slide, shape_type, x, y, w, h, fill, line=None, transparency=0):
    item = slide.shapes.add_shape(shape_type, inches(x), inches(y), inches(w), inches(h))
    item.fill.solid()
    item.fill.fore_color.rgb = fill
    item.fill.transparency = transparency
    if line:
        item.line.color.rgb = line
        item.line.width = Pt(1)
    else:
        item.line.fill.background()
    return item


def text_box(
    slide,
    text,
    x,
    y,
    w,
    h,
    size=24,
    color=NAVY,
    bold=False,
    italic=False,
    align=PP_ALIGN.LEFT,
    valign=MSO_ANCHOR.TOP,
):
    box = slide.shapes.add_textbox(inches(x), inches(y), inches(w), inches(h))
    frame = box.text_frame
    frame.clear()
    frame.margin_left = inches(0.06)
    frame.margin_right = inches(0.06)
    frame.margin_top = inches(0.02)
    frame.margin_bottom = inches(0.02)
    frame.word_wrap = True
    frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    frame.vertical_anchor = valign

    paragraph = frame.paragraphs[0]
    paragraph.alignment = align
    run = paragraph.add_run()
    run.text = text
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return box


def bullets(slide, items, x, y, w, h, size=21, color=SLATE, accent=ORANGE):
    box = slide.shapes.add_textbox(inches(x), inches(y), inches(w), inches(h))
    frame = box.text_frame
    frame.clear()
    frame.margin_left = inches(0.04)
    frame.margin_right = inches(0.04)
    frame.margin_top = inches(0.02)
    frame.margin_bottom = inches(0.02)
    frame.word_wrap = True
    frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

    for index, item in enumerate(items):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.space_after = Pt(11)
        paragraph.line_spacing = 1.05
        marker = paragraph.add_run()
        marker.text = "• "
        marker.font.name = FONT
        marker.font.size = Pt(size + 2)
        marker.font.bold = True
        marker.font.color.rgb = accent

        body = paragraph.add_run()
        body.text = item
        body.font.name = FONT
        body.font.size = Pt(size)
        body.font.color.rgb = color
    return box


def header(slide, title, subtitle=None, dark=False, section=""):
    color = WHITE if dark else NAVY
    sub_color = RGBColor(222, 230, 240) if dark else SLATE
    text_box(slide, title, 0.72, 0.38, 8.8, 0.55, size=29, color=color, bold=True)
    shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 0.78, 1.07, 1.15, 0.08, ORANGE)
    if subtitle:
        text_box(slide, subtitle, 0.72, 1.23, 9.2, 0.38, size=15, color=sub_color)
    if section:
        pill = shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, 10.55, 0.42, 1.95, 0.38, ORANGE)
        pill.text_frame.clear()
        pill.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        paragraph = pill.text_frame.paragraphs[0]
        paragraph.alignment = PP_ALIGN.CENTER
        run = paragraph.add_run()
        run.text = section.upper()
        run.font.name = FONT
        run.font.size = Pt(10)
        run.font.bold = True
        run.font.color.rgb = WHITE


def footer(slide, number, dark=False):
    color = RGBColor(172, 185, 202) if dark else GREY
    text_box(slide, f"{number:02d}", 12.05, 6.86, 0.55, 0.22, size=10, color=color, bold=True, align=PP_ALIGN.RIGHT)


def card(slide, x, y, w, h, title, body, accent=ORANGE, fill=WHITE):
    shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, x + 0.05, y + 0.07, w, h, RGBColor(214, 222, 232), transparency=35)
    panel = shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, x, y, w, h, fill, RGBColor(225, 231, 238))
    shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, x, y, 0.14, h, accent)
    text_box(slide, title, x + 0.32, y + 0.23, w - 0.55, 0.43, size=17, color=NAVY, bold=True)
    text_box(slide, body, x + 0.32, y + 0.84, w - 0.55, h - 1.02, size=14, color=SLATE)
    return panel


def metric_card(slide, x, y, w, h, number, label, fill=NAVY, accent=YELLOW):
    panel = shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, x, y, w, h, fill)
    panel.text_frame.clear()
    text_box(slide, number, x + 0.22, y + 0.2, w - 0.44, 0.45, size=30, color=accent, bold=True, align=PP_ALIGN.CENTER)
    text_box(slide, label, x + 0.22, y + 0.82, w - 0.44, 0.58, size=13, color=WHITE, bold=True, align=PP_ALIGN.CENTER)


def icon_book(slide, x, y, scale=1):
    colors = [BLUE, ORANGE, GREEN]
    for index, color in enumerate(colors):
        shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, x + index * 0.36 * scale, y - index * 0.08 * scale, 0.3 * scale, 1.45 * scale, color)
        shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, x + 0.06 * scale + index * 0.36 * scale, y + 0.2 * scale - index * 0.08 * scale, 0.18 * scale, 0.035 * scale, WHITE)
        shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, x + 0.06 * scale + index * 0.36 * scale, y + 0.36 * scale - index * 0.08 * scale, 0.18 * scale, 0.035 * scale, WHITE)
    shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, x - 0.1 * scale, y + 1.42 * scale, 1.45 * scale, 0.1 * scale, SLATE)


def icon_heart(slide, x, y, scale=1):
    shape(slide, MSO_AUTO_SHAPE_TYPE.HEART, x, y, 1.15 * scale, 1.0 * scale, ORANGE)
    shape(slide, MSO_AUTO_SHAPE_TYPE.OVAL, x + 0.34 * scale, y + 0.32 * scale, 0.28 * scale, 0.28 * scale, WHITE)


def icon_growth(slide, x, y, scale=1):
    for index, height in enumerate([0.45, 0.75, 1.05]):
        shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, x + index * 0.42 * scale, y + (1.1 - height) * scale, 0.28 * scale, height * scale, [ORANGE, BLUE, GREEN][index])
    shape(slide, MSO_AUTO_SHAPE_TYPE.RIGHT_ARROW, x + 0.15 * scale, y + 0.05 * scale, 1.45 * scale, 0.28 * scale, YELLOW)


def title_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 0, 0, 13.333, 7.5, DEEP_NAVY)
    shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 0, 6.52, 13.333, 0.98, ORANGE)
    shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 8.15, 0, 5.2, 7.5, RGBColor(28, 49, 78))
    shape(slide, MSO_AUTO_SHAPE_TYPE.ARC, 7.55, -0.35, 5.8, 5.8, RGBColor(56, 86, 121), transparency=30)
    shape(slide, MSO_AUTO_SHAPE_TYPE.ARC, 8.35, 1.25, 4.65, 4.65, RGBColor(68, 104, 143), transparency=45)

    text_box(slide, "Бедността", 0.78, 1.05, 6.3, 0.74, size=47, color=WHITE, bold=True)
    text_box(slide, "отвъд празния джоб", 0.78, 1.86, 7.4, 0.72, size=38, color=YELLOW, bold=True)
    text_box(slide, "Материални, интелектуални и духовни измерения", 0.84, 2.92, 6.5, 0.56, size=20, color=RGBColor(224, 232, 244))
    text_box(slide, "Кратка презентация за видимата и невидимата бедност", 0.84, 5.72, 6.9, 0.35, size=13, color=RGBColor(224, 232, 244))

    # Clean contrast illustration: modern building beside a small old house.
    shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 9.05, 1.08, 1.34, 4.72, RGBColor(198, 216, 232), RGBColor(226, 236, 246))
    shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 10.7, 1.55, 1.1, 4.25, RGBColor(225, 234, 242), RGBColor(236, 242, 248))
    for row in range(6):
        for col in range(2):
            shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 9.28 + col * 0.48, 1.36 + row * 0.66, 0.22, 0.18, BLUE)
    for row in range(5):
        shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 10.95, 1.9 + row * 0.7, 0.42, 0.15, RGBColor(86, 109, 136))
    shape(slide, MSO_AUTO_SHAPE_TYPE.ISOSCELES_TRIANGLE, 8.62, 5.07, 2.05, 0.92, RGBColor(121, 78, 56))
    shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 8.9, 5.58, 1.5, 0.76, RGBColor(191, 158, 119))
    shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 9.48, 5.77, 0.34, 0.57, NAVY)
    shape(slide, MSO_AUTO_SHAPE_TYPE.OVAL, 8.55, 6.13, 3.75, 0.32, RGBColor(13, 30, 52), transparency=25)
    footer(slide, 1, dark=True)


def definition_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 0, 0, 13.333, 7.5, LIGHT_BG)
    header(slide, "Какво е бедност?", "Бедността не е само липса на пари.", section="дефиниция")

    card(
        slide,
        0.82,
        1.98,
        5.65,
        3.25,
        "Тя е липса на:",
        "възможности за развитие\n\nдостъп до здравеопазване и образование\n\nдостоен и сигурен начин на живот",
        accent=ORANGE,
    )
    metric_card(slide, 0.95, 5.66, 5.35, 0.85, "764 лв.", "линия на бедност в България за 2024 г.")

    shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, 7.35, 1.74, 4.85, 4.82, WHITE, RGBColor(225, 231, 238))
    text_box(slide, "Материална бедност", 7.78, 2.1, 3.95, 0.36, size=22, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
    shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 8.85, 3.1, 1.72, 1.14, RGBColor(222, 190, 145))
    shape(slide, MSO_AUTO_SHAPE_TYPE.ISOSCELES_TRIANGLE, 8.65, 2.42, 2.12, 0.94, RGBColor(123, 80, 58))
    shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 9.48, 3.52, 0.42, 0.72, NAVY)
    shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 9.04, 3.45, 0.26, 0.22, YELLOW)
    shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 10.13, 3.45, 0.26, 0.22, YELLOW)
    shape(slide, MSO_AUTO_SHAPE_TYPE.CLOUD, 8.1, 4.78, 1.1, 0.62, RGBColor(219, 226, 236))
    shape(slide, MSO_AUTO_SHAPE_TYPE.OVAL, 8.2, 5.55, 3.25, 0.3, RGBColor(215, 224, 235))
    text_box(slide, "Това е най-видимата форма,\nно не и единствената.", 7.82, 6.0, 3.95, 0.5, size=14, color=SLATE, align=PP_ALIGN.CENTER)
    footer(slide, 2)


def mind_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 0, 0, 13.333, 7.5, WHITE)
    header(slide, "Бедността на ума", "Знанието е капитал, който не се губи.", section="знание")

    bullets(
        slide,
        [
            "Липсата на образование създава цикъл на бедност.",
            "Недостигът на знания ни прави лесни за манипулация.",
            "Образованието отваря избори, професии и глас.",
        ],
        0.86,
        2.03,
        6.0,
        2.25,
        size=21,
        accent=BLUE,
    )
    card(
        slide,
        0.86,
        5.18,
        6.45,
        1.05,
        "Статистика",
        "Рискът от бедност при висшистите е около 10 пъти по-нисък, отколкото при хората без образование.",
        accent=BLUE,
        fill=RGBColor(248, 251, 255),
    )

    shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, 8.02, 1.85, 3.9, 4.42, PALE_BLUE, RGBColor(210, 223, 238))
    icon_book(slide, 8.55, 3.08, scale=1.18)
    icon_growth(slide, 10.05, 3.03, scale=1.1)
    text_box(slide, "учене → възможности", 8.45, 5.7, 3.1, 0.34, size=17, color=GREEN, bold=True, align=PP_ALIGN.CENTER)
    footer(slide, 3)


def spiritual_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 0, 0, 13.333, 7.5, DEEP_NAVY)
    shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 0, 6.82, 13.333, 0.68, ORANGE)
    header(slide, "Духовната бедност", "Празният дух в пълния костюм.", dark=True, section="ценности")

    bullets(
        slide,
        [
            "Липса на емпатия, ценности и чувство за отговорност.",
            "Можеш да имаш милиони, но да си просяк в душата си.",
            "Симптоми: безразличие, егоизъм, липса на културни потребности.",
        ],
        0.86,
        2.02,
        6.45,
        3.05,
        size=21,
        color=RGBColor(230, 236, 246),
        accent=YELLOW,
    )

    shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, 8.05, 1.7, 3.82, 4.82, RGBColor(24, 45, 75), RGBColor(74, 97, 127))
    shape(slide, MSO_AUTO_SHAPE_TYPE.OVAL, 9.38, 2.02, 1.16, 1.16, RGBColor(226, 205, 181))
    shape(slide, MSO_AUTO_SHAPE_TYPE.TRAPEZOID, 8.87, 3.26, 2.18, 2.46, RGBColor(45, 52, 61))
    shape(slide, MSO_AUTO_SHAPE_TYPE.ISOSCELES_TRIANGLE, 9.38, 3.3, 1.1, 1.0, WHITE)
    shape(slide, MSO_AUTO_SHAPE_TYPE.ISOSCELES_TRIANGLE, 9.78, 3.5, 0.3, 0.46, NAVY)
    icon_heart(slide, 9.45, 4.78, scale=1.05)
    text_box(slide, "богатство без човечност = празнота", 8.4, 5.98, 3.05, 0.35, size=13, color=RGBColor(229, 236, 246), bold=True, align=PP_ALIGN.CENTER)
    footer(slide, 4, dark=True)


def paradox_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 0, 0, 13.333, 7.5, LIGHT_BG)
    header(slide, "Парадоксът на богатството", "Материалното и духовното богатство невинаги вървят заедно.", section="контраст")

    card(
        slide,
        0.82,
        1.95,
        5.35,
        3.2,
        "Материално беден,\nно духовно богат",
        "Хора с малко средства, но с висок морал, доброта и готовност да помагат.\n\nПример: Дядо Добри.",
        accent=GREEN,
    )
    card(
        slide,
        7.15,
        1.95,
        5.35,
        3.2,
        "Материално богат,\nно духовно беден",
        "Хора с огромни възможности, но без капка човечност, състрадание или културна чувствителност.",
        accent=RED,
    )
    shape(slide, MSO_AUTO_SHAPE_TYPE.LEFT_RIGHT_ARROW, 6.23, 3.12, 0.86, 0.52, ORANGE)
    quote = shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, 1.72, 5.58, 9.9, 0.84, NAVY)
    quote.text_frame.clear()
    quote.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    paragraph = quote.text_frame.paragraphs[0]
    paragraph.alignment = PP_ALIGN.CENTER
    run = paragraph.add_run()
    run.text = '"Най-бедният човек е този, който има само пари."'
    run.font.name = FONT
    run.font.size = Pt(24)
    run.font.italic = True
    run.font.bold = True
    run.font.color.rgb = YELLOW
    footer(slide, 5)


def action_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 0, 0, 13.333, 7.5, WHITE)
    header(slide, "Как се борим с бедността?", "Трите измерения изискват три вида действие.", section="действие")

    actions = [
        ("1", "Четене и учене", "срещу бедността на ума", BLUE),
        ("2", "Доброволчество и помощ", "срещу духовната бедност", ORANGE),
        ("3", "Икономическа активност", "срещу материалната бедност", GREEN),
    ]
    for index, (num, title, body, color) in enumerate(actions):
        x = 0.86 + index * 4.08
        shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, x, 2.0, 3.35, 3.9, RGBColor(242, 246, 250), RGBColor(222, 230, 238))
        shape(slide, MSO_AUTO_SHAPE_TYPE.OVAL, x + 1.23, 2.32, 0.88, 0.88, color)
        text_box(slide, num, x + 1.48, 2.48, 0.38, 0.32, size=21, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        if index == 0:
            icon_book(slide, x + 1.02, 3.2, scale=0.65)
        elif index == 1:
            icon_heart(slide, x + 1.18, 3.28, scale=0.8)
        else:
            icon_growth(slide, x + 0.98, 3.35, scale=0.75)
        text_box(slide, title, x + 0.32, 4.72, 2.72, 0.42, size=18, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
        text_box(slide, body, x + 0.45, 5.25, 2.45, 0.35, size=13, color=SLATE, align=PP_ALIGN.CENTER)
    footer(slide, 6)


def closing_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 0, 0, 13.333, 7.5, DEEP_NAVY)
    shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 0, 0, 13.333, 1.18, ORANGE)
    text_box(slide, "Заключение", 0.82, 0.35, 4.4, 0.44, size=31, color=WHITE, bold=True)
    text_box(
        slide,
        "Истинското богатство е в това, което даваш,\nа не само в това, което притежаваш.",
        1.18,
        2.0,
        10.95,
        1.22,
        size=31,
        color=WHITE,
        bold=True,
        align=PP_ALIGN.CENTER,
    )
    shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, 1.35, 4.25, 10.62, 1.18, RGBColor(28, 52, 85), RGBColor(84, 111, 144))
    text_box(
        slide,
        "Коя бедност ви плаши повече - тази на портфейла или тази на сърцето?",
        1.82,
        4.58,
        9.68,
        0.46,
        size=22,
        color=YELLOW,
        bold=True,
        align=PP_ALIGN.CENTER,
    )
    icon_heart(slide, 5.55, 6.02, scale=0.72)
    shape(slide, MSO_AUTO_SHAPE_TYPE.OVAL, 6.78, 6.04, 0.58, 0.58, RGBColor(225, 234, 244))
    text_box(slide, "?", 6.92, 6.13, 0.28, 0.24, size=22, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
    footer(slide, 7, dark=True)


def build_deck():
    prs = Presentation()
    prs.slide_width = inches(13.333)
    prs.slide_height = inches(7.5)

    title_slide(prs)
    definition_slide(prs)
    mind_slide(prs)
    spiritual_slide(prs)
    paradox_slide(prs)
    action_slide(prs)
    closing_slide(prs)

    prs.core_properties.title = "Бедността: отвъд празния джоб"
    prs.core_properties.subject = "Материални, интелектуални и духовни измерения"
    prs.core_properties.language = "bg-BG"
    prs.core_properties.keywords = "Google Slides, Bulgarian, poverty, presentation"
    prs.save(GOOGLE_SLIDES_FILE)

    # Keep the original filename as an alias so previously shared links still work.
    Path(LEGACY_FILE).write_bytes(Path(GOOGLE_SLIDES_FILE).read_bytes())


if __name__ == "__main__":
    build_deck()
