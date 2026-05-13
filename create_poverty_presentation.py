from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE, MSO_ANCHOR
from pptx.util import Inches, Pt


OUTFILE = "poverty_presentation_bg.pptx"

NAVY = RGBColor(20, 36, 60)
DEEP_NAVY = RGBColor(12, 25, 45)
SLATE = RGBColor(67, 78, 94)
LIGHT_BG = RGBColor(244, 247, 250)
WHITE = RGBColor(255, 255, 255)
ORANGE = RGBColor(238, 143, 43)
WARM_YELLOW = RGBColor(248, 196, 89)
GREY = RGBColor(128, 137, 148)
DARK_GREY = RGBColor(58, 65, 73)
GREEN = RGBColor(73, 143, 93)
RED = RGBColor(176, 70, 66)


def add_shape(slide, shape_type, x, y, w, h, fill, line=None, radius=False):
    shape = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = Pt(1.2)
    return shape


def add_text(
    slide,
    text,
    x,
    y,
    w,
    h,
    size=24,
    color=NAVY,
    bold=False,
    align=PP_ALIGN.LEFT,
    font="Aptos",
    italic=False,
):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    box.text_frame.clear()
    box.text_frame.margin_left = Inches(0.08)
    box.text_frame.margin_right = Inches(0.08)
    box.text_frame.margin_top = Inches(0.03)
    box.text_frame.margin_bottom = Inches(0.03)
    box.text_frame.word_wrap = True
    box.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    paragraph = box.text_frame.paragraphs[0]
    paragraph.alignment = align
    run = paragraph.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return box


def add_bullets(slide, bullets, x, y, w, h, size=22, color=DARK_GREY, accent=ORANGE):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.margin_left = Inches(0.08)
    tf.margin_right = Inches(0.08)
    tf.margin_top = Inches(0.02)
    tf.margin_bottom = Inches(0.02)
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    for index, bullet in enumerate(bullets):
        paragraph = tf.paragraphs[0] if index == 0 else tf.add_paragraph()
        paragraph.level = 0
        paragraph.space_after = Pt(10)
        paragraph.line_spacing = 1.08
        paragraph.text = ""
        marker = paragraph.add_run()
        marker.text = "• "
        marker.font.name = "Aptos"
        marker.font.size = Pt(size + 1)
        marker.font.bold = True
        marker.font.color.rgb = accent
        body = paragraph.add_run()
        body.text = bullet
        body.font.name = "Aptos"
        body.font.size = Pt(size)
        body.font.color.rgb = color
    return box


def add_header(slide, title, subtitle=None, dark=False):
    title_color = WHITE if dark else NAVY
    sub_color = RGBColor(215, 223, 234) if dark else SLATE
    add_text(slide, title, 0.7, 0.45, 8.8, 0.6, size=28, color=title_color, bold=True)
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 0.72, 1.16, 1.1, 0.06, ORANGE)
    if subtitle:
        add_text(slide, subtitle, 0.7, 1.28, 9.2, 0.45, size=15, color=sub_color)


def add_footer(slide, number, dark=False):
    color = RGBColor(180, 190, 205) if dark else GREY
    add_text(slide, f"{number:02d}", 12.0, 6.85, 0.55, 0.25, size=10, color=color, bold=True, align=PP_ALIGN.RIGHT)


def add_fact_badge(slide, text, x, y, w, h):
    badge = add_shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, x, y, w, h, NAVY)
    badge.text_frame.clear()
    badge.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    paragraph = badge.text_frame.paragraphs[0]
    paragraph.alignment = PP_ALIGN.CENTER
    run = paragraph.add_run()
    run.text = text
    run.font.name = "Aptos"
    run.font.size = Pt(20)
    run.font.bold = True
    run.font.color.rgb = WHITE
    return badge


def add_card(slide, x, y, w, h, title, body, accent=ORANGE):
    shadow = add_shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, x + 0.06, y + 0.06, w, h, RGBColor(218, 225, 234))
    shadow.fill.transparency = 30
    card = add_shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, x, y, w, h, WHITE, RGBColor(225, 231, 238))
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, x, y, 0.12, h, accent)
    add_text(slide, title, x + 0.25, y + 0.22, w - 0.45, 0.34, size=17, color=NAVY, bold=True)
    add_text(slide, body, x + 0.25, y + 0.74, w - 0.45, h - 0.9, size=14, color=DARK_GREY)
    return card


def add_icon_books(slide, x, y):
    colors = [NAVY, ORANGE, GREEN]
    for i, color in enumerate(colors):
        book = add_shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, x + i * 0.38, y - i * 0.1, 0.32, 1.55, color)
        add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, x + 0.06 + i * 0.38, y + 0.18 - i * 0.1, 0.2, 0.04, WHITE)
        add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, x + 0.06 + i * 0.38, y + 0.34 - i * 0.1, 0.2, 0.04, WHITE)
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, x - 0.1, y + 1.52, 1.5, 0.12, SLATE)


def add_icon_heart(slide, x, y):
    heart = add_shape(slide, MSO_AUTO_SHAPE_TYPE.HEART, x, y, 1.6, 1.35, ORANGE)
    add_text(slide, "ценности", x + 0.18, y + 0.45, 1.25, 0.3, size=13, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.LIGHTNING_BOLT, x + 0.65, y + 0.1, 0.3, 1.1, WHITE)
    return heart


def add_slide_1(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 0, 0, 13.333, 7.5, DEEP_NAVY)
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 8.4, 0, 4.95, 7.5, RGBColor(38, 54, 80))
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 0, 6.65, 13.333, 0.85, ORANGE)
    add_text(slide, "Бедността:", 0.75, 1.15, 6.3, 0.8, size=44, color=WHITE, bold=True)
    add_text(slide, "отвъд празния джоб", 0.75, 2.0, 7.4, 0.75, size=39, color=WARM_YELLOW, bold=True)
    add_text(slide, "Материални, интелектуални и духовни измерения", 0.8, 3.05, 6.7, 0.6, size=20, color=RGBColor(220, 228, 238))
    add_text(slide, "Контрастът показва, че бедността не е само финансова.", 0.82, 5.82, 7.2, 0.35, size=13, color=RGBColor(225, 232, 242))

    # Abstract contrast: glass tower beside an old house.
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 9.15, 1.0, 1.5, 4.95, RGBColor(195, 211, 226), RGBColor(225, 235, 245))
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 10.9, 1.55, 1.15, 4.4, RGBColor(218, 227, 236), RGBColor(235, 241, 247))
    for row in range(6):
        for col in range(2):
            add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 9.38 + col * 0.5, 1.35 + row * 0.7, 0.24, 0.2, RGBColor(76, 101, 130))
    for row in range(5):
        add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 11.17, 1.9 + row * 0.72, 0.46, 0.17, RGBColor(86, 109, 136))
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.ISOSCELES_TRIANGLE, 8.75, 5.0, 2.2, 1.0, RGBColor(121, 75, 55))
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 9.05, 5.55, 1.6, 0.75, RGBColor(190, 158, 119))
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 9.65, 5.75, 0.35, 0.55, DARK_GREY)
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 9.2, 5.75, 0.28, 0.22, RGBColor(242, 213, 161))
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 10.18, 5.75, 0.28, 0.22, RGBColor(242, 213, 161))
    add_footer(slide, 1, dark=True)


def add_slide_2(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 0, 0, 13.333, 7.5, LIGHT_BG)
    add_header(slide, "Какво е бедност?", "Тя е повече от липса на пари.")
    add_text(slide, "Бедността е липса на:", 0.85, 1.9, 4.3, 0.42, size=24, color=NAVY, bold=True)
    add_bullets(
        slide,
        [
            "възможности за развитие",
            "достъп до здравеопазване и образование",
            "достоен и сигурен начин на живот",
        ],
        0.85,
        2.55,
        5.6,
        2.2,
        size=23,
    )
    add_fact_badge(slide, "Линия на бедност в България, 2024 г.: 764 лв.", 0.9, 5.45, 6.3, 0.7)
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, 7.7, 1.55, 4.3, 4.75, WHITE, RGBColor(226, 232, 239))
    add_text(slide, "Материална бедност", 8.05, 2.0, 3.6, 0.35, size=22, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.CLOUD, 8.45, 2.75, 1.1, 0.78, RGBColor(225, 231, 238))
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.LINE_INVERSE, 8.78, 3.15, 0.05, 0.92, RGBColor(225, 231, 238), GREY)
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 9.62, 2.85, 1.35, 1.2, RGBColor(223, 190, 141))
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.ISOSCELES_TRIANGLE, 9.45, 2.35, 1.7, 0.8, RGBColor(126, 81, 62))
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 10.08, 3.35, 0.35, 0.7, DARK_GREY)
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.OVAL, 8.35, 4.45, 2.7, 0.35, RGBColor(210, 218, 228))
    add_text(slide, "Най-видимата форма, но не и единствената.", 8.12, 5.1, 3.5, 0.6, size=15, color=SLATE, align=PP_ALIGN.CENTER)
    add_footer(slide, 2)


def add_slide_3(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 0, 0, 13.333, 7.5, WHITE)
    add_header(slide, "Бедността на ума", "Знанието е единственият капитал, който не се губи.")
    add_bullets(
        slide,
        [
            "Липсата на образование създава цикъл на бедност.",
            "Недостигът на знания ни прави лесни за манипулация.",
            "Образованието отваря избори, професии и глас.",
        ],
        0.85,
        2.05,
        6.15,
        2.55,
        size=22,
        accent=NAVY,
    )
    add_card(
        slide,
        0.9,
        5.35,
        6.5,
        0.9,
        "Статистика",
        "Рискът от бедност при висшистите е около 10 пъти по-нисък, отколкото при хората без образование.",
        accent=NAVY,
    )
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, 8.05, 1.65, 3.8, 4.75, RGBColor(239, 244, 249), RGBColor(222, 230, 238))
    add_icon_books(slide, 8.55, 3.05)
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.LINE_INVERSE, 10.3, 5.1, 1.3, -1.35, WHITE, GREEN)
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.ISOSCELES_TRIANGLE, 11.34, 3.53, 0.35, 0.35, GREEN)
    add_text(slide, "учене → възможности", 8.55, 5.65, 2.9, 0.35, size=17, color=GREEN, bold=True, align=PP_ALIGN.CENTER)
    add_footer(slide, 3)


def add_slide_4(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 0, 0, 13.333, 7.5, DEEP_NAVY)
    add_header(slide, "Духовната бедност", "Празният дух в пълния костюм.", dark=True)
    add_bullets(
        slide,
        [
            "Липса на емпатия, ценности и чувство за отговорност.",
            "Можеш да имаш милиони, но да си просяк в душата си.",
            "Симптоми: безразличие, егоизъм, липса на културни потребности.",
        ],
        0.85,
        2.05,
        6.4,
        3.1,
        size=22,
        color=RGBColor(230, 235, 243),
        accent=WARM_YELLOW,
    )
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, 8.0, 1.7, 3.95, 4.85, RGBColor(25, 44, 72), RGBColor(72, 91, 116))
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.OVAL, 9.32, 2.05, 1.25, 1.25, RGBColor(226, 205, 180))
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.TRAPEZOID, 8.82, 3.35, 2.25, 2.55, DARK_GREY)
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.ISOSCELES_TRIANGLE, 9.33, 3.35, 1.2, 1.15, WHITE)
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.ISOSCELES_TRIANGLE, 9.76, 3.55, 0.3, 0.48, NAVY)
    add_icon_heart(slide, 9.14, 4.62)
    add_text(slide, "Богатство без човечност остава празно.", 8.45, 6.05, 3.05, 0.35, size=14, color=RGBColor(225, 231, 239), align=PP_ALIGN.CENTER)
    add_footer(slide, 4, dark=True)


def add_slide_5(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 0, 0, 13.333, 7.5, LIGHT_BG)
    add_header(slide, "Парадоксът на богатството", "Материалното и духовното богатство невинаги вървят заедно.")
    add_card(
        slide,
        0.85,
        1.9,
        5.45,
        3.15,
        "Материално беден, но духовно богат",
        "Хора с малко средства, но с висок морал, доброта и готовност да помагат. Пример: Дядо Добри.",
        accent=GREEN,
    )
    add_card(
        slide,
        7.05,
        1.9,
        5.45,
        3.15,
        "Материално богат, но духовно беден",
        "Хора с огромни възможности, но без капка човечност, състрадание или културна чувствителност.",
        accent=RED,
    )
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.LEFT_RIGHT_ARROW, 6.23, 3.05, 0.85, 0.55, ORANGE)
    quote = add_shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, 2.0, 5.55, 9.35, 0.78, NAVY)
    quote.text_frame.clear()
    quote.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    paragraph = quote.text_frame.paragraphs[0]
    paragraph.alignment = PP_ALIGN.CENTER
    run = paragraph.add_run()
    run.text = "„Най-бедният човек е този, който има само пари.“"
    run.font.name = "Aptos"
    run.font.size = Pt(24)
    run.font.italic = True
    run.font.bold = True
    run.font.color.rgb = WARM_YELLOW
    add_footer(slide, 5)


def add_slide_6(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 0, 0, 13.333, 7.5, WHITE)
    add_header(slide, "Как се борим с бедността?", "Трите измерения изискват три вида действие.")
    actions = [
        ("1", "Четене и учене", "Борба с бедността на ума чрез знания, критично мислене и образование.", NAVY),
        ("2", "Доброволчество и помощ", "Борба с духовната бедност чрез емпатия, грижа и общност.", ORANGE),
        ("3", "Икономическа активност", "Борба с материалната бедност чрез труд, умения и предприемчивост.", GREEN),
    ]
    for i, (num, title, body, color) in enumerate(actions):
        x = 0.95 + i * 4.05
        add_shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, x, 2.0, 3.35, 3.9, RGBColor(242, 246, 250), RGBColor(222, 230, 238))
        add_shape(slide, MSO_AUTO_SHAPE_TYPE.OVAL, x + 1.23, 2.32, 0.86, 0.86, color)
        add_text(slide, num, x + 1.47, 2.48, 0.38, 0.3, size=21, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
        add_text(slide, title, x + 0.32, 3.45, 2.72, 0.6, size=21, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
        add_text(slide, body, x + 0.42, 4.32, 2.55, 0.95, size=15, color=DARK_GREY, align=PP_ALIGN.CENTER)
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.RIGHT_ARROW, 3.78, 3.62, 0.5, 0.3, WARM_YELLOW)
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.RIGHT_ARROW, 7.83, 3.62, 0.5, 0.3, WARM_YELLOW)
    add_footer(slide, 6)


def add_slide_7(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 0, 0, 13.333, 7.5, DEEP_NAVY)
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, 0, 0, 13.333, 1.2, ORANGE)
    add_text(slide, "Заключение", 0.85, 0.35, 4.5, 0.45, size=31, color=WHITE, bold=True)
    add_text(
        slide,
        "Истинското богатство е в това, което даваш,\nа не само в това, което притежаваш.",
        1.2,
        2.0,
        10.8,
        1.25,
        size=32,
        color=WHITE,
        bold=True,
        align=PP_ALIGN.CENTER,
    )
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, 1.5, 4.25, 10.35, 1.15, RGBColor(29, 51, 82), RGBColor(86, 111, 141))
    add_text(
        slide,
        "Коя бедност ви плаши повече – тази на портфейла или тази на сърцето?",
        1.85,
        4.58,
        9.65,
        0.45,
        size=23,
        color=WARM_YELLOW,
        bold=True,
        align=PP_ALIGN.CENTER,
    )
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.HEART, 5.8, 5.95, 0.7, 0.6, ORANGE)
    add_shape(slide, MSO_AUTO_SHAPE_TYPE.OVAL, 6.75, 5.97, 0.65, 0.65, RGBColor(215, 224, 235))
    add_text(slide, "?", 6.92, 6.08, 0.28, 0.28, size=24, color=NAVY, bold=True, align=PP_ALIGN.CENTER)
    add_footer(slide, 7, dark=True)


def build_deck():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    add_slide_1(prs)
    add_slide_2(prs)
    add_slide_3(prs)
    add_slide_4(prs)
    add_slide_5(prs)
    add_slide_6(prs)
    add_slide_7(prs)
    prs.core_properties.title = "Бедността: отвъд празния джоб"
    prs.core_properties.subject = "Материални, интелектуални и духовни измерения"
    prs.core_properties.language = "bg-BG"
    prs.save(OUTFILE)


if __name__ == "__main__":
    build_deck()
