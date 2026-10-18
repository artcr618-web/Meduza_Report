#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Редизайн: медуза крупно справа, текст слева на чистом фоне.
Ничего не перекрывает объект. Шрифты: Trebuchet MS + Calibri + Arial Black."""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn
from PIL import Image

NAVY  = RGBColor(6, 42, 82)
DEEP  = RGBColor(3, 22, 48)
CYAN  = RGBColor(53, 208, 255)
SUN   = RGBColor(255, 210, 63)
CORAL = RGBColor(230, 80, 60)
WHITE = RGBColor(255, 255, 255)
LIGHT = RGBColor(225, 240, 255)
STEEL = RGBColor(150, 190, 220)
LIME  = RGBColor(170, 230, 120)
PINK  = RGBColor(255, 150, 200)

TITLE_FONT = 'Trebuchet MS'
BODY_FONT = 'Calibri'
DISPLAY_FONT = 'Arial Black'

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
IMG = "doklad_medusy/images/"

# ---------------- helpers ----------------
def fill_alpha(shape, rgb, alpha_pct):
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb
    spPr = shape._element.find(qn('p:spPr'))
    solid = spPr.find(qn('a:solidFill'))
    clr = solid.find(qn('a:srgbClr'))
    a = OxmlElement('a:alpha'); a.set('val', str(int(alpha_pct * 1000)))
    clr.append(a)
    shape.line.fill.background()

def sample_bg(path, factor=0.45):
    im = Image.open(path).convert('RGB')
    w, h = im.size
    strip = im.crop((0, int(h * 0.25), max(1, int(w * 0.15)), int(h * 0.75))).resize((1, 1))
    r, g, b = strip.getpixel((0, 0))
    return RGBColor(int(r * factor), int(g * factor), int(b * factor))

def sidebar(slide, path, six=0.5):
    """Фото справа (объект в точке ~9.3), слева чистый фон + растушёвка."""
    color = sample_bg(path)
    f = slide.background.fill; f.solid(); f.fore_color.rgb = color
    iw, ih = Image.open(path).size
    ai = iw / ih
    W = max(14.6, 8.0 * ai)
    H = W / ai
    x0 = 9.3 - W * six
    y0 = -(H - 7.5) / 2
    slide.shapes.add_picture(path, Inches(x0), Inches(y0), Inches(W), Inches(H))
    # глухая вуаль под текстом + растушёвка к фото
    veil = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(-0.5), Inches(-0.3), Inches(5.7), Inches(8.1))
    fill_alpha(veil, color, 86)
    for i, a in enumerate([58, 32, 12]):
        st = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5.2 + i * 0.75), Inches(-0.3), Inches(0.78), Inches(8.1))
        fill_alpha(st, color, a)
    return color

def tb(slide, l, t, w, h):
    tx = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tx.text_frame; tf.word_wrap = True
    return tf

def _style_run(r, size, bold, color, font, ls, sh):
    r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color; r.font.name = font
    rPr = r._r.get_or_add_rPr()
    if ls:
        rPr.set('spc', str(ls))
    if sh:
        lst = OxmlElement('a:effectLst')
        e = OxmlElement('a:outerShdw')
        e.set('blurRad', '45000'); e.set('dist', '20000'); e.set('dir', '5400000')
        e.set('algn', 'ctr'); e.set('rotWithShape', '0')
        clr = OxmlElement('a:srgbClr'); clr.set('val', '000000')
        al = OxmlElement('a:alpha'); al.set('val', '60000')
        clr.append(al); e.append(clr); lst.append(e); rPr.append(lst)

def para(tf, text, size=20, bold=False, color=WHITE, align=PP_ALIGN.LEFT, space=6,
         first=False, font=BODY_FONT, ls=None, sh=False):
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    p.alignment = align; p.space_after = Pt(space); p.space_before = Pt(0)
    r = p.add_run(); r.text = text
    _style_run(r, size, bold, color, font, ls, sh)
    return p

def bullets(tf, items, size=21, color=WHITE, gap=10, bullet="— ", bcolor=CYAN,
            font=BODY_FONT, lcolor=None):
    for i, it in enumerate(items):
        empty_first = (i == 0 and len(tf.paragraphs) == 1 and tf.paragraphs[0].text == "")
        p = tf.paragraphs[0] if empty_first else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT; p.space_after = Pt(gap); p.space_before = Pt(0)
        lead_c = lcolor or color
        if isinstance(it, tuple):
            b, rest = it
            r = p.add_run(); r.text = bullet
            _style_run(r, size, True, bcolor, font, None, False)
            r2 = p.add_run(); r2.text = b
            _style_run(r2, size, True, lead_c, font, None, False)
            r3 = p.add_run(); r3.text = rest
            _style_run(r3, size, False, color, font, None, False)
        else:
            r = p.add_run(); r.text = bullet + it
            _style_run(r, size, False, color, font, None, False)

def eyebrow_title(tf, eyebrow, title, ecolor=CYAN, tsize=40):
    para(tf, eyebrow, size=17, bold=True, color=ecolor, font=TITLE_FONT, ls=250, first=True, space=6)
    para(tf, title, size=tsize, bold=True, color=WHITE, font=TITLE_FONT, space=12)

def num(slide, n):
    tf = tb(slide, 0.75, 6.92, 2.0, 0.4)
    para(tf, f"{n} / 21", size=14, bold=True, color=STEEL, font=TITLE_FONT, first=True)

def pill(slide, l, t, w, h, text, fill, tcolor, size=19):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = fill
    sh.line.fill.background()
    sh.adjustments[0] = 0.25
    tf = tb(slide, l + 0.15, t + 0.08, w - 0.3, h - 0.16)
    para(tf, text, size=size, bold=True, color=tcolor, align=PP_ALIGN.CENTER, first=True)

def diagram_slide(heading, img, title, key, n):
    s = prs.slides.add_slide(BLANK)
    f = s.background.fill; f.solid(); f.fore_color.rgb = WHITE
    iw, ih = Image.open(IMG + img).size
    H = 6.0; W = H * iw / ih
    s.shapes.add_picture(IMG + img, Inches((13.333 - W) / 2), Inches(0.05), Inches(W), Inches(H))
    band = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(6.15), Inches(13.333), Inches(1.35))
    band.fill.solid(); band.fill.fore_color.rgb = NAVY
    band.line.fill.background()
    tf = tb(s, 0.75, 6.3, 9.5, 1.05)
    para(tf, title, size=26, bold=True, color=WHITE, font=TITLE_FONT, first=True, space=2)
    para(tf, key, size=17, color=CYAN)
    tf2 = tb(s, 11.6, 6.85, 1.2, 0.4)
    para(tf2, f"{n} / 21", size=14, bold=True, color=STEEL, font=TITLE_FONT, align=PP_ALIGN.RIGHT, first=True)
    return s

# ================= 1. TITLE (без подложки) =================
s = prs.slides.add_slide(BLANK)
iw, ih = Image.open(IMG + "pres_swarm.jpg").size
pic = s.shapes.add_picture(IMG + "pres_swarm.jpg", Inches(0), Inches(0), Inches(13.333), Inches(7.5))
at, ai = 13.333 / 7.5, iw / ih
if ai > at:
    c = (1 - at / ai) / 2; pic.crop_left = c; pic.crop_right = c
elif ai < at:
    c = (1 - ai / at) / 2; pic.crop_top = c; pic.crop_bottom = c
tf = tb(s, 0.9, 1.3, 8.5, 5.5)
para(tf, "БИОЛОГИЯ • 8 «Б» КЛАСС", size=20, bold=True, color=CYAN, font=TITLE_FONT, ls=300, sh=True, first=True, space=8)
para(tf, "МЕДУЗЫ", size=100, bold=True, color=WHITE, font=DISPLAY_FONT, sh=True, space=4)
para(tf, "загадочные стрекающие океана", size=30, color=LIGHT, font=TITLE_FONT, sh=True, space=14)
para(tf, "Автор: ученик(ца) 8 «Б» класса", size=20, color=LIGHT, sh=True)
tf2 = tb(s, 11.9, 6.9, 1.0, 0.4)
para(tf2, "01 / 21", size=14, bold=True, color=WHITE, font=TITLE_FONT, align=PP_ALIGN.RIGHT, sh=True, first=True)

# ================= 2. WHO =================
s = prs.slides.add_slide(BLANK)
sidebar(s, IMG + "pres_inside.jpg", six=0.5)
tf = tb(s, 0.75, 0.7, 4.6, 6.0)
eyebrow_title(tf, "КЛАССИФИКАЦИЯ", "Кто такие медузы?", tsize=38)
bullets(tf, [
    ("Царство — ", "Животные"),
    ("Тип — ", "Стрекающие (Cnidaria)"),
    ("Медуза — ", "плавающая стадия"),
    ("Полип — ", "сидячая стадия"),
    ("Классов — ", "три: гидроидные, сцифоидные, кубомедузы"),
], size=20, gap=9)
para(tf, "Полип сидит — медуза плывёт!", size=20, bold=True, color=SUN, font=TITLE_FONT)
num(s, "02")

# ================= 3. ANATOMY =================
diagram_slide(None, "anatomy_aurelia_ru.jpg", "Как устроена аурелия",
              "Найди на схеме: гонады • стрекательные клетки • ропалии", "03")

# ================= 4. MACRO =================
s = prs.slides.add_slide(BLANK)
sidebar(s, IMG + "pres_macro_arms.jpg", six=0.5)
tf = tb(s, 0.75, 0.7, 4.6, 6.0)
eyebrow_title(tf, "КРУПНЫЙ ПЛАН", "Заглянем внутрь", ecolor=SUN, tsize=38)
bullets(tf, [
    ("«Лапки» — ", "кружевные ротовые лопасти"),
    ("Гонады — ", "4 светящиеся подковы"),
    ("Каналы — ", "разносят пищу по куполу"),
    ("Край — ", "щупальца и органы чувств"),
], size=20, gap=10, bcolor=SUN)
para(tf, "Всё просвечивает сквозь купол!", size=20, bold=True, color=CYAN, font=TITLE_FONT)
num(s, "04")

# ================= 5. WEAPON =================
s = prs.slides.add_slide(BLANK)
sidebar(s, IMG + "pres_tentacles.jpg", six=0.5)
tf = tb(s, 0.75, 0.7, 4.6, 6.0)
eyebrow_title(tf, "СУПЕРОРУЖИЕ", "Выстрел за 0,003 секунды", ecolor=PINK, tsize=34)
bullets(tf, [
    ("Книдоцит — ", "живая гарпунная пушка"),
    ("Ускорение — ", "5 000 000 g!"),
    ("Яд — ", "парализует мгновенно"),
    ("Срабатывает — ", "сама от прикосновения"),
], size=20, gap=10, bcolor=PINK)
num(s, "05")

# ================= 6. MOVE =================
s = prs.slides.add_slide(BLANK)
sidebar(s, IMG + "pres_hunt.jpg", six=0.45)
tf = tb(s, 0.75, 0.7, 4.6, 6.0)
eyebrow_title(tf, "ДВИЖЕНИЕ И ОХОТА", "Живой реактивный двигатель", tsize=34)
bullets(tf, [
    ("Сжал купол — ", "вытолкнул воду — толчок!"),
    ("Щупальца — ", "рыболовная сеть"),
    ("Все медузы — ", "хищники"),
    ("Рот один на всё — ", "и ест, и выводит отходы"),
], size=20, gap=10)
num(s, "06")

# ================= 7. LIFECYCLE =================
diagram_slide(None, "lifecycle_ru.jpg", "6 превращений за жизнь",
              "медуза → яйца → планула → полип → стробиляция → эфира", "07")

# ================= 8. HYDRO =================
s = prs.slides.add_slide(BLANK)
sidebar(s, IMG + "pres_hydroid.jpg", six=0.55)
tf = tb(s, 0.75, 0.7, 4.6, 5.9)
eyebrow_title(tf, "ГИДРОИДНЫЕ", "Крошки и колонии", ecolor=LIME, tsize=38)
bullets(tf, [
    ("Рост — ", "всего 1–4 см"),
    ("Кораблик — ", "целая колония под парусом!"),
    ("Щупальца — ", "до 30 метров вглубь"),
], size=20, gap=10, bcolor=LIME)
para(tf, "★ Turritopsis умеет молодеть!", size=20, bold=True, color=SUN, font=TITLE_FONT)
num(s, "08")

# ================= 9. SCYPH =================
s = prs.slides.add_slide(BLANK)
sidebar(s, IMG + "pres_cyanea.jpg", six=0.35)
tf = tb(s, 0.75, 0.7, 4.0, 5.9)
eyebrow_title(tf, "СЦИФОИДНЫЕ", "Гиганты морей", ecolor=SUN, tsize=38)
bullets(tf, [
    ("Аурелия — ", "4 колечка, безопасна"),
    ("Корнерот — ", "кружево Чёрного моря"),
    ("Живут — ", "только в морях"),
], size=20, gap=10, bcolor=SUN)
para(tf, "РЕКОРД: цианея 37 м — длиннее кита!", size=20, bold=True, color=SUN, font=TITLE_FONT)
num(s, "09")

# ================= 10. BOX =================
s = prs.slides.add_slide(BLANK)
sidebar(s, IMG + "pres_box.jpg", six=0.64)
tf = tb(s, 0.75, 0.6, 4.6, 5.2)
eyebrow_title(tf, "КУБОМЕДУЗЫ", "Опасные красавицы", tsize=36)
bullets(tf, [
    ("Зонтик — ", "кубический!"),
    ("Скорость — ", "6 км/ч"),
    ("Глаза — ", "24, есть с хрусталиком"),
], size=20, gap=10)
pill(s, 0.75, 5.65, 4.4, 0.95, "Морская оса убивает за 2–5 минут!", CORAL, WHITE, 18)
num(s, "10")

# ================= 11. HARMLESS =================
s = prs.slides.add_slide(BLANK)
sidebar(s, IMG + "pres_aurelia.jpg", six=0.5)
tf = tb(s, 0.75, 0.6, 4.6, 5.2)
eyebrow_title(tf, "БЕЗОПАСНОСТЬ", "Безобидные медузы", ecolor=LIME, tsize=36)
bullets(tf, [
    ("Аурелия — ", "почти не жжётся"),
    ("Корнерот — ", "как крапива"),
    ("Гребневики — ", "не стрекаются вовсе!"),
], size=20, gap=10, bcolor=LIME)
pill(s, 0.75, 5.65, 4.4, 0.95, "Незнакомую медузу не трогай!", LIME, DEEP, 18)
num(s, "11")

# ================= 12. DANGER =================
s = prs.slides.add_slide(BLANK)
sidebar(s, IMG + "pres_macro_edge.jpg", six=0.5)
tf = tb(s, 0.75, 0.5, 4.6, 5.3)
eyebrow_title(tf, "ПЕРВАЯ ПОМОЩЬ", "Ожог! Что делать?", ecolor=CORAL, tsize=36)
bullets(tf, [
    ("1 — ", "выйди из воды, не три кожу"),
    ("2 — ", "щупальца убери картой, не руками"),
    ("3 — ", "промой уксусом или морской водой"),
    ("4 — ", "боль не проходит — к врачу"),
], size=19, gap=8, bcolor=CORAL, lcolor=CORAL)
pill(s, 0.75, 5.85, 4.4, 0.8, "НЕЛЬЗЯ: пресная вода • песок • спирт", CORAL, WHITE, 16)
num(s, "12")

# ================= 13. BENEFIT =================
s = prs.slides.add_slide(BLANK)
sidebar(s, IMG + "pres_turtle.jpg", six=0.6)
tf = tb(s, 0.75, 0.7, 4.6, 6.0)
eyebrow_title(tf, "ПОЛЬЗА", "Чем полезны?", tsize=40)
bullets(tf, [
    ("Еда — ", "салат из медуз, 1700 лет в Азии"),
    ("Медицина — ", "коллаген и лекарства из яда"),
    ("Наука — ", "Нобелевка-2008 за свечение"),
    ("Природа — ", "еда черепах и рыб"),
], size=20, gap=11)
num(s, "13")

# ================= 14. FACT DINO =================
s = prs.slides.add_slide(BLANK)
sidebar(s, IMG + "fact_dino.jpg", six=0.5)
tf = tb(s, 0.75, 0.7, 4.6, 6.0)
eyebrow_title(tf, "ФАКТ 1 ИЗ 7", "Старше динозавров!", ecolor=SUN, tsize=38)
bullets(tf, [
    ("Медузам — ", "650 000 000 лет"),
    ("Динозаврам — ", "«всего» 230 миллионов"),
    ("Пережили — ", "5 массовых вымираний!"),
], size=20, gap=10, bcolor=SUN)
num(s, "14")

# ================= 15. FACT WATER =================
s = prs.slides.add_slide(BLANK)
sidebar(s, IMG + "fact_water.jpg", six=0.5)
tf = tb(s, 0.75, 0.7, 4.6, 6.0)
eyebrow_title(tf, "ФАКТ 2 ИЗ 7", "Тело — почти вода", tsize=38)
bullets(tf, [
    ("Вода — ", "до 98% тела!"),
    ("Остальное — ", "чуть-чуть белка-желе"),
    ("На берегу — ", "тает, как мороженое"),
], size=20, gap=10)
num(s, "21")

# ================= 16. FACT WHALE =================
s = prs.slides.add_slide(BLANK)
sidebar(s, IMG + "fact_whale.jpg", six=0.5)
tf = tb(s, 0.75, 0.7, 4.6, 6.0)
eyebrow_title(tf, "ФАКТ 3 ИЗ 7", "Длиннее синего кита!", ecolor=SUN, tsize=36)
bullets(tf, [
    ("Цианея — ", "щупальца до 37 метров"),
    ("Синий кит — ", "до 33 метров"),
    ("Вывод — ", "медуза длиннее великана!"),
], size=20, gap=10, bcolor=SUN)
num(s, "16")

# ================= 17. FACT EYES =================
s = prs.slides.add_slide(BLANK)
sidebar(s, IMG + "fact_eyes.jpg", six=0.5)
tf = tb(s, 0.75, 0.6, 4.6, 3.4)
eyebrow_title(tf, "ФАКТ 4 ИЗ 7", "24 глаза. Считаем!", tsize=36)
bullets(tf, [
    ("У куба — ", "4 стороны"),
    ("На каждой — ", "6 глаз"),
    ("Итого — ", "4 × 6 = 24"),
], size=20, gap=8)
dot_colors = [CYAN, SUN, PINK, LIME]
for g in range(4):
    for k in range(6):
        dot = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.85 + k * 0.42), Inches(4.35 + g * 0.42), Inches(0.3), Inches(0.3))
        dot.fill.solid(); dot.fill.fore_color.rgb = dot_colors[g]
        dot.line.fill.background()
tf2 = tb(s, 0.75, 6.15, 4.6, 0.5)
para(tf2, "4 цвета — 4 стороны куба. Пересчитай!", size=17, bold=True, color=LIGHT, font=TITLE_FONT, first=True)
num(s, "17")

# ================= 18. FACT IMMORTAL =================
s = prs.slides.add_slide(BLANK)
sidebar(s, IMG + "fact_immortal.jpg", six=0.56)
tf = tb(s, 0.75, 0.7, 4.6, 6.0)
eyebrow_title(tf, "ФАКТ 5 ИЗ 7", "Бессмертная медуза", ecolor=PINK, tsize=36)
bullets(tf, [
    ("Беда? — ", "превращается обратно в полип"),
    ("Клетки — ", "обновляются и молодеют"),
    ("Цикл — ", "начинается заново. И так без конца!"),
    ("Учёные — ", "изучают её секрет старения"),
], size=20, gap=10, bcolor=PINK)
num(s, "18")

# ================= 19. FACT STORM =================
s = prs.slides.add_slide(BLANK)
sidebar(s, IMG + "fact_storm.jpg", six=0.5)
tf = tb(s, 0.75, 0.7, 4.6, 6.0)
eyebrow_title(tf, "ФАКТ 6 ИЗ 7", "Живой барометр", tsize=40)
bullets(tf, [
    ("Слышит — ", "гул шторма за 15 часов"),
    ("Как — ", "тело ловит инфразвук волн"),
    ("Что делает — ", "уходит на глубину"),
], size=20, gap=10)
num(s, "19")

# ================= 20. FACT SPACE =================
s = prs.slides.add_slide(BLANK)
sidebar(s, IMG + "fact_space.jpg", six=0.6)
tf = tb(s, 0.75, 0.7, 4.6, 6.0)
eyebrow_title(tf, "ФАКТ 7 ИЗ 7", "Летали в космос!", ecolor=PINK, tsize=40)
bullets(tf, [
    ("1991 год — ", "шаттл «Колумбия», 2500 полипов"),
    ("В невесомости — ", "полипы размножались"),
    ("Но — ", "«космические» медузы не держали равновесие: кружилась голова!"),
], size=20, gap=10, bcolor=PINK)
num(s, "20")

# ================= 21. FINALE =================
s = prs.slides.add_slide(BLANK)
color = sample_bg(IMG + "diversity_bg.jpg")
f = s.background.fill; f.solid(); f.fore_color.rgb = color
iw, ih = Image.open(IMG + "diversity_bg.jpg").size
H = 8.1; W = H * iw / ih
x0 = 13.33 - W + 0.4
s.shapes.add_picture(IMG + "diversity_bg.jpg", Inches(x0), Inches(-0.3), Inches(W), Inches(H))
veil = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(-0.5), Inches(-0.3), Inches(7.5), Inches(8.1))
fill_alpha(veil, color, 88)
for i, a in enumerate([60, 34, 12]):
    st = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.0 + i * 0.75), Inches(-0.3), Inches(0.78), Inches(8.1))
    fill_alpha(st, color, a)
tf = tb(s, 0.75, 1.3, 5.6, 5.2)
para(tf, "Спасибо за внимание!", size=46, bold=True, color=SUN, font=TITLE_FONT, first=True, space=14)
para(tf, "Без мозга и сердца — а живут 650 млн лет", size=22, color=WHITE, space=8)
para(tf, "Хищники с гарпунами быстрее пули", size=22, color=WHITE, space=8)
para(tf, "Красота океана, которую надо беречь", size=22, color=WHITE, space=16)
para(tf, "Вопросы?", size=30, bold=True, color=CYAN, font=TITLE_FONT)
num(s, "21")

prs.save("doklad_medusy/Prezentaciya_Meduzy_8klass.pptx")
print("pptx saved, slides:", len(prs.slides.__iter__.__self__._sldIdLst))
