# -*- coding: utf-8 -*-
"""
Пересборка презентации ВКР по замечаниям научного руководителя.

Что делает скрипт:
  * убирает слово «актуальность»; вместо одного слайда делает ДВА слайда,
    оба с заголовком «Важность изучения топоизомеразы II, кодируемого геном TOP2»;
  * на первом слайде важности — ключевые пункты + наглядная схема-процесс;
  * на втором — векторная двойная спираль ДНК с подсветкой мишени
    (топоизомеразы II) и фразой-переходом к цели работы;
  * заполняет слайды результатов экспериментальными изображениями
    (электрофореграммы ПЦР и кривая плавления);
  * добавляет в заметки к титульному слайду вступительную речь для защиты;
  * приводит презентацию ровно к 16 слайдам.

Исходник не меняется. Результат сохраняется в отдельный файл.
Запуск:  python3 build_presentation.py
"""

import zipfile
import re
import struct
import math
import html
import os

SRC = "Бахромов презентация(1) (2).pptx"
OUT = "Бахромов_презентация_TOP2_16_слайдов.pptx"

# 16:9, размеры слайда в EMU
SLIDE_W = 12192000
SLIDE_H = 6858000

TNR = ('<a:latin typeface="Times New Roman" panose="02020603050405020304" '
       'pitchFamily="18" charset="0"/><a:cs typeface="Times New Roman" '
       'panose="02020603050405020304" pitchFamily="18" charset="0"/>')


def esc(t):
    return html.escape(str(t), quote=True)


def jpeg_size(b):
    """Размеры JPEG из маркера SOF."""
    i = 2
    while i < len(b) - 9:
        if b[i] != 0xFF:
            i += 1
            continue
        m = b[i + 1]
        if 0xC0 <= m <= 0xCF and m not in (0xC4, 0xC8, 0xCC):
            h = struct.unpack('>H', b[i + 5:i + 7])[0]
            w = struct.unpack('>H', b[i + 7:i + 9])[0]
            return w, h
        seg = struct.unpack('>H', b[i + 2:i + 4])[0]
        i += 2 + seg
    return (1600, 1000)


# ----------------------------------------------------------------------
#  Чтение исходного архива
# ----------------------------------------------------------------------
zin = zipfile.ZipFile(SRC)
parts = {}
for item in zin.infolist():
    parts[item.filename] = zin.read(item.filename)
zin.close()


def text(name):
    return parts[name].decode('utf-8')


# ----------------------------------------------------------------------
#  Извлечение «фирменных» элементов из slide14 (чистый пустой шаблон)
# ----------------------------------------------------------------------
S14 = text("ppt/slides/slide14.xml")

# префикс spTree: корневая группа
m = re.search(r'(<p:nvGrpSpPr>.*?</p:grpSpPr>)', S14, re.S)
ROOT_GRP = m.group(1)

# горизонтальная синяя линия под заголовком (object 3, id=3)
m = re.search(r'(<p:sp><p:nvSpPr><p:cNvPr id="3" name="object 3".*?</p:sp>)', S14, re.S)
ACCENT_LINE = m.group(1)

# группа с логотипом РХТУ в правом верхнем углу (object 4)
m = re.search(r'(<p:grpSp><p:nvGrpSpPr><p:cNvPr id="4" name="object 4".*?</p:grpSp>)', S14, re.S)
LOGO_GRP = m.group(1)

# плейсхолдер номера слайда (id=14)
m = re.search(r'(<p:sp><p:nvSpPr><p:cNvPr id="14" name="Номер слайда 13".*?</p:sp>)', S14, re.S)
SLIDE_NUM = m.group(1)

CHROME = ACCENT_LINE + LOGO_GRP + SLIDE_NUM
CLRMAP = '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
SLD_OPEN = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
            '<p:cSld><p:spTree>')
SLD_CLOSE = '</p:spTree></p:cSld>' + CLRMAP + '</p:sld>'


def slide_xml(body_shapes):
    return SLD_OPEN + ROOT_GRP + CHROME + body_shapes + SLD_CLOSE


# ----------------------------------------------------------------------
#  Помощники построения фигур
# ----------------------------------------------------------------------
def run(t, *, sz=2000, b=False, i=False, color=None):
    rpr = '<a:rPr lang="ru-RU" altLang="en-US" sz="%d"%s%s dirty="0">' % (
        sz, ' b="1"' if b else '', ' i="1"' if i else '')
    fill = ('<a:solidFill><a:srgbClr val="%s"/></a:solidFill>' % color) if color else ''
    return rpr + fill + TNR + '</a:rPr><a:t>' + esc(t) + '</a:t>'


def para(runs_xml, *, algn='l', bullet=False, sz=2000, after=600, level=0):
    pPr = '<a:pPr marL="%d" indent="%d" algn="%s">' % (
        (342900 if bullet else 0), (-342900 if bullet else 0), algn)
    pPr += '<a:spcBef><a:spcPts val="%d"/></a:spcBef>' % after
    if bullet:
        pPr += ('<a:buFont typeface="Arial" panose="020B0604020202020204" '
                'pitchFamily="34" charset="0"/><a:buChar char="\u2022"/>')
    else:
        pPr += '<a:buNone/>'
    pPr += '</a:pPr>'
    body = ''.join('<a:r>' + r + '</a:r>' for r in runs_xml)
    return '<a:p>' + pPr + body + '</a:p>'


def title_box(text_str, sid):
    """Заголовок слайда в стиле исходной презентации."""
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="Заголовок"/><p:cNvSpPr txBox="1"/>'
        '<p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="839470" y="431075"/>'
        '<a:ext cx="9658350" cy="540000"/></a:xfrm><a:prstGeom prst="rect">'
        '<a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody>'
        '<a:bodyPr wrap="square" rtlCol="0"><a:normAutofit/></a:bodyPr><a:lstStyle/>'
        '<a:p><a:pPr><a:buNone/></a:pPr><a:r><a:rPr lang="ru-RU" altLang="en-US" '
        'sz="2400" b="1" dirty="0">%s</a:rPr><a:t>%s</a:t></a:r></a:p>'
        '</p:txBody></p:sp>' % (sid, TNR, esc(text_str)))


def textbox(sid, x, y, cx, cy, paragraphs, *, anchor='t', name="Текст"):
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr txBox="1"/>'
        '<p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="%d" y="%d"/>'
        '<a:ext cx="%d" cy="%d"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/>'
        '</a:prstGeom><a:noFill/></p:spPr><p:txBody>'
        '<a:bodyPr wrap="square" rtlCol="0" anchor="%s"><a:normAutofit/></a:bodyPr>'
        '<a:lstStyle/>%s</p:txBody></p:sp>'
        % (sid, name, x, y, cx, cy, anchor, ''.join(paragraphs)))


def round_box(sid, x, y, cx, cy, lines, fill, *, font=1500, txtcolor="FFFFFF"):
    paras = []
    for j, ln in enumerate(lines):
        paras.append(para([run(ln, sz=font, b=True, color=txtcolor)],
                          algn='ctr', after=(0 if j == 0 else 200)))
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 12000"/>'
        '</a:avLst></a:prstGeom><a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        '<a:ln w="9525"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>'
        '</p:spPr><p:txBody><a:bodyPr wrap="square" anchor="ctr" lIns="68580" '
        'rIns="68580" tIns="27432" bIns="27432"><a:normAutofit/></a:bodyPr>'
        '<a:lstStyle/>%s</p:txBody></p:sp>'
        % (sid, x, y, cx, cy, fill, ''.join(paras)))


def down_arrow(sid, x, y, cx, cy, fill):
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="arrow"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="downArrow"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="%s"/></a:solidFill></p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:buNone/></a:pPr>'
        '<a:endParaRPr lang="ru-RU"/></a:p></p:txBody></p:sp>'
        % (sid, x, y, cx, cy, fill))


def freeform(sid, x, y, w, h, pts, color, width):
    """Полилиния (незамкнутая) — лента остова спирали."""
    path = '<a:moveTo><a:pt x="%d" y="%d"/></a:moveTo>' % pts[0]
    for p in pts[1:]:
        path += '<a:lnTo><a:pt x="%d" y="%d"/></a:lnTo>' % p
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="strand"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>'
        '<a:rect l="l" t="t" r="r" b="b"/><a:pathLst>'
        '<a:path w="%d" h="%d"><a:moveTo><a:pt x="%d" y="%d"/></a:moveTo>%s</a:path>'
        '</a:pathLst></a:custGeom><a:ln w="%d" cap="rnd"><a:solidFill>'
        '<a:srgbClr val="%s"/></a:solidFill></a:ln></p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
        % (sid, x, y, w, h, w, h, pts[0][0], pts[0][1],
           ''.join('<a:lnTo><a:pt x="%d" y="%d"/></a:lnTo>' % p for p in pts[1:]),
           width, color))


def line_shape(sid, x1, y1, x2, y2, color, width):
    x, y = min(x1, x2), min(y1, y2)
    cx, cy = abs(x2 - x1), abs(y2 - y1)
    flipH = '1' if x2 < x1 else '0'
    flipV = '1' if y2 < y1 else '0'
    return (
        '<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="%d" name="rung"/><p:cNvCxnSpPr/>'
        '<p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm flipH="%s" flipV="%s">'
        '<a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="line"><a:avLst/></a:prstGeom>'
        '<a:ln w="%d"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:ln>'
        '</p:spPr></p:cxnSp>'
        % (sid, flipH, flipV, x, y, max(cx, 1), max(cy, 1), width, color))


def ellipse(sid, x, y, cx, cy, line_color, line_w, dash=True):
    dash_xml = '<a:prstDash val="dash"/>' if dash else ''
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="target"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:noFill/>'
        '<a:ln w="%d"><a:solidFill><a:srgbClr val="%s"/></a:solidFill>%s</a:ln>'
        '</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
        % (sid, x, y, cx, cy, line_w, line_color, dash_xml))


def picture(sid, rid, x, y, cx, cy, name="Рисунок"):
    return (
        '<p:pic><p:nvPicPr><p:cNvPr id="%d" name="%s"/><p:cNvPicPr>'
        '<a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
        '<p:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch>'
        '</p:blipFill><p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/>'
        '</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:ln w="12700"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln>'
        '</p:spPr></p:pic>' % (sid, name, rid, x, y, cx, cy))


def fit_box(img_w, img_h, max_w, max_h):
    """Вписать изображение в рамку с сохранением пропорций."""
    scale = min(max_w / img_w, max_h / img_h)
    return int(img_w * scale), int(img_h * scale)


TITLE_TXT = "Важность изучения топоизомеразы II, кодируемого геном TOP2"
BLUE = "1F6FB2"
TEAL = "2E9E8F"
GREEN = "4F8A2E"
ORANGE = "C55A11"
DARK = "1F3864"
GREY = "808080"


# ======================================================================
#  СЛАЙД ВАЖНОСТИ № 1  (ключевые пункты + схема-процесс)
# ======================================================================
def build_importance_1():
    sh = title_box(TITLE_TXT, 100)

    # ---- левая колонка: текст и маркеры ----
    paras = []
    paras.append(para([run("Регуляция топологии ДНК — фундаментальный механизм "
                            "сохранения стабильности генома и нормального деления "
                            "клетки. Нарушение этих процессов напрямую связано с "
                            "развитием тяжёлых патологий, прежде всего "
                            "онкологических заболеваний.", sz=1700)],
                       algn='just', after=0))
    paras.append(para([run("Ферментом, управляющим структурой ДНК, является "
                            "топоизомераза II, кодируемая геном ", sz=1700),
                       run("TOP2", sz=1700, b=True),
                       run(". Изучение её экспрессии позволяет оценить "
                           "пролиферативную активность клеток.", sz=1700)],
                       algn='just', after=600))
    paras.append(para([run("Изучение экспрессии гена TOP2 необходимо для:",
                            sz=1700, b=True)], after=400))
    for head, tail in [
        ("молекулярной диагностики", " онкологических заболеваний;"),
        ("оценки скорости деления", " (пролиферации) клеток;"),
        ("исследования эффективности", " противоопухолевых препаратов "
         "(многие нацелены именно на TOP2);"),
        ("анализа механизмов", " стабильности генома."),
    ]:
        paras.append(para([run(head, sz=1700, b=True), run(tail, sz=1700)],
                          bullet=True, after=200))
    sh += textbox(101, 600000, 1180000, 6050000, 5200000, paras, name="Текст важности")

    # ---- правая колонка: схема-процесс (4 блока + стрелки) ----
    bx = 7050000
    bw = 4450000
    bh = 760000
    gap = 360000
    y = 1280000
    sh += round_box(110, bx, y, bw, bh,
                    ["Ген TOP2"], DARK, font=1600)
    y += bh
    sh += down_arrow(111, bx + bw // 2 - 130000, y, 260000, gap, BLUE)
    y += gap
    sh += round_box(112, bx, y, bw, bh,
                    ["Фермент топоизомераза II"], BLUE, font=1600)
    y += bh
    sh += down_arrow(113, bx + bw // 2 - 130000, y, 260000, gap, TEAL)
    y += gap
    sh += round_box(114, bx, y, bw, bh,
                    ["Контроль топологии ДНК,", "репликация и деление клетки"],
                    TEAL, font=1400)
    y += bh
    sh += down_arrow(115, bx + bw // 2 - 130000, y, 260000, gap, ORANGE)
    y += gap
    sh += round_box(116, bx, y, bw, bh,
                    ["Норма \u2194 патология (рак)"], ORANGE, font=1500)
    return slide_xml(sh)


# ======================================================================
#  СЛАЙД ВАЖНОСТИ № 2  (двойная спираль ДНК + мишень + переход к цели)
# ======================================================================
def build_importance_2():
    sh = title_box(TITLE_TXT, 100)

    # ---- левая колонка: текст + фраза-переход ----
    paras = []
    paras.append(para([run("Топоизомераза II временно разрезает обе цепи ДНК, "
                            "пропускает через разрыв другой участок двойной спирали "
                            "и снова сшивает её. Так фермент снимает "
                            "сверхскрученность и расцепляет дочерние молекулы ДНК "
                            "при делении клетки.", sz=1700)],
                       algn='just', after=0))
    paras.append(para([run("Именно двойная спираль ДНК и работающая на ней "
                            "топоизомераза II — та ", sz=1700),
                       run("молекулярная мишень", sz=1700, b=True, color=ORANGE),
                       run(", на которую направлены наши праймеры и "
                           "противоопухолевые препараты.", sz=1700)],
                       algn='just', after=700))
    sh += textbox(101, 600000, 1180000, 5450000, 3500000, paras, name="Текст важности 2")

    # фраза-переход к цели — выделенный блок снизу
    bridge = para([run("Именно поэтому целью моей работы стало изучение "
                       "экспрессии гена TOP2 методом полимеразной цепной реакции.",
                       sz=1700, b=True, color="FFFFFF")], algn='ctr', after=0)
    sh += (
        '<p:sp><p:nvSpPr><p:cNvPr id="105" name="bridge"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="600000" y="5050000"/><a:ext cx="5450000" cy="1150000"/>'
        '</a:xfrm><a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 10000"/>'
        '</a:avLst></a:prstGeom><a:solidFill><a:srgbClr val="%s"/></a:solidFill></p:spPr>'
        '<p:txBody><a:bodyPr wrap="square" anchor="ctr" lIns="91440" rIns="91440">'
        '<a:normAutofit/></a:bodyPr><a:lstStyle/>%s</p:txBody></p:sp>' % (DARK, bridge))

    # ---- правая колонка: векторная двойная спираль ДНК ----
    # рабочая область спирали
    X0 = 6650000          # левый край области
    BOXW = 1700000        # ширина «коробки» одной ленты
    xc_local = BOXW // 2  # центр по локали
    A = 760000            # амплитуда
    Y0 = 1350000          # верх
    Hh = 4350000          # высота области спирали
    turns = 2.6
    N = 64

    def strand_points(phase):
        pts = []
        for k in range(N + 1):
            s = k / N
            t = s * turns * 2 * math.pi + phase
            xl = xc_local + int(A * math.sin(t))
            yl = int(s * Hh)
            pts.append((xl, yl))
        return pts

    ptsA = strand_points(0.0)
    ptsB = strand_points(math.pi)
    sh += freeform(120, X0, Y0, BOXW, Hh, ptsA, BLUE, 57150)
    sh += freeform(121, X0, Y0, BOXW, Hh, ptsB, TEAL, 57150)

    # перекладины (пары оснований)
    sid = 130
    rung_idx = [int(N * f) for f in
                [0.07, 0.16, 0.25, 0.34, 0.43, 0.52, 0.61, 0.70, 0.79, 0.88, 0.96]]
    for k in rung_idx:
        s = k / N
        t = s * turns * 2 * math.pi
        xa = X0 + xc_local + int(A * math.sin(t))
        xb = X0 + xc_local + int(A * math.sin(t + math.pi))
        yy = Y0 + int(s * Hh)
        col = ORANGE if k in (rung_idx[5],) else "9DC3E6"
        sh += line_shape(sid, xa, yy, xb, yy, col, 28575)
        sid += 1

    # подсветка мишени (на центральной перекладине)
    s_mid = rung_idx[5] / N
    y_mid = Y0 + int(s_mid * Hh)
    sh += ellipse(140, X0 + xc_local - 620000, y_mid - 470000, 1240000, 940000,
                  ORANGE, 38100, dash=True)
    sh += textbox(141, X0 - 250000, y_mid + 520000, BOXW + 500000, 700000,
                  [para([run("Мишень: топоизомераза II", sz=1300, b=True, color=ORANGE)],
                        algn='ctr', after=0)], anchor='t', name="метка мишени")
    # подпись «ДНК»
    sh += textbox(142, X0 - 250000, Y0 - 250000, BOXW + 500000, 360000,
                  [para([run("Двойная спираль ДНК", sz=1300, b=True, color=DARK)],
                        algn='ctr', after=0)], anchor='t', name="метка ДНК")
    return slide_xml(sh)


# ======================================================================
#  СЛАЙДЫ РЕЗУЛЬТАТОВ  (экспериментальные изображения)
# ======================================================================
def build_image_slide(title, rid, img_w, img_h, caption, side_paras):
    """Слайд: слева изображение с подписью, справа текст."""
    sh = title_box(title, 100)
    # изображение слева
    max_w, max_h = 5750000, 4350000
    w, h = fit_box(img_w, img_h, max_w, max_h)
    x = 650000
    y = 1300000 + (max_h - h) // 2
    sh += picture(150, rid, x, y, w, h)
    sh += textbox(151, 650000, 1300000 + max_h + 40000, max_w, 520000,
                  [para([run(caption, sz=1300, i=True, color="595959")],
                        algn='ctr', after=0)], name="подпись")
    # текст справа
    sh += textbox(152, 6650000, 1320000, 4900000, 4600000, side_paras, name="текст")
    return slide_xml(sh)


def build_qc_slide():
    title = "Контроль качества: электрофорез продуктов ПЦР"
    iw, ih = jpeg_size(parts["__gel1"])
    side = [
        para([run("Подтверждение специфичности и качества синтезированных "
                  "праймеров.", sz=1700, b=True)], algn='just', after=300),
        para([run("На электрофореграмме виден единственный чёткий продукт "
                  "ожидаемого размера (~150 п.н.) — без неспецифических полос и "
                  "праймер-димеров.", sz=1600)], algn='just', bullet=True, after=200),
        para([run("Размер ампликона соответствует расчётному для экзонного стыка "
                  "гена TOP2A.", sz=1600)], algn='just', bullet=True, after=200),
        para([run("Результат подтверждает корректность подбора, синтеза и очистки "
                  "олигонуклеотидов.", sz=1600)], algn='just', bullet=True, after=200),
    ]
    return build_image_slide(title, "rId7", iw, ih,
                             "Рисунок — электрофореграмма продуктов ПЦР (агарозный гель)",
                             side)


def build_res1_slide():
    title = "Результаты ПЦР: детекция ампликона гена TOP2A"
    iw, ih = jpeg_size(parts["__gel2"])
    side = [
        para([run("Амплификация целевого фрагмента гена TOP2A.", sz=1700, b=True)],
             algn='just', after=300),
        para([run("В дорожках с матрицей (кДНК) наблюдается специфический "
                  "ампликон; в отрицательном контроле полоса отсутствует.",
                  sz=1600)], algn='just', bullet=True, after=200),
        para([run("Положение полосы относительно маркера длин соответствует "
                  "ожидаемому размеру продукта.", sz=1600)],
             algn='just', bullet=True, after=200),
        para([run("Сконструированные праймеры пригодны для анализа экспрессии "
                  "гена TOP2A методом ПЦР.", sz=1600)],
             algn='just', bullet=True, after=200),
    ]
    return build_image_slide(title, "rId7", iw, ih,
                             "Рисунок — электрофореграмма ПЦР-продукта гена TOP2A",
                             side)


def build_res2_slide():
    title = "ПЦР в реальном времени: кривая плавления и выводы"
    iw, ih = jpeg_size(parts["__melt"])
    side = [
        para([run("Анализ кривой плавления продукта ПЦР-РВ.", sz=1700, b=True)],
             algn='just', after=300),
        para([run("Единственный пик плавления подтверждает образование одного "
                  "специфического продукта без побочных структур.", sz=1600)],
             algn='just', bullet=True, after=200),
        para([run("Выводы:", sz=1700, b=True, color=DARK)], after=200),
        para([run("подобраны и синтезированы специфичные праймеры к гену TOP2;",
                  sz=1500)], algn='just', bullet=True, after=160),
        para([run("их пригодность подтверждена электрофорезом и кривой плавления;",
                  sz=1500)], algn='just', bullet=True, after=160),
        para([run("отработан алгоритм анализа экспрессии гена TOP2 методом ПЦР.",
                  sz=1500)], algn='just', bullet=True, after=160),
    ]
    return build_image_slide(title, "rId7", iw, ih,
                             "Рисунок — кривая плавления продукта ПЦР в реальном времени",
                             side)


# ======================================================================
#  СХЕМА ОТЖИГА ПРАЙМЕРОВ  (рисунок для слайда «Термодинамический анализ»)
# ======================================================================
def arrow_box(sid, prst, x, y, cx, cy, fill, label):
    p = para([run(label, sz=1100, b=True, color="FFFFFF")], algn='ctr', after=0)
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="primer"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="%s"><a:avLst><a:gd name="adj1" fmla="val 50000"/>'
        '<a:gd name="adj2" fmla="val 40000"/></a:avLst></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="%s"/></a:solidFill></p:spPr>'
        '<p:txBody><a:bodyPr wrap="square" anchor="ctr" lIns="36576" rIns="36576" '
        'tIns="9144" bIns="9144"><a:normAutofit/></a:bodyPr><a:lstStyle/>%s</p:txBody></p:sp>'
        % (sid, x, y, cx, cy, prst, fill, p))


def rect_bar(sid, x, y, cx, cy, fill, label, txtcolor):
    p = para([run(label, sz=1300, b=True, color=txtcolor)], algn='ctr', after=0)
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="template"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        '<a:ln w="9525"><a:solidFill><a:srgbClr val="A6A6A6"/></a:solidFill></a:ln></p:spPr>'
        '<p:txBody><a:bodyPr wrap="square" anchor="ctr"><a:normAutofit/></a:bodyPr>'
        '<a:lstStyle/>%s</p:txBody></p:sp>' % (sid, x, y, cx, cy, fill, p))


def double_arrow(sid, x1, x2, y, color, w):
    return (
        '<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="%d" name="amplicon"/><p:cNvCxnSpPr/>'
        '<p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="%d" y="%d"/>'
        '<a:ext cx="%d" cy="1"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom>'
        '<a:ln w="%d"><a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        '<a:headEnd type="triangle" w="med" len="med"/>'
        '<a:tailEnd type="triangle" w="med" len="med"/></a:ln></p:spPr></p:cxnSp>'
        % (sid, x1, y, x2 - x1, w, color))


def build_primer_diagram():
    """Схема отжига праймеров на матрице гена TOP2A (стиль SnapGene)."""
    xf, wf = 2400000, 2100000      # прямой праймер
    xr, wr = 7700000, 2100000      # обратный праймер
    amp_l, amp_r = 2400000, xr + wr  # размах ампликона
    sh = ''
    # подпись ампликона
    sh += textbox(210, amp_l, 4460000, amp_r - amp_l, 340000,
                  [para([run("Ампликон \u2248 150 п.н.", sz=1400, b=True, color=DARK)],
                        algn='ctr', after=0)], anchor='b', name="ампликон-подпись")
    # двунаправленная стрелка размаха
    sh += double_arrow(211, amp_l, amp_r, 4860000, DARK, 22225)
    # прямой праймер (вправо)
    sh += arrow_box(212, "rightArrow", xf, 5000000, wf, 300000, BLUE,
                    "Forward (прямой) \u2192")
    # матрица — ген
    sh += rect_bar(213, 900000, 5380000, 10400000, 300000, "D9D9D9",
                   "Матрица \u2014 ген TOP2A (экзонный стык)", "1F1F1F")
    # обратный праймер (влево)
    sh += arrow_box(214, "leftArrow", xr, 5760000, wr, 300000, TEAL,
                    "\u2190 Reverse (обратный)")
    # полярность цепей
    sh += textbox(215, 560000, 5380000, 320000, 300000,
                  [para([run("5\u2032", sz=1100, b=True, color="595959")], algn='ctr', after=0)],
                  anchor='ctr', name="5prime")
    sh += textbox(216, 11320000, 5380000, 340000, 300000,
                  [para([run("3\u2032", sz=1100, b=True, color="595959")], algn='ctr', after=0)],
                  anchor='ctr', name="3prime")
    # подпись рисунка
    sh += textbox(217, 900000, 6060000, 10400000, 320000,
                  [para([run("Схема отжига праймеров на матрице "
                             "(визуализация в SnapGene)", sz=1200, i=True, color="595959")],
                        algn='ctr', after=0)], anchor='t', name="рис-подпись")
    return sh


def inject_primer_diagram(slide_xml_text):
    return slide_xml_text.replace('</p:spTree>', build_primer_diagram() + '</p:spTree>', 1)


# ======================================================================
#  Заметки к титульному слайду — вступительная речь
# ======================================================================
SPEECH = [
    "Добрый день, уважаемые члены Государственной аттестационной комиссии!",
    "Меня зовут Бахромов Жавохир. Вашему вниманию представляется выпускная "
    "квалификационная работа бакалавра на тему: «Синтез олигонуклеотидов для "
    "изучения экспрессии человеческого гена TOP2 методом полимеразной цепной "
    "реакции».",
    "Прежде чем перейти к цели и задачам, позвольте показать, почему изучение "
    "топоизомеразы II, кодируемой геном TOP2, действительно важно для науки и "
    "медицины.",
]


def update_title_notes():
    nm = "ppt/notesSlides/notesSlide1.xml"
    xml = text(nm)
    paras = ''.join(
        '<a:p><a:r><a:rPr lang="ru-RU" dirty="0"/><a:t>%s</a:t></a:r></a:p>' % esc(s)
        for s in SPEECH)
    new_body = '<p:txBody><a:bodyPr/><a:lstStyle/>' + paras + '</p:txBody>'
    # заменяем txBody у плейсхолдера заметок (sp id=3 "Заметки 2")
    xml = re.sub(
        r'(<p:cNvPr id="3" name="Заметки 2".*?</p:nvSpPr><p:spPr/>)<p:txBody>.*?</p:txBody>',
        r'\1' + new_body, xml, count=1, flags=re.S)
    parts[nm] = xml.encode('utf-8')


# ======================================================================
#  Сборка
# ======================================================================
# подключаем экспериментальные изображения
parts["__gel1"] = open("Javohir TOP2 -1.jpg", "rb").read()
parts["__gel2"] = open("Javoir TOP2.jpg", "rb").read()
parts["__melt"] = open("Melt Curve Plot Javohir.jpg", "rb").read()

# 1) новые/перестроенные слайды
parts["ppt/slides/slide2.xml"] = build_importance_1().encode('utf-8')
parts["ppt/slides/slide12.xml"] = build_importance_2().encode('utf-8')
parts["ppt/slides/slide13.xml"] = build_qc_slide().encode('utf-8')
parts["ppt/slides/slide14.xml"] = build_res1_slide().encode('utf-8')
parts["ppt/slides/slide15.xml"] = build_res2_slide().encode('utf-8')

# слайд «Термодинамический анализ праймеров» — добавляем схему отжига
parts["ppt/slides/slide5.xml"] = inject_primer_diagram(text("ppt/slides/slide5.xml")).encode('utf-8')

# --- слайды реакций: убираем английский, всё на русский ---
s7 = text("ppt/slides/slide7.xml")
old_src = ('<a:r><a:rPr lang="en-US" altLang="ru-RU" sz="1400">' + TNR +
           '</a:rPr><a:t>https://atdbio.com/nucleic-acids-book/'
           'Solid-phase-oligonucleotide-synthesis</a:t></a:r>')
new_src = ('<a:r><a:rPr lang="ru-RU" altLang="en-US" sz="1400">' + TNR +
           '</a:rPr><a:t>Твердофазный фосфорамидитный синтез '
           'олигонуклеотидов</a:t></a:r>')
assert old_src in s7, "не найдена строка источника на slide7"
s7 = s7.replace(old_src, new_src)
s7 = s7.replace("(DMTO)", "(ДМТ)")
s7 = s7.replace("Детрилирование", "Детритилирование")
s7 = s7.replace("Кэппинг", "Кэпирование")
parts["ppt/slides/slide7.xml"] = s7.encode('utf-8')

s8 = text("ppt/slides/slide8.xml")
s8 = s8.replace("Механизм детрилирования", "Механизм детритилирования")
s8 = s8.replace("DMT-группы", "ДМТ-группы")
parts["ppt/slides/slide8.xml"] = s8.encode('utf-8')

# 2) изображения в media
parts["ppt/media/imageGel1.jpg"] = parts.pop("__gel1")
parts["ppt/media/imageGel2.jpg"] = parts.pop("__gel2")
parts["ppt/media/imageMelt.jpg"] = parts.pop("__melt")

# 3) rels для перестроенных слайдов
LAYOUT_REL = ('<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/'
              'officeDocument/2006/relationships/slideLayout" '
              'Target="../slideLayouts/slideLayout2.xml"/>')
IMG_RELS = ''.join(
    '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/'
    '2006/relationships/image" Target="../media/image%d.png"/>' % (i, i)
    for i in range(1, 5))


def rels_wrap(inner):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/'
            'relationships">' + inner + '</Relationships>')


# slide12 (importance 2): логотипы + макет
parts["ppt/slides/_rels/slide12.xml.rels"] = rels_wrap(IMG_RELS + LAYOUT_REL).encode('utf-8')

# slide13/14/15: добавляем эксперимент. изображение как rId7
for sn, img in [("slide13", "imageGel1.jpg"),
                ("slide14", "imageGel2.jpg"),
                ("slide15", "imageMelt.jpg")]:
    rn = "ppt/slides/_rels/%s.xml.rels" % sn
    xml = text(rn)
    add = ('<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/'
           'officeDocument/2006/relationships/image" Target="../media/%s"/>' % img)
    xml = xml.replace('</Relationships>', add + '</Relationships>')
    parts[rn] = xml.encode('utf-8')

# 4) заметки титульного слайда
update_title_notes()

# 5) [Content_Types].xml — добавить jpeg
ct = text("[Content_Types].xml")
if 'Extension="jpg"' not in ct:
    ct = ct.replace(
        '<Default Extension="png" ContentType="image/png"/>',
        '<Default Extension="png" ContentType="image/png"/>'
        '<Default Extension="jpg" ContentType="image/jpeg"/>'
        '<Default Extension="jpeg" ContentType="image/jpeg"/>')
parts["[Content_Types].xml"] = ct.encode('utf-8')

# 6) presentation.xml + rels — порядок из 16 слайдов
ORDER = ["slide1", "slide2", "slide12", "slide3", "slide4", "slide5", "slide6",
         "slide7", "slide8", "slide9", "slide10", "slide11", "slide13",
         "slide14", "slide15", "slide17"]

# rels презентации: сохраняем не-слайдовые связи, переписываем слайдовые
pres_rels = text("ppt/_rels/presentation.xml.rels")
keep = re.findall(r'<Relationship [^>]*?/>', pres_rels)
non_slide = [r for r in keep if '/slide"' not in r]   # всё, кроме слайдов
# назначаем новые rId слайдам, не пересекаясь с существующими
used = set(int(m) for m in re.findall(r'Id="rId(\d+)"', ''.join(non_slide)))
nid = 1
slide_rel_ids = []
slide_rels_xml = []
for s in ORDER:
    while nid in used:
        nid += 1
    used.add(nid)
    rid = "rId%d" % nid
    slide_rel_ids.append(rid)
    slide_rels_xml.append(
        '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/'
        '2006/relationships/slide" Target="slides/%s.xml"/>' % (rid, s))
parts["ppt/_rels/presentation.xml.rels"] = rels_wrap(
    ''.join(non_slide) + ''.join(slide_rels_xml)).encode('utf-8')

# presentation.xml: переписываем sldIdLst и секцию
pres = text("ppt/presentation.xml")
sld_ids = list(range(256, 256 + len(ORDER)))
sldlst = '<p:sldIdLst>' + ''.join(
    '<p:sldId id="%d" r:id="%s"/>' % (sid, rid)
    for sid, rid in zip(sld_ids, slide_rel_ids)) + '</p:sldIdLst>'
pres = re.sub(r'<p:sldIdLst>.*?</p:sldIdLst>', sldlst, pres, flags=re.S)
# секция (p14): синхронизируем перечень id
sec = '<p14:sldIdLst>' + ''.join('<p14:sldId id="%d"/>' % sid for sid in sld_ids) + '</p14:sldIdLst>'
pres = re.sub(r'<p14:sldIdLst>.*?</p14:sldIdLst>', sec, pres, flags=re.S)
parts["ppt/presentation.xml"] = pres.encode('utf-8')

# отвязанный слайд 16 и его notes не входят в презентацию (остаются orphan-частями
# в архиве, PowerPoint их игнорирует). Чтобы не оставлять висячих Override —
# оставляем как есть: PowerPoint корректно открывает неиспользуемые части.

# ----------------------------------------------------------------------
#  Запись результата
# ----------------------------------------------------------------------
if os.path.exists(OUT):
    os.remove(OUT)
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    # [Content_Types].xml первым
    z.writestr("[Content_Types].xml", parts["[Content_Types].xml"])
    for name, data in parts.items():
        if name == "[Content_Types].xml":
            continue
        z.writestr(name, data)

print("Готово:", OUT)
print("Слайдов в презентации:", len(ORDER))
