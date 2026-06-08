# -*- coding: utf-8 -*-
"""
Генератор .docx на чистой стандартной библиотеке Python (без сторонних пакетов).
DOCX — это ZIP-архив с XML по стандарту Office Open XML.

Оформление по ГОСТ 7.32 / 7.0.5:
  - шрифт Times New Roman, 14 пт;
  - межстрочный интервал 1.5;
  - выравнивание по ширине, отступ первой строки 1.25 см;
  - поля: левое 30 мм, правое 15 мм, верхнее/нижнее 20 мм;
  - нумерация страниц снизу по центру;
  - автоматически собираемое оглавление (поле TOC).
"""

import zipfile
import html
import datetime


def esc(text):
    """Экранирование спецсимволов XML."""
    return html.escape(str(text), quote=True)


# Единицы: 1 см = 567 twips; 1 пункт = 20 twips; EMU: 914400 на дюйм.
CM = 567
PT = 20


class DocxBuilder:
    def __init__(self):
        self.body = []          # список XML-строк (абзацы/таблицы)
        self.images = {}        # имя -> bytes (для будущих растровых вставок)
        self._img_counter = 0
        self._rels = []         # дополнительные связи (изображения)

    # ---------- низкоуровневые помощники ----------

    def _run(self, text, *, bold=False, italic=False, size=28, sub=False,
             sup=False, color=None, font="Times New Roman", underline=False):
        """Один текстовый run. size в полупунктах (28 = 14 пт)."""
        rpr = ['<w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/>' % (font, font, font)]
        if bold:
            rpr.append('<w:b/>')
        if italic:
            rpr.append('<w:i/>')
        if underline:
            rpr.append('<w:u w:val="single"/>')
        if color:
            rpr.append('<w:color w:val="%s"/>' % color)
        rpr.append('<w:sz w:val="%d"/><w:szCs w:val="%d"/>' % (size, size))
        if sub:
            rpr.append('<w:vertAlign w:val="subscript"/>')
        if sup:
            rpr.append('<w:vertAlign w:val="superscript"/>')
        return ('<w:r><w:rPr>%s</w:rPr><w:t xml:space="preserve">%s</w:t></w:r>'
                % (''.join(rpr), esc(text)))

    def _runs_from_segments(self, segments, base_size=28):
        """segments: список (text, {опции}) -> склеенные run-ы.
        Позволяет вставлять подстрочные индексы, курсив и т.п. внутри абзаца."""
        out = []
        for seg in segments:
            if isinstance(seg, str):
                out.append(self._run(seg, size=base_size))
            else:
                text, opts = seg
                o = dict(opts)
                o.setdefault('size', base_size)
                out.append(self._run(text, **o))
        return ''.join(out)

    # ---------- публичные методы добавления контента ----------

    def page_break(self):
        self.body.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')

    def heading(self, text, level=1, number=None, page_break_before=True):
        """Заголовок с привязкой к стилю Heading{level} (нужно для оглавления)."""
        style = "Heading%d" % level
        ppr = ['<w:pStyle w:val="%s"/>' % style]
        if page_break_before and level == 1:
            ppr.append('<w:pageBreakBefore/>')
        # выравнивание: главы по центру, подзаголовки слева
        if level == 1:
            ppr.append('<w:jc w:val="center"/>')
            ppr.append('<w:spacing w:before="240" w:after="240" w:line="360" w:lineRule="auto"/>')
        else:
            ppr.append('<w:jc w:val="left"/>')
            ppr.append('<w:spacing w:before="200" w:after="160" w:line="360" w:lineRule="auto"/>')
            ppr.append('<w:ind w:firstLine="709"/>')
        label = (number + " " if number else "") + text
        size = 32 if level == 1 else 28
        run = self._run(label, bold=True, size=size)
        self.body.append('<w:p><w:pPr>%s</w:pPr>%s</w:p>' % (''.join(ppr), run))

    def para(self, text_or_segments, *, justify=True, indent=True, bold=False,
             italic=False, size=28, center=False, spacing_after=0, line=360):
        ppr = []
        if center:
            ppr.append('<w:jc w:val="center"/>')
        elif justify:
            ppr.append('<w:jc w:val="both"/>')
        if indent and not center:
            ppr.append('<w:ind w:firstLine="709"/>')  # 1.25 см
        ppr.append('<w:spacing w:after="%d" w:line="%d" w:lineRule="auto"/>'
                   % (spacing_after, line))
        if isinstance(text_or_segments, str):
            runs = self._run(text_or_segments, bold=bold, italic=italic, size=size)
        else:
            runs = self._runs_from_segments(text_or_segments, base_size=size)
        self.body.append('<w:p><w:pPr>%s</w:pPr>%s</w:p>' % (''.join(ppr), runs))

    def bullet(self, text_or_segments, *, numbered=False, size=28):
        num_id = 2 if numbered else 1
        ppr = ['<w:pStyle w:val="ListParagraph"/>',
               '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="%d"/></w:numPr>' % num_id,
               '<w:jc w:val="both"/>',
               '<w:spacing w:after="0" w:line="360" w:lineRule="auto"/>']
        if isinstance(text_or_segments, str):
            runs = self._run(text_or_segments, size=size)
        else:
            runs = self._runs_from_segments(text_or_segments, base_size=size)
        self.body.append('<w:p><w:pPr>%s</w:pPr>%s</w:p>' % (''.join(ppr), runs))

    def table(self, rows, *, widths=None, header=True, font_size=24,
              shade_header="D9E2F3", align_center=False):
        """rows: список строк, каждая строка — список ячеек (str или segments).
        widths: список ширин в twips (сумма ~ 9000 для книжной A4 с полями)."""
        ncol = max(len(r) for r in rows)
        if widths is None:
            total = 9300
            widths = [total // ncol] * ncol
        grid = ''.join('<w:gridCol w:w="%d"/>' % w for w in widths)
        out = ['<w:tbl>']
        out.append(
            '<w:tblPr>'
            '<w:tblStyle w:val="TableGrid"/>'
            '<w:tblW w:w="0" w:type="auto"/>'
            '<w:tblBorders>'
            '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
            '</w:tblBorders>'
            '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" '
            'w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
            '</w:tblPr>')
        out.append('<w:tblGrid>%s</w:tblGrid>' % grid)
        for ri, row in enumerate(rows):
            out.append('<w:tr>')
            for ci in range(ncol):
                cell = row[ci] if ci < len(row) else ""
                w = widths[ci] if ci < len(widths) else widths[-1]
                shade = ''
                is_head = header and ri == 0
                if is_head and shade_header:
                    shade = '<w:shd w:val="clear" w:color="auto" w:fill="%s"/>' % shade_header
                tcpr = ('<w:tcPr><w:tcW w:w="%d" w:type="dxa"/>%s'
                        '<w:vAlign w:val="center"/></w:tcPr>' % (w, shade))
                jc = 'center' if (is_head or align_center) else 'both'
                ppr = ('<w:pPr><w:jc w:val="%s"/>'
                       '<w:spacing w:after="40" w:line="276" w:lineRule="auto"/></w:pPr>' % jc)
                if isinstance(cell, str):
                    runs = self._run(cell, bold=is_head, size=font_size)
                else:
                    runs = self._runs_from_segments(cell, base_size=font_size)
                    if is_head:
                        # делаем заголовок жирным принудительно
                        runs = self._run(
                            cell if isinstance(cell, str) else
                            ''.join(s if isinstance(s, str) else s[0] for s in cell),
                            bold=True, size=font_size)
                out.append('<w:tc>%s<w:p>%s%s</w:p></w:tc>' % (tcpr, ppr, runs))
            out.append('</w:tr>')
        out.append('</w:tbl>')
        # пустой абзац после таблицы
        out.append('<w:p><w:pPr><w:spacing w:after="120"/></w:pPr></w:p>')
        self.body.append(''.join(out))

    def figure_caption(self, text, *, number=None):
        label = ("Рисунок %s — " % number) if number else ""
        ppr = ('<w:pPr><w:jc w:val="center"/>'
               '<w:spacing w:before="60" w:after="200" w:line="276" w:lineRule="auto"/></w:pPr>')
        run = self._run(label + text, size=24, italic=False)
        self.body.append('<w:p>%s%s</w:p>' % (ppr, run))

    def schematic_box(self, title, lines, *, fill="EFEFEF"):
        """Однострочная схематичная рамка (как блок на блок-схеме)."""
        cell_lines = [title] + lines
        # одна ячейка-таблица с заливкой
        out = ['<w:tbl>']
        out.append(
            '<w:tblPr><w:tblW w:w="9300" w:type="dxa"/>'
            '<w:jc w:val="center"/>'
            '<w:tblBorders>'
            '<w:top w:val="single" w:sz="8" w:space="0" w:color="404040"/>'
            '<w:left w:val="single" w:sz="8" w:space="0" w:color="404040"/>'
            '<w:bottom w:val="single" w:sz="8" w:space="0" w:color="404040"/>'
            '<w:right w:val="single" w:sz="8" w:space="0" w:color="404040"/>'
            '</w:tblBorders></w:tblPr>')
        out.append('<w:tblGrid><w:gridCol w:w="9300"/></w:tblGrid>')
        out.append('<w:tr><w:tc><w:tcPr><w:tcW w:w="9300" w:type="dxa"/>'
                   '<w:shd w:val="clear" w:color="auto" w:fill="%s"/>'
                   '<w:vAlign w:val="center"/></w:tcPr>' % fill)
        # заголовок блока
        out.append('<w:p><w:pPr><w:jc w:val="center"/>'
                   '<w:spacing w:after="20" w:line="276" w:lineRule="auto"/></w:pPr>%s</w:p>'
                   % self._run(title, bold=True, size=24))
        for ln in lines:
            out.append('<w:p><w:pPr><w:jc w:val="center"/>'
                       '<w:spacing w:after="0" w:line="259" w:lineRule="auto"/></w:pPr>%s</w:p>'
                       % self._run(ln, size=22))
        out.append('</w:tc></w:tr></w:tbl>')
        self.body.append(''.join(out))

    def arrow_down(self, label=""):
        run = self._run("↓ " + label if label else "↓", size=24, bold=True)
        self.body.append('<w:p><w:pPr><w:jc w:val="center"/>'
                         '<w:spacing w:after="0" w:before="0"/></w:pPr>%s</w:p>' % run)

    def toc(self):
        """Поле автоматического оглавления (обновляется в Word при открытии)."""
        self.body.append(
            '<w:p><w:pPr><w:jc w:val="center"/>'
            '<w:spacing w:after="240" w:line="360" w:lineRule="auto"/></w:pPr>'
            '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            '<w:b/><w:sz w:val="32"/></w:rPr><w:t>СОДЕРЖАНИЕ</w:t></w:r></w:p>')
        self.body.append(
            '<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/>'
            '<w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>'
            '<w:p><w:pPr><w:pStyle w:val="TOC1"/><w:tabs>'
            '<w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs>'
            '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            '<w:noProof/><w:sz w:val="28"/></w:rPr></w:pPr>'
            '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            '<w:r><w:rPr><w:noProof/></w:rPr><w:t>Обновите поле: правый клик → «Обновить поле».</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
            '</w:sdtContent></w:sdt>')

    # ---------- сборка файла ----------

    def _content_types(self):
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Default Extension="png" ContentType="image/png"/>'
                '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
                '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
                '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
                '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
                '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
                '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
                '</Types>')

    def _root_rels(self):
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
                '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
                '</Relationships>')

    def _document_rels(self):
        rels = [
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>',
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>',
            '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>',
        ]
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                + ''.join(rels) + '</Relationships>')

    def _styles(self):
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                '<w:docDefaults><w:rPrDefault><w:rPr>'
                '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
                '<w:sz w:val="28"/><w:szCs w:val="28"/><w:lang w:val="ru-RU"/>'
                '</w:rPr></w:rPrDefault>'
                '<w:pPrDefault><w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
                '</w:docDefaults>'
                # Normal
                '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
                '<w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
                '<w:sz w:val="28"/></w:rPr></w:style>'
                # Heading 1
                '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>'
                '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>'
                '<w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr>'
                '<w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
                # Heading 2
                '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/>'
                '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>'
                '<w:pPr><w:keepNext/><w:outlineLvl w:val="1"/></w:pPr>'
                '<w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>'
                # Heading 3
                '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/>'
                '<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>'
                '<w:pPr><w:keepNext/><w:outlineLvl w:val="2"/></w:pPr>'
                '<w:rPr><w:b/><w:i/><w:sz w:val="28"/></w:rPr></w:style>'
                # TOC styles
                '<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/>'
                '<w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="60"/></w:pPr></w:style>'
                '<w:style w:type="paragraph" w:styleId="TOC2"><w:name w:val="toc 2"/>'
                '<w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="240"/><w:spacing w:after="60"/></w:pPr></w:style>'
                '<w:style w:type="paragraph" w:styleId="TOC3"><w:name w:val="toc 3"/>'
                '<w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="480"/><w:spacing w:after="60"/></w:pPr></w:style>'
                # ListParagraph
                '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/>'
                '<w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/></w:pPr></w:style>'
                # TableGrid
                '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/>'
                '<w:tblPr><w:tblBorders>'
                '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
                '</w:tblBorders></w:tblPr></w:style>'
                # Footer
                '<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/>'
                '<w:basedOn w:val="Normal"/></w:style>'
                '</w:styles>')

    def _numbering(self):
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                # маркированный список (abstractNum 0)
                '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0">'
                '<w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="—"/>'
                '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
                '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr></w:lvl></w:abstractNum>'
                # нумерованный список (abstractNum 1)
                '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0">'
                '<w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1)"/>'
                '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>'
                # нумерованный список для списка литературы (abstractNum 2)
                '<w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0">'
                '<w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>'
                '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="567" w:hanging="567"/></w:pPr>'
                '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr></w:lvl></w:abstractNum>'
                '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'
                '<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>'
                '<w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num>'
                '</w:numbering>')

    def _settings(self):
        # updateFields=true заставит Word предложить обновить оглавление при открытии
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                '<w:updateFields w:val="true"/>'
                '<w:defaultTabStop w:val="708"/>'
                '<w:characterSpacingControl w:val="doNotCompress"/>'
                '</w:settings>')

    def _footer(self):
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                '<w:p><w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="center"/></w:pPr>'
                '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr>'
                '<w:fldChar w:fldCharType="begin"/></w:r>'
                '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr>'
                '<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
                '<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/></w:rPr>'
                '<w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>')

    def _core(self, title, author):
        now = datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
                'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
                'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
                '<dc:title>%s</dc:title><dc:creator>%s</dc:creator>'
                '<cp:lastModifiedBy>%s</cp:lastModifiedBy>'
                '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>'
                '<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>'
                '</cp:coreProperties>' % (esc(title), esc(author), esc(author), now, now))

    def _app(self):
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
                '<Application>Microsoft Office Word</Application></Properties>')

    def _document(self):
        sect = (
            '<w:sectPr>'
            '<w:footerReference w:type="default" r:id="rId4"/>'
            '<w:pgSz w:w="11906" w:h="16838"/>'
            # поля: левое 30мм=1701, правое 15мм=850, верх/низ 20мм=1134
            '<w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" '
            'w:header="720" w:footer="720" w:gutter="0"/>'
            '<w:pgNumType w:start="1"/>'
            '</w:sectPr>')
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
                'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
                'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
                'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
                '<w:body>' + ''.join(self.body) + sect + '</w:body></w:document>')

    def save(self, path, title="ВКР", author="Автор"):
        with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
            z.writestr('[Content_Types].xml', self._content_types())
            z.writestr('_rels/.rels', self._root_rels())
            z.writestr('word/document.xml', self._document())
            z.writestr('word/_rels/document.xml.rels', self._document_rels())
            z.writestr('word/styles.xml', self._styles())
            z.writestr('word/numbering.xml', self._numbering())
            z.writestr('word/settings.xml', self._settings())
            z.writestr('word/footer1.xml', self._footer())
            z.writestr('docProps/core.xml', self._core(title, author))
            z.writestr('docProps/app.xml', self._app())
        return path
