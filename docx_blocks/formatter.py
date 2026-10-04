"""
Единая точка применения FormatSpec к элементам docx:
параграфам, run-ам, ячейкам, строкам, таблицам.
"""

import re

from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from .format import FormatSpec


# CSS-стиль → OOXML-стиль границы
_BORDER_STYLE_MAP = {
    'solid': 'single',
    'dashed': 'dashed',
    'dotted': 'dotted',
    'double': 'double',
    'none': 'none',
    'hidden': 'none',
    'single': 'single',
    'dashdot': 'dashDot',
    'dashdotdot': 'dashDotDot',
    'dashsmallgap': 'dashSmallGap',
    'dotdash': 'dotDash',
    'dotdotdash': 'dotDotDash',
    'wave': 'wave',
    'triple': 'triple',
    'thick': 'thick',
}

_ALIGN_PARAGRAPH = {
    'left': WD_ALIGN_PARAGRAPH.LEFT,
    'center': WD_ALIGN_PARAGRAPH.CENTER,
    'right': WD_ALIGN_PARAGRAPH.RIGHT,
    'justify': WD_ALIGN_PARAGRAPH.JUSTIFY,
}

_ALIGN_OOXML = {
    'left': 'left',
    'center': 'center',
    'right': 'right',
    'justify': 'both',
}

_VALIGN_OOXML = {
    'top': 'top',
    'middle': 'center',
    'bottom': 'bottom',
}

_TABLE_ALIGN = {
    'left': WD_TABLE_ALIGNMENT.LEFT,
    'center': WD_TABLE_ALIGNMENT.CENTER,
    'right': WD_TABLE_ALIGNMENT.RIGHT,
}


class DocxFormatter:
    """
    Применяет FormatSpec к элементам docx.

    Публичные методы:
        parse_size(value)                   -> Length | None
        apply_paragraph(paragraph, spec)    -> None
        apply_run(run, spec)                -> None
        apply_cell(cell, spec)              -> None
        apply_row(row, spec)                -> None
        apply_table(table, spec, default_border=...) -> None
    """

    # ------------------------------------------------------------------
    # Размеры
    # ------------------------------------------------------------------
    @staticmethod
    def parse_size(value):
        """
        '14pt' / '12px' / '2cm' / '1in' / '14' -> Length.
        Проценты ('100%') — возвращают None (для таблиц учитывается отдельно).
        """
        if value is None:
            return None
        value = str(value).strip().lower()
        if not value:
            return None
        try:
            if value.endswith('pt'):
                return Pt(float(value[:-2]))
            if value.endswith('px'):
                return Pt(float(value[:-2]) * 0.75)
            if value.endswith('in'):
                return Inches(float(value[:-2]))
            if value.endswith('cm'):
                return Cm(float(value[:-2]))
            return Pt(float(value))
        except (ValueError, TypeError):
            return None

    # ------------------------------------------------------------------
    # Параграф
    # ------------------------------------------------------------------
    def apply_paragraph(self, paragraph, spec: FormatSpec):
        if not spec:
            return

        # padding (общий padding(l,t,r,b) + отдельные стороны)
        paddings = spec.padding()
        for side in ('left', 'top', 'right', 'bottom'):
            if side not in paddings:
                continue
            size = self.parse_size(paddings[side])
            if size is None:
                continue
            if side == 'left':
                paragraph.paragraph_format.left_indent = size
            elif side == 'right':
                paragraph.paragraph_format.right_indent = size
            elif side == 'top':
                paragraph.paragraph_format.space_before = size
            elif side == 'bottom':
                paragraph.paragraph_format.space_after = size

        if 'line-height' in spec:
            self._apply_line_height(paragraph, spec.arg('line-height'))

        if 'bg-color' in spec:
            self._set_paragraph_shading(paragraph, spec.arg('bg-color'))

        if 'align' in spec:
            a = _ALIGN_PARAGRAPH.get(str(spec.arg('align')).lower())
            if a is not None:
                paragraph.alignment = a

    def _apply_line_height(self, paragraph, value):
        if value is None:
            return
        value = str(value).strip().lower()
        if value.endswith(('pt', 'px', 'in', 'cm')):
            size = self.parse_size(value)
            if size is not None:
                paragraph.paragraph_format.line_spacing_rule = (
                    WD_LINE_SPACING.EXACTLY)
                paragraph.paragraph_format.line_spacing = size
        else:
            try:
                paragraph.paragraph_format.line_spacing = float(value)
            except ValueError:
                pass

    def _set_paragraph_shading(self, paragraph, color):
        cs = self.normalize_color(color)
        if not cs:
            return
        pPr = paragraph._p.get_or_add_pPr()
        for old in pPr.findall(qn('w:shd')):
            pPr.remove(old)
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), cs)
        pPr.append(shd)

    # ------------------------------------------------------------------
    # Run
    # ------------------------------------------------------------------
    def apply_run(self, run, spec: FormatSpec):
        if not spec:
            return

        styles = spec.text_styles()
        if styles.get('bold'):
            run.bold = True
        if styles.get('italic'):
            run.italic = True
        if styles.get('underline'):
            run.underline = True

        if 'txt-color' in spec:
            try:
                cs = self.normalize_color(spec.arg('txt-color'))
                if cs:
                    run.font.color.rgb = RGBColor.from_string(cs)
            except Exception:
                pass

        if 'font-size' in spec:
            size = self.parse_size(spec.arg('font-size'))
            if size is not None:
                run.font.size = size

    # ------------------------------------------------------------------
    # Ячейка
    # ------------------------------------------------------------------
    def apply_cell(self, cell, spec: FormatSpec):
        if not spec:
            return
        tcPr = cell._tc.get_or_add_tcPr()

        # bg-color — после tcBorders, до noWrap/tcMar/vAlign
        if 'bg-color' in spec:
            cs = self.normalize_color(spec.arg('bg-color'))
            if cs:
                for old in tcPr.findall(qn('w:shd')):
                    tcPr.remove(old)
                shd = OxmlElement('w:shd')
                shd.set(qn('w:val'), 'clear')
                shd.set(qn('w:color'), 'auto')
                shd.set(qn('w:fill'), cs)
                self._insert_tcpr(
                    tcPr, shd,
                    before=('w:noWrap', 'w:tcMar', 'w:textDirection',
                            'w:tcFitText', 'w:vAlign', 'w:hideMark'))

        # padding → w:tcMar (до vAlign)
        self._apply_cell_padding(tcPr, spec)

        # valign — после tcMar
        if 'valign' in spec:
            v = _VALIGN_OOXML.get(str(spec.arg('valign')).lower())
            if v:
                for old in tcPr.findall(qn('w:vAlign')):
                    tcPr.remove(old)
                va = OxmlElement('w:vAlign')
                va.set(qn('w:val'), v)
                self._insert_tcpr(
                    tcPr, va,
                    before=('w:hideMark',))

        # width — ставим в правильную позицию через python-docx
        if 'width' in spec:
            size = self.parse_size(spec.arg('width'))
            if size is not None:
                cell.width = size

        # ВАЖНО: align здесь НЕ обрабатываем — w:jc в tcPr невалиден.
        # Он применяется на уровне параграфов (см. _fill_cell).

    @staticmethod
    def _insert_tcpr(tcPr, element, before=()):
        """Вставляет element в tcPr перед первым из before; иначе в конец."""
        for tag in before:
            el = tcPr.find(qn(tag))
            if el is not None:
                el.addprevious(element)
                return
        tcPr.append(element)

    def _apply_cell_padding(self, tcPr, spec: FormatSpec):
        paddings = spec.padding()
        if not paddings:
            return
        for old in tcPr.findall(qn('w:tcMar')):
            tcPr.remove(old)
        mar = OxmlElement('w:tcMar')
        for side, tag in (('left', 'w:left'), ('top', 'w:top'),
                          ('right', 'w:right'), ('bottom', 'w:bottom')):
            if side not in paddings:
                continue
            size = self.parse_size(paddings[side])
            if size is None:
                continue
            el = OxmlElement(tag)
            el.set(qn('w:w'), str(int(size.twips)))
            el.set(qn('w:type'), 'dxa')
            mar.append(el)
        if len(mar):
            tcPr.append(mar)

    # ------------------------------------------------------------------
    # Строка
    # ------------------------------------------------------------------
    def apply_row(self, row, spec: FormatSpec):
        if not spec:
            return
        if 'height' in spec:
            size = self.parse_size(spec.arg('height'))
            if size is not None:
                row.height = size
        if 'header' in spec and self._truthy(spec.arg('header')):
            trPr = row._tr.get_or_add_trPr()
            for old in trPr.findall(qn('w:tblHeader')):
                trPr.remove(old)
            th = OxmlElement('w:tblHeader')
            th.set(qn('w:val'), 'true')
            trPr.append(th)

    @staticmethod
    def _truthy(v):
        return str(v).strip().lower() in ('true', '1', 'yes', 'on', 'да')

    # ------------------------------------------------------------------
    # Таблица
    # ------------------------------------------------------------------
    def apply_table(self, table, spec: FormatSpec,
                    default_border=('1pt', 'single', '#999999')):
        """
        spec           — FormatSpec (может быть пустым).
        default_border — применяется, если в spec нет border(...).
                         None — не применять дефолт.
        """
        spec = spec or FormatSpec()

        # align
        if 'align' in spec:
            a = _TABLE_ALIGN.get(str(spec.arg('align')).lower())
            if a is not None:
                table.alignment = a

        # width (в т.ч. '100%')
        if 'width' in spec:
            w = str(spec.arg('width')).strip()
            if w.endswith('%'):
                # Процентная ширина — применяем к таблице через tblW
                self._set_table_percent_width(table, w)
            else:
                size = self.parse_size(w)
                if size is not None:
                    table.autofit = False
                    for row in table.rows:
                        for cell in row.cells:
                            cell.width = size

        # border
        b = spec.border()
        if b is None:
            b = default_border
        if b is not None:
            self._set_table_borders(table, *b)

    def _set_table_percent_width(self, table, percent_str):
        """percent_str вида '100%' -> w:tblW w:type='pct'."""
        try:
            pct = float(percent_str.rstrip('%').strip())
        except ValueError:
            return
        # OOXML: pct в 50-х долях процента → 100% = 5000
        val = str(int(pct * 50))

        tblPr = table._tbl.tblPr
        for old in tblPr.findall(qn('w:tblW')):
            tblPr.remove(old)
        tblW = OxmlElement('w:tblW')
        tblW.set(qn('w:w'), val)
        tblW.set(qn('w:type'), 'pct')
        # tblW идёт сразу после tblStyle/tblpPr, до tblBorders
        successor = None
        for tag in ('w:jc', 'w:tblCellSpacing', 'w:tblInd',
                    'w:tblBorders', 'w:shd', 'w:tblLayout',
                    'w:tblCellMar', 'w:tblLook'):
            el = tblPr.find(qn(tag))
            if el is not None:
                successor = el
                break
        if successor is not None:
            successor.addprevious(tblW)
        else:
            tblPr.append(tblW)

    def _set_table_borders(self, table, size_str='1pt',
                           style='single', color='#000000'):
        try:
            pt = float(re.sub(r'[^\d.]', '', str(size_str)) or '1')
        except ValueError:
            pt = 1.0
        sz = str(int(pt * 8))  # 1/8 pt

        color = self.normalize_color(color) or '000000'
        ooxml_style = _BORDER_STYLE_MAP.get(str(style).lower(), 'single')

        tblPr = table._tbl.tblPr
        for old in tblPr.findall(qn('w:tblBorders')):
            tblPr.remove(old)

        borders = OxmlElement('w:tblBorders')
        for edge in ('top', 'left', 'bottom', 'right',
                     'insideH', 'insideV'):
            el = OxmlElement(f'w:{edge}')
            el.set(qn('w:val'), ooxml_style)
            el.set(qn('w:sz'), sz)
            el.set(qn('w:space'), '0')
            el.set(qn('w:color'), color)
            borders.append(el)

        # Правильная позиция в CT_TblPr:
        # ... tblW, jc, tblCellSpacing, tblInd, tblBorders, shd,
        #     tblLayout, tblCellMar, tblLook ...
        successors = (
            'w:shd', 'w:tblLayout', 'w:tblCellMar', 'w:tblLook',
            'w:tblCaption', 'w:tblDescription', 'w:tblPrChange',
        )
        try:
            tblPr.insert_element_before(borders, *successors)
        except AttributeError:
            successor = None
            for tag in successors:
                el = tblPr.find(qn(tag))
                if el is not None:
                    successor = el
                    break
            if successor is not None:
                successor.addprevious(borders)
            else:
                tblPr.append(borders)
                
    @staticmethod
    def normalize_color(value):
        """
        '#333'    -> '333333'
        '#333333' -> '333333'
        'abc'     -> 'aabbcc'
        'abcdef'  -> 'abcdef'
        None/''   -> None
        """
        if value is None:
            return None
        s = str(value).strip().lstrip('#').lower()
        if not s:
            return None
        # 3-символьный CSS-hex: #abc -> #aabbcc
        if len(s) == 3 and all(c in '0123456789abcdef' for c in s):
            return s[0]*2 + s[1]*2 + s[2]*2
        # 6-символьный — как есть
        if len(s) == 6 and all(c in '0123456789abcdef' for c in s):
            return s
        # не hex — вернём как есть, вдруг это 'auto' или именованный
        return s