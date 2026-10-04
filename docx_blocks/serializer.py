# docx_serializer.py
import re
import base64
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Length
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.table import Table

from .table import DocxTableSerializer
from .format import FormatSpec

class DocxBlockSerializer:
    """
    Обратный парсер: docx -> текст в формате блоков
    {block ...}...{/block} со встроенными {text ...}...{/text}
    и {image ...}...{/image}.
    """

    def __init__(self):
        self.table_serializer = DocxTableSerializer(self)

    # ------------------------------------------------------------------
    # Публичный метод
    # ------------------------------------------------------------------
    def serialize(self, docx_input) -> str:
        if isinstance(docx_input, str):
            doc = Document(docx_input)
        else:
            doc = docx_input

        blocks = []
        for child in doc.element.body:
            if child.tag == qn('w:p'):
                p = Paragraph(child, doc)
                blocks.append(self._serialize_paragraph(p))
            elif child.tag == qn('w:tbl'):
                t = Table(child, doc)
                blocks.append(self.table_serializer.serialize_table(t))
        return '\n'.join(blocks)

    # ------------------------------------------------------------------
    # Сериализация одного абзаца
    # ------------------------------------------------------------------
    def _serialize_paragraph(self, paragraph) -> str:
        type_str = self._detect_block_type(paragraph)
        block_format_dict = self._extract_block_format_dict(paragraph)
        block_format_str = block_format_dict.to_string()
        content = self._serialize_content(paragraph, block_format_dict)

        params = []
        if type_str:
            params.append(f'type:"{type_str}"')
        if block_format_str:
            params.append(f'format:"{block_format_str}"')
        params_str = ' '.join(params)

        if params_str:
            return f'{{block {params_str}}}{content}{{/block}}'
        return f'{{block}}{content}{{/block}}'

    # ------------------------------------------------------------------
    # Определение типа блока
    # ------------------------------------------------------------------
    def _detect_block_type(self, paragraph) -> str | None:
        style_name = paragraph.style.name if paragraph.style else ''
        if style_name.startswith('Heading '):
            try:
                level = int(style_name.split(' ')[1])
                return f'h:{level}'
            except (ValueError, IndexError):
                return None

        if style_name.startswith('List Bullet'):
            parts = style_name.split(' ')
            level = 1
            if len(parts) > 2:
                try:
                    level = int(parts[2])
                except ValueError:
                    pass
            return f'ul:{level}'

        if style_name.startswith('List Number'):
            parts = style_name.split(' ')
            level = 1
            if len(parts) > 2:
                try:
                    level = int(parts[2])
                except ValueError:
                    pass
            return f'ol:{level}'

        return None  # обычный параграф (p)

    # ------------------------------------------------------------------
    # Извлечение форматирования блока (paragraph-level + базовая run-формат)
    # ------------------------------------------------------------------
    def _extract_block_format_dict(self, paragraph) -> FormatSpec:
        spec = FormatSpec()
        pf = paragraph.paragraph_format

        if pf.left_indent is not None:
            spec.set('padding-left', f'{self._emu_to_pt(pf.left_indent)}pt')
        if pf.right_indent is not None:
            spec.set('padding-right', f'{self._emu_to_pt(pf.right_indent)}pt')
        if pf.space_before is not None:
            spec.set('padding-top', f'{self._emu_to_pt(pf.space_before)}pt')
        if pf.space_after is not None:
            spec.set('padding-bottom', f'{self._emu_to_pt(pf.space_after)}pt')

        if pf.line_spacing is not None:
            if isinstance(pf.line_spacing, Length):
                spec.set('line-height',
                         f'{round(pf.line_spacing.pt, 2)}pt')
            else:
                spec.set('line-height', str(pf.line_spacing))

        bg = self._get_paragraph_shading(paragraph)
        if bg:
            spec.set('bg-color', f'#{bg}')

        if pf.alignment is not None and pf.alignment != WD_ALIGN_PARAGRAPH.LEFT:
            a = {
                WD_ALIGN_PARAGRAPH.CENTER: 'center',
                WD_ALIGN_PARAGRAPH.RIGHT: 'right',
                WD_ALIGN_PARAGRAPH.JUSTIFY: 'justify',
            }.get(pf.alignment)
            if a:
                spec.set('align', a)

        for run in paragraph.runs:
            if run.text:
                spec = spec.merge(self._extract_run_format_dict(run))
                break

        return spec

    def _get_paragraph_shading(self, paragraph) -> str | None:
        pPr = paragraph._p.find(qn('w:pPr'))
        if pPr is None:
            return None
        shd = pPr.find(qn('w:shd'))
        if shd is None:
            return None
        fill = shd.get(qn('w:fill'))
        if fill and fill.lower() not in ('auto', 'ffffff'):
            return fill
        return None

    # ------------------------------------------------------------------
    # Сериализация содержимого (runs, {text}, {image})
    # ------------------------------------------------------------------
    def _serialize_content(self, paragraph, block_format_dict: dict) -> str:
        result = []
        for run in paragraph.runs:
            # --- изображения ---
            images = self._extract_images(run)
            if images:
                for img in images:
                    result.append(self._serialize_image(img))
                if run.text:
                    result.append(run.text)
                continue

            # --- обычный текст ---
            text = run.text
            if not text:
                continue

            run_format_dict = self._extract_run_format_dict(run)
            if run_format_dict and self._run_format_differs(run_format_dict, block_format_dict):
                run_format_str = run_format_dict.to_string()
                if run_format_str:
                    result.append(
                        f'{{text format:"{run_format_str}"}}{text}{{/text}}'
                    )
                    continue
            result.append(text)

        return ''.join(result)

    def _extract_run_format_dict(self, run) -> FormatSpec:
        spec = FormatSpec()

        styles = []
        if run.bold:
            styles.append('bold')
        if run.italic:
            styles.append('italic')
        if run.underline:
            styles.append('underline')
        if styles:
            spec.set('text-style', styles)

        try:
            color = run.font.color
            if color is not None and color.rgb is not None:
                spec.set('txt-color', f'#{color.rgb}')
        except Exception:
            pass

        if run.font.size is not None:
            spec.set('font-size', f'{run.font.size.pt}pt')

        return spec

    def _run_format_differs(self, run_spec: FormatSpec, block_spec: FormatSpec) -> bool:
        for key in ('text-style', 'txt-color', 'font-size'):
            if run_spec.get(key) != block_spec.get(key):
                return True
        return False

    # ------------------------------------------------------------------
    # Работа с изображениями
    # ------------------------------------------------------------------
    def _extract_images(self, run) -> list:
        images = []
        for drawing in run._element.findall('.//' + qn('w:drawing')):
            for blip in drawing.findall('.//' + qn('a:blip')):
                rId = blip.get(qn('r:embed'))
                if not rId:
                    continue
                try:
                    image_part = run.part.related_parts[rId]
                    images.append({
                        'content_type': image_part.content_type,
                        'blob': image_part.blob,
                    })
                except KeyError:
                    pass
        return images

    def _serialize_image(self, img: dict) -> str:
        b64 = base64.b64encode(img['blob']).decode('ascii')
        content_type = img['content_type']
        data_uri = f'data:{content_type};base64,{b64}'
        return f'{{image text:""}}{data_uri}{{/image}}'

    # ------------------------------------------------------------------
    # Утилиты
    # ------------------------------------------------------------------
    def _emu_to_pt(self, emu) -> float:
        return round(emu / 12700, 2)