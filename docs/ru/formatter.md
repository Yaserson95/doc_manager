# `DocxFormatter`

Модуль: `docx_blocks/formatter.py`

Единая точка применения `FormatSpec` к элементам docx:
параграфам, run-ам, ячейкам, строкам, таблицам.

## Импорт

```python
from docx_blocks import DocxFormatter
# или
from docx_blocks.formatter import DocxFormatter
```

## Публичные методы

### `parse_size(value) -> Length | None`

Преобразует строку размера в объект `docx.shared.Length`.

| Вход | Результат |
|---|---|
| `'14pt'` | `Pt(14)` |
| `'12px'` | `Pt(9)` (приблизительно) |
| `'1in'` | `Inches(1)` |
| `'2cm'` | `Cm(2)` |
| `'14'` | `Pt(14)` |
| `'100%'` | `None` (обрабатывается отдельно в таблицах) |

### `apply_paragraph(paragraph, spec)`

Применяет `FormatSpec` к абзацу:

- `padding(l,t,r,b)` и `padding-left/top/right/bottom`;
- `line-height(...)`;
- `bg-color(#hex)`;
- `align(...)`.

### `apply_run(run, spec)`

Применяет к текстовому фрагменту:

- `text-style(bold,italic,underline)`;
- `txt-color(#hex)`;
- `font-size(...)`.

### `apply_cell(cell, spec)`

Применяет к ячейке таблицы:

- `bg-color(#hex)` → `w:shd` в `tcPr`;
- `valign(top|middle|bottom)` → `w:vAlign`;
- `width(...)` → `tcW`;
- `padding(...)` → `w:tcMar`.

Порядок вставки в `tcPr` соблюдается согласно схеме OOXML
(`shd → noWrap → tcMar → vAlign`); иначе Word молча
игнорирует элементы.

**Важно:** `align(...)` в ячейке на уровне `tcPr` невалиден
(`w:jc` там нет в схеме `CT_TcPr`). Горизонтальное выравнивание
применяется к параграфам внутри ячейки — этим занимается
`DocxTableParser._fill_cell`, прокидывая `align` в формат
вложенных блоков.

### `apply_row(row, spec)`

Применяет к строке таблицы:

- `height(...)` → `trHeight`;
- `header(...)` → `w:tblHeader` (повторяющийся заголовок).

Значения `header`: `true` / `1` / `yes` / `on` / `да`.

### `apply_table(table, spec, default_border=('1pt','single','#999999'))`

Применяет к таблице:

- `align(left|center|right)` → `w:jc`;
- `width(...)` (в т.ч. `100%`) → `w:tblW` (`type='pct'` для процентов);
- `border(size,style,color)` → `w:tblBorders`.

Если `border` в spec отсутствует — применяется `default_border`
(тонкая серая сетка). Чтобы отключить это, передай `default_border=None`.

## Стили границ

CSS-имя приводится к OOXML:

| CSS | OOXML |
|---|---|
| `solid` | `single` |
| `dashed` | `dashed` |
| `dotted` | `dotted` |
| `double` | `double` |
| `none` / `hidden` | `none` |

Всё, что не найдено в таблице, приводится к `single`.

## Цвета

`normalize_color(value)` принимает:

| Вход | Результат |
|---|---|
| `'#333'` | `'333333'` |
| `'#333333'` | `'333333'` |
| `'abc'` | `'aabbcc'` |
| `'ABCDEF'` | `'abcdef'` |
| `None` / `''` | `None` |

Короткая CSS-форма (`#333`) разворачивается до `RRGGBB` —
иначе Word игнорирует `w:fill` и `w:color`.

## Внутренние методы

| Метод | Назначение |
|---|---|
| `_apply_line_height(p, v)` | межстрочный интервал |
| `_set_paragraph_shading(p, c)` | заливка абзаца |
| `_apply_cell_padding(tcPr, spec)` | `w:tcMar` в ячейке |
| `_set_table_percent_width(t, s)` | `w:tblW` с `type='pct'` |
| `_set_table_borders(t, *args)` | `w:tblBorders` с правильной позицией |

## Использование вне парсера

```python
from docx import Document
from docx_blocks import DocxFormatter, FormatSpec

doc = Document()
table = doc.add_table(rows=1, cols=1)
cell = table.cell(0, 0)

spec = FormatSpec.parse('align(center);bg-color(#eef);text-style(bold)')
DocxFormatter().apply_cell(cell, spec)
doc.save('cell.docx')
```