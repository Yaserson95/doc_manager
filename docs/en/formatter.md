# `DocxFormatter`

Module: `docx_blocks/formatter.py`

A single entry point for applying `FormatSpec` to docx elements:
paragraphs, runs, cells, rows, tables.

## Import

```python
from docx_blocks import DocxFormatter
# or
from docx_blocks.formatter import DocxFormatter
```

## Public methods

### `parse_size(value) -> Length | None`

Converts a size string into a `docx.shared.Length`.

| Input | Result |
|---|---|
| `'14pt'` | `Pt(14)` |
| `'12px'` | `Pt(9)` (approximate) |
| `'1in'` | `Inches(1)` |
| `'2cm'` | `Cm(2)` |
| `'14'` | `Pt(14)` |
| `'100%'` | `None` (handled separately for tables) |

### `apply_paragraph(paragraph, spec)`

Applies a `FormatSpec` to a paragraph:

- `padding(l,t,r,b)` and `padding-left/top/right/bottom`;
- `line-height(...)`;
- `bg-color(#hex)`;
- `align(...)`.

### `apply_run(run, spec)`

Applies to a text run:

- `text-style(bold,italic,underline)`;
- `txt-color(#hex)`;
- `font-size(...)`.

### `apply_cell(cell, spec)`

Applies to a table cell:

- `bg-color(#hex)` → `w:shd` in `tcPr`;
- `valign(top|middle|bottom)` → `w:vAlign`;
- `width(...)` → `tcW`;
- `padding(...)` → `w:tcMar`.

Insertion order in `tcPr` follows the OOXML schema
(`shd → noWrap → tcMar → vAlign`); otherwise Word silently
ignores the elements.

**Note:** `align(...)` inside a cell is invalid at the `tcPr`
level (`w:jc` is not part of `CT_TcPr`). Horizontal alignment is
applied to paragraphs inside the cell — this is done by
`DocxTableParser._fill_cell`, which forwards `align` into the
nested block format.

### `apply_row(row, spec)`

Applies to a table row:

- `height(...)` → `trHeight`;
- `header(...)` → `w:tblHeader` (repeating header row).

`header` values: `true` / `1` / `yes` / `on` / `да`.

### `apply_table(table, spec, default_border=('1pt','single','#999999'))`

Applies to a table:

- `align(left|center|right)` → `w:jc`;
- `width(...)` (including `100%`) → `w:tblW` (`type='pct'` for percentages);
- `border(size,style,color)` → `w:tblBorders`.

If `border` is missing from the spec, `default_border` is applied
(thin grey grid). Pass `default_border=None` to disable.

## Border styles

CSS name → OOXML:

| CSS | OOXML |
|---|---|
| `solid` | `single` |
| `dashed` | `dashed` |
| `dotted` | `dotted` |
| `double` | `double` |
| `none` / `hidden` | `none` |

Anything not in the map falls back to `single`.

## Colors

`normalize_color(value)` accepts:

| Input | Result |
|---|---|
| `'#333'` | `'333333'` |
| `'#333333'` | `'333333'` |
| `'abc'` | `'aabbcc'` |
| `'ABCDEF'` | `'abcdef'` |
| `None` / `''` | `None` |

Short CSS form (`#333`) is expanded to `RRGGBB` — otherwise Word
ignores `w:fill` and `w:color`.

## Internal methods

| Method | Purpose |
|---|---|
| `_apply_line_height(p, v)` | line height |
| `_set_paragraph_shading(p, c)` | paragraph shading |
| `_apply_cell_padding(tcPr, spec)` | `w:tcMar` in a cell |
| `_set_table_percent_width(t, s)` | `w:tblW` with `type='pct'` |
| `_set_table_borders(t, *args)` | `w:tblBorders` at the correct position |

## Using outside the parser

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