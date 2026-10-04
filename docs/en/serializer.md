# `DocxBlockSerializer`

Module: `docx_blocks/serializer.py`

Converts `.docx` into text in block format.

## Import

```python
from docx_blocks import DocxBlockSerializer
# or
from docx_blocks.serializer import DocxBlockSerializer
```

## Public API

### `serialize(docx_input) -> str`

**Parameters**

- `docx_input` — `str` (path) or `docx.Document`.

**Returns** — `str`; blocks separated by `\n`.

**Example**

```python
ser = DocxBlockSerializer()

text = ser.serialize('example.docx')

from docx import Document
text = ser.serialize(Document('example.docx'))
```

### Attributes

| Attribute | Type | Purpose |
|---|---|---|
| `table_serializer` | `DocxTableSerializer` | Table serialization |

## What is restored

| docx element | Text output |
|---|---|
| `Heading N` | `type:"h:N"` |
| `List Bullet` / `List Bullet 2` | `type:"ul:1"` / `type:"ul:2"` |
| `List Number` | `type:"ol:N"` |
| Plain paragraph | no `type` |
| Table | `{block type:"table"}...{/block}` |
| Indents (spacing before/after, left/right) | `padding-top/bottom/left/right(...)` |
| Line height | `line-height(...)` |
| Paragraph shading | `bg-color(#hex)` |
| Paragraph alignment | `align(...)` |
| Bold / italic / underline | `text-style(bold,...)` |
| Text color | `txt-color(#hex)` |
| Font size | `font-size(Npt)` |
| Run with non-default format | `{text format:"..."}...{/text}` |
| Image | `{image text:""}data:image/...;base64,...{/image}` |

## Internal methods

| Method | Purpose |
|---|---|
| `_serialize_paragraph(p)` | one paragraph → block string |
| `_detect_block_type(p)` | detect type (h / ul / ol / p) |
| `_extract_block_format_dict(p)` | paragraph format → `FormatSpec` |
| `_serialize_content(p, spec)` | content with `{text}` / `{image}` |
| `_extract_run_format_dict(run)` | run format → `FormatSpec` |
| `_run_format_differs(run, block)` | whether to wrap in `{text}` |
| `_extract_images(run)` | images from a run |
| `_serialize_image(img)` | image → `{image ...}` |
| `_get_paragraph_shading(p)` | paragraph shading color |
| `_emu_to_pt(emu)` | EMU → pt |

The body is walked via `w:p` and `w:tbl`, preserving paragraph
and table order.

## Limitations

- Headers, footers, footnotes and sections are ignored.
- `padding(...)` is expanded into separate `padding-*`.
- Image `format` (wrap/positioning) is not restored.
- Table `border(...)` is restored generically:
  `border(1pt,single,#000000)`.