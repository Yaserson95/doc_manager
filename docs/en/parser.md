# `DocxBlockParser`

Module: `docx_blocks/parser.py`

Converts text in block format into a `docx.Document`.

## Import

```python
from docx_blocks import DocxBlockParser
# or
from docx_blocks.parser import DocxBlockParser
```

## Public API

### `parse(input_text: str) -> Document`

Parses a string and returns a ready `docx.Document`.

**Example**

```python
parser = DocxBlockParser()
doc = parser.parse(open('input.txt', encoding='utf-8').read())
doc.save('output.docx')
```

### Attributes

| Attribute | Type | Purpose |
|---|---|---|
| `formatter` | `DocxFormatter` | Applies formatting |
| `table_parser` | `DocxTableParser` | Handles `{block type:"table"}` |

The parser is stateless between calls — you can reuse a single
instance for multiple documents.

## Supported elements

| Element | Supported |
|---|---|
| `type:"p"` | plain paragraph |
| `type:"h:N"` | heading level N |
| `type:"ul:N"` | bulleted list |
| `type:"ol:N"` | numbered list |
| `type:"table"` | table (see [table.md](table.md)) |
| `format:"text-style(...)"` | bold / italic / underline |
| `format:"bg-color(#hex)"` | paragraph shading |
| `format:"txt-color(#hex)"` | text color |
| `format:"font-size(Npt)"` | font size |
| `format:"line-height(...)"` | line height |
| `format:"align(...)"` | paragraph alignment |
| `format:"padding(...)"` | padding |
| `{text ...}...{/text}` | inline formatting |
| `{image ...}...{/image}` | image (path or base64) |

## Internal methods

| Method | Purpose |
|---|---|
| `_parse_params(s)` | parse `key:"value"` pairs |
| `_parse_block_type(s)` | `"h:2"` → `('h', 2)` |
| `_parse_block_format(s)` | string → `FormatSpec` |
| `_parse_block(s)` | parse one block |
| `_parse_content(s)` | split content into text and inline tags |
| `_parse_size(s)` | `"14pt"` → `Pt(14)` (delegates to `DocxFormatter`) |
| `_apply_block_format(p, spec)` | apply formatting (delegates) |
| `_apply_run_format(run, b, t)` | apply formatting (delegates) |
| `_add_paragraph_with_type(c, t, lvl)` | create paragraph in `Document` or `_Cell` |
| `_process_block_content(p, c, spec)` | process block content |
| `_add_block_to_container(c, s)` | handle a single block |

## Errors

Does not raise its own exceptions. Invalid blocks are silently
skipped; invalid sizes/colors are silently ignored.

## Limitations

- Headers, footers, footnotes and sections are not supported.
- Image wrapping (inline / floating) is not implemented —
  all images are inserted as inline.