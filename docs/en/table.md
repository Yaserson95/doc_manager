# Tables

Module: `docx_blocks/table.py`

Classes:

- `DocxTableParser` — parses `{block type:"table"}...{/block}` and adds
  the table into a `Document` or a cell (`_Cell`);
- `DocxTableSerializer` — serializes a `docx.Table` back to block format.

## Table format

    {block type:"table" format:"..."}
    {row height:"24pt" header:"true"}
    {cell colspan:"2" rowspan:"1" format:"align(center);bg-color(#eef)"}
    {block}...{/block}
    {/cell}
    {/row}
    {/block}

### Table params (`format`)

| Key | Example | Values |
|---|---|---|
| `align` | `align(center)` | left / center / right |
| `width` | `width(100%)` or `width(400pt)` | size or percent |
| `border` | `border(1pt,solid,#333)` | size, style, color |

If `border` is absent, the default `1pt single #999999` is applied.

### Row params (regular params)

| Key | Example |
|---|---|
| `height` | `height:"24pt"` |
| `header` | `header:"true"` (repeat on page break) |

### Cell params

Outside `format` — only `colspan` and `rowspan`:

    {cell colspan:"3" rowspan:"2" format:"..."}...{/cell}

Inside `format` — same as a block, plus cell-specific keys:

| Key | Example | Where it goes |
|---|---|---|
| `align` | `align(center)` | content alignment |
| `valign` | `valign(middle)` | top / middle / bottom |
| `width` | `width(40pt)` | cell width |
| `bg-color` | `bg-color(#eef)` | cell background |
| `padding` | `padding(6pt,3pt,6pt,3pt)` | padding l,t,r,b |
| `padding-left` | `padding-left(6pt)` | single side |
| `text-style` | `text-style(bold)` | default for nested blocks |
| `txt-color` | `txt-color(#222)` | same |
| `font-size` | `font-size(12pt)` | same |
| `line-height` | `line-height(1.3)` | same |

**Keys `text-style`, `txt-color`, `font-size`, `line-height`, `align`**
do not format the cell itself — they become defaults for nested blocks.

### Cell content

One or more `{block}...{/block}`. Nested tables are allowed:

    {cell}
    {block type:"table"}...{/block}
    {/cell}

## Examples

Simple table with a header:

    {block type:"table" format:"align(center);width(100%);border(1pt,solid,#333)"}
    {row header:"true" height:"24pt"}
    {cell format:"align(center);bg-color(#eef);text-style(bold)"}
    {block}Parameter{/block}
    {/cell}
    {cell format:"align(center);bg-color(#eef);text-style(bold)"}
    {block}Value{/block}
    {/cell}
    {/row}
    {row}
    {cell}Price{/cell}
    {cell format:"align(right)"}{block}100{/block}{/cell}
    {/row}
    {/block}

Merged cells:

    {block type:"table"}
    {row}
    {cell colspan:"2" format:"align(center)"}
    {block}Two-column heading{/block}
    {/cell}
    {/row}
    {row}
    {cell rowspan:"2" format:"valign(middle)"}
    {block}Shared cell{/block}
    {/cell}
    {cell}A{/cell}
    {/row}
    {row}
    {cell}B{/cell}
    {/row}
    {/block}

## Programmatic API

### `DocxTableParser(block_parser)`

Takes a reference to `DocxBlockParser` (uses its `formatter`,
`_parse_params`, `_parse_block`, `_add_paragraph_with_type`,
`_process_block_content`).

Method:

    add_table(container, table_params: dict, table_content: str) -> Table | None

- `container` — `Document` or `_Cell`;
- `table_params` — block `params` (keys `type`, `format`);
- `table_content` — block content (inside `{block ...}`).

### `DocxTableSerializer(block_serializer)`

Takes a reference to `DocxBlockSerializer` (uses its
`_serialize_paragraph`).

Method:

    serialize_table(table) -> str

## Limitations

- `align` in `tcPr` is invalid per the OOXML schema — it is applied
  to paragraphs inside the cell (see [formatter.md](formatter.md)).
- On serialization, `border` is restored generically
  (`1pt,single,#000000`); exact edge parameters are not preserved.
- Image positioning inside cells is not supported (no wrapping).