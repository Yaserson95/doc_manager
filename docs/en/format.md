# Block format

## Block

    {block [params]}content{/block}

- Params are optional.
- Params are `key:"value"` pairs separated by spaces.
- Content may include arbitrary text, inline `{text}` / `{image}` tags,
  and nested `{block}` (e.g. inside table cells).

### Block params

| Key | Format | Description |
|---|---|---|
| `type` | `"p"` / `"h:N"` / `"ul:N"` / `"ol:N"` / `"table"` | Block type. Default: `p`. `N` — level |
| `format` | string `key(value);key(value);...` | Formatting |

### Formatting (`format`)

| Key | Example | What it does |
|---|---|---|
| `text-style(...)` | `text-style(bold,underline)` | bold / italic / underline |
| `bg-color(#hex)` | `bg-color(#fff2cc)` | Background color |
| `txt-color(#hex)` | `txt-color(#0000ff)` | Text color |
| `font-size(...)` | `font-size(14pt)` | Font size |
| `line-height(...)` | `line-height(1.5)` or `line-height(18pt)` | Line height |
| `align(...)` | `align(center)` | Alignment: left / center / right / justify |
| `padding(l,t,r,b)` | `padding(10pt,0,10pt,0)` | Padding (left, top, right, bottom) |
| `padding-left(...)` | `padding-left(12pt)` | Left padding |
| `padding-top(...)` | `padding-top(6pt)` | Top padding |
| `padding-right(...)` | `padding-right(12pt)` | Right padding |
| `padding-bottom(...)` | `padding-bottom(6pt)` | Bottom padding |

Sizes: `pt`, `px`, `in`, `cm`, or a bare number (treated as `pt`).

## Inline tags

### `{text}`

    {text format:"..."}text{/text}

Same keys as block `format`, except `padding*`.

### `{image}`

    {image format:"..." text:"alt"}[path | base64]{/image}

- `format` — image fitting params (partially supported).
- `text` — alternative text.
- Content — a file path or a data-URI (`data:image/png;base64,...`).

### Tables

    {block type:"table" format:"..."}
    {row height:"24pt" header:"true"}
    {cell colspan:"2" format:"align(center);bg-color(#eef)"}
    {block}...{/block}
    {/cell}
    {/row}
    {/block}

See [table.md](table.md) for details.

## Examples

Simple paragraph:

    {block}Just a paragraph.{/block}

Formatted heading:

    {block type:"h:1" format:"text-style(bold,underline);txt-color(#0000ff);padding-bottom(16pt)"}
    Blue underlined bold heading
    {/block}

Paragraph with inline styling:

    {block}
    Some text with {text format:"txt-color(#ffff00)"}yellow{/text} highlight.
    {/block}

Lists:

    {block type:"ul:1"}First item{/block}
    {block type:"ul:1"}Second item{/block}

## Format syntax details

- Parameter separator: `;`.
- Argument separator inside parens: `,`.
- Whitespace around names and arguments is ignored.
- Duplicates: the last one wins.
- Unknown parameters are preserved and do not break parsing.
- The form `key:(value)` (with colon) is **not** supported — use `key(value)`.