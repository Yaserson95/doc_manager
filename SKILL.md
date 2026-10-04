---
name: docx-blocks
description: >
  Create and edit Word .docx documents via a plain-text block format.
  Use for generating .docx files with headings, paragraphs, lists,
  tables or images; modifying an existing document (append, prepend,
  replace, insert, delete); or extracting document structure as text.
license: MIT
---

# docx-blocks

## What it is

A `.docx` document is described as a sequence of blocks in a plain-text
format. The `doc-cli` command-line tool converts blocks into `.docx`
and back.

The tool is available as `doc-cli` (console script) or
`python -m doc_cli`.

## When to use

- The user asks to **create** a `.docx` with text, headings, lists,
  tables or images.
- The user asks to **modify** an existing `.docx`: append, prepend,
  replace, insert or delete blocks.
- You need to **read** the contents of a `.docx` as text.

**Do not use** when the `.docx` requires headers, footers, footnotes,
sections/page layout, or precise image positioning — these are not
supported.

## Workflow

### Step 1. Inspect what is already in the file

If the file exists, serialize it to see the structure:

```bash
doc-cli path/to/file.docx --serialize
```

The output is a list of blocks in text format. Each block can be
addressed by index (1-based): `--insert N`, `--insert-after N`,
`--insert-before N`.

### Step 2. Prepare blocks

Write blocks to a temp file (easier than an inline string — no quote
escaping needed):

```bash
cat > /tmp/blocks.txt <<'EOF'
{block type:"h:1" format:"align(center);text-style(bold)"}
Document title
{/block}
{block}
A regular paragraph with {text format:"text-style(bold)"}emphasis{/text}.
{/block}
{block type:"ul:1"}Item 1{/block}
{block type:"ul:1"}Item 2{/block}
EOF
```

### Step 3. Apply a command

```bash
# create from scratch (or overwrite)
doc-cli output.docx --rewrite --file /tmp/blocks.txt

# append at the end
doc-cli output.docx --append --file /tmp/blocks.txt

# prepend at the beginning
doc-cli output.docx --prepend --file /tmp/blocks.txt

# replace block 2
doc-cli output.docx --insert 2 --file /tmp/blocks.txt

# insert after block 2
doc-cli output.docx --insert-after 2 --file /tmp/blocks.txt

# insert before block 1
doc-cli output.docx --insert-before 1 --file /tmp/blocks.txt
```

Alternative — pass blocks via stdin:

```bash
echo '{block}Hello{/block}' | doc-cli output.docx --append --stdin
```

### Step 4. Verify

```bash
doc-cli output.docx --serialize
```

If the structure matches what you intended, the task is done.

## Block syntax (quick reference)

### Paragraph

```
{block}Just text{/block}
```

### Heading

```
{block type:"h:1"}First-level heading{/block}
{block type:"h:2"}Second-level{/block}
```

`h:1` … `h:6`.

### List

```
{block type:"ul:1"}Bulleted item{/block}
{block type:"ol:1"}Numbered item{/block}
```

### Formatting

The `format` parameter is a `key(value)` string, entries separated by `;`.

```
{block format:"align(center);text-style(bold,italic);font-size(14pt)"}
Text
{/block}
```

| Key | Values | Effect |
|---|---|---|
| `text-style(...)` | `bold`, `italic`, `underline` (comma-separated) | font style |
| `txt-color(#hex)` | `#RRGGBB` or `#RGB` | text color |
| `bg-color(#hex)` | `#RRGGBB` or `#RGB` | background color |
| `font-size(14pt)` | `pt`, `px`, `in`, `cm` or a bare number | font size |
| `line-height(1.5)` | number or size | line spacing |
| `align(center)` | `left`, `center`, `right`, `justify` | alignment |
| `padding(10pt,0,10pt,0)` | `l,t,r,b` | padding |
| `padding-left(10pt)` | size | single side |

### Inline formatting

Inside a block, a fragment can be styled individually:

```
{block}
Regular text and {text format:"txt-color(#c00);text-style(bold)"}red bold{/text}.
{/block}
```

Same keys as above, except `padding*`.

### Table

```
{block type:"table" format:"align(center);width(100%);border(1pt,solid,#333)"}
{row header:"true" height:"24pt"}
{cell format:"align(center);bg-color(#eef);text-style(bold)"}
{block}Header A{/block}
{/cell}
{cell format:"align(center);bg-color(#eef);text-style(bold)"}
{block}Header B{/block}
{/cell}
{/row}
{row}
{cell}Value 1{/cell}
{cell format:"align(right)"}{block}42{/block}{/cell}
{/row}
{/block}
```

Merging — outside `format`:

```
{cell colspan:"2" rowspan:"1" format:"..."}...{/cell}
```

Cell `format` keys: `align`, `valign`, `width`, `bg-color`, `padding*`,
plus `text-style`, `txt-color`, `font-size`, `line-height` as defaults
for nested blocks.

### Image

```
{block}
{image text:"Diagram"}/path/to/image.png{/image}
{/block}
```

Or base64:

```
{image text:""}data:image/png;base64,iVBORw0KGgo...{/image}
```

## CLI commands

| Command | What it does |
|---|---|
| `--rewrite "..."` | overwrite the whole file |
| `--append "..."` | append blocks at the end |
| `--prepend "..."` | prepend blocks at the beginning |
| `--insert N "..."` | replace block N |
| `--insert-after N "..."` | insert after block N |
| `--insert-before N "..."` | insert before block N |
| `--serialize` | print structure to stdout |
| `--serialize out.txt` | save structure to a file |
| `--file FILE` / `-f FILE` | read blocks from a file |
| `--stdin` | read blocks from stdin |
| `--quiet` / `-q` | suppress the `OK: ...` line |
| `--json` | print result as JSON |

Without a command, `--rewrite` is used.

## Success check

After each command run:

```bash
doc-cli output.docx --serialize --quiet
```

and confirm the blocks were added, replaced or removed as expected.

If `--json` is enabled, use:

```bash
doc-cli output.docx --insert 2 --file /tmp/b.txt --json
```

Response:

```json
{"ok": true, "file": "output.docx", "blocks_before": 3, "blocks_after": 3}
```

## Exit codes

| Code | Meaning |
|---|---|
| 0 | success |
| 1 | logical error (N out of range) |
| 2 | bad arguments or read error |
| 3 | write error |

## Common pitfalls

- **Do not escape quotes.** If passing blocks inline, use single
  quotes around the string in bash: `'{block type:"h:1"}...'`.
  Better — always write blocks via `--file`.
- **`--insert*` requires an existing file.** If the file does not
  exist yet, use `--rewrite` or `--append` — they create it.
- **N is 1-based.** The first block is `1`, not `0`.
- **Colors.** Both `#RRGGBB` and short `#RGB` are accepted.
- **Background color** requires text inside the block, otherwise
  it is not visible.
- **Cell `align`** is applied to the paragraphs inside the cell,
  not to the cell itself (an OOXML limitation) — but visually it
  works as expected.

## Full example

User: "Create a document with a heading, two paragraphs and a 3×2 table."

```bash
cat > /tmp/doc.txt <<'EOF'
{block type:"h:1" format:"align(center)"}
Quarterly report
{/block}
{block}
First paragraph — introduction.
{/block}
{block}
Second paragraph with {text format:"text-style(bold)"}emphasis{/text}.
{/block}
{block type:"table" format:"border(1pt,solid,#333)"}
{row header:"true"}
{cell format:"text-style(bold);align(center)"}{block}Month{/block}{/cell}
{cell format:"text-style(bold);align(center)"}{block}Revenue{/block}{/cell}
{/row}
{row}
{cell}January{/cell}
{cell format:"align(right)"}{block}100,000{/block}{/cell}
{/row}
{row}
{cell}February{/cell}
{cell format:"align(right)"}{block}120,000{/block}{/cell}
{/row}
{/block}
EOF

doc-cli report.docx --rewrite --file /tmp/doc.txt --json
doc-cli report.docx --serialize
```

## Localization

Russian version: [SKILL_ru.md](SKILL_ru.md).