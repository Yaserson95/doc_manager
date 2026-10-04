[RU](README_ru.md) | [EN]

# Docx Blocks

A set of tools for working with `.docx` through a plain-text block format.

## What it is

A document is described as a sequence of blocks:

    {block [params]}content{/block}

Content can contain inline tags:

    {text format:"..."}text{/text}
    {image format:"..." text:"alt"}[path | base64]{/image}

Tables are supported:

    {block type:"table" format:"..."}
    {row}{cell format:"..."}{block}...{/block}{/cell}{/row}
    {/block}

The format is easy to read, write and edit by hand, while conversion
to and from `.docx` is fully automated.

## Project layout

    doc_manager/
    ├── docx_blocks/              # library
    │   ├── __init__.py           # public API
    │   ├── format.py             # FormatSpec
    │   ├── formatter.py          # DocxFormatter
    │   ├── parser.py             # DocxBlockParser
    │   ├── serializer.py         # DocxBlockSerializer
    │   ├── table.py              # DocxTableParser / DocxTableSerializer
    │   └── tags.py               # nested tag search
    ├── doc_cli.py                # CLI (console script doc-cli)
    ├── doc_editor.py             # interactive editor (doc-editor)
    ├── docs/                     # documentation
    ├── examples/                 # examples
    ├── tests/                    # tests
    ├── pyproject.toml
    └── README.md

## Installation

```bash
pip install -e .
```

This installs the `doc-cli` and `doc-editor` console scripts.
You can also run the scripts directly:

```bash
python doc_cli.py --help
python doc_editor.py --help
```

Requirements: Python 3.10+, `python-docx`.

## Quick start

### Parse text into .docx

```python
from docx_blocks import DocxBlockParser

text = '{block type:"h:1"}Header{/block}'
doc = DocxBlockParser().parse(text)
doc.save('out.docx')
```

### Serialize .docx into text

```python
from docx_blocks import DocxBlockSerializer

text = DocxBlockSerializer().serialize('out.docx')
print(text)
```

### CLI

```bash
doc-cli example.docx --append "{block}One more paragraph{/block}"
doc-cli example.docx --serialize out.txt
```

### Interactive editor

```bash
doc-editor example.docx
```

## Documentation

- [format.md](docs/en/format.md) — block format syntax
- [parser.md](docs/en/parser.md) — `DocxBlockParser`
- [serializer.md](docs/en/serializer.md) — `DocxBlockSerializer`
- [formatter.md](docs/en/formatter.md) — `DocxFormatter`
- [table.md](docs/en/table.md) — tables
- [tags.md](docs/en/tags.md) — tag search utilities
- [cli.md](docs/en/cli.md) — `doc-cli`
- [editor.md](docs/en/editor.md) — `doc-editor`

## Limitations

- Paragraphs and tables are processed; headers, footers, footnotes
  and sections are not.
- Image wrap/positioning is not restored.
- `padding(l,t,r,b)` is expanded into separate
  `padding-left/top/right/bottom` on serialization.