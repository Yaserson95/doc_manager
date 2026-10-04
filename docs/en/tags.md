# `docx_tags`

Module: `docx_blocks/tags.py`

Utilities for finding top-level tags with proper nesting awareness.
Needed because simple regexes such as `\{block\b.*?\{/block\}` do not
distinguish nested blocks:

    {block}
      {block}nested{/block}     <-- regex closes here
    {/block}

## Public API

### `find_top_level(text, open_re, close_tag) -> list[tuple[int, int]]`

Returns a list of `(start, end)` for top-level occurrences of the
opening tag and its matching closing tag.

**Parameters**

- `text` — string;
- `open_re` — regex of the opening tag (no groups);
- `close_tag` — closing tag string.

**Example**

```python
from docx_blocks.tags import find_top_level

ranges = find_top_level(
    '{block}a{block}b{/block}c{/block}',
    r'\{block\b',
    '{/block}',
)
# [(0, 32)] — one top-level block
```

### `find_blocks(text) -> list[str]`

Top-level `{block ...}...{/block}`.

### `find_rows(text) -> list[str]`

Top-level `{row ...}...{/row}`.

### `find_cells(text) -> list[str]`

Top-level `{cell ...}...{/cell}`.

## How it works

Simple algorithm: iterate over all matches of `open_re | close_tag`,
track depth. Every time depth returns to zero, record the pair.

## Usage

Inside the project:

```python
from docx_blocks.tags import find_blocks

blocks = find_blocks(text)   # list of strings, one per whole block
```

This is the correct way to split text into blocks — otherwise
nested blocks (e.g. inside table cells) will close at the wrong
position.