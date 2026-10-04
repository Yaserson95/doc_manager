# Editor `doc-editor`

File: `doc_editor.py` (installed as the `doc-editor` command).

A vim-like interactive editor for `.docx` with a cursor on the
current block.

## Launch

```bash
doc-editor example.docx
```

If the file does not exist, it will be created on save.

## Concept

- There is a **cursor** — the current block index (0-based, shown
  as `[N/total]` in the prompt).
- Most commands operate on the **current block**, not on a number.
- Prompt: `[3/7]> ` — block 3 of 7.

## Commands

### Core

| Command | Alias | What it does |
|---|---|---|
| `format` | `f` | set current block params; creates a block if none exist |
| `edit` | `e` | edit current block content (multiline) |
| `delete` | `d` | delete current block (with confirmation) |
| `append` | `a` | create a new block **after** the current one and move to it |
| `prepend` | | create a new block **before** the current one and move to it |

### Navigation

| Command | Alias | What it does |
|---|---|---|
| `moveto N` | `m N` | go to block N |
| `move M` | `mv M` | move current block to position M |
| `swap M` | `sw M` | swap current block with block M |
| `start` | `home` | jump to the beginning |
| `end` | | jump to the end |
| `next` | `n` | next block |
| `prev` | | previous block |

### Viewing

| Command | Alias | What it does |
|---|---|---|
| `list` | `l` | short list (current marked with `>`) |
| `show` | `s` | show current block in detail |
| `raw` | `p` | print all blocks in raw form |

### File and exit

| Command | Alias | What it does |
|---|---|---|
| `save` | `w` | save `.docx` (creates file/directories) |
| `reload` | | re-read the file |
| `quit` | `q` | exit (asks about unsaved changes) |
| `wq` | | save and exit |

### Built into `cmd.Cmd`

| Command | What it does |
|---|---|
| `help` / `?` | command list |
| `help <cmd>` | help for a command |
| `Ctrl-D` | same as `quit` |
| `Ctrl-C` | soft exit |

## Multiline input

`edit`, `replace`, `insert`, `append` use multiline input.
Terminate with a line containing only a dot:

    > First line
    > Second with {text format:"..."}inline{/text}
    > .
    Block updated.

## Sample session

```
$ doc-editor example.docx
Loaded: example.docx — 2 blocks

[1/2]> list
>  1 [h:1]    First-level heading
   2          Paragraph with text

[1/2]> edit
Current content of block 1:
First-level heading
Enter text. A single '.' ends input.
> New heading
> .
Content of block 1 updated.

[1/2]> append
Created block 2 (cursor on it).

[2/2]> edit
...
> Paragraph with {text format:"text-style(bold)"}bold{/text}
> .

[2/2]> moveto 1
--- Block 1 ---
Params: type:"h:1"
Content: New heading
--- end ---

[1/2]> move 2
Block moved to position 2.
--- Block 2 ---
...

[2/2]> start
--- Block 1 ---
...

[1/2]> wq
Saved: example.docx — 2 blocks
```

## `move` and `swap` semantics

- `move M` — final position semantics (like `:move` in vim).
  The cursor stays on the moved block, now under a new number.
- `swap M` — the cursor stays on the same logical block, which
  is now located at position `M`.

## Command history

Saved via the `readline` module (Linux/macOS out of the box).
On Windows install `pyreadline3`, but the editor works without it.

## Limitations

- Inline tags `{text}` / `{image}` are part of block content;
  there are no dedicated commands for them.
- Tables are edited as regular blocks
  (`{block type:"table"}...`) without special commands.
- On `save`, the file is fully rewritten.