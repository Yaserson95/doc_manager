# CLI `doc-cli`

File: `doc_cli.py` (installed as the `doc-cli` command).

A command-line tool for fine-grained editing of `.docx` and for
serializing to the text format.

## Synopsis

    doc-cli FILE.docx [command] [args]

Without a command, `--rewrite` is used.

## Commands

| Command | What it does | N |
|---|---|---|
| `--rewrite "..."` | overwrite the whole file | — |
| `--append "..."` | append blocks at the end | — |
| `--prepend "..."` | prepend blocks at the beginning | — |
| `--insert N "..."` | replace block N | `1..total` |
| `--insert-after N "..."` | insert after block N | `1..total` |
| `--insert-before N "..."` | insert before block N | `1..total+1` |
| `--serialize` | print structure to stdout | — |
| `--serialize out.txt` | save structure to file | — |

Blocks can be passed inline or via `--file` / `-f`:

    doc-cli example.docx --append "..."           # inline
    doc-cli example.docx --append --file blocks.txt
    doc-cli example.docx --insert 2 --file blocks.txt
    doc-cli example.docx --file blocks.txt         # == --rewrite

Or read from stdin:

    echo "{block}Hi{/block}" | doc-cli out.docx --append --stdin

The positional form without a command is equivalent to `--rewrite`:

    doc-cli example.docx "{block}...{/block}"

## Modifiers

| Flag | Meaning |
|---|---|
| `--stdin` | read blocks from stdin |
| `-q` / `--quiet` | suppress the `OK: ...` line |
| `--json` | print the result as JSON |

## Examples

```bash
# overwrite
doc-cli example.docx \
    "{block type:\"h:1\"}Header{/block}{block}Paragraph{/block}"

# append from a file
doc-cli example.docx --append --file blocks.txt

# replace block 2
doc-cli example.docx --insert 2 "{block type:\"p\"}New{/block}"

# insert before block 1
doc-cli example.docx --insert-before 1 "{block}Title{/block}"

# serialize to stdout
doc-cli example.docx --serialize

# serialize to a file
doc-cli example.docx --serialize out.txt

# JSON output (useful for tool-calling)
doc-cli example.docx --append --file blocks.txt --json --quiet
```

## Behavior

- **File creation**: `--rewrite`, `--append`, `--prepend` create the
  file if it does not exist. `--insert*` require an existing file.
- **N indexing** — 1-based.
- **Out of range** — error message, file is left untouched.
- **Mutual exclusion**: commands from the group
  `--rewrite / --append / --prepend / --insert* / --serialize`
  cannot be combined.
- **`--serialize`** works standalone; `--file` is ignored.

## Exit codes

| Code | Meaning |
|---|---|
| 0 | success |
| 1 | logical error (N out of range, serialization error) |
| 2 | bad arguments or file read error |
| 3 | write error |

## Dependencies

- `docx_blocks.parser.DocxBlockParser`
- `docx_blocks.serializer.DocxBlockSerializer`
- `docx_blocks.tags.find_blocks`
- `python-docx`

## Limitations

- On save, the `.docx` is fully rewritten; headers, footers,
  footnotes and sections are not preserved.
- An invalid block in the argument leads to an empty block list
  (with `--rewrite` this clears the document).