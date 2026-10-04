# Docx Blocks

Набор утилит для работы с `.docx` через текстовый блочный формат.

## Что это

Документ описывается как последовательность блоков:

    {block [параметры]}содержимое{/block}

Внутри содержимого могут быть встроенные теги:

    {text format:"..."}текст{/text}
    {image format:"..." text:"alt"}[путь | base64]{/image}

Поддерживаются таблицы:

    {block type:"table" format:"..."}
    {row}{cell format:"..."}{block}...{/block}{/cell}{/row}
    {/block}

Формат легко читать, писать и править вручную, а конвертация
в `.docx` и обратно полностью автоматизирована.

## Структура проекта

    doc_manager/
    ├── docx_blocks/              # библиотека
    │   ├── __init__.py           # публичный API
    │   ├── format.py             # FormatSpec
    │   ├── formatter.py          # DocxFormatter
    │   ├── parser.py             # DocxBlockParser
    │   ├── serializer.py         # DocxBlockSerializer
    │   ├── table.py              # DocxTableParser / DocxTableSerializer
    │   └── tags.py               # поиск вложенных тегов
    ├── doc_cli.py                # CLI (пакетная точка входа doc-cli)
    ├── doc_editor.py             # интерактивный редактор (doc-editor)
    ├── docs/                     # документация
    ├── examples/                 # примеры
    ├── tests/                    # тесты
    ├── pyproject.toml
    └── README.md

## Установка

```bash
pip install -e .
```

Появятся команды `doc-cli` и `doc-editor`; либо вызывай напрямую:

```bash
python doc_cli.py --help
python doc_editor.py --help
```

Зависимости: Python 3.10+, `python-docx`.

## Быстрый старт

### Парсинг текста в .docx

```python
from docx_blocks import DocxBlockParser

text = '{block type:"h:1"}Заголовок{/block}'
doc = DocxBlockParser().parse(text)
doc.save('out.docx')
```

### Сериализация .docx в текст

```python
from docx_blocks import DocxBlockSerializer

text = DocxBlockSerializer().serialize('out.docx')
print(text)
```

### CLI

```bash
doc-cli example.docx --append "{block}Ещё абзац{/block}"
doc-cli example.docx --serialize out.txt
```

### Интерактивный редактор

```bash
doc-editor example.docx
```

## Документация

- [format.md](docs/format.md) — синтаксис блочного формата
- [parser.md](docs/parser.md) — `DocxBlockParser`
- [serializer.md](docs/serializer.md) — `DocxBlockSerializer`
- [formatter.md](docs/formatter.md) — `DocxFormatter`
- [table.md](docs/table.md) — таблицы
- [tags.md](docs/tags.md) — утилиты поиска тегов
- [cli.md](docs/cli.md) — CLI `doc-cli`
- [editor.md](docs/editor.md) — редактор `doc-editor`

## Ограничения

- Обрабатываются параграфы и таблицы; колонтитулы, сноски и
  секции — нет.
- У `{image}` не восстанавливаются параметры обтекания/положения.
- `padding(l,t,r,b)` при обратной сериализации разворачивается
  в отдельные `padding-left/top/right/bottom`.