# CLI `doc-cli`

Файл: `doc_cli.py` (устанавливается как команда `doc-cli`).

Утилита командной строки для точечного редактирования `.docx`
и сериализации в текстовый формат.

## Синопсис

    doc-cli FILE.docx [команда] [аргументы]

Без команды выполняется `--rewrite`.

## Команды

| Команда | Что делает | N |
|---|---|---|
| `--rewrite "..."` | перезаписать весь файл | — |
| `--append "..."` | добавить блоки в конец | — |
| `--prepend "..."` | добавить блоки в начало | — |
| `--insert N "..."` | заменить блок N | `1..total` |
| `--insert-after N "..."` | вставить после блока N | `1..total` |
| `--insert-before N "..."` | вставить перед блоком N | `1..total+1` |
| `--serialize` | вывести структуру в консоль | — |
| `--serialize out.txt` | сохранить структуру в файл | — |

Блоки можно передать inline-строкой или через `--file`/`-f`:

    doc-cli example.docx --append "..."           # inline
    doc-cli example.docx --append --file blocks.txt
    doc-cli example.docx --insert 2 --file blocks.txt
    doc-cli example.docx --file blocks.txt         # == --rewrite

Позиционная форма без команды эквивалентна `--rewrite`:

    doc-cli example.docx "{block}...{/block}"

## Примеры

```bash
# перезаписать
doc-cli example.docx \
    "{block type:\"h:1\"}Заголовок{/block}{block}Абзац{/block}"

# добавить в конец из файла
doc-cli example.docx --append --file blocks.txt

# заменить 2-й блок
doc-cli example.docx --insert 2 "{block type:\"p\"}Новый{/block}"

# вставить перед 1-м
doc-cli example.docx --insert-before 1 "{block}Титул{/block}"

# сериализация в консоль
doc-cli example.docx --serialize

# сериализация в файл
doc-cli example.docx --serialize out.txt
```

## Поведение

- **Создание файла**: `--rewrite`, `--append`, `--prepend` создают
  файл, если его нет. `--insert*` требуют существующий файл.
- **N-индексация** — 1-based.
- **Выход за диапазон** — сообщение об ошибке, файл не меняется.
- **Взаимоисключение**: команды из группы
  `--rewrite / --append / --prepend / --insert* / --serialize`
  нельзя указывать вместе.
- **`--serialize`** работает и без остальных флагов; `--file`
  игнорируется.

## Коды возврата

| Код | Значение |
|---|---|
| 0 | успех |
| 1 | логическая ошибка (N вне диапазона, ошибка сериализации) |
| 2 | ошибка аргументов или чтения файла |
| 3 | ошибка записи |

## Зависимости

- `docx_blocks.parser.DocxBlockParser`
- `docx_blocks.serializer.DocxBlockSerializer`
- `docx_blocks.tags.find_blocks`
- `python-docx`

## Ограничения

- При сохранении `.docx` файл перезаписывается целиком;
  колонтитулы, секции и сноски не сохраняются.
- Некорректный блок в аргументе приводит к пустому списку
  блоков (при `--rewrite` это очистит документ).