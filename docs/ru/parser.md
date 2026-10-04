# `DocxBlockParser`

Модуль: `docx_blocks/parser.py`

Превращает текст в формате блоков в объект `docx.Document`.

## Импорт

```python
from docx_blocks import DocxBlockParser
# или
from docx_blocks.parser import DocxBlockParser
```

## Публичный API

### `parse(input_text: str) -> Document`

Парсит строку и возвращает готовый `docx.Document`.

**Пример**

```python
parser = DocxBlockParser()
doc = parser.parse(open('input.txt', encoding='utf-8').read())
doc.save('output.docx')
```

### Атрибуты

| Атрибут | Тип | Назначение |
|---|---|---|
| `formatter` | `DocxFormatter` | Применение форматов |
| `table_parser` | `DocxTableParser` | Обработка `{block type:"table"}` |

Можно переиспользовать один парсер для нескольких документов —
внутреннее состояние между вызовами `parse()` не сохраняется.

## Поддерживаемые элементы

| Элемент | Поддержка |
|---|---|
| `type:"p"` | обычный абзац |
| `type:"h:N"` | заголовок N-го уровня |
| `type:"ul:N"` | маркированный список |
| `type:"ol:N"` | нумерованный список |
| `type:"table"` | таблица (см. [table.md](table.md)) |
| `format:"text-style(...)"` | bold / italic / underline |
| `format:"bg-color(#hex)"` | заливка абзаца |
| `format:"txt-color(#hex)"` | цвет текста |
| `format:"font-size(Npt)"` | размер шрифта |
| `format:"line-height(...)"` | межстрочный интервал |
| `format:"align(...)"` | выравнивание |
| `format:"padding(...)"` | отступы |
| `{text ...}...{/text}` | встроенное форматирование |
| `{image ...}...{/image}` | картинка (путь или base64) |

## Внутренние методы

| Метод | Назначение |
|---|---|
| `_parse_params(s)` | разбор пар `ключ:"значение"` |
| `_parse_block_type(s)` | `"h:2"` → `('h', 2)` |
| `_parse_block_format(s)` | строка → `FormatSpec` |
| `_parse_block(s)` | разбор одного блока |
| `_parse_content(s)` | разбор содержимого на текст и встроенные теги |
| `_parse_size(s)` | `"14pt"` → `Pt(14)` (делегирует `DocxFormatter`) |
| `_apply_block_format(p, spec)` | применить формат (делегирует `DocxFormatter`) |
| `_apply_run_format(run, b, t)` | применить формат (делегирует) |
| `_add_paragraph_with_type(c, t, lvl)` | создать абзац в `Document` или `_Cell` |
| `_process_block_content(p, c, spec)` | обработка содержимого блока |
| `_add_block_to_container(c, s)` | обработка одного блока |

## Ошибки

Собственных исключений не бросает. Некорректные блоки пропускаются,
некорректные размеры/цвета игнорируются «молча».

## Ограничения

- Не поддерживаются колонтитулы, сноски и секции.
- Обтекание картинок (inline / floating) не учитывается —
  все картинки вставляются как inline.