# `DocxBlockSerializer`

Модуль: `docx_blocks/serializer.py`

Превращает `.docx` в текст в формате блоков.

## Импорт

```python
from docx_blocks import DocxBlockSerializer
# или
from docx_blocks.serializer import DocxBlockSerializer
```

## Публичный API

### `serialize(docx_input) -> str`

**Параметры**

- `docx_input` — `str` (путь к файлу) или `docx.Document`.

**Возвращает** — `str`; блоки разделены `\n`.

**Пример**

```python
ser = DocxBlockSerializer()

text = ser.serialize('example.docx')

from docx import Document
text = ser.serialize(Document('example.docx'))
```

### Атрибуты

| Атрибут | Тип | Назначение |
|---|---|---|
| `table_serializer` | `DocxTableSerializer` | Сериализация таблиц |

## Что восстанавливается

| Элемент docx | Что получается в тексте |
|---|---|
| Заголовок `Heading N` | `type:"h:N"` |
| `List Bullet` / `List Bullet 2` | `type:"ul:1"` / `type:"ul:2"` |
| `List Number` | `type:"ol:N"` |
| Обычный абзац | без `type` |
| Таблица | `{block type:"table"}...{/block}` |
| Отступы до/после/слева/справа | `padding-top/bottom/left/right(...)` |
| Межстрочный интервал | `line-height(...)` |
| Заливка абзаца | `bg-color(#hex)` |
| Выравнивание абзаца | `align(...)` |
| Полужирный / курсив / подчёркнутый | `text-style(bold,...)` |
| Цвет текста | `txt-color(#hex)` |
| Размер шрифта | `font-size(Npt)` |
| Run с отличающимся форматом | `{text format:"..."}...{/text}` |
| Картинка | `{image text:""}data:image/...;base64,...{/image}` |

## Внутренние методы

| Метод | Назначение |
|---|---|
| `_serialize_paragraph(p)` | один абзац → строка блока |
| `_detect_block_type(p)` | тип блока (h / ul / ol / p) |
| `_extract_block_format_dict(p)` | формат абзаца → `FormatSpec` |
| `_serialize_content(p, spec)` | содержимое абзаца с `{text}` / `{image}` |
| `_extract_run_format_dict(run)` | формат фрагмента → `FormatSpec` |
| `_run_format_differs(run, block)` | нужно ли оборачивать в `{text}` |
| `_extract_images(run)` | картинки из run |
| `_serialize_image(img)` | картинка → `{image ...}` |
| `_get_paragraph_shading(p)` | цвет заливки абзаца |
| `_emu_to_pt(emu)` | EMU → pt |

Обход `body` идёт по `w:p` и `w:tbl`, чтобы сохранить порядок
абзацев и таблиц.

## Ограничения

- Колонтитулы, сноски, секции игнорируются.
- `padding(...)` разворачивается в отдельные `padding-*`.
- У `{image}` не восстанавливается `format` (обтекание/положение).
- `border(...)` у таблиц восстанавливается обобщённо:
  `border(1pt,single,#000000)`.