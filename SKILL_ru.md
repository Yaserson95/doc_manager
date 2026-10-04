---
name: docx-blocks
description: Создание и редактирование Word-документов (.docx) через текстовый блочный формат. Используй, когда нужно сгенерировать .docx с заголовками, абзацами, списками, таблицами или картинками; изменить существующий документ; или извлечь его структуру в текст.
license: MIT
---

# docx-blocks

## Что это

Документ `.docx` описывается как последовательность блоков в текстовом
формате. CLI `doc-cli` конвертирует блоки в `.docx` и обратно.

Инструмент командной строки — `doc-cli`. Он установлен глобально либо
доступен как `python -m doc_cli`.

## Когда использовать

- Пользователь просит **создать** `.docx` с текстом, заголовками,
  списками, таблицами, картинками.
- Пользователь просит **изменить** существующий `.docx`: добавить,
  заменить, вставить или удалить блок.
- Нужно **прочитать** содержимое `.docx` в текстовом виде.

**Не используй**, если нужен `.docx` с колонтитулами, сносками,
полями/секциями или точным позиционированием картинок — эти
возможности не поддерживаются.

## Рабочий процесс

### Шаг 1. Понять, что уже есть в файле

Если файл существует — сериализуй его, чтобы увидеть структуру:

```bash
doc-cli path/to/file.docx --serialize
```

Вывод — список блоков в текстовом формате. Каждый блок можно
адресовать по индексу (1-based): `--insert N`, `--insert-after N`,
`--insert-before N`.

### Шаг 2. Подготовить блоки

Запиши блоки в temp-файл (удобнее, чем inline-строкой — не нужно
экранировать кавычки):

```bash
cat > /tmp/blocks.txt <<'EOF'
{block type:"h:1" format:"align(center);text-style(bold)"}
Заголовок документа
{/block}
{block}
Обычный абзац с {text format:"text-style(bold)"}выделением{/text}.
{/block}
{block type:"ul:1"}Пункт 1{/block}
{block type:"ul:1"}Пункт 2{/block}
EOF
```

### Шаг 3. Применить команду

```bash
# создать с нуля (или перезаписать)
doc-cli output.docx --rewrite --file /tmp/blocks.txt

# добавить в конец
doc-cli output.docx --append --file /tmp/blocks.txt

# добавить в начало
doc-cli output.docx --prepend --file /tmp/blocks.txt

# заменить 2-й блок
doc-cli output.docx --insert 2 --file /tmp/blocks.txt

# вставить после 2-го
doc-cli output.docx --insert-after 2 --file /tmp/blocks.txt

# вставить перед 1-м
doc-cli output.docx --insert-before 1 --file /tmp/blocks.txt
```

### Шаг 4. Проверить

```bash
doc-cli output.docx --serialize
```

Если структура совпадает с ожидаемой — задача выполнена.

## Синтаксис блоков (кратко)

### Абзац

```
{block}Просто текст{/block}
```

### Заголовок

```
{block type:"h:1"}Заголовок первого уровня{/block}
{block type:"h:2"}Второго{/block}
```

`h:1` … `h:6`.

### Список

```
{block type:"ul:1"}Маркированный пункт{/block}
{block type:"ol:1"}Нумерованный пункт{/block}
```

### Форматирование

Параметр `format` — строка из `key(value)`, разделённых `;`.

```
{block format:"align(center);text-style(bold,italic);font-size(14pt)"}
Текст
{/block}
```

| Ключ | Значения | Что делает |
|---|---|---|
| `text-style(...)` | `bold`, `italic`, `underline` (можно несколько через `,`) | стиль шрифта |
| `txt-color(#hex)` | `#RRGGBB` или `#RGB` | цвет текста |
| `bg-color(#hex)` | `#RRGGBB` или `#RGB` | цвет фона |
| `font-size(14pt)` | `pt`, `px`, `in`, `cm` или число | размер шрифта |
| `line-height(1.5)` | число или размер | межстрочный интервал |
| `align(center)` | `left`, `center`, `right`, `justify` | выравнивание |
| `padding(10pt,0,10pt,0)` | `l,t,r,b` | отступы |
| `padding-left(10pt)` | размер | отдельная сторона |

### Встроенное форматирование

Внутри блока можно выделить фрагмент:

```
{block}
Обычный текст и {text format:"txt-color(#c00);text-style(bold)"}красный жирный{/text}.
{/block}
```

Поддерживаются те же ключи, кроме `padding*`.

### Таблица

```
{block type:"table" format:"align(center);width(100%);border(1pt,solid,#333)"}
{row header:"true" height:"24pt"}
{cell format:"align(center);bg-color(#eef);text-style(bold)"}
{block}Заголовок A{/block}
{/cell}
{cell format:"align(center);bg-color(#eef);text-style(bold)"}
{block}Заголовок B{/block}
{/cell}
{/row}
{row}
{cell}Значение 1{/cell}
{cell format:"align(right)"}{block}42{/block}{/cell}
{/row}
{/block}
```

Объединения — вне `format`:

```
{cell colspan:"2" rowspan:"1" format:"..."}...{/cell}
```

Ключи `format` ячейки: `align`, `valign`, `width`, `bg-color`,
`padding*`, а также `text-style`, `txt-color`, `font-size`,
`line-height` — как значения по умолчанию для вложенных блоков.

### Картинка

```
{block}
{image text:"Диаграмма"}/path/to/image.png{/image}
{/block}
```

Или base64:

```
{image text:""}data:image/png;base64,iVBORw0KGgo...{/image}
```

## Команды CLI

| Команда | Что делает |
|---|---|
| `--rewrite "..."` | перезаписать весь файл |
| `--append "..."` | добавить блоки в конец |
| `--prepend "..."` | добавить блоки в начало |
| `--insert N "..."` | заменить N-й блок |
| `--insert-after N "..."` | вставить после N-го |
| `--insert-before N "..."` | вставить перед N-м |
| `--serialize` | вывести структуру в консоль |
| `--serialize out.txt` | сохранить структуру в файл |
| `--file FILE` / `-f FILE` | читать блоки из файла |
| `--stdin` | читать блоки из stdin |
| `--quiet` / `-q` | не печатать `OK: ...` |
| `--json` | выводить результат в JSON |

Без команды — `--rewrite`.

## Проверка успеха

После каждой команды:

```bash
doc-cli output.docx --serialize --quiet
```

и убедись, что блоки появились/изменились/удалились как ожидалось.

Если добавлена опция `--json`, используй:

```bash
doc-cli output.docx --insert 2 --file /tmp/b.txt --json
```

Ответ:

```json
{"ok": true, "file": "output.docx", "blocks_before": 3, "blocks_after": 3}
```

## Коды возврата

| Код | Значение |
|---|---|
| 0 | успех |
| 1 | логическая ошибка (N вне диапазона) |
| 2 | ошибка аргументов или чтения файла |
| 3 | ошибка записи |

## Частые ошибки и как их избегать

- **Не экранируй кавычки.** Если пишешь блоки inline, используй
  одинарные кавычки вокруг строки в bash: `'{block type:"h:1"}...'`.
  Лучше — всегда писать блоки через `--file`.
- **`--insert*` требует существующий файл.** Если файла ещё нет —
  используй `--rewrite` или `--append`, они создают файл.
- **N — 1-based.** Первый блок — `1`, не `0`.
- **Цвета.** Поддерживаются `#RRGGBB` и короткие `#RGB`.
- **Цвет фона абзаца** требует наличия текста в блоке, иначе
  визуально его не видно.
- **`align` ячейки** применяется к параграфам внутри неё, а не
  к самой ячейке (ограничение OOXML) — но визуально работает как
  ожидается.

## Пример полного сценария

Пользователь: «Сделай документ с заголовком, двумя абзацами и таблицей
3×2».

```bash
cat > /tmp/doc.txt <<'EOF'
{block type:"h:1" format:"align(center)"}
Отчёт за квартал
{/block}
{block}
Первый абзац — вводная часть.
{/block}
{block}
Второй абзац с {text format:"text-style(bold)"}выделением{/text}.
{/block}
{block type:"table" format:"border(1pt,solid,#333)"}
{row header:"true"}
{cell format:"text-style(bold);align(center)"}{block}Месяц{/block}{/cell}
{cell format:"text-style(bold);align(center)"}{block}Выручка{/block}{/cell}
{/row}
{row}
{cell}Январь{/cell}
{cell format:"align(right)"}{block}100 000 ₽{/block}{/cell}
{/row}
{row}
{cell}Февраль{/cell}
{cell format:"align(right)"}{block}120 000 ₽{/block}{/cell}
{/row}
{/block}
EOF

doc-cli report.docx --rewrite --file /tmp/doc.txt --json
doc-cli report.docx --serialize
```