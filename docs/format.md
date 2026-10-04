# Блочный формат

## Блок

    {block [параметры]}содержимое{/block}

- Параметры опциональны.
- Параметры — пары `ключ:"значение"`, разделены пробелами.
- Содержимое может включать любой текст, встроенные теги
  `{text}` / `{image}`, а также вложенные `{block}` (например,
  внутри ячеек таблицы).

### Параметры блока

| Ключ | Формат | Описание |
|---|---|---|
| `type` | `"p"` / `"h:N"` / `"ul:N"` / `"ol:N"` / `"table"` | Тип блока. По умолчанию `p`. `N` — уровень |
| `format` | строка `key(value);key(value);...` | Форматирование |

### Форматирование (`format`)

| Ключ | Пример | Что делает |
|---|---|---|
| `text-style(...)` | `text-style(bold,underline)` | bold / italic / underline |
| `bg-color(#hex)` | `bg-color(#fff2cc)` | Цвет фона |
| `txt-color(#hex)` | `txt-color(#0000ff)` | Цвет текста |
| `font-size(...)` | `font-size(14pt)` | Размер шрифта |
| `line-height(...)` | `line-height(1.5)` или `line-height(18pt)` | Межстрочный интервал |
| `align(...)` | `align(center)` | Выравнивание: left / center / right / justify |
| `padding(l,t,r,b)` | `padding(10pt,0,10pt,0)` | Отступы (left, top, right, bottom) |
| `padding-left(...)` | `padding-left(12pt)` | Отступ слева |
| `padding-top(...)` | `padding-top(6pt)` | Отступ сверху |
| `padding-right(...)` | `padding-right(12pt)` | Отступ справа |
| `padding-bottom(...)` | `padding-bottom(6pt)` | Отступ снизу |

Размеры: `pt`, `px`, `in`, `cm` или число (трактуется как `pt`).

## Встроенные теги

### `{text}`

    {text format:"..."}текст{/text}

Поддерживаются те же ключи, что и в `format` блока, кроме `padding*`.

### `{image}`

    {image format:"..." text:"alt"}[путь | base64]{/image}

- `format` — параметры вписывания (в текущей версии используется частично).
- `text` — альтернативный текст.
- Содержимое — путь к файлу или data-URI (`data:image/png;base64,...`).

### Таблицы

    {block type:"table" format:"..."}
    {row height:"24pt" header:"true"}
    {cell colspan:"2" format:"align(center);bg-color(#eef)"}
    {block}...{/block}
    {/cell}
    {/row}
    {/block}

Подробно — см. [table.md](table.md).

## Примеры

Простой абзац:

    {block}Это просто абзац.{/block}

Заголовок с форматированием:

    {block type:"h:1" format:"text-style(bold,underline);txt-color(#0000ff);padding-bottom(16pt)"}
    Синий подчёркнутый полужирный заголовок
    {/block}

Абзац с выделением:

    {block}
    Текст с {text format:"txt-color(#ffff00)"}жёлтым{/text} выделением.
    {/block}

Список:

    {block type:"ul:1"}Первый пункт{/block}
    {block type:"ul:1"}Второй пункт{/block}

## Синтаксис format (детали)

- Разделитель параметров — `;`.
- Разделитель аргументов внутри скобок — `,`.
- Пробелы вокруг имён и аргументов игнорируются.
- Дубликаты: побеждает последний.
- Неизвестные параметры сохраняются и не ломают разбор.
- Форма `key:(value)` (с двоеточием) **не поддерживается** —
  только `key(value)`.