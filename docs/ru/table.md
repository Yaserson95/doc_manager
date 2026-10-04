# Таблицы

Модуль: `docx_blocks/table.py`

Классы:

- `DocxTableParser` — разбор `{block type:"table"}...{/block}` и добавление
  таблицы в `Document` или в ячейку (`_Cell`);
- `DocxTableSerializer` — сериализация `docx.Table` обратно в блочный формат.

## Формат таблицы

    {block type:"table" format:"..."}
    {row height:"24pt" header:"true"}
    {cell colspan:"2" rowspan:"1" format:"align(center);bg-color(#eef)"}
    {block}...{/block}
    {/cell}
    {/row}
    {/block}

### Параметры таблицы (`format`)

| Ключ | Пример | Значения |
|---|---|---|
| `align` | `align(center)` | left / center / right |
| `width` | `width(100%)` или `width(400pt)` | размер или проценты |
| `border` | `border(1pt,solid,#333)` | размер, стиль, цвет |

Если `border` не указан — применяется дефолт `1pt single #999999`.

### Параметры строки (обычные параметры)

| Ключ | Пример |
|---|---|
| `height` | `height:"24pt"` |
| `header` | `header:"true"` (повторять при разрыве страницы) |

### Параметры ячейки

Вне `format` — только `colspan` и `rowspan`:

    {cell colspan:"3" rowspan:"2" format:"..."}...{/cell}

Внутри `format` — то же, что у блока, плюс ячеечные ключи:

| Ключ | Пример | Куда идёт |
|---|---|---|
| `align` | `align(center)` | выравнивание содержимого |
| `valign` | `valign(middle)` | top / middle / bottom |
| `width` | `width(40pt)` | ширина ячейки |
| `bg-color` | `bg-color(#eef)` | фон ячейки |
| `padding` | `padding(6pt,3pt,6pt,3pt)` | отступы l,t,r,b |
| `padding-left` | `padding-left(6pt)` | отдельная сторона |
| `text-style` | `text-style(bold)` | по умолчанию для вложенных блоков |
| `txt-color` | `txt-color(#222)` | то же |
| `font-size` | `font-size(12pt)` | то же |
| `line-height` | `line-height(1.3)` | то же |

**Ключи `text-style`, `txt-color`, `font-size`, `line-height`,
`align`** не применяются к самой ячейке — они становятся
значениями по умолчанию для вложенных блоков.

### Содержимое ячейки

Один или несколько `{block}...{/block}`. Может быть вложенная
таблица:

    {cell}
    {block type:"table"}...{/block}
    {/cell}

## Примеры

Простая таблица с шапкой:

    {block type:"table" format:"align(center);width(100%);border(1pt,solid,#333)"}
    {row header:"true" height:"24pt"}
    {cell format:"align(center);bg-color(#eef);text-style(bold)"}
    {block}Параметр{/block}
    {/cell}
    {cell format:"align(center);bg-color(#eef);text-style(bold)"}
    {block}Значение{/block}
    {/cell}
    {/row}
    {row}
    {cell}Цена{/cell}
    {cell format:"align(right)"}{block}100{/block}{/cell}
    {/row}
    {/block}

Объединения:

    {block type:"table"}
    {row}
    {cell colspan:"2" format:"align(center)"}
    {block}Заголовок на две колонки{/block}
    {/cell}
    {/row}
    {row}
    {cell rowspan:"2" format:"valign(middle)"}
    {block}Общая ячейка{/block}
    {/cell}
    {cell}A{/cell}
    {/row}
    {row}
    {cell}B{/cell}
    {/row}
    {/block}

## Программный API

### `DocxTableParser(block_parser)`

Принимает ссылку на `DocxBlockParser` (использует его
`formatter`, `_parse_params`, `_parse_block`, `_add_paragraph_with_type`,
`_process_block_content`).

Метод:

    add_table(container, table_params: dict, table_content: str) -> Table | None

- `container` — `Document` или `_Cell`;
- `table_params` — `params` блока (с ключами `type`, `format`);
- `table_content` — содержимое блока (внутри `{block ...}`).

### `DocxTableSerializer(block_serializer)`

Принимает ссылку на `DocxBlockSerializer` (использует его
`_serialize_paragraph`).

Метод:

    serialize_table(table) -> str

## Ограничения

- `align` в `tcPr` невалиден по схеме OOXML — применяется к
  параграфам ячейки (см. [formatter.md](formatter.md)).
- `border` при обратной сериализации восстанавливается
  обобщённо (`1pt,single,#000000`), точные параметры рёбер
  не сохраняются.
- Позиционирование картинок внутри ячеек — как и везде,
  без обтекания.