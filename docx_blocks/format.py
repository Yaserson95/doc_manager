"""
Разбор и представление строки format.

Синтаксис:
    параметр1; параметр2_со_свойствами(св1, св2); параметр3(arg)

- Параметры разделяются точкой с запятой ';'.
- Параметр без аргументов — просто ключ: `align`.
- Параметр с аргументами — `key(a, b, c)`.
- Пробелы вокруг имён и аргументов игнорируются.
- Дубликаты: побеждает последний.
- Нераспознанные фрагменты молча игнорируются.

Примеры:
    FormatSpec.parse('text-style(bold,underline);txt-color(#0000ff)')
    FormatSpec.parse('align(center)')
    FormatSpec.parse('padding(6pt,0,6pt,0);bg-color(#eef)')
    FormatSpec.parse('bold')
"""

import re
from collections import OrderedDict


_PARAM_RE = re.compile(r'''
    \s*
    (?P<name>[A-Za-z_][\w-]*)            # имя параметра
    (?:\((?P<args>[^)]*)\))?             # опциональные (args)
    \s*
''', re.VERBOSE)


class FormatSpec:
    """
    Разобранное представление format-строки.

    Хранит OrderedDict[str, list[str]]:
        ключ -> список аргументов (пустой список, если параметр без args).
    """

    __slots__ = ('_items',)

    def __init__(self, items=None):
        self._items = OrderedDict()
        if items:
            for k, v in items:
                self._items[k] = list(v)

    # ------------------------------------------------------------------
    # Разбор / сериализация
    # ------------------------------------------------------------------
    @classmethod
    def parse(cls, text: str) -> 'FormatSpec':
        spec = cls()
        if not text:
            return spec
        for part in str(text).split(';'):
            part = part.strip()
            if not part:
                continue
            m = _PARAM_RE.fullmatch(part)
            if not m:
                continue
            name = m.group('name')
            args_raw = m.group('args')
            args = []
            if args_raw is not None:
                args = [a.strip() for a in args_raw.split(',') if a.strip()]
            spec._items[name] = args
        return spec

    def to_string(self) -> str:
        parts = []
        for name, args in self._items.items():
            if args:
                parts.append(f'{name}({",".join(args)})')
            else:
                parts.append(name)
        return ';'.join(parts)

    def __str__(self):
        return self.to_string()

    def __repr__(self):
        return f'FormatSpec({self.to_string()!r})'

    # ------------------------------------------------------------------
    # Доступ
    # ------------------------------------------------------------------
    def __contains__(self, key) -> bool:
        return key in self._items

    def __iter__(self):
        return iter(self._items)

    def __len__(self) -> int:
        return len(self._items)

    def __bool__(self):
        return bool(self._items)

    def __eq__(self, other):
        if not isinstance(other, FormatSpec):
            return NotImplemented
        return self._items == other._items

    def keys(self):
        return self._items.keys()

    def items(self):
        return self._items.items()

    def has(self, key: str) -> bool:
        return key in self._items

    def get(self, key: str, default=None):
        return self._items.get(key, default)

    def args(self, key: str, default=None):
        """Список аргументов параметра или default."""
        return self._items.get(key, default)

    def arg(self, key: str, index: int = 0, default=None):
        """N-й аргумент параметра или default."""
        vals = self._items.get(key)
        if not vals or index >= len(vals):
            return default
        return vals[index]

    # ------------------------------------------------------------------
    # Изменение
    # ------------------------------------------------------------------
    def set(self, key: str, args=None) -> 'FormatSpec':
        """args — None | строка 'a,b' | список строк."""
        if args is None:
            self._items[key] = []
        elif isinstance(args, str):
            self._items[key] = [a.strip() for a in args.split(',') if a.strip()]
        else:
            self._items[key] = list(args)
        return self

    def unset(self, key: str) -> 'FormatSpec':
        self._items.pop(key, None)
        return self

    def merge(self, other: 'FormatSpec') -> 'FormatSpec':
        """Новый spec: поверх self накладываются значения other."""
        result = FormatSpec(self._items.items())
        if other is not None:
            for k, v in other._items.items():
                result._items[k] = list(v)
        return result

    def copy(self) -> 'FormatSpec':
        return FormatSpec(self._items.items())

    # ------------------------------------------------------------------
    # Типизированные разборы
    # ------------------------------------------------------------------
    def text_styles(self) -> dict:
        """{'bold': True, 'italic': True, ...} из text-style(bold,...)."""
        result = {}
        for s in self._items.get('text-style', []):
            result[s.strip().lower()] = True
        return result

    def padding(self) -> dict:
        """
        {'left': '10pt', 'top': '0', ...} с учётом padding(l,t,r,b)
        и отдельных padding-left/top/right/bottom.
        """
        result = {}
        vals = self._items.get('padding')
        if vals and len(vals) == 4:
            result['left'] = vals[0]
            result['top'] = vals[1]
            result['right'] = vals[2]
            result['bottom'] = vals[3]
        for side in ('left', 'top', 'right', 'bottom'):
            k = f'padding-{side}'
            if k in self._items and self._items[k]:
                result[side] = self._items[k][0]
        return result

    def border(self):
        """(size, style, color) или None."""
        if 'border' not in self._items:
            return None
        a = self._items['border']
        size = a[0] if len(a) > 0 else '1pt'
        style = a[1] if len(a) > 1 else 'single'
        color = a[2] if len(a) > 2 else '#000000'
        return size, style, color