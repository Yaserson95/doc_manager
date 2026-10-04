"""
Утилиты для поиска верхнеуровневых тегов с учётом вложенности.
"""

import re


def find_top_level(text, open_re, close_tag):
    """
    Возвращает список (start, end) для верхнеуровневых вхождений
    open-тега и соответствующего close-тега. open_re — regex
    открывающего тега, close_tag — строка закрывающего.
    """
    pattern = re.compile(f'({open_re}|{re.escape(close_tag)})')
    depth = 0
    start = None
    result = []
    for m in pattern.finditer(text):
        if m.group(0) == close_tag:
            if depth > 0:
                depth -= 1
                if depth == 0:
                    result.append((start, m.end()))
                    start = None
        else:
            if depth == 0:
                start = m.start()
            depth += 1
    return result


def find_blocks(text):
    """Верхнеуровневые {block ...}...{/block}."""
    return [text[s:e] for s, e in
            find_top_level(text, r'\{block\b', '{/block}')]


def find_rows(text):
    """Верхнеуровневые {row ...}...{/row}."""
    return [text[s:e] for s, e in
            find_top_level(text, r'\{row\b', '{/row}')]


def find_cells(text):
    """Верхнеуровневые {cell ...}...{/cell}."""
    return [text[s:e] for s, e in
            find_top_level(text, r'\{cell\b', '{/cell}')]