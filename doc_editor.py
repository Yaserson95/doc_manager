# doc_editor.py
"""
Простой vim-подобный редактор блоков .docx с курсором на текущем блоке.

Запуск:
    python doc_editor.py example.docx

Команды:
    format    — задать параметры текущего блока (создаёт блок, если их нет)
    edit      — изменить содержимое текущего блока
    delete    — удалить текущий блок
    append    — создать новый блок после текущего
    prepend   — создать новый блок перед текущим
    moveto N  — перейти к N-му блоку
    list, show, raw, next, prev, save, reload, quit, wq
"""

import argparse
import cmd
import re
import sys
from pathlib import Path

try:
    import readline  # noqa: F401 — история команд на *nix
except ImportError:
    pass

from docx import Document
from docx.opc.exceptions import PackageNotFoundError

from docx_blocks import DocxBlockParser
from docx_blocks import DocxBlockSerializer


BLOCK_RE = re.compile(r'\{block\s*(.*?)\}(.*?)\{\/block\}', re.DOTALL)
BLOCK_PATTERN = r'\{block\s*.*?\}.*?\{\/block\}'


def split_blocks(text: str) -> list[str]:
    if not text:
        return []
    return re.findall(BLOCK_PATTERN, text, re.DOTALL)


def parse_block_parts(block_str: str) -> tuple[str, str]:
    m = BLOCK_RE.match(block_str)
    if not m:
        return '', block_str
    return m.group(1).strip(), m.group(2)


def compose_block(params: str, content: str) -> str:
    params = params.strip()
    if params:
        return f'{{block {params}}}{content}{{/block}}'
    return f'{{block}}{content}{{/block}}'


class DocEditor(cmd.Cmd):
    intro = (
        "Простой редактор блоков .docx (vim-like, с курсором)\n"
        "Команды: format, edit, delete, append, prepend,\n"
        "         moveto N, move M, swap M, start, end,\n"
        "         next, prev, list, show, raw, save, reload, quit, wq\n"
        "Многострочный ввод завершается строкой '.'.\n"
        "Введите 'help' или '?' для подробностей.\n"
    )
    def __init__(self, path: str):
        super().__init__()
        self.path = path
        self.blocks: list[str] = []
        self.current: int = -1  # 0-based индекс текущего блока, -1 если нет
        self.modified = False
        self._load()

    # ------------------------------------------------------------------
    # prompt
    # ------------------------------------------------------------------
    @property
    def prompt(self):
        if not self.blocks:
            return "(no blocks)> "
        if self.current < 0:
            return f"(-/{len(self.blocks)})> "
        return f"[{self.current + 1}/{len(self.blocks)}]> "

    # ------------------------------------------------------------------
    # I/O
    # ------------------------------------------------------------------
    def _load(self):
        try:
            doc = Document(self.path)
        except (FileNotFoundError, PackageNotFoundError):
            self.blocks = []
            self.current = -1
            print(f"Файла {self.path} нет — начинаем с пустого документа.")
            print("Создайте первый блок командой 'format' или 'append'.")
            return
        ser = DocxBlockSerializer()
        self.blocks = split_blocks(ser.serialize(doc))
        self.current = 0 if self.blocks else -1
        print(f"Загружено: {self.path} — блоков: {len(self.blocks)}")
        if self.blocks:
            self._print_current()

    def _save(self):
        # Создаём директорию при необходимости
        Path(self.path).parent.mkdir(parents=True, exist_ok=True)
        parser = DocxBlockParser()
        doc = parser.parse('\n'.join(self.blocks))
        doc.save(self.path)
        self.modified = False
        print(f"Сохранено: {self.path} — блоков: {len(self.blocks)}")

    # ------------------------------------------------------------------
    # Утилиты
    # ------------------------------------------------------------------
    def _print_current(self):
        if not self.blocks:
            print("(нет блоков)")
            return
        if self.current < 0:
            print("(курсор не установлен — используйте moveto N)")
            return
        params, content = parse_block_parts(self.blocks[self.current])
        print(f"--- Блок {self.current + 1} ---")
        print(f"Параметры: {params or '(нет)'}")
        one = content.replace('\n', ' ').strip()
        if len(one) > 200:
            one = one[:197] + '...'
        print(f"Содержимое: {one if one else '(пусто)'}")
        print("--- конец ---")

    def _read_multiline(self, prompt: str = '> ') -> str:
        print("Введите текст. Строка '.' — конец ввода.")
        lines: list[str] = []
        while True:
            try:
                line = input(prompt)
            except EOFError:
                break
            if line == '.':
                break
            lines.append(line)
        return '\n'.join(lines)

    def _confirm(self, question: str) -> bool:
        ans = input(f"{question} [y/N] ").strip().lower()
        return ans in ('y', 'yes', 'д', 'да')

    # ------------------------------------------------------------------
    # Основные команды
    # ------------------------------------------------------------------
    def do_format(self, arg):
        """format — задать параметры текущего блока.
Если блоков нет — создаётся первый блок и курсор ставится на него.
Пример: type:"h:1" format:"text-style(bold);txt-color(#ff0000)"
"""
        if not self.blocks:
            self.blocks.append(compose_block('', ''))
            self.current = 0
            self.modified = True
            print("Создан первый блок (курсор на нём).")
        elif self.current < 0:
            self.current = 0

        params, content = parse_block_parts(self.blocks[self.current])
        print(f"Текущие параметры: {params or '(нет)'}")
        print('Пример: type:"h:1" format:"text-style(bold);txt-color(#ff0000)"')
        try:
            new_params = input("Новые параметры (пусто — без изменений): ").strip()
        except EOFError:
            return
        if not new_params:
            print("Отменено.")
            return
        self.blocks[self.current] = compose_block(new_params, content)
        self.modified = True
        print(f"Параметры блока {self.current + 1} обновлены.")

    def do_edit(self, arg):
        """edit — изменить содержимое текущего блока (многострочно, '.' — конец)"""
        if not self.blocks:
            print("Нет блоков. Создайте блок: format, append или prepend.")
            return
        if self.current < 0:
            self.current = 0
        params, content = parse_block_parts(self.blocks[self.current])
        print(f"Текущее содержимое блока {self.current + 1}:")
        print(content if content else '(пусто)')
        new_content = self._read_multiline()
        self.blocks[self.current] = compose_block(params, new_content)
        self.modified = True
        print(f"Содержимое блока {self.current + 1} обновлено.")

    def do_delete(self, arg):
        """delete — удалить текущий блок"""
        if not self.blocks or self.current < 0:
            print("Нечего удалять.")
            return
        if not self._confirm(f"Удалить блок {self.current + 1}?"):
            print("Отменено.")
            return
        del self.blocks[self.current]
        self.modified = True
        if not self.blocks:
            self.current = -1
            print("Блок удалён. Блоков больше нет.")
        elif self.current >= len(self.blocks):
            self.current = len(self.blocks) - 1
            print(f"Блок удалён. Текущий: {self.current + 1}.")
        else:
            print(f"Блок удалён. Текущий: {self.current + 1}.")

    def do_append(self, arg):
        """append — создать новый (пустой) блок после текущего и перейти на него.
Если блоков нет — создаётся первый блок.
"""
        if not self.blocks:
            self.blocks.append(compose_block('', ''))
            self.current = 0
            self.modified = True
            print("Создан первый блок (курсор на нём).")
            return
        if self.current < 0:
            self.current = 0
        idx = self.current + 1
        self.blocks.insert(idx, compose_block('', ''))
        self.current = idx
        self.modified = True
        print(f"Создан блок {self.current + 1} (курсор на нём).")

    def do_prepend(self, arg):
        """prepend — создать новый (пустой) блок перед текущим и перейти на него.
Если блоков нет — создаётся первый блок.
"""
        if not self.blocks:
            self.blocks.append(compose_block('', ''))
            self.current = 0
            self.modified = True
            print("Создан первый блок (курсор на нём).")
            return
        if self.current < 0:
            self.current = 0
        idx = self.current
        self.blocks.insert(idx, compose_block('', ''))
        self.current = idx
        self.modified = True
        print(f"Создан блок {self.current + 1} (курсор на нём).")

    def do_moveto(self, arg):
        """moveto N — перейти к N-му блоку (1-based)"""
        try:
            n = int(arg.strip())
        except (ValueError, AttributeError):
            print("Использование: moveto N")
            return
        if not self.blocks:
            print("Нет блоков.")
            return
        if n < 1 or n > len(self.blocks):
            print(f"Ошибка: N должно быть в диапазоне 1..{len(self.blocks)}")
            return
        self.current = n - 1
        self._print_current()
        
    # ------------------------------------------------------------------
    # Перемещение текущего блока
    # ------------------------------------------------------------------
    def do_move(self, arg):
        """move M — переместить текущий блок на позицию M (1-based).
            Курсор остаётся на перемещённом блоке.
            """
        if not self.blocks or self.current < 0:
            print("Нет текущего блока.")
            return
        try:
            m = int(arg.strip())
        except (ValueError, AttributeError):
            print("Использование: move M")
            return
        if m < 1 or m > len(self.blocks):
            print(f"Ошибка: M должно быть в диапазоне 1..{len(self.blocks)}")
            return
        if m == self.current + 1:
            print("Блок уже на этой позиции.")
            return
        blk = self.blocks.pop(self.current)
        dst = m - 1
        self.blocks.insert(dst, blk)
        self.current = dst
        self.modified = True
        print(f"Блок перемещён на позицию {m}.")
        self._print_current()

    def do_swap(self, arg):
        """swap M — поменять текущий блок и блок M местами.
        Курсор остаётся на текущем блоке (он теперь на позиции M).
        """
        if not self.blocks or self.current < 0:
            print("Нет текущего блока.")
            return
        try:
            m = int(arg.strip())
        except (ValueError, AttributeError):
            print("Использование: swap M")
            return
        if m < 1 or m > len(self.blocks):
            print(f"Ошибка: M должно быть в диапазоне 1..{len(self.blocks)}")
            return
        j = m - 1
        if j == self.current:
            print("Это тот же самый блок.")
            return
        self.blocks[self.current], self.blocks[j] = (
            self.blocks[j], self.blocks[self.current]
        )
        self.current = j
        self.modified = True
        print(f"Блоки поменяны местами.")
        self._print_current()

    # ------------------------------------------------------------------
    # Навигация к краям
    # ------------------------------------------------------------------
    def do_start(self, arg):
        """start — перейти в начало документа (к первому блоку)"""
        if not self.blocks:
            print("Нет блоков.")
            return
        self.current = 0
        self._print_current()

    def do_end(self, arg):
        """end — перейти в конец документа (к последнему блоку)"""
        if not self.blocks:
            print("Нет блоков.")
            return
        self.current = len(self.blocks) - 1
        self._print_current()

    # ------------------------------------------------------------------
    # Навигация / просмотр
    # ------------------------------------------------------------------
    def do_list(self, arg):
        """list — краткий список блоков; '>' отмечает текущий"""
        if not self.blocks:
            print("(нет блоков)")
            return
        for i, b in enumerate(self.blocks, 1):
            params, content = parse_block_parts(b)
            m = re.search(r'type:"([^"]*)"', params)
            tag = f"[{m.group(1)}]" if m else ""
            one = content.replace('\n', ' ').strip()
            if len(one) > 60:
                one = one[:57] + '...'
            mark = '>' if (i - 1) == self.current else ' '
            print(f"{mark}{i:3} {tag:8} {one}")

    def do_show(self, arg):
        """show — показать текущий блок подробно"""
        self._print_current()

    def do_next(self, arg):
        """next — перейти к следующему блоку"""
        if not self.blocks:
            print("Нет блоков.")
            return
        if self.current < len(self.blocks) - 1:
            self.current += 1
        self._print_current()

    def do_prev(self, arg):
        """prev — перейти к предыдущему блоку"""
        if not self.blocks:
            print("Нет блоков.")
            return
        if self.current > 0:
            self.current -= 1
        self._print_current()

    def do_raw(self, arg):
        """raw — показать все блоки в исходном текстовом виде"""
        if not self.blocks:
            print("(пусто)")
            return
        print('\n'.join(self.blocks))

    # ------------------------------------------------------------------
    # Файл / выход
    # ------------------------------------------------------------------
    def do_save(self, arg):
        """save — сохранить изменения в .docx (создаёт файл при необходимости)"""
        try:
            self._save()
        except Exception as e:
            print(f"Ошибка сохранения: {e}")

    def do_reload(self, arg):
        """reload — перечитать файл, отбросив несохранённые изменения"""
        if self.modified and not self._confirm(
                "Есть несохранённые изменения. Перечитать?"):
            print("Отменено.")
            return
        self._load()
        self.modified = False

    def do_quit(self, arg):
        """quit — выйти (при несохранённых изменениях спросит)"""
        if self.modified:
            if self._confirm("Есть несохранённые изменения. Сохранить?"):
                try:
                    self._save()
                except Exception as e:
                    print(f"Ошибка сохранения: {e}")
                    return False
        return True

    def do_EOF(self, arg):
        """Ctrl-D — выход (как quit)"""
        print()
        return self.do_quit(arg)

    # ------------------------------------------------------------------
    # Псевдонимы
    # ------------------------------------------------------------------
    def do_l(self, arg):     return self.do_list(arg)
    def do_e(self, arg):     return self.do_edit(arg)
    def do_a(self, arg):     return self.do_append(arg)
    def do_d(self, arg):     return self.do_delete(arg)
    def do_s(self, arg):     return self.do_show(arg)
    def do_w(self, arg):     return self.do_save(arg)
    def do_q(self, arg):     return self.do_quit(arg)
    def do_p(self, arg):     return self.do_raw(arg)
    def do_f(self, arg):     return self.do_format(arg)
    def do_n(self, arg):     return self.do_next(arg)
    def do_m(self, arg):     return self.do_moveto(arg)
    def do_mv(self, arg):    return self.do_move(arg)
    def do_sw(self, arg):    return self.do_swap(arg)
    def do_home(self, arg):  return self.do_start(arg)   # удобно в пару к end

    def do_wq(self, arg):
        """wq — сохранить и выйти"""
        try:
            self._save()
        except Exception as e:
            print(f"Ошибка сохранения: {e}")
            return False
        return True

    # ------------------------------------------------------------------
    def emptyline(self):
        return

    def default(self, line):
        if line.strip() and not line.startswith('#'):
            print(f"Неизвестная команда: {line!r}. Введите 'help'.")


def main(argv=None) -> int:
    p = argparse.ArgumentParser(
        description="Простой vim-подобный редактор блоков .docx"
    )
    p.add_argument('file', help='Путь к .docx файлу (создаётся при сохранении)')
    args = p.parse_args(argv)

    editor = DocEditor(args.file)
    try:
        editor.cmdloop()
    except KeyboardInterrupt:
        print()
        editor.do_quit('')
    return 0


if __name__ == '__main__':
    sys.exit(main())