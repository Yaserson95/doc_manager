# doc_cli.py
"""
CLI для редактирования .docx через блочный формат.

Использование:
    python doc_cli.py example.docx "{block}...{/block}"           # --rewrite
    python doc_cli.py example.docx --rewrite   "{block}...{/block}"
    python doc_cli.py example.docx --append    "{block}...{/block}"
    python doc_cli.py example.docx --prepend   "{block}...{/block}"
    python doc_cli.py example.docx --insert        N "{block}...{/block}"
    python doc_cli.py example.docx --insert-after  N "{block}...{/block}"
    python doc_cli.py example.docx --insert-before N "{block}...{/block}"

Любую команду можно использовать с --file вместо inline-строки:
    python doc_cli.py example.docx --append --file blocks.txt
    python doc_cli.py example.docx --insert 2 --file blocks.txt
    python doc_cli.py example.docx --file blocks.txt        # == --rewrite

Или читать блоки из stdin:
    echo "{block}Привет{/block}" | python doc_cli.py out.docx --append --stdin

Сериализация (без изменения файла):
    python doc_cli.py example.docx --serialize              # в консоль
    python doc_cli.py example.docx --serialize out.txt      # в файл

N — номер блока (1-индексация).
"""

import argparse
import json
import subprocess
import sys

from docx import Document
from docx.opc.exceptions import PackageNotFoundError

from docx_blocks import DocxBlockParser, DocxBlockSerializer
from docx_blocks.tags import find_blocks


_FROM_FILE = object()  # маркер: значение брать из --file / --stdin


# ----------------------------------------------------------------------
# Утилиты чтения / записи
# ----------------------------------------------------------------------
def _read_stdin() -> str:
    return sys.stdin.read()


def load_blocks(docx_path: str, allow_missing: bool = False) -> list[str]:
    """Открывает .docx и возвращает список блоков в текстовом формате."""
    try:
        doc = Document(docx_path)
    except (FileNotFoundError, PackageNotFoundError):
        if allow_missing:
            return []
        raise
    serializer = DocxBlockSerializer()
    return find_blocks(serializer.serialize(doc))


def save_blocks(docx_path: str, blocks: list[str]) -> None:
    """Парсит блоки и сохраняет результат в .docx (перезаписывая файл)."""
    parser = DocxBlockParser()
    text = '\n'.join(blocks)
    doc = parser.parse(text)
    doc.save(docx_path)


def read_blocks_file(path: str) -> str:
    with open(path, 'r', encoding='utf-8') as f:
        return f.read()


# ----------------------------------------------------------------------
# Сериализация
# ----------------------------------------------------------------------
def do_serialize(docx_path: str, output,
                 as_json: bool = False,
                 quiet: bool = False) -> int:
    """
    output:
        True      — печатать в stdout;
        'path'    — сохранить в файл.
    as_json — печатать JSON со списком блоков (только с output=True).
    """
    try:
        doc = Document(docx_path)
    except (FileNotFoundError, PackageNotFoundError):
        print(f'Файл не найден: {docx_path}', file=sys.stderr)
        return 2
    except Exception as e:
        print(f'Не удалось открыть файл: {e}', file=sys.stderr)
        return 2

    try:
        serializer = DocxBlockSerializer()
        text = serializer.serialize(doc)
    except Exception as e:
        print(f'Ошибка сериализации: {e}', file=sys.stderr)
        return 1

    if as_json and output is True:
        print(json.dumps({'ok': True, 'blocks': find_blocks(text)},
                         ensure_ascii=False))
        return 0

    if output is True:
        print(text)
        return 0

    try:
        with open(output, 'w', encoding='utf-8') as f:
            f.write(text)
            if not text.endswith('\n'):
                f.write('\n')
    except OSError as e:
        print(f'Не удалось записать в {output}: {e}', file=sys.stderr)
        return 3

    if not quiet:
        print(f'OK: структура сохранена в {output}')
    return 0


# ----------------------------------------------------------------------
# Применение команды к списку блоков
# ----------------------------------------------------------------------
def apply_command(existing: list[str], cmd: tuple) -> list[str]:
    kind = cmd[0]
    total = len(existing)

    if kind == 'rewrite':
        return find_blocks(cmd[1] or '')

    if kind == 'append':
        return existing + find_blocks(cmd[1])

    if kind == 'prepend':
        return find_blocks(cmd[1]) + existing

    if kind == 'insert':
        n, text = cmd[1], cmd[2]
        idx = n - 1
        if not (0 <= idx < total):
            raise IndexError(
                f'Блок №{n} вне диапазона (в файле {total} блок(ов), '
                f'допустимо 1..{total})'
            )
        return existing[:idx] + find_blocks(text) + existing[idx + 1:]

    if kind == 'insert-after':
        n, text = cmd[1], cmd[2]
        if not (1 <= n <= total):
            raise IndexError(
                f'Блок №{n} вне диапазона для insert-after '
                f'(допустимо 1..{total})'
            )
        return existing[:n] + find_blocks(text) + existing[n:]

    if kind == 'insert-before':
        n, text = cmd[1], cmd[2]
        if not (1 <= n <= total + 1):
            raise IndexError(
                f'Блок №{n} вне диапазона для insert-before '
                f'(допустимо 1..{total + 1})'
            )
        idx = n - 1
        return existing[:idx] + find_blocks(text) + existing[idx:]

    raise ValueError(f'Неизвестная команда: {kind}')


# ----------------------------------------------------------------------
# CLI
# ----------------------------------------------------------------------
def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog='doc_cli.py',
        description='Редактирование .docx через блочный формат '
                    '{block ...}...{/block}.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            'Примеры:\n'
            '  python doc_cli.py example.docx "{block}...{/block}"\n'
            '  python doc_cli.py example.docx --append "{block}...{/block}"\n'
            '  python doc_cli.py example.docx --append --file blocks.txt\n'
            '  python doc_cli.py example.docx --insert 2 "{block}...{/block}"\n'
            '  python doc_cli.py example.docx --insert 2 --file blocks.txt\n'
            '  python doc_cli.py example.docx --insert-before 1 --file blocks.txt\n'
            '  echo "{block}...{/block}" | python doc_cli.py out.docx --append --stdin\n'
            '  python doc_cli.py example.docx --serialize\n'
            '  python doc_cli.py example.docx --serialize out.txt\n'
        ),
    )
    p.add_argument('docx', metavar='FILE.docx',
                   help='Путь к .docx файлу')
    p.add_argument('blocks', nargs='?', metavar='BLOCKS',
                   help='Блоки для полной перезаписи (действие по умолчанию)')
    p.add_argument('-f', '--file', dest='blocks_file', metavar='FILE',
                   help='Файл с блоками (альтернатива inline-строке)')

    # Взаимоисключающие команды — только они
    g = p.add_mutually_exclusive_group()
    g.add_argument('--rewrite', nargs='?', const=_FROM_FILE, default=None,
                   metavar='BLOCKS',
                   help='Перезаписать весь файл (по умолчанию)')
    g.add_argument('--append', nargs='?', const=_FROM_FILE, default=None,
                   metavar='BLOCKS',
                   help='Добавить блоки в конец файла')
    g.add_argument('--prepend', nargs='?', const=_FROM_FILE, default=None,
                   metavar='BLOCKS',
                   help='Добавить блоки в начало файла')
    g.add_argument('--insert', nargs='+', metavar='N [BLOCKS]',
                   help='Заменить N-й блок')
    g.add_argument('--insert-after', nargs='+', metavar='N [BLOCKS]',
                   help='Вставить блоки после N-го')
    g.add_argument('--insert-before', nargs='+', metavar='N [BLOCKS]',
                   help='Вставить блоки перед N-м')
    g.add_argument('--serialize', nargs='?', const=True, default=None,
                   metavar='OUT',
                   help='Вывести структуру документа: без OUT — в консоль, '
                        'с OUT — сохранить в указанный файл')

    # Модификаторы — вне группы, комбинируются с любой командой
    p.add_argument('--stdin', action='store_true',
                   help='Читать блоки из stdin (вместо inline/--file)')
    p.add_argument('-q', '--quiet', action='store_true',
                   help='Не печатать итоговые сообщения')
    p.add_argument('--json', action='store_true',
                   help='Печатать результат в JSON')
    return p


def _as_int(value: str) -> int:
    try:
        return int(value)
    except ValueError:
        raise ValueError(f'Ожидалось целое число, получено: {value!r}')


def _pick_blocks_source(inline, blocks_file, from_stdin=False) -> str:
    """Ровно один источник блоков: inline, --file или --stdin."""
    sources = sum([bool(inline), bool(blocks_file), bool(from_stdin)])
    if sources > 1:
        raise ValueError(
            'Укажите ровно один источник блоков: inline, --file или --stdin'
        )
    if from_stdin:
        return _read_stdin()
    if blocks_file:
        try:
            return read_blocks_file(blocks_file)
        except OSError as e:
            raise ValueError(f'Не удалось прочитать {blocks_file}: {e}')
    if inline:
        return inline
    raise ValueError(
        'Блоки не указаны: передайте inline-строку, --file FILE или --stdin'
    )


def _resolve_command(args) -> tuple:
    """Превращает аргументы argparse в кортеж-команду."""
    for name, attr in (('insert', 'insert'),
                       ('insert-after', 'insert_after'),
                       ('insert-before', 'insert_before')):
        val = getattr(args, attr)
        if val is None:
            continue
        n = _as_int(val[0])
        inline = ' '.join(val[1:]) if len(val) > 1 else None
        blocks = _pick_blocks_source(inline, args.blocks_file, args.stdin)
        return (name, n, blocks)

    for name in ('rewrite', 'append', 'prepend'):
        val = getattr(args, name)
        if val is None:
            continue
        if val is _FROM_FILE:
            # без inline-значения: берём из --file или --stdin
            blocks = _pick_blocks_source(None, args.blocks_file, args.stdin)
            return (name, blocks)
        if args.blocks_file or args.stdin:
            raise ValueError(
                f'Нельзя одновременно использовать --{name} "<inline>" '
                f'и --file / --stdin'
            )
        return (name, val)

    # без команды — как --rewrite
    blocks = _pick_blocks_source(args.blocks, args.blocks_file, args.stdin)
    return ('rewrite', blocks or '')


def call_docx_tool(args: dict):
    """
    Хелпер для tool-calling: собирает argv и вызывает doc-cli.
    Пример args:
        {'docx_path': 'out.docx', 'command': 'append',
         'blocks_file': '/tmp/b.txt'}
    """
    argv = ['doc-cli', args['docx_path']]
    cmd = args['command']
    if cmd == 'serialize':
        argv.append('--serialize')
        if 'serialize_output' in args:
            argv.append(args['serialize_output'])
    else:
        argv.append(f'--{cmd}')
        if cmd in ('insert', 'insert-after', 'insert-before'):
            argv.append(str(args['n']))
        if 'blocks_file' in args:
            argv += ['--file', args['blocks_file']]
        elif 'blocks' in args:
            argv.append(args['blocks'])
    return subprocess.run(argv, capture_output=True, text=True)


def main(argv=None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    # --serialize не изменяет файл — обрабатываем отдельно и раньше всего
    if args.serialize is not None:
        return do_serialize(args.docx, args.serialize,
                            as_json=args.json, quiet=args.quiet)

    try:
        command = _resolve_command(args)
    except ValueError as e:
        print(f'Ошибка параметров: {e}', file=sys.stderr)
        return 2

    can_create = command[0] in ('rewrite', 'append', 'prepend')

    try:
        existing = load_blocks(args.docx, allow_missing=can_create)
    except (FileNotFoundError, PackageNotFoundError):
        print(
            f'Файл не найден: {args.docx}. '
            f'Команда --{command[0]} требует существующий файл.',
            file=sys.stderr,
        )
        return 2
    except Exception as e:
        print(f'Не удалось открыть файл: {e}', file=sys.stderr)
        return 2

    if command[0].startswith('insert') and not existing:
        print(
            f'В файле {args.docx} нет блоков — '
            f'команда --{command[0]} не применима.',
            file=sys.stderr,
        )
        return 1

    try:
        new_blocks = apply_command(existing, command)
    except (IndexError, ValueError) as e:
        print(f'Ошибка: {e}', file=sys.stderr)
        return 1

    try:
        save_blocks(args.docx, new_blocks)
    except Exception as e:
        print(f'Не удалось сохранить файл: {e}', file=sys.stderr)
        return 3

    if args.json:
        print(json.dumps({
            'ok': True,
            'file': args.docx,
            'blocks_before': len(existing),
            'blocks_after': len(new_blocks),
        }, ensure_ascii=False))
    elif not args.quiet:
        print(f'OK: {args.docx} — блоков: {len(existing)} -> {len(new_blocks)}')
    return 0


if __name__ == '__main__':
    sys.exit(main())