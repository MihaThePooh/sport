#!/usr/bin/env python3.12
"""Собирает из data/log.csv читаемую книгу Excel.

Запускать python3.12 — openpyxl стоит только там (python3 на этом VPS
десятый и без него).

Таблица широкая: строка — занятие, колонки — упражнения. Так видно
и подходы, и сумму, и чем этот день отличался от соседних.

Зелёная заливка ставится сама: раньше Михаил красил рекордные ячейки
руками, а рекорд — это просто максимум по колонке на текущий момент.
Считаем нарастающим итогом, то есть красим тот день, когда рекорд был
поставлен, а не все, что его позже повторили.
"""
import csv
import os
import sys
import tempfile
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
# Ту же переменную читает server.py: тесты и боевой дневник не пересекаются.
DATA = Path(os.environ.get("SPORT_DATA") or (BASE / "data"))
LOG = DATA / "log.csv"
OUT = DATA / "тренировки.xlsx"

MOVES = ["подтягивания", "отжимания", "приседания"]
# Тот же знак, что и в дневнике: внутри серии подходы через «+»,
# между сериями — «\». Две мио-серии за день видно одним взглядом.
SERIES_SEP = "\\"
NOTE_HEAD = "комментарий"

GREEN = PatternFill("solid", fgColor="C6E7D2")   # рекорд за 12 месяцев
GOLD = PatternFill("solid", fgColor="F3E3B0")    # рекорд за всё время
YEAR = 365
HEADFILL = PatternFill("solid", fgColor="2F6F4E")
STRIPE = PatternFill("solid", fgColor="F7F6F4")  # через строку, чтобы глаз не сползал
THIN = Side(style="thin", color="E3DED7")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def load():
    """CSV → {дата: {упражнение: [подходы]}}, отсортировано по дате."""
    days = {}
    if not LOG.exists():
        return days
    with LOG.open(encoding="utf-8", newline="") as f:
        for r in csv.DictReader(f):
            when = (r.get("дата") or "").strip()
            move = (r.get("упражнение") or "").strip()
            if not when or move not in MOVES:
                continue
            # Подходы разложены по сериям: за день их может быть несколько,
            # и в ячейке они разделяются тем же знаком, что и в дневнике.
            series = []
            for part in str(r.get("подходы", "")).split(SERIES_SEP):
                nums = [int(x) for x in part.split() if x.isdigit()]
                if nums:
                    series.append(nums)
            sets = [x for one in series for x in one]
            total = str(r.get("всего", "")).strip()
            total = int(total) if total.isdigit() else sum(sets)
            # Формат занятия. Колонки может не быть вовсе (старый файл) —
            # тогда это обычная тренировка.
            kind = str(r.get("формат") or "").strip()
            note = str(r.get("комментарий") or "").strip()
            # Занятие без подходов, но с суммой, — тоже занятие: в старой
            # таблице есть дни, где записана только она.
            if sets or total:
                days.setdefault(when, {})[move] = {"подходы": sets, "серии": series,
                                                   "всего": total, "формат": kind,
                                                   "комментарий": note}
    return dict(sorted(days.items()))


def shown(cell):
    """Сколько знаков видно в ячейке. Дата хранится числом, а на экране
    это ДД.ММ.ГГГГ — по str() получилось бы «2026-09-13 00:00:00»
    и колонка вышла бы вдвое шире нужного."""
    v = cell.value
    if v is None:
        return 0
    if isinstance(v, date):
        return len("ДД.ММ.ГГГГ")
    return max(len(line) for line in str(v).split("\n"))


def autofit(ws, last_col, pad=2.0, bold_pad=0.9):
    """Ширина колонки по самому длинному содержимому — то же, что даёт
    двойной клик по границе заголовка в Excel.

    Объединённые ячейки пропускаем: название упражнения растянуто на две
    колонки, и если считать его длину, «всего» раздуется без нужды.
    Excel при автоподборе ведёт себя так же.

    Жирный шрифт шире обычного при той же длине строки, поэтому такой
    ячейке добавляется запас — иначе сумма упирается в рамку."""
    merged = {c for rng in ws.merged_cells.ranges for c in rng.cells}
    for i in range(1, last_col + 1):
        width = 0
        for row in range(1, ws.max_row + 1):
            if (row, i) in merged:
                continue
            cell = ws.cell(row, i)
            need = shown(cell)
            if need and cell.font and cell.font.bold:
                need += bold_pad
            width = max(width, need)
        ws.column_dimensions[get_column_letter(i)].width = width + pad


def build():
    days = load()
    wb = Workbook()
    ws = wb.active
    ws.title = "Тренировки"

    # Шапка в два яруса: над каждым упражнением — «подходы» и «всего».
    ws.cell(1, 1, "дата")
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
    col = 2
    where = {}
    for m in MOVES:
        ws.cell(1, col, m)
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 1)
        ws.cell(2, col, "подходы")
        ws.cell(2, col + 1, "всего")
        where[m] = col
        col += 2

    # Комментарий — один на день, поэтому шапка у него в два яруса, как
    # у даты: делить его по упражнениям незачем.
    note_col = col
    ws.cell(1, note_col, NOTE_HEAD)
    ws.merge_cells(start_row=1, start_column=note_col, end_row=2, end_column=note_col)
    col += 1

    for row in (1, 2):
        for c in range(1, col):
            cell = ws.cell(row, c)
            cell.fill = HEADFILL
            cell.font = Font(color="FFFFFF", bold=True, size=11)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = BOX

    # Рекорды копим по ходу: красится день прорыва, а не его повторы.
    #
    # Два вида отметок, и это не украшательство. Абсолютные рекорды
    # Михаил поставил осенью 2022-го, когда делал по 7 подходов за раз.
    # Сейчас подходов 3 — сумма за занятие физически не догонит ту,
    # и зелёного в таблице не было бы с 2022 года. Поэтому основная,
    # живая отметка — рекорд за последние 12 месяцев на момент занятия:
    # он честно говорит «лучше, чем ты был весь этот год». Абсолютный
    # пик отмечается отдельно, золотым, как памятка.
    peak = {m: {"подход": 0, "всего": 0} for m in MOVES}
    r = 3

    def year_best(m, upto):
        """Лучшее за 12 месяцев ДО этого дня (сам день не считаем)."""
        edge = upto - timedelta(days=YEAR)
        rep = tot = 0
        for d, mv in days.items():
            day = date.fromisoformat(d)
            if not (edge <= day < upto) or m not in mv:
                continue
            if mv[m]["подходы"]:
                rep = max(rep, max(mv[m]["подходы"]))
            tot = max(tot, mv[m]["всего"])
        return rep, tot

    for when, moves in days.items():
        d = date.fromisoformat(when)
        cell = ws.cell(r, 1, d)
        cell.number_format = "DD.MM.YYYY"
        cell.alignment = Alignment(horizontal="center")

        for m in MOVES:
            c = where[m]
            got = moves.get(m)
            if not got:
                for k in (c, c + 1):
                    ws.cell(r, k, "").alignment = Alignment(horizontal="center")
                continue
            sets, total = got["подходы"], got["всего"]
            # Мио-серия читается иначе обычных подходов: первое число —
            # активационный подход, дальше короткие круги. Помечаем прямо
            # в ячейке, иначе через полгода «10 + 4 + 4 + 3» выглядит как
            # неудачная тренировка, хотя это другой метод.
            # Внутри серии подходы через «+», серии между собой — через
            # «\». Строка «9 + 3 + 3 \ 5 + 2 + 2» читается сразу: два
            # захода за день, а не один длинный.
            series = got.get("серии") or ([sets] if sets else [])
            запись = (" %s " % SERIES_SEP).join(
                " + ".join(str(x) for x in one) for one in series) or "—"
            if sets and got.get("формат") == "мио":
                запись = "мио " + запись
            a = ws.cell(r, c, запись)
            b = ws.cell(r, c + 1, total)
            a.alignment = Alignment(horizontal="center")
            b.alignment = Alignment(horizontal="center")
            b.font = Font(bold=True)

            yrep, ytot = year_best(m, d)
            if sets and max(sets) > peak[m]["подход"]:
                a.fill = GOLD                      # рекорд за всё время
                peak[m]["подход"] = max(sets)
            elif sets and yrep and max(sets) > yrep:
                a.fill = GREEN                     # лучше, чем весь прошедший год
            if total > peak[m]["всего"]:
                b.fill = GOLD
                peak[m]["всего"] = total
            elif ytot and total > ytot:
                b.fill = GREEN

        # Комментарий пишется к упражнению, а колонка в книге одна на
        # день: собираем непустые по порядку движений.
        notes = [moves[m]["комментарий"] for m in MOVES
                 if m in moves and moves[m].get("комментарий")]
        if notes:
            note = ws.cell(r, note_col, "; ".join(notes))
            note.alignment = Alignment(horizontal="left", vertical="center")

        if r % 2:
            for c in range(1, col):
                if not ws.cell(r, c).fill.fgColor.rgb or \
                        ws.cell(r, c).fill.fill_type is None:
                    ws.cell(r, c).fill = STRIPE
        for c in range(1, col):
            ws.cell(r, c).border = BOX
        r += 1

    # Итоговой строки в книге нет намеренно: дневник заканчивается
    # последним занятием. Суммы за всё время ничего не говорят о том,
    # как идут дела сейчас, а лишняя строка мешает — при прокрутке
    # вниз глаз ждёт свежую дату, а натыкается на «итого».

    autofit(ws, col - 1)
    ws.freeze_panes = "B3"          # шапка и дата не уезжают при прокрутке

    OUT.parent.mkdir(parents=True, exist_ok=True)
    # Сохраняем через временный файл и подменяем одним движением.
    # Книга пересобирается после каждой записи, и без этого можно было
    # бы скачать её ровно в тот момент, когда openpyxl дописывает архив,
    # — получился бы битый .xlsx без единой ошибки на экране.
    fd, tmp = tempfile.mkstemp(dir=str(OUT.parent), suffix=".xlsx")
    os.close(fd)
    try:
        wb.save(tmp)
        # mkstemp отдаёт 0600 — своим же файлом потом не поделишься.
        os.chmod(tmp, 0o644)
        os.replace(tmp, OUT)
    except BaseException:
        if os.path.exists(tmp):
            os.unlink(tmp)
        raise
    return len(days)


if __name__ == "__main__":
    n = build()
    print("собрано занятий: %d → %s" % (n, OUT))
    if not n:
        sys.exit(1)
