#!/usr/bin/env python3
"""Дневник тренировок: веб-форма для ввода подходов + CSV как источник правды.

Только стандартная библиотека — сервису нечего ставить и нечему ломаться
при обновлении зависимостей. Красивый .xlsx собирает отдельный export.py,
он и только он знает про openpyxl.

Данные лежат в data/log.csv одной строкой на упражнение за день:

    дата,упражнение,подходы,всего,формат
    2026-01-15,подтягивания,10 8 6,24,
    2026-09-10,подтягивания,10 4 4 3,21,мио

Почему CSV, а не сразу .xlsx: сервер, который переписывает книгу Excel
на каждое нажатие кнопки, рано или поздно её испортит — формулы, стили,
объединённые ячейки. CSV испортить нечем. Красота живёт в export.py и
пересобирается из CSV в любой момент.
"""
import csv
import hashlib
import json
import os
import re
import shutil
import subprocess
import tempfile
import threading
from datetime import date, timedelta
from urllib.parse import quote
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from pathlib import Path

BASE = Path(__file__).resolve().parent
# Папку данных и порт можно увести в сторону: тесты пишут к себе, а не
# в живой дневник. По умолчанию — рядом с кодом.
DATA = Path(os.environ.get("SPORT_DATA") or (BASE / "data"))
PAGE = BASE / "index.html"

# Ключ доступа. Своего у дневника раньше не было вовсе: он жил на
# localhost, и снаружи до него было не дотянуться. Как только перед ним
# встал nginx, любой прохожий мог и читать записи, и дописывать свои.
#
# Устройство то же, что в тудушке: ключ — 256 бит случайности, никакого
# имени и пароля. Проверять его не с чем и не надо — любой ключ нужной
# формы открывает СВОЙ дневник. Чужой не подберёшь: это 2^256 вариантов.
# Ключ лежит в папке данных, а не рядом с кодом: data/ и так вне git,
# а тесты, которым папку подменяют, получают свой ключ и не трогают
# боевой.
TOKEN_FILE = DATA / "token"
KEY_RE = re.compile(r"^[0-9a-fA-F]{64}$")
LISTS = DATA / "lists"

HOST = "127.0.0.1"
PORT = int(os.environ.get("SPORT_PORT") or 8790)

# Порядок важен: в таком виде упражнения показываются в форме и в таблице.
MOVES = ["подтягивания", "отжимания", "приседания"]

# Колонка «всего» нужна не для красоты: в старой таблице есть занятия,
# где записана только сумма, а подходы утеряны. Без неё такой день
# пропал бы из истории совсем.
#
# Колонка «формат» появилась позже и потому пустая у всех старых
# записей — так и задумано: пусто читается как «обычная тренировка».
# Добавлять её значением по умолчанию в старые строки не нужно, файл
# и так разбирается по именам колонок, а не по их числу.
HEAD = ["дата", "упражнение", "подходы", "всего", "формат", "комментарий"]

# Разделитель серий внутри колонки «подходы»:
#
#   9 3 3 3 3 3 \ 5 2 2 2 2 2
#
# Две мио-серии подтягиваний за день — это одно занятие, а не два, и
# ключ записи остаётся прежним: дата плюс движение. Поэтому серии
# живут в одной ячейке, а не в отдельных строках: так не пришлось
# трогать ни ключ, ни старые записи. Файл без этого знака читается
# ровно как раньше — одна серия.
SERIES_SEP = "\\"
MAX_SERIES = 6       # больше серий за занятие — уже не тренировка, а опечатка

# Форматы занятия.
#
#   обычная  подходы до усталости, отдых сколько нужно
#   мио      мио-серия: первый подход почти до отказа («активационный»),
#            дальше короткие круги по 3-5 повторений через 20-30 секунд
#
# Хранятся оба одинаково — списком чисел. Разница в СМЫСЛЕ первого
# числа и в длине пауз, а не в структуре, поэтому отдельной таблицы
# мио-серии не заводят: это тот же дневник, просто помеченный.
NORMAL = ""
MYO = "мио"
KINDS = (NORMAL, MYO)
# Чем мио-серия подписывает себя в комментарии. Отдельно от MYO: в
# колонке «формат» лежит короткий признак для машины, а это — текст,
# который Михаил читает в книге через полгода.
MYO_NOTE = "мио-серии"
MIN_MYO_SETS = 2     # активационный подход плюс хотя бы один круг

# openpyxl стоит только под 3.12; если его нет — выгрузка честно откажет,
# а не свалит весь сервис.
EXPORT_PY = shutil.which("python3.12") or "python3"

MAX_SETS = 12        # больше подходов в серии — почти наверняка опечатка
MAX_REPS = 500       # приседаний бывает много, но не столько
MAX_NOTE = 200       # комментарий — пометка на полях, а не дневниковая запись

def token():
    """Ключ хозяина. Нет файла — заводим при первом обращении."""
    if not TOKEN_FILE.exists():
        DATA.mkdir(parents=True, exist_ok=True)
        TOKEN_FILE.write_text(os.urandom(32).hex(), encoding="utf-8")
        TOKEN_FILE.chmod(0o600)
    return TOKEN_FILE.read_text(encoding="utf-8").strip()


def home_for(key):
    """Папка дневника по ключу.

    Ключ хозяина ведёт в саму data/ — там записи лежали до того, как
    появились ключи, и трогать их с места незачем. Остальные получают
    подпапку с именем из хеша: сам ключ в именах файлов не светится,
    поэтому список папок ничего не выдаёт."""
    if key.lower() == token().lower():
        return DATA
    return LISTS / hashlib.sha256(key.encode()).hexdigest()


def log_of(home):
    return home / "log.csv"


def book_of(home):
    return home / "тренировки.xlsx"

# Статика отдаётся по белому списку, а не по пути из запроса. Так к
# обходу каталогов (../../etc/passwd) просто нечего приложить.
STATIC = {
    "/manifest.webmanifest": ("manifest.webmanifest",
                              "application/manifest+json; charset=utf-8"),
    "/sw.js": ("sw.js", "application/javascript; charset=utf-8"),
    "/icons/icon-192.png": ("icons/icon-192.png", "image/png"),
    "/icons/icon-512.png": ("icons/icon-512.png", "image/png"),
    "/icons/icon-mask.png": ("icons/icon-mask.png", "image/png"),
}

# Пересборка книги идёт в фоне, и одновременно её запускать нельзя:
# два процесса писали бы в один файл. Замок пропускает одного, а
# «пришло ещё» запоминается флагом — после текущей сборки будет ровно
# одна догоняющая, а не очередь из десяти.
#
# Замок теперь на каждый дневник свой: чужая сборка не должна заставлять
# ждать твою. Словарь замков растёт по числу заходивших ключей, и его
# самого тоже надо от гонки прикрыть — этим занят _locks_lock.
_locks = {}
_locks_lock = threading.Lock()


def lock_for(home):
    """Замок и флаг «пришло ещё» для одного дневника."""
    with _locks_lock:
        pair = _locks.get(home)
        if pair is None:
            pair = (threading.Lock(), threading.Event())
            _locks[home] = pair
        return pair


def read(home):
    """Весь дневник списком словарей. Нет файла — пустой дневник."""
    log = log_of(home)
    if not log.exists():
        return []
    with log.open(encoding="utf-8", newline="") as f:
        rows = [r for r in csv.DictReader(f) if r.get("дата")]
    # Старый файл колонки «формат» не имеет вовсе, а csv кладёт в такие
    # места None. Дальше по коду это значение сравнивают со строками
    # и пишут обратно в файл — пусть оно с самого начала будет строкой.
    for r in rows:
        r["формат"] = kind_of(r)
        r["комментарий"] = note_of(r)
    return rows


def kind_of(row):
    """Формат занятия строкой. Всё, что не «мио», — обычная тренировка."""
    return MYO if str(row.get("формат") or "").strip() == MYO else NORMAL


def write(home, rows):
    """Пишем через временный файл: обрыв на середине не оставит огрызок."""
    home.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=str(home), suffix=".csv")
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="") as f:
            w = csv.DictWriter(f, fieldnames=HEAD)
            w.writeheader()
            w.writerows(rows)
        os.replace(tmp, log_of(home))  # атомарная замена, без промежуточного состояния
    except BaseException:
        if os.path.exists(tmp):
            os.unlink(tmp)
        raise


def sets_of(row):
    """Все подходы занятия одним списком, границы серий стёрты. В таком
    виде считаются сумма и рекорды: лучший подход остаётся лучшим
    независимо от того, в какой серии он сделан."""
    return [int(x) for x in str(row.get("подходы", "")).split() if x.isdigit()]


def series_of(row):
    """Подходы, разложенные по сериям: [[9,3,3,3],[5,2,2]].

    Занятие без разделителя — это одна серия, поэтому у старых записей
    список всегда из одного элемента, и вызывающему коду не нужно знать,
    новая перед ним запись или четырёхлетней давности."""
    out = []
    for part in str(row.get("подходы", "")).split(SERIES_SEP):
        nums = [int(x) for x in part.split() if x.isdigit()]
        if nums:
            out.append(nums)
    return out


def note_of(row):
    """Комментарий к занятию. У старых записей колонки нет вовсе."""
    return str(row.get("комментарий") or "").strip()


def total_of(row):
    """Сумма за занятие. Подходы известны — считаем; нет — берём колонку."""
    s = sets_of(row)
    if s:
        return sum(s)
    v = str(row.get("всего", "")).strip()
    return int(v) if v.isdigit() else 0


YEAR = 365


def best(rows, move, since=None):
    """Лучший подход и лучшая сумма за занятие. since — считать только
    занятия не старше этой даты."""
    rep, total = 0, 0
    for r in rows:
        if r.get("упражнение") != move:
            continue
        if since and r.get("дата", "") < since:
            continue
        s = sets_of(r)
        if s:
            rep = max(rep, max(s))
        total = max(total, total_of(r))
    return {"подход": rep, "занятие": total}


def year_ago():
    return (date.today() - timedelta(days=YEAR)).isoformat()


def state(home):
    """То, что нужно форме: рекорды и последние занятия по каждому движению."""
    rows = read(home)
    out = {"сегодня": date.today().isoformat(), "движения": {}}
    for m in MOVES:
        mine = [r for r in rows if r.get("упражнение") == m]
        mine.sort(key=lambda r: r["дата"])
        last = mine[-1] if mine else None
        out["движения"][m] = {
            # Два рекорда, а не один. Абсолютный поставлен в 2022-м, когда
            # занятие было по 7 подходов; сейчас их 3, и он не побьётся
            # никогда — как мерка мотивации он мёртв. Живая мерка —
            # рекорд за последние 12 месяцев: он обновляется несколько раз
            # в год и говорит про нынешнюю форму, а не про позапрошлую.
            "рекорд": best(rows, m),
            "загод": best(rows, m, since=year_ago()),
            "последнее": ({"дата": last["дата"], "подходы": sets_of(last),
                           "серии": series_of(last),
                           "сумма": total_of(last),
                           "формат": kind_of(last),
                           "комментарий": note_of(last)} if last else None),
            "занятий": len(mine),
        }
    out["дни"] = recent(rows)
    return out


TABLE_DAYS = 15


def recent(rows, n=TABLE_DAYS):
    """Последние дни тренировок в том же виде, что и в книге: строка —
    день, в ней по каждому движению подходы и сумма, плюс общий
    комментарий.

    Форма показывает эту таблицу внизу страницы, и она же служит списком
    для правки — поэтому отдаётся вместе с состоянием, одним запросом.
    Иначе после каждой записи пришлось бы ходить на сервер дважды."""
    days = {}
    for r in rows:
        when, move = r.get("дата"), r.get("упражнение")
        if not when or move not in MOVES:
            continue
        days.setdefault(when, {})[move] = {
            "серии": series_of(r), "сумма": total_of(r),
            "формат": kind_of(r), "комментарий": note_of(r)}
    out = []
    for when in sorted(days)[-n:]:
        # Комментарии дня собираются в одну строку с именем движения
        # впереди — ровно так, как они лежат в колонке книги.
        notes = [f"{m}: {days[when][m]['комментарий']}" for m in MOVES
                 if m in days[when] and days[when][m]["комментарий"]]
        out.append({"дата": when, "движения": days[when],
                    "комментарий": (" %s " % SERIES_SEP).join(notes)})
    return out


def as_series(sets):
    """Привести присланные подходы к списку серий.

    Форма умеет слать и то и другое: плоский список [10, 4, 4] — это
    одна серия, вложенный [[9, 3, 3], [5, 2, 2]] — две. Старый вызов с
    плоским списком должен работать как раньше, поэтому разбираем оба
    вида здесь, а дальше по коду серия всегда одна и та же структура."""
    if sets and all(isinstance(x, list) for x in sets):
        return [list(s) for s in sets]
    return [list(sets)]


def add(home, move, sets, when, kind=NORMAL, note="", append=False):
    """Добавить занятие. Запись за тот же день и то же движение заменяется:
    правишь опечатку — не плодишь дубль.

    sets — либо подходы одной серии, либо список серий. Две мио-серии
    подтягиваний за день — это одно занятие с двумя сериями, а не две
    записи: иначе вторая затирала бы первую по ключу.

    append — дописать серию к тому, что за этот день уже есть. Так
    работают обе кнопки записи в форме: человек делает серию, жмёт
    «записать», поля чистятся, и следующая серия ложится рядом, а не
    вместо. Без этого флага запись дня заменяется целиком — это режим
    правки, когда нужно исправить неверно введённое.

    kind — формат занятия: обычная тренировка или мио-серия. На хранение
    он не влияет (числа те же), но влияет на то, как запись читают: в
    мио-серии первое число — активационный подход, остальные — круги."""
    if move not in MOVES:
        raise ValueError("неизвестное упражнение")
    kind = str(kind or NORMAL).strip()
    if kind not in KINDS:
        raise ValueError("неизвестный формат занятия")
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", when or ""):
        raise ValueError("дата не в формате ГГГГ-ММ-ДД")

    series = as_series(sets)
    if not series or len(series) > MAX_SERIES:
        raise ValueError("серий должно быть от 1 до %d" % MAX_SERIES)
    for one in series:
        # Раньше тут был int(x), и он молча округлял: 7.5 записывалось как 7.
        # Тихо менять присланное число — худший вид ошибки: в дневнике потом
        # стоит правдоподобная цифра, которой не было. Лучше честный отказ.
        # bool проверяем отдельно: в питоне True — это тоже int, и он
        # превратился бы в один повтор.
        for x in one:
            if isinstance(x, bool) or not isinstance(x, int):
                raise ValueError("повторения должны быть целыми числами")
        if not one or len(one) > MAX_SETS:
            raise ValueError("подходов в серии должно быть от 1 до %d" % MAX_SETS)
        if any(x < 1 or x > MAX_REPS for x in one):
            raise ValueError("повторения — от 1 до %d" % MAX_REPS)
        # Мио-серия из одного числа — это обычный подход, назвать его серией
        # нельзя: весь смысл метода в кругах после активационного подхода.
        if kind == MYO and len(one) < MIN_MYO_SETS:
            raise ValueError("в мио-серии нужен активационный подход и хотя бы один круг")

    # Перевод строки в комментарии сломал бы CSV на чтении, а обратный
    # слеш — разбор серий и склейку комментариев дня в книге.
    note = re.sub(r"[\r\n\\]+", " ", str(note or "")).strip()[:MAX_NOTE]

    rows = read(home)
    было = next((r for r in rows
                 if r["дата"] == when and r["упражнение"] == move), None)
    if append and было:
        # Серии дня складываются: сделал ещё один заход — он встаёт
        # рядом с прежними, а не вместо них.
        series = series_of(было) + series
        if len(series) > MAX_SERIES:
            raise ValueError("серий за день должно быть не больше %d" % MAX_SERIES)
        # Комментарий пустым не перетираем: кнопка записи его не
        # передаёт, и молчание не должно стирать уже написанное.
        note = note or note_of(было)
        # Пометка «мио» единожды поставлена — она про весь день.
        kind = MYO if MYO in (kind, kind_of(было)) else kind

    # Мио-серия помечает себя сама, одним словом в комментарии: отдельной
    # колонки для этого в книге нет, а через полгода «9 + 3 + 3» без
    # пометки не отличить от неудачной обычной тренировки. Пишем только
    # в пустой комментарий — свой текст затирать нельзя.
    if kind == MYO and not note:
        note = MYO_NOTE

    rows = [r for r in rows
            if not (r["дата"] == when and r["упражнение"] == move)]
    rows.append({"дата": when, "упражнение": move,
                 "подходы": (" %s " % SERIES_SEP).join(
                     " ".join(str(x) for x in one) for one in series),
                 "всего": sum(sum(one) for one in series),
                 "формат": kind, "комментарий": note})
    rows.sort(key=lambda r: (r["дата"], MOVES.index(r["упражнение"])
                             if r["упражнение"] in MOVES else 99))
    write(home, rows)


def note_set(home, move, when, note):
    """Переписать комментарий занятия, не трогая подходы.

    Отдельная операция, потому что кнопка «комментарий» открывает поле
    с тем, что уже написано, и сохраняет отредактированный текст целиком.
    Гонять через add() пришлось бы вместе с подходами — лишний повод
    их испортить."""
    if move not in MOVES:
        raise ValueError("неизвестное упражнение")
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", when or ""):
        raise ValueError("дата не в формате ГГГГ-ММ-ДД")
    rows = read(home)
    цель = next((r for r in rows
                 if r["дата"] == when and r["упражнение"] == move), None)
    if цель is None:
        raise LookupError("нет такого занятия")
    цель["комментарий"] = re.sub(r"[\r\n\\]+", " ",
                                 str(note or "")).strip()[:MAX_NOTE]
    write(home, rows)


def drop(home, move, when):
    """Убрать занятие. Ключ тот же, что у add: дата плюс движение.

    Если такой записи нет — это не ошибка ввода, а расхождение с тем,
    что человек видел на экране: список у него мог устареть. Возвращаем
    честный отказ, чтобы интерфейс перечитал журнал, а не молчал."""
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", when or ""):
        raise ValueError("дата не в формате ГГГГ-ММ-ДД")
    rows = read(home)
    left = [r for r in rows
            if not (r["дата"] == when and r["упражнение"] == move)]
    if len(left) == len(rows):
        raise LookupError("такой записи уже нет")
    write(home, left)


def build_book(home):
    """Пересобрать .xlsx из CSV. Возвращает (получилось, пояснение).

    Экспорт живёт отдельным скриптом под python3.12: openpyxl стоит
    только там. Книга собирается ЦЕЛИКОМ заново, а не правится —
    поэтому испортить её нечем.

    Папку экспорт получает переменной окружения — тем же способом, каким
    её задают тесты, так что чужой дневник собирается ровно так же."""
    try:
        done = subprocess.run(
            [EXPORT_PY, str(BASE / "export.py")],
            capture_output=True, text=True, timeout=120,
            env={**os.environ, "SPORT_DATA": str(home)})
    except Exception as e:                           # noqa: BLE001
        return False, "не смог запустить сборку: %s" % e
    if done.returncode != 0 or not book_of(home).exists():
        hint = (done.stderr or done.stdout or "").strip().splitlines()
        return False, "сборка не удалась: %s" % (hint[-1] if hint else "без подробностей")
    return True, ""


def rebuild_soon(home):
    """Пересобрать книгу после записи, не заставляя браузер ждать.

    Смысл в том, чтобы тренировки.xlsx всегда лежала свежей: зашёл
    на сервер, забрал файл — и не гадаешь, всё ли туда попало."""
    lock, again = lock_for(home)
    if not lock.acquire(blocking=False):
        again.set()                  # уже собираем — попросим повторить
        return
    try:
        while True:
            again.clear()
            ok, why = build_book(home)
            if not ok:
                print("книга не пересобралась:", why, flush=True)
            if not again.is_set():
                break
    finally:
        lock.release()


class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def log_message(self, *a):
        pass                      # своя тишина вместо шума в journalctl

    def send(self, code, body, ctype="application/json; charset=utf-8"):
        if isinstance(body, (dict, list)):
            body = json.dumps(body, ensure_ascii=False)
        blob = body.encode("utf-8") if isinstance(body, str) else body
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(blob)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(blob)

    def key(self):
        """Ключ из заголовка или None.

        Правильность проверять не с чем: любой ключ нужной формы открывает
        свой собственный дневник. Форму проверяем строго — иначе строка
        вроде «../..» доехала бы до имени папки."""
        got = self.headers.get("Authorization", "")
        if not got.startswith("Bearer "):
            return None
        k = got[7:].strip()
        return k if KEY_RE.match(k) else None

    def home(self):
        """Папка дневника для этого запроса. Нет ключа — вернём 401 и None."""
        k = self.key()
        if not k:
            self.deny(401, "нужен ключ доступа")
            return None
        return home_for(k)

    def deny(self, code, why):
        """Отказать — и обязательно дочитать тело запроса.

        Соединение живёт дальше: HTTP/1.1 держит его открытым под
        следующий запрос. Непрочитанные байты тела сервер примет за
        начало этого следующего запроса, тот развалится, и клиент
        получит 501 на ровном месте. Ошибка коварная: одиночный
        отказ выглядит правильным, ломается только следующий за ним.

        Слишком длинное тело не вычитываем: заявленную длину задаёт
        клиент, и читать по его команде гигабайт — плохая идея. Такому
        просто закрываем соединение."""
        try:
            n = int(self.headers.get("Content-Length") or 0)
        except ValueError:
            n = 0
        if n > 65536:
            self.close_connection = True
        else:
            while n > 0:
                part = self.rfile.read(min(n, 8192))
                if not part:
                    break
                n -= len(part)
        self.send(code, {"ошибка": why})

    def do_GET(self):
        if self.path.split("?")[0] in ("/", "/index.html"):
            return self.send(200, PAGE.read_bytes(), "text/html; charset=utf-8")
        # Всё под /api/ — только с ключом. Страница и иконки открыты:
        # без них некому будет ключ предъявить.
        if self.path.startswith("/api/"):
            home = self.home()
            if home is None:
                return
            if self.path == "/api/state":
                return self.send(200, state(home))
            if self.path == "/api/log":
                return self.send(200, read(home))
            if self.path == "/api/xlsx":
                return self.xlsx(home)
            return self.send(404, {"ошибка": "нет такого метода"})
        got = STATIC.get(self.path.split("?")[0])
        if got:
            f = BASE / got[0]
            if not f.exists():
                return self.send(404, {"ошибка": "файл не собран: %s" % got[0]})
            return self.send(200, f.read_bytes(), got[1])
        self.send(404, {"ошибка": "нет такой страницы"})

    def xlsx(self, home):
        """Пересобирает книгу и отдаёт её файлом.

        Пересобираем перед отдачей, хотя она и так обновляется после
        каждой записи: замок держим на время сборки, чтобы не читать
        файл, пока фоновая пересборка его пишет."""
        lock, _ = lock_for(home)
        with lock:
            ok, why = build_book(home)
            if not ok:
                return self.send(500, {"ошибка": why})
            blob = book_of(home).read_bytes()
        self.send_response(200)
        self.send_header(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        # Кириллица в имени файла живёт только в filename*; обычный
        # filename оставляем латиницей как запасной.
        self.send_header("Content-Disposition",
                         "attachment; filename=\"trenirovki.xlsx\"; "
                         "filename*=UTF-8''%s" % quote("тренировки.xlsx"))
        self.send_header("Content-Length", str(len(blob)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(blob)

    def do_POST(self):
        if self.path not in ("/api/add", "/api/del", "/api/note"):
            return self.deny(404, "нет такого метода")
        home = self.home()
        if home is None:
            return
        try:
            n = int(self.headers.get("Content-Length") or 0)
            if n > 4096:
                self.close_connection = True
                raise ValueError("слишком длинный запрос")
            body = json.loads(self.rfile.read(n) or b"{}")
            if self.path == "/api/del":
                drop(home, body.get("упражнение"),
                     body.get("дата") or "")
            elif self.path == "/api/note":
                note_set(home, body.get("упражнение"),
                         body.get("дата") or date.today().isoformat(),
                         body.get("комментарий") or "")
            else:
                add(home, body.get("упражнение"), body.get("подходы") or [],
                    body.get("дата") or date.today().isoformat(),
                    body.get("формат") or NORMAL,
                    body.get("комментарий") or "",
                    bool(body.get("дописать")))
        except LookupError as e:
            return self.send(404, {"ошибка": str(e)})
        except ValueError as e:
            return self.send(400, {"ошибка": str(e)})
        except Exception as e:                       # noqa: BLE001 — наружу не пускаем
            return self.send(500, {"ошибка": "не смог записать: %s" % e})
        # Ответ уходит сразу, книга дособирается следом. Ждать сборки
        # тут нельзя: openpyxl на большой таблице думает секунду-другую,
        # и кнопка «Записать» подвисала бы на ровном месте.
        threading.Thread(target=rebuild_soon, args=(home,), daemon=True).start()
        self.send(200, state(home))


if __name__ == "__main__":
    DATA.mkdir(parents=True, exist_ok=True)
    token()                   # заводим ключ хозяина, если его ещё нет
    print("дневник тренировок: http://%s:%d" % (HOST, PORT))
    ThreadingHTTPServer((HOST, PORT), Handler).serve_forever()
