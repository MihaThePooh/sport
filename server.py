#!/usr/bin/env python3
"""Дневник тренировок: веб-форма для ввода подходов + CSV как источник правды.

Только стандартная библиотека — сервису нечего ставить и нечему ломаться
при обновлении зависимостей. Красивый .xlsx собирает отдельный export.py,
он и только он знает про openpyxl.

Данные лежат в data/log.csv одной строкой на упражнение за день:

    дата,упражнение,подходы
    2026-01-15,подтягивания,10 8 6

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
HEAD = ["дата", "упражнение", "подходы", "всего"]

# openpyxl стоит только под 3.12; если его нет — выгрузка честно откажет,
# а не свалит весь сервис.
EXPORT_PY = shutil.which("python3.12") or "python3"

MAX_SETS = 12        # больше подходов за раз — почти наверняка опечатка
MAX_REPS = 500       # приседаний бывает много, но не столько

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
        return [r for r in csv.DictReader(f) if r.get("дата")]


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
    return [int(x) for x in str(row.get("подходы", "")).split() if x.isdigit()]


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
                           "сумма": total_of(last)} if last else None),
            "занятий": len(mine),
        }
    return out


def add(home, move, sets, when):
    """Добавить занятие. Запись за тот же день и то же движение заменяется:
    правишь опечатку — не плодишь дубль."""
    if move not in MOVES:
        raise ValueError("неизвестное упражнение")
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", when or ""):
        raise ValueError("дата не в формате ГГГГ-ММ-ДД")
    # Раньше тут был int(x), и он молча округлял: 7.5 записывалось как 7.
    # Тихо менять присланное число — худший вид ошибки: в дневнике потом
    # стоит правдоподобная цифра, которой не было. Лучше честный отказ.
    # bool проверяем отдельно: в питоне True — это тоже int, и он
    # превратился бы в один повтор.
    for x in sets:
        if isinstance(x, bool) or not isinstance(x, int):
            raise ValueError("повторения должны быть целыми числами")
    sets = list(sets)
    if not sets or len(sets) > MAX_SETS:
        raise ValueError("подходов должно быть от 1 до %d" % MAX_SETS)
    if any(x < 1 or x > MAX_REPS for x in sets):
        raise ValueError("повторения — от 1 до %d" % MAX_REPS)

    rows = [r for r in read(home)
            if not (r["дата"] == when and r["упражнение"] == move)]
    rows.append({"дата": when, "упражнение": move,
                 "подходы": " ".join(str(x) for x in sets),
                 "всего": sum(sets)})
    rows.sort(key=lambda r: (r["дата"], MOVES.index(r["упражнение"])
                             if r["упражнение"] in MOVES else 99))
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
        if self.path not in ("/api/add", "/api/del"):
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
            else:
                add(home, body.get("упражнение"), body.get("подходы") or [],
                    body.get("дата") or date.today().isoformat())
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
