#!/usr/bin/env python3
"""Рисует иконку дневника тренировок в PNG. Только стандартная библиотека.

Pillow, cairo и pip на этой машине отсутствуют, поэтому всё вручную:
фигура описывается функцией «эта точка внутри?», считается с
supersampling'ом ради сглаживания и пакуется в PNG через zlib.

Сделано по образцу ~/tools/todo/icon.py — приложения соседние, иконки
должны быть из одного набора, а не «кто во что горазд».

    ./icon.py            все размеры в icons/
    ./icon.py лист       один лист с вариантами, чтобы выбрать
"""
import pathlib
import struct
import sys
import zlib

BASE = pathlib.Path(__file__).resolve().parent
OUT = BASE / "icons"
SS = 3          # каждый пиксель усредняется из SS x SS проб

PAPER = (0xF7, 0xF6, 0xF4)
GREEN = (0x2F, 0x6F, 0x4E)
WHITE = (0xFF, 0xFF, 0xFF)


# --- геометрия ------------------------------------------------------------
# Каждая функция принимает координаты в квадрате 0..1 и отвечает «внутри?».

def squircle(x, y, n=4.0):
    """Суперэллипс |x|^n + |y|^n <= 1 — форма, под которую режут Android и iOS."""
    dx, dy = abs(x - .5) * 2, abs(y - .5) * 2
    return dx ** n + dy ** n <= 1.0


def seg(x, y, x1, y1, x2, y2, w):
    """Толстая линия с круглыми концами: расстояние от точки до отрезка."""
    vx, vy = x2 - x1, y2 - y1
    px, py = x - x1, y - y1
    L = vx * vx + vy * vy
    t = 0.0 if L == 0 else max(0.0, min(1.0, (px * vx + py * vy) / L))
    dx, dy = px - vx * t, py - vy * t
    return dx * dx + dy * dy <= (w / 2) ** 2


def plate(x, y, cx, h, w):
    """Блин гантели: вертикальный прямоугольник со скруглением."""
    return seg(x, y, cx, .5 - h / 2 + w / 2, cx, .5 + h / 2 - w / 2, w)


def dumbbell(x, y, s=1.0):
    """Гантель: гриф и по два блина с каждой стороны.

    Внешние блины короче внутренних — так силуэт читается гантелью
    даже в 48 px, когда детали сливаются."""
    c = .5
    if seg(x, y, c - .30 * s, c, c + .30 * s, c, .085 * s):      # гриф
        return True
    for cx, h in ((c - .34 * s, .30 * s), (c + .34 * s, .30 * s),
                  (c - .22 * s, .46 * s), (c + .22 * s, .46 * s)):
        if plate(x, y, cx, h, .105 * s):
            return True
    return False


# --- варианты -------------------------------------------------------------
# Возвращают цвет точки либо None — «прозрачно».

def v1(x, y):
    """Зелёная плитка, белая гантель."""
    if not squircle(x, y):
        return None
    return WHITE if dumbbell(x, y) else GREEN


def v2(x, y):
    """Светлая плитка, зелёная гантель — под светлую тему системы."""
    if not squircle(x, y):
        return None
    return GREEN if dumbbell(x, y) else PAPER


def v3(x, y):
    """Гантель без плитки — для maskable, где систему интересует только рисунок."""
    return WHITE if dumbbell(x, y, .86) else None


VARIANTS = {"1": v1, "2": v2, "3": v3}


# --- отрисовка ------------------------------------------------------------

def render(fn, size, pad=0.0):
    """RGBA-байты построчно. pad ужимает рисунок внутри холста."""
    px = bytearray()
    step = 1.0 / (size * SS)
    scale = 1.0 - 2 * pad
    for row in range(size):
        line = bytearray()
        for col in range(size):
            r = g = b = a = 0
            for sy in range(SS):
                for sx in range(SS):
                    x = (col * SS + sx + .5) * step
                    y = (row * SS + sy + .5) * step
                    c = fn((x - pad) / scale, (y - pad) / scale) if scale else None
                    if c:
                        r += c[0]; g += c[1]; b += c[2]; a += 255
            n = SS * SS
            if a:
                # Цвет усредняем только по закрытым пробам, иначе края
                # уходят в чёрный вместо того, чтобы держать свой тон.
                k = a // 255
                line += bytes((r // k, g // k, b // k, a // n))
            else:
                line += b"\0\0\0\0"
        px += line
    return bytes(px)


def png(path, pixels, size):
    """Минимальный писатель PNG: подпись, IHDR, IDAT, IEND."""
    raw = bytearray()
    stride = size * 4
    for row in range(size):
        raw += b"\0"                     # тип фильтра 0 для каждой строки
        raw += pixels[row * stride:(row + 1) * stride]

    def chunk(tag, data):
        return (struct.pack(">I", len(data)) + tag + data
                + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))

    head = struct.pack(">IIBBBBB", size, size, 8, 6, 0, 0, 0)
    path.write_bytes(b"\x89PNG\r\n\x1a\n"
                     + chunk(b"IHDR", head)
                     + chunk(b"IDAT", zlib.compress(bytes(raw), 9))
                     + chunk(b"IEND", b""))


def build():
    OUT.mkdir(parents=True, exist_ok=True)
    made = []
    for name, fn, size, pad in (
        ("icon-192.png", v1, 192, 0.0),
        ("icon-512.png", v1, 512, 0.0),
        # Маскируемая: система обрежет её под свою форму и может съесть
        # до 20% с каждого края, поэтому рисунок ужат внутрь холста.
        ("icon-mask.png", v1, 512, 0.14),
    ):
        p = OUT / name
        png(p, render(fn, size, pad), size)
        made.append("%s (%d байт)" % (name, p.stat().st_size))
    return made


def sheet(size=232, gap=26):
    """Все варианты рядом на сером поле — чтобы выбрать."""
    keys = sorted(VARIANTS)
    w = gap + (size + gap) * len(keys)
    h = size + gap * 2
    tiles = {k: render(VARIANTS[k], size) for k in keys}
    px = bytearray()
    for row in range(h):
        line = bytearray()
        for col in range(w):
            c = (0x9A, 0x9A, 0x9A, 255)
            for i, k in enumerate(keys):
                x0 = gap + (size + gap) * i
                if x0 <= col < x0 + size and gap <= row < gap + size:
                    o = ((row - gap) * size + (col - x0)) * 4
                    t = tiles[k][o:o + 4]
                    if t[3]:
                        c = tuple(t)
            line += bytes(c)
        px += line
    # Писатель PNG рассчитан на квадрат, поэтому лист собираем отдельно.
    raw = bytearray()
    for row in range(h):
        raw += b"\0" + px[row * w * 4:(row + 1) * w * 4]

    def chunk(tag, data):
        return (struct.pack(">I", len(data)) + tag + data
                + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))

    OUT.mkdir(parents=True, exist_ok=True)
    p = OUT / "лист.png"
    p.write_bytes(b"\x89PNG\r\n\x1a\n"
                  + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 6, 0, 0, 0))
                  + chunk(b"IDAT", zlib.compress(bytes(raw), 9))
                  + chunk(b"IEND", b""))
    return p


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] in ("лист", "sheet"):
        print("лист:", sheet())
    else:
        for line in build():
            print(" ", line)
