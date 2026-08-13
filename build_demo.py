#!/usr/bin/env python3
"""Собирает автономную копию дневника — один файл, который открывается
с диска и работает без сервера.

Нужна, чтобы потрогать интерфейс на телефоне до того, как приложение
выставлено наружу. Внутрь зашиваются настоящие записи из data/log.csv,
а вместо сети подставляется заглушка: те же два адреса, но считает их
браузер, а записи оседают в localStorage.

Запуск: python3 build_demo.py [куда.html]
"""
import json
import sys
from pathlib import Path

import server                      # переиспользуем чтение и разбор CSV

BASE = Path(__file__).resolve().parent
OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else BASE / "тренировки-демо.html"

# Заглушка сети. Логика повторяет server.py: те же рекорды, то же
# правило «запись за тот же день и движение заменяет прежнюю».
SHIM = """
<script>
/* --- ДЕМО-РЕЖИМ ---
   Сервера нет: fetch перехвачен, данные лежат прямо здесь и в памяти
   браузера. Всё, что запишешь, останется только на этом устройстве
   и на боевой дневник не влияет. */
(function () {
  'use strict';
  var MOVES = %(moves)s;
  var СЕГОДНЯ = %(today)s;
  var КЛЮЧ = 'спорт-демо-записи';

  var rows;
  try { rows = JSON.parse(localStorage.getItem(КЛЮЧ) || 'null'); } catch (e) { rows = null; }
  if (!rows) { rows = %(rows)s; }

  function сохранить() {
    try { localStorage.setItem(КЛЮЧ, JSON.stringify(rows)); } catch (e) {}
  }

  function подходы(r) {
    return String(r['подходы'] || '').split(/\\s+/)
      .filter(function (x) { return /^\\d+$/.test(x); })
      .map(Number);
  }
  function сумма(r) {
    var s = подходы(r);
    if (s.length) { return s.reduce(function (a, b) { return a + b; }, 0); }
    return Number(r['всего']) || 0;
  }

  function лучшее(move, since) {
    var rep = 0, tot = 0;
    rows.forEach(function (r) {
      if (r['упражнение'] !== move) { return; }
      if (since && r['дата'] < since) { return; }
      var s = подходы(r);
      if (s.length) { rep = Math.max(rep, Math.max.apply(null, s)); }
      tot = Math.max(tot, сумма(r));
    });
    return { 'подход': rep, 'занятие': tot };
  }

  function годНазад() {
    var d = new Date(СЕГОДНЯ + 'T00:00:00');
    d.setDate(d.getDate() - 365);
    return d.toISOString().slice(0, 10);
  }

  function состояние() {
    var out = { 'сегодня': СЕГОДНЯ, 'движения': {} };
    MOVES.forEach(function (m) {
      var mine = rows.filter(function (r) { return r['упражнение'] === m; })
        .sort(function (a, b) { return a['дата'] < b['дата'] ? -1 : 1; });
      var last = mine[mine.length - 1];
      out['движения'][m] = {
        'рекорд': лучшее(m),
        'загод': лучшее(m, годНазад()),
        'последнее': last ? {
          'дата': last['дата'], 'подходы': подходы(last), 'сумма': сумма(last)
        } : null,
        'занятий': mine.length
      };
    });
    return out;
  }

  function добавить(body) {
    var move = body['упражнение'], sets = body['подходы'] || [], when = body['дата'];
    if (MOVES.indexOf(move) < 0) { throw new Error('неизвестное упражнение'); }
    if (!/^\\d{4}-\\d{2}-\\d{2}$/.test(when || '')) { throw new Error('дата не в формате ГГГГ-ММ-ДД'); }
    if (!sets.length || sets.length > 12) { throw new Error('подходов должно быть от 1 до 12'); }
    rows = rows.filter(function (r) {
      return !(r['дата'] === when && r['упражнение'] === move);
    });
    rows.push({
      'дата': when, 'упражнение': move,
      'подходы': sets.join(' '),
      'всего': sets.reduce(function (a, b) { return a + b; }, 0)
    });
    rows.sort(function (a, b) { return a['дата'] < b['дата'] ? -1 : 1; });
    сохранить();
  }

  function ответ(code, data) {
    return Promise.resolve({
      ok: code < 400, status: code,
      json: function () { return Promise.resolve(data); }
    });
  }

  window.fetch = function (url, opts) {
    url = String(url);
    if (url.indexOf('/api/state') >= 0) { return ответ(200, состояние()); }
    if (url.indexOf('/api/log') >= 0) { return ответ(200, rows); }
    if (url.indexOf('/api/xlsx') >= 0) {
      // Книгу собирает python на сервере — в отдельном файле некому.
      return ответ(400, { 'ошибка': 'в демо таблица не собирается' });
    }
    if (url.indexOf('/api/add') >= 0) {
      try {
        добавить(JSON.parse((opts && opts.body) || '{}'));
      } catch (e) {
        return ответ(400, { 'ошибка': e.message });
      }
      return ответ(200, состояние());
    }
    return ответ(404, { 'ошибка': 'нет такого адреса' });
  };
}());
</script>
"""


def main():
    rows = server.read()
    page = (BASE / "index.html").read_text(encoding="utf-8")

    shim = SHIM % {
        "moves": json.dumps(server.MOVES, ensure_ascii=False),
        "today": json.dumps(server.date.today().isoformat()),
        "rows": json.dumps(rows, ensure_ascii=False),
    }

    # Заглушка должна встать ДО основного скрипта: он на старте зовёт fetch.
    mark = "<script>"
    at = page.index(mark)
    page = page[:at] + shim + "\n" + page[at:]

    page = page.replace(
        "<footer>Таблица пересобирается из этих записей — <code>export.py</code></footer>",
        "<footer>Демо-копия: записи остаются только в этом браузере "
        "и на дневник не влияют.</footer>")
    page = page.replace("<title>Тренировки</title>",
                        "<title>Тренировки — демо</title>")

    OUT.write_text(page, encoding="utf-8")
    print("собрано: %s (%d записей, %.0f КБ)" %
          (OUT, len(rows), OUT.stat().st_size / 1024))


if __name__ == "__main__":
    main()
