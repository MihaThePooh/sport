/* Минимальный service worker. Нужен, чтобы Chrome считал страницу
   устанавливаемым приложением, а не ярлыком на сайт. Заодно даёт
   оффлайн: без сети открывается последняя сохранённая страница.

   Стратегия — сеть вперёд, кэш как запасной вариант. Наоборот нельзя:
   страница правится часто, и кэш-вперёд показывал бы старую версию,
   пока её кто-нибудь не сбросит вручную. */
/* Версию поднимаем при замене иконок и прочей статики: старый кэш
   чистится на активации, и лаунчер заберёт новые файлы, а не те,
   что уже лежат. */
var CACHE = 'sport-v2';
var SHELL = ['/', '/icons/icon-192.png', '/icons/icon-512.png'];

self.addEventListener('install', function (e) {
  e.waitUntil(
    caches.open(CACHE).then(function (c) { return c.addAll(SHELL); })
      .then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener('activate', function (e) {
  // Чистим кэши прошлых версий и берём управление сразу, без перезагрузки.
  e.waitUntil(
    caches.keys().then(function (names) {
      return Promise.all(names.map(function (n) {
        return n === CACHE ? null : caches.delete(n);
      }));
    }).then(function () { return self.clients.claim(); })
  );
});

self.addEventListener('fetch', function (e) {
  var url = new URL(e.request.url);
  // Данные и выгрузку не кэшируем никогда: дневник должен быть
  // настоящим, а скачанная таблица — сегодняшней.
  if (e.request.method !== 'GET' || url.pathname.indexOf('/api/') === 0) { return; }
  if (url.origin !== self.location.origin) { return; }

  e.respondWith(
    fetch(e.request).then(function (res) {
      var copy = res.clone();
      caches.open(CACHE).then(function (c) { c.put(e.request, copy); });
      return res;
    }).catch(function () {
      return caches.match(e.request).then(function (hit) {
        return hit || caches.match('/');
      });
    })
  );
});
