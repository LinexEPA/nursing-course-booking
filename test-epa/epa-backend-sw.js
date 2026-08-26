const EPA_API_URL = 'https://script.google.com/macros/s/AKfycbx8UPej4nsrJbwaN3VXDc3kDTDf4PTu1Amxzd4b7g_FzS25gEKucOwQI1Gag4I3TpR7/exec';
const CORE_FILE = 'epa-medication-core.html';

self.addEventListener('install', event => {
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(self.clients.claim());
});

self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  if (!url.pathname.endsWith('/' + CORE_FILE)) return;

  event.respondWith((async () => {
    try {
      const response = await fetch(event.request, { cache: 'no-store' });
      let html = await response.text();

      html = html.replace(
        "// 後端網址會在下一步連接既有 EPA Apps Script Web App。\nconst EPA_API_URL='';",
        "// 已連接 EPA Apps Script 後端。\nconst EPA_API_URL='" + EPA_API_URL + "';"
      );

      return new Response(html, {
        status: response.status,
        statusText: response.statusText,
        headers: {
          'Content-Type': 'text/html; charset=utf-8',
          'Cache-Control': 'no-store, max-age=0'
        }
      });
    } catch (error) {
      return new Response(
        '<!doctype html><meta charset="utf-8"><title>EPA 載入失敗</title><body style="font-family:sans-serif;padding:24px">EPA 工具載入失敗，請返回 Learning Hub 後再試一次。</body>',
        { headers: { 'Content-Type': 'text/html; charset=utf-8' } }
      );
    }
  })());
});
