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

      html = html.replace(
        '<button id="pdfBtn" class="pdfBtn" onclick="saveAsPdf()">列印／儲存為 PDF ↓</button>',
        '<button id="pdfBtn" class="pdfBtn" onclick="saveAsPdf()">下載學習紀錄 ↓</button>'
      );

      html = html.replace(
        '<div id="pdfHint" class="pdfHint">PDF 只輸出思考紀錄，不包含網站介面。</div>',
        '<div id="pdfHint" class="pdfHint">會下載一份可重新開啟的學習紀錄；系統紀錄仍會同步保存。</div>'
      );

      const saveOverride = `
<script>
function saveAsPdf(){
  try{
    const nameEl=document.getElementById('resultName');
    const hintEl=document.getElementById('pdfHint');
    const rawName=(nameEl&&nameEl.textContent?nameEl.textContent:'學員').trim();
    const safeName=rawName.replace(/[\\/:*?"<>|]/g,'_')||'學員';
    const now=new Date();
    const p=n=>String(n).padStart(2,'0');
    const stamp=now.getFullYear()+p(now.getMonth()+1)+p(now.getDate())+'_'+p(now.getHours())+p(now.getMinutes());

    let recordHtml=buildPdf();
    recordHtml=recordHtml.replace('<div class="toolbar"><button onclick="window.print()">列印／儲存為 PDF</button></div>','');
    recordHtml=recordHtml.replace('.toolbar{position:sticky;top:0;padding:10px;background:#fff;border-bottom:1px solid #eee;text-align:center}.toolbar button{padding:11px 18px;border:0;border-radius:12px;background:#739F99;color:white;font-weight:800}','');
    recordHtml=recordHtml.replace('@media print{.toolbar{display:none}}','');

    const blob=new Blob([recordHtml],{type:'text/html;charset=utf-8'});
    const fileName='EPA_給藥思考紀錄_'+safeName+'_'+stamp+'.html';
    const objectUrl=URL.createObjectURL(blob);
    const link=document.createElement('a');
    link.href=objectUrl;
    link.download=fileName;
    link.style.display='none';
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(()=>URL.revokeObjectURL(objectUrl),3000);

    if(hintEl){
      hintEl.textContent='✓ 已下載學習紀錄。檔案可直接用瀏覽器重新開啟；系統紀錄也會另外保存。';
    }
  }catch(error){
    console.error(error);
    alert('目前無法下載學習紀錄，請稍後再試。');
  }
}
<\/script>
`;

      const closingTag = '</body></html>';
      const closingIndex = html.lastIndexOf(closingTag);
      if (closingIndex >= 0) {
        html = html.slice(0, closingIndex) + saveOverride + html.slice(closingIndex);
      }

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
