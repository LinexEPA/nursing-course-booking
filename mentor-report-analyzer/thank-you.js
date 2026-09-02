(() => {
  'use strict';

  const status = document.getElementById('status');
  if (!status) return;

  const style = document.createElement('style');
  style.textContent = `
    .download-thanks{
      position:fixed;right:24px;bottom:24px;z-index:9999;width:min(360px,calc(100vw - 32px));
      padding:18px 18px 16px;border:1px solid #eadcdf;border-radius:20px;
      background:rgba(255,253,252,.98);box-shadow:0 18px 45px rgba(110,81,91,.16);
      color:#625a60;opacity:0;transform:translateY(14px) scale(.98);pointer-events:none;
      transition:opacity .24s ease,transform .24s ease;
      backdrop-filter:blur(8px);
    }
    .download-thanks.show{opacity:1;transform:translateY(0) scale(1);pointer-events:auto}
    .download-thanks-head{display:flex;align-items:flex-start;gap:12px;padding-right:24px}
    .download-thanks-icon{
      display:flex;align-items:center;justify-content:center;flex:0 0 auto;width:42px;height:42px;
      border-radius:14px;background:#f3e4e8;color:#a97482;font-size:21px;font-weight:700;
    }
    .download-thanks h3{margin:1px 0 5px;font-size:17px;line-height:1.35;color:#725f67}
    .download-thanks p{margin:0;color:#83777d;font-size:13.5px;line-height:1.75}
    .download-thanks-close{
      position:absolute;right:12px;top:10px;border:0;background:transparent;color:#ad9da3;
      font-size:20px;line-height:1;cursor:pointer;padding:5px;border-radius:50%;
    }
    .download-thanks-close:hover{background:#faf1f3;color:#8f747d}
    .download-thanks-action{
      margin:14px 0 0 54px;border:0;border-radius:999px;padding:7px 15px;
      background:#c995a2;color:white;font-size:13px;font-weight:700;cursor:pointer;
      box-shadow:0 5px 12px rgba(169,116,130,.13);
    }
    .download-thanks-action:hover{background:#b98290}
    @media(max-width:640px){
      .download-thanks{right:16px;bottom:16px;width:calc(100vw - 32px)}
    }
  `;
  document.head.appendChild(style);

  const card = document.createElement('aside');
  card.className = 'download-thanks';
  card.setAttribute('role','status');
  card.setAttribute('aria-live','polite');
  card.innerHTML = `
    <button class="download-thanks-close" type="button" aria-label="關閉">×</button>
    <div class="download-thanks-head">
      <div class="download-thanks-icon">♡</div>
      <div>
        <h3>下載完成</h3>
        <p>謝謝你完成本次輔導紀錄整理。<br>系統已協助先行彙整，最後仍請以人工審閱與確認為準。</p>
      </div>
    </div>
    <button class="download-thanks-action" type="button">完成</button>
  `;
  document.body.appendChild(card);

  let hideTimer = null;
  const hide = () => {
    card.classList.remove('show');
    if (hideTimer) { clearTimeout(hideTimer); hideTimer = null; }
  };
  const show = () => {
    if (hideTimer) clearTimeout(hideTimer);
    requestAnimationFrame(() => card.classList.add('show'));
    hideTimer = setTimeout(hide, 8000);
  };

  card.querySelector('.download-thanks-close')?.addEventListener('click', hide);
  card.querySelector('.download-thanks-action')?.addEventListener('click', hide);
  document.addEventListener('keydown', e => { if (e.key === 'Escape') hide(); });

  let lastText = '';
  const checkStatus = () => {
    const text = (status.textContent || '').trim();
    if (text === lastText) return;
    lastText = text;
    if (/^(細緻分析版|主管審閱版)已產生。?$/.test(text)) show();
  };

  checkStatus();
  new MutationObserver(checkStatus).observe(status,{childList:true,characterData:true,subtree:true});
})();
