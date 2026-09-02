(() => {
  'use strict';

  const input=document.getElementById('feedbackFileInput');
  const meta=document.getElementById('feedbackMeta');
  const status=document.getElementById('feedbackStatus');
  const results=document.getElementById('feedbackResults');
  const countBox=document.getElementById('feedbackCount');
  if (!input || !meta || !status || !results || typeof ExcelJS==='undefined') return;

  const cellText=v=>{
    if(v==null) return '';
    if(v instanceof Date) return v.toISOString().slice(0,10);
    if(typeof v==='object'){
      if(Array.isArray(v.richText)) return v.richText.map(x=>x.text||'').join('');
      if('result' in v) return cellText(v.result);
      if('text' in v) return String(v.text??'');
      if('hyperlink' in v) return String(v.text??v.hyperlink??'');
    }
    return String(v).trim();
  };
  const norm=v=>cellText(v).replace(/\s+/g,'').replace(/／/g,'/').trim();
  const clip=(s,n=180)=>{s=String(s||'').replace(/\s+/g,' ').trim();return s.length>n?`${s.slice(0,n)}…`:s;};
  const unique=xs=>[...new Set(xs.map(x=>String(x||'').trim()).filter(Boolean))];
  const escapeHtml=s=>String(s||'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));

  function findHeaderRow(ws){
    const max=Math.min(ws.rowCount,12);
    for(let r=1;r<=max;r++){
      const vals=[];
      ws.getRow(r).eachCell({includeEmpty:true},(c,col)=>vals[col]=norm(c.value));
      if(vals.includes('導師') && vals.includes('教委會感謝回饋')) return {row:r,headers:vals};
    }
    return null;
  }

  function buildFeedback(group){
    const mentor=group.mentor;
    const students=unique(group.items.map(x=>x['學員']));
    const reasons=unique(group.items.map(x=>x['為什麼值得肯定']));
    const types=unique(group.items.map(x=>x['教學亮點類型']));
    const studentText=students.length ? students.join('、') : '新進同仁';
    const reasonText=reasons.length ? clip(reasons.slice(0,2).join('；'),220) : '您在臨床陪伴中留下了具體且值得肯定的教學行動。';
    const typeText=types.length ? `在${types.slice(0,2).join('、')}的陪伴上，` : '';

    return `${mentor}老師您好：\n\n感謝您在繁忙的臨床工作中，仍願意投入新進同仁的陪伴與教學。\n\n從本月的輔導紀錄中，我們看見您陪伴${studentText}學習的過程。${typeText}${reasonText}\n\n謝謝您願意把臨床經驗轉化成實際的引導與支持，讓學員能在一次次練習與回饋中，逐步建立能力與信心。\n\n感謝您的付出，也謝謝您與我們一起，陪著新進同仁一步一步走得更穩，讓臨床教學這條路一起走得更遠。\n\n護理部教學委員會 敬上`;
  }

  function render(groups){
    results.innerHTML='';
    if(!groups.length){
      countBox.hidden=true;
      status.className='feedback-status feedback-warn';
      status.textContent='目前沒有標記為「納入」的教學亮點。可回到主管審閱版的「教學亮點附件」，在「教委會感謝回饋」欄選擇「納入」後再帶入。';
      return;
    }

    countBox.hidden=false;
    countBox.innerHTML=`本月有 <strong>${groups.length}</strong> 位教師值得被看見 <span>♡</span>`;
    status.className='feedback-status feedback-ok';
    status.textContent='已整理完成。可先閱讀每位教師的教學亮點，再視需要修改文字並一鍵複製。';

    groups.forEach((group,index)=>{
      const students=unique(group.items.map(x=>x['學員']));
      const summaries=unique(group.items.map(x=>x['具體教學內容摘要'])).slice(0,3);
      const types=unique(group.items.map(x=>x['教學亮點類型']));
      const text=buildFeedback(group);
      const card=document.createElement('article');
      card.className='feedback-card';
      card.innerHTML=`
        <div class="feedback-card-head">
          <div>
            <div class="feedback-kicker">教委會感恩有您 ♡</div>
            <h3>${escapeHtml(group.mentor)}老師</h3>
            <div class="feedback-meta-line">${students.length?`陪伴學員：${escapeHtml(students.join('、'))}`:'本月教學亮點'}${types.length?`　｜　${escapeHtml(types.join('、'))}`:''}</div>
          </div>
          <div class="feedback-badge">${group.items.length} 筆亮點</div>
        </div>
        <div class="feedback-highlight">
          <strong>本月值得被看見的教學</strong>
          ${summaries.map(s=>`<p>${escapeHtml(clip(s,210))}</p>`).join('')}
        </div>
        <label class="feedback-copy-label" for="feedbackText${index}">感謝回饋內容</label>
        <textarea id="feedbackText${index}" class="feedback-textarea">${escapeHtml(text)}</textarea>
        <div class="feedback-actions">
          <button type="button" class="feedback-copy" data-copy-target="feedbackText${index}">複製回饋內容</button>
          <span class="feedback-copy-note">可直接在上方修改後再複製</span>
        </div>
      `;
      results.appendChild(card);
    });

    results.querySelectorAll('[data-copy-target]').forEach(btn=>{
      btn.addEventListener('click',async()=>{
        const ta=document.getElementById(btn.dataset.copyTarget);
        if(!ta) return;
        try{
          await navigator.clipboard.writeText(ta.value);
          const old=btn.textContent;
          btn.textContent='已複製 ✓';
          btn.classList.add('copied');
          setTimeout(()=>{btn.textContent=old;btn.classList.remove('copied');},1600);
        }catch(_){
          ta.focus();ta.select();
          document.execCommand('copy');
          btn.textContent='已複製 ✓';
          setTimeout(()=>btn.textContent='複製回饋內容',1600);
        }
      });
    });
  }

  async function processFile(file){
    meta.textContent=file.name;
    status.className='feedback-status';
    status.textContent='正在整理本月教學回饋…';
    countBox.hidden=true;
    results.innerHTML='';

    try{
      const wb=new ExcelJS.Workbook();
      await wb.xlsx.load(await file.arrayBuffer());
      const ws=wb.getWorksheet('教學亮點附件');
      if(!ws) throw new Error('找不到「教學亮點附件」頁籤，請帶入由本系統產生的主管審閱版 Excel。');
      const found=findHeaderRow(ws);
      if(!found) throw new Error('找不到「教委會感謝回饋」欄位，請先使用新版主管審閱檔案。');

      const rows=[];
      for(let r=found.row+1;r<=ws.rowCount;r++){
        const obj={};
        found.headers.forEach((name,c)=>{if(name)obj[name]=cellText(ws.getRow(r).getCell(c).value);});
        if(!obj['導師']) continue;
        if(norm(obj['教委會感謝回饋'])!=='納入') continue;
        rows.push(obj);
      }

      const map=new Map();
      rows.forEach(row=>{
        const mentor=String(row['導師']||'').trim();
        if(!mentor) return;
        if(!map.has(mentor)) map.set(mentor,{mentor,items:[]});
        map.get(mentor).items.push(row);
      });
      render([...map.values()]);
    }catch(err){
      console.error(err);
      status.className='feedback-status feedback-warn';
      status.textContent=err.message||String(err);
    }
  }

  input.addEventListener('change',()=>{
    const file=input.files?.[0];
    if(file) processFile(file);
  });
})();
