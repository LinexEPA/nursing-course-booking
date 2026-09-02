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
  const escapeHtml=s=>String(s||'').replace(/[&<>\"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','\"':'&quot;'}[c]));

  function findHeaderRow(ws){
    const max=Math.min(ws.rowCount,12);
    for(let r=1;r<=max;r++){
      const vals=[];
      ws.getRow(r).eachCell({includeEmpty:true},(c,col)=>vals[col]=norm(c.value));
      if(vals.includes('導師') && vals.includes('教委會感謝回饋')) return {row:r,headers:vals};
    }
    return null;
  }

  function stableIndex(text,length){
    let h=0;
    for(let i=0;i<text.length;i++) h=((h<<5)-h+text.charCodeAt(i))|0;
    return Math.abs(h)%length;
  }

  function teachingFocus(group){
    const all=[
      ...group.items.map(x=>x['教學亮點類型']),
      ...group.items.map(x=>x['為什麼值得肯定']),
      ...group.items.map(x=>x['具體教學內容摘要'])
    ].join(' ');
    if(/臨床推理|判斷|思考|優先|情境|原因|評估/.test(all)) return 'reasoning';
    if(/情緒|壓力|支持|陪伴|信心|安心|適應/.test(all)) return 'support';
    if(/示範|操作|技術|步驟|帶著做|練習/.test(all)) return 'practice';
    if(/回饋|提醒|修正|追問|觀察|追蹤/.test(all)) return 'feedback';
    if(/獨立|自主|放手|信任|承擔/.test(all)) return 'autonomy';
    return 'general';
  }

  function buildFeedback(group){
    const mentor=group.mentor;
    const students=unique(group.items.map(x=>x['學員']));
    const reasons=unique(group.items.map(x=>x['為什麼值得肯定']));
    const types=unique(group.items.map(x=>x['教學亮點類型']));
    const studentText=students.length ? students.join('、') : '新進同仁';
    const reasonText=reasons.length ? clip(reasons.slice(0,2).join('；'),220) : '您在臨床陪伴中留下了具體且值得肯定的教學行動。';
    const typeText=types.length ? types.slice(0,2).join('、') : '臨床陪伴';
    const seed=`${mentor}|${studentText}|${typeText}|${reasonText}`;

    const openings=[
      `感謝您在繁忙的臨床工作中，仍願意為新進同仁保留一段可以學習、提問與被回饋的空間。`,
      `從本月的輔導紀錄裡，可以看見您不是只把工作教完，而是持續陪著新進同仁把每一步做得更穩。`,
      `謝謝您願意把日常臨床中的經驗留下來，成為新進同仁可以理解、可以練習，也可以再往前走的方向。`,
      `臨床教學常發生在最忙碌的時候，而您仍願意停下來看見學員的需要，這份投入很值得被肯定。`
    ];

    const focusMessages={
      reasoning:[
        `在${typeText}的陪伴上，您不只是提供答案，也幫助${studentText}理解判斷背後的原因，讓經驗慢慢轉化成可以帶走的思考方式。`,
        `您把臨床上的判斷與思路說得更具體，讓${studentText}有機會從「知道怎麼做」走向「理解為什麼這樣做」。`
      ],
      support:[
        `您在${typeText}的陪伴中，既看見學習任務，也留意學員當下的感受與承受度；這樣的支持能讓${studentText}更有餘裕累積信心。`,
        `除了教會事情本身，您也讓${studentText}在需要時知道有人可以討論、有人願意陪著整理，這份安全感很珍貴。`
      ],
      practice:[
        `在${typeText}的教學裡，您把經驗拆成可以實際操作與反覆練習的步驟，讓${studentText}能從跟著做，逐漸走向自己完成。`,
        `您願意示範、觀察，再把練習機會交還給學員，讓${studentText}不是只看過，而是真的有機會把能力做出來。`
      ],
      feedback:[
        `您給出的提醒與回饋具有具體方向，讓${studentText}知道哪裡已經做得不錯、下一步又可以怎麼調整。`,
        `從紀錄中能看見您持續觀察學員的變化，也願意在適當時機給出提醒，讓${studentText}的進步有跡可循。`
      ],
      autonomy:[
        `您在支持與放手之間拿捏得很細緻，讓${studentText}可以在需要時獲得協助，也逐步練習承擔與獨立完成。`,
        `您沒有急著替學員把事情做好，而是留下思考與嘗試的空間，讓${studentText}能逐漸建立自己的判斷與信心。`
      ],
      general:[
        `在${typeText}的陪伴上，您把自己的臨床經驗轉化成具體可理解的引導，讓${studentText}有機會在實際工作中一步一步累積能力。`,
        `您的教學不只存在於當下的提醒，也透過持續觀察與陪伴，讓${studentText}知道下一步可以怎麼做得更好。`
      ]
    };

    const closings=[
      `這樣的教學未必轟轟烈烈，卻會在一次次被理解、被提醒與被相信的經驗裡，慢慢成為學員成長的一部分。謝謝您的投入。`,
      `謝謝您讓臨床經驗不只停留在自己身上，而是成為下一位護理師可以接得住、用得上的能力。`,
      `能被好好陪著學習，是新進同仁建立專業感的重要過程。謝謝您持續做這件不容易、但很有價值的事。`,
      `謝謝您願意在每天的工作裡多做一點觀察、多留一點回饋，這些看似細小的陪伴，往往正是學員走穩的重要力量。`
    ];

    const opening=openings[stableIndex(seed,openings.length)];
    const focusSet=focusMessages[teachingFocus(group)];
    const focus=focusSet[stableIndex(`${seed}|focus`,focusSet.length)];
    const closing=closings[stableIndex(`${seed}|closing`,closings.length)];

    return `${mentor}老師您好：\n\n${opening}\n\n本月特別值得被看見的是：${reasonText}\n\n${focus}\n\n${closing}\n\n台中慈濟醫院\n護理部｜教學委員會 ♡`;
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
