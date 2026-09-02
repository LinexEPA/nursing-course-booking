(() => {
  'use strict';

  const input=document.getElementById('feedbackFileInput');
  const meta=document.getElementById('feedbackMeta');
  const status=document.getElementById('feedbackStatus');
  const results=document.getElementById('feedbackResults');
  const countBox=document.getElementById('feedbackCount');
  if (!input || !meta || !status || !results || typeof ExcelJS==='undefined') return;

  // 第二部分採單行換行，讓訊息在一個畫面內更容易讀完。
  const compactStyle=document.createElement('style');
  compactStyle.textContent=`
    .feedback-card{padding:17px 18px !important}
    .feedback-card-head{margin-bottom:10px !important}
    .feedback-highlight{padding:10px 12px !important;margin-bottom:11px !important}
    .feedback-highlight p{margin-top:5px !important}
    .feedback-copy-label{margin-bottom:5px !important}
    .feedback-textarea{min-height:150px !important;max-height:205px;line-height:1.5 !important;padding:11px 14px !important}
    .feedback-actions{margin-top:8px !important}
    @media(max-width:720px){.feedback-textarea{min-height:175px !important;max-height:225px}}
  `;
  document.head.appendChild(compactStyle);

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
  const escapeHtml=s=>String(s||'').replace(/[&<>\"]/g,c=>c==='&'?'&amp;':c==='<'?'&lt;':c==='>'?'&gt;':'&quot;');

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
    const studentText=students.length ? students.slice(0,2).join('、') : '新進同仁';
    const reasonText=reasons.length ? clip(reasons[0],78) : '您把臨床經驗轉化成具體且可以被學員帶走的教學。';
    const typeText=types.length ? types.slice(0,2).join('、') : '臨床陪伴';
    const seed=`${mentor}|${studentText}|${typeText}|${reasonText}`;

    const praiseIntros=[
      '本月特別想謝謝您：',
      '這個月特別想肯定您的投入：',
      '從本月紀錄裡，我們特別感受到您的用心：',
      '這個月，我們特別想謝謝您的陪伴與投入：'
    ];

    const focusMessages={
      reasoning:[
        `您把判斷背後的原因說清楚，讓${studentText}不只知道怎麼做，也更理解為什麼。`,
        `您把臨床思路說得更具體，讓${studentText}能把一次經驗慢慢轉成自己的判斷。`
      ],
      support:[
        `您在教學之外也留意學員的狀態，讓${studentText}能在被支持的情況下慢慢建立信心。`,
        `您讓${studentText}知道遇到困難時有人可以討論，也有人願意陪著整理，這份支持很珍貴。`
      ],
      practice:[
        `您願意示範、觀察再讓學員練習，讓${studentText}有機會把能力真正做出來。`,
        `您把經驗拆成可以練習的步驟，讓${studentText}從跟著做，逐步走向自己完成。`
      ],
      feedback:[
        `您的提醒有具體方向，讓${studentText}知道哪裡做得好、下一步又能怎麼調整。`,
        `您持續觀察學員的變化，也在適當時機給回饋，讓${studentText}的進步有跡可循。`
      ],
      autonomy:[
        `您在協助與放手之間保留空間，讓${studentText}逐步練習判斷與獨立完成。`,
        `您沒有急著替學員完成，而是留下思考與嘗試的空間，讓${studentText}慢慢建立自己的信心。`
      ],
      general:[
        `您把經驗轉成具體引導，讓${studentText}在實際工作中一步一步把能力建立起來。`,
        `您的陪伴不只停在當下提醒，也讓${studentText}更清楚下一次可以怎麼做得更好。`
      ]
    };

    const closings=[
      '謝謝您持續把經驗留給下一位護理師。',
      '這些看似細小的陪伴，正是學員走穩的重要力量。謝謝您。',
      '謝謝您讓好的臨床經驗，能一點一點被學員接住。',
      '謝謝您持續做這件不容易、但很有價值的事。'
    ];

    const intro=praiseIntros[stableIndex(seed,praiseIntros.length)];
    const focusSet=focusMessages[teachingFocus(group)];
    const focus=focusSet[stableIndex(`${seed}|focus`,focusSet.length)];
    const closing=closings[stableIndex(`${seed}|closing`,closings.length)];

    // 每句只換一行，不再插入空白段落。
    return `${mentor}老師您好：\n${intro}${reasonText}\n${focus}\n${closing}\n台中慈濟醫院\n護理部｜教學委員會 ♡`;
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
    status.textContent='已整理完成。回饋採單行換行，可直接修改後複製。';

    groups.forEach((group,index)=>{
      const students=unique(group.items.map(x=>x['學員']));
      const summaries=unique(group.items.map(x=>x['具體教學內容摘要'])).slice(0,1);
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
          ${summaries.map(s=>`<p>${escapeHtml(clip(s,135))}</p>`).join('')}
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