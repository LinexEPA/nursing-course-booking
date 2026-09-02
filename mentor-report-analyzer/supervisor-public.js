(() => {
  'use strict';

  const $ = id => document.getElementById(id);
  const supervisorBtn = $('supervisorBtn');
  const fileInput = $('fileInput');
  const status = $('status');
  if (!supervisorBtn || !fileInput || typeof ExcelJS === 'undefined' || !window.MentorRules) return;

  const cellText = v => {
    if (v == null) return '';
    if (v instanceof Date) return v.toISOString().slice(0,10);
    if (typeof v === 'object') {
      if (Array.isArray(v.richText)) return v.richText.map(x => x.text || '').join('');
      if ('result' in v) return cellText(v.result);
      if ('text' in v) return String(v.text ?? '');
      if ('hyperlink' in v) return String(v.text ?? v.hyperlink ?? '');
    }
    return String(v).trim();
  };
  const h = v => cellText(v).replace(/\s+/g,'').replace(/／/g,'/').trim();
  const safe = v => String(v || '').replace(/[\\/:*?"<>|]/g,'_').replace(/\s+/g,'').slice(0,50);
  const border = {
    top:{style:'thin',color:{argb:'FFD8E2E8'}}, bottom:{style:'thin',color:{argb:'FFD8E2E8'}},
    left:{style:'thin',color:{argb:'FFD8E2E8'}}, right:{style:'thin',color:{argb:'FFD8E2E8'}}
  };
  const fill = argb => ({type:'pattern',pattern:'solid',fgColor:{argb}});
  const unique = xs => [...new Set(xs.filter(Boolean))];
  const clip = (s,n=220) => { s=String(s||'').trim(); return s.length>n ? `${s.slice(0,n)}…` : s; };

  function styleRow(row, opts={}) {
    row.eachCell({includeEmpty:true}, c => {
      c.alignment={vertical:'top',wrapText:true,...(opts.align||{})};
      c.border=border;
      if (opts.fill) c.fill=fill(opts.fill);
      if (opts.font) c.font={...opts.font};
    });
    if (opts.height) row.height=opts.height;
  }
  const setWidths = (ws,widths) => widths.forEach((w,i)=>ws.getColumn(i+1).width=w);
  function mergeTitle(ws, range, value) {
    ws.mergeCells(range);
    const c=ws.getCell(range.split(':')[0]);
    c.value=value; c.fill=fill('FF203B63'); c.font={bold:true,color:{argb:'FFFFFFFF'},size:14};
    c.alignment={vertical:'middle',horizontal:'left',wrapText:true};
  }
  function addHeader(ws, headers) {
    const r=ws.addRow(headers);
    styleRow(r,{fill:'FF0D8F98',font:{bold:true,color:{argb:'FFFFFFFF'}},align:{horizontal:'center'},height:36});
  }

  async function download(wb, filename) {
    const buffer=await wb.xlsx.writeBuffer();
    const blob=new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a'); a.href=url; a.download=filename; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(()=>URL.revokeObjectURL(url),1500);
  }

  async function parseCurrentFile() {
    const file=fileInput.files?.[0];
    if (!file) throw new Error('請先選擇 Excel。');
    const wb=new ExcelJS.Workbook();
    await wb.xlsx.load(await file.arrayBuffer());
    const ws=wb.worksheets[0];
    if (!ws) throw new Error('找不到工作表。');
    const headers=[];
    ws.getRow(1).eachCell({includeEmpty:true},(c,col)=>headers[col]=h(c.value));
    const rows=[];
    for (let r=2;r<=ws.rowCount;r++) {
      const obj={};
      headers.forEach((name,c)=>{ if(name) obj[name]=cellText(ws.getRow(r).getCell(c).value); });
      if (!(obj['學員姓名'] || obj['導師姓名'])) continue;
      rows.push({sourceRow:r,data:obj});
    }
    if (!rows.length) throw new Error('沒有讀到可分析的輔導紀錄。');
    const analyses=rows.map(x=>MentorRules.analyzeRecord(x.data,x.sourceRow));
    return {file,rows,analyses};
  }

  function isPureTypo(a) {
    if (!a.typoAlert) return false;
    const types=String(a.textType||'').split('＋').map(x=>x.trim()).filter(Boolean);
    return types.length>0 && types.every(x=>x==='明顯錯字');
  }
  function nonTypoReview(a) {
    if (!a.textIssue) return false;
    return !isPureTypo(a);
  }

  function monthLabel(parsed) {
    const first=parsed.rows[0].data;
    const y=String(first['年度']||'').replace(/年$/,'');
    const m=String(first['月份']||first['輔導月份']||'').replace(/月份$/,'').replace(/月$/,'');
    return `${y?y+'年':''}${m?m+'月':''}` || '本月';
  }

  function buildWorkbook(parsed) {
    const analyses=parsed.analyses;
    const label=monthLabel(parsed);
    const typoCases=analyses.filter(a=>a.typoAlert);
    const reviewCases=analyses.filter(nonTypoReview);
    const approvalCases=reviewCases.filter(a=>a.needsApproval);
    const mgmtCases=analyses.filter(a=>a.management);
    const highlights=analyses.filter(a=>a.highlight);

    const wb=new ExcelJS.Workbook();

    // 1. 主管月報：只留結果摘要，不顯示內部審閱規則或分流說明。
    const m=wb.addWorksheet('主管月報');
    setWidths(m,[18,18,18,18,18]);
    mergeTitle(m,'A1:E1',`${label}｜新手導師計畫 主管審閱版`);
    m.addRow([]);
    const hr=m.addRow(['總紀錄','語句／內容審批','管理確認','教學亮點','錯字附件']);
    styleRow(hr,{fill:'FFE9F3F4',font:{bold:true,color:{argb:'FF203B63'}},align:{horizontal:'center'},height:34});
    const vr=m.addRow([analyses.length,approvalCases.length,mgmtCases.length,highlights.length,typoCases.length]);
    styleRow(vr,{font:{bold:true,size:16,color:{argb:'FF203B63'}},align:{horizontal:'center'},height:30});
    m.views=[{state:'frozen',ySplit:1}];

    // 2. 退件審批：純錯字不放在主要流程；只留語句／內容需要主管判斷的案件。
    const rj=wb.addWorksheet('退件審批');
    setWidths(rj,[10,9,12,12,24,46,34,30,16,14,18,24]);
    mergeTitle(rj,'A1:L1','退件審批｜語句／內容');
    rj.addRow([]);
    addHeader(rj,['單位','案例','學員','導師','問題類型','問題原句／摘要','為什麼需要處理','護理部初審建議','是否需主管審批','涉及管理確認','主管審批結果','主管備註']);
    reviewCases.forEach(a=>{
      const row=rj.addRow([
        a.row['單位']||'待補',a.caseId,a.row['學員姓名']||'',a.row['導師姓名']||'',a.textType||'',
        clip(a.row['陪伴訓練紀錄']||'',260),a.textReason||'',a.textSuggestion||'',a.needsApproval?'是':'否',a.management?'是':'否','',''
      ]);
      styleRow(row,{height:72});
    });
    rj.autoFilter={from:'A3',to:'L3'}; rj.views=[{state:'frozen',ySplit:3}];

    // 3. 管理確認
    const mg=wb.addWorksheet('管理確認');
    setWidths(mg,[10,9,12,12,26,34,40,18,14,24]);
    mergeTitle(mg,'A1:J1','管理確認');
    mg.addRow([]);
    addHeader(mg,['單位','案例','學員','導師','管理議題','現有紀錄狀態','護理部待確認事項','主管確認結果','下月追蹤','主管備註']);
    mgmtCases.forEach(a=>{
      const row=mg.addRow([a.row['單位']||'待補',a.caseId,a.row['學員姓名']||'',a.row['導師姓名']||'',a.managementIssue||'',a.managementState||'',a.managementQuestion||'','','','']);
      styleRow(row,{height:64});
    });
    mg.autoFilter={from:'A3',to:'J3'}; mg.views=[{state:'frozen',ySplit:3}];

    // 4. 教學亮點附件
    const hi=wb.addWorksheet('教學亮點附件');
    setWidths(hi,[10,9,12,12,26,46,36,28,22,24]);
    mergeTitle(hi,'A1:J1','教學亮點附件');
    hi.addRow([]);
    addHeader(hi,['單位','案例','學員','導師','教學亮點類型','具體教學內容摘要','為什麼值得肯定','護理部建議回饋方式','主管確認回饋方式','主管備註']);
    highlights.forEach(a=>{
      const row=hi.addRow([a.row['單位']||'待補',a.caseId,a.row['學員姓名']||'',a.row['導師姓名']||'',a.highlightType||'',clip(a.row['陪伴訓練紀錄']||'',280),a.highlightReason||'',a.highlightFeedback||'','','']);
      styleRow(row,{height:68});
    });
    hi.autoFilter={from:'A3',to:'J3'}; hi.views=[{state:'frozen',ySplit:3}];

    // 5. 錯字附件：主管可選擇查看，不混入主要審批頁。
    const ty=wb.addWorksheet('錯字附件');
    setWidths(ty,[10,9,12,12,48,32,30,24]);
    mergeTitle(ty,'A1:H1','錯字附件｜供需要時查閱');
    ty.addRow([]);
    addHeader(ty,['單位','案例','學員','導師','原始紀錄摘要','明顯錯字提醒','護理部建議處理','主管備註']);
    typoCases.forEach(a=>{
      const row=ty.addRow([a.row['單位']||'待補',a.caseId,a.row['學員姓名']||'',a.row['導師姓名']||'',clip(a.row['陪伴訓練紀錄']||'',280),a.typoAlert||'','修正明顯錯字後再送出','']);
      styleRow(row,{height:68});
    });
    ty.autoFilter={from:'A3',to:'H3'}; ty.views=[{state:'frozen',ySplit:3}];

    return {wb,label,counts:{total:analyses.length,review:reviewCases.length,approval:approvalCases.length,mgmt:mgmtCases.length,highlight:highlights.length,typo:typoCases.length}};
  }

  async function makeSupervisor() {
    try {
      if (status) { status.className='status'; status.textContent='正在產生主管審閱版…'; }
      const parsed=await parseCurrentFile();
      const {wb,label}=buildWorkbook(parsed);
      await download(wb,`${safe(label)}_新手導師計畫_主管審閱.xlsx`);
      if (status) { status.className='status ok'; status.textContent='主管審閱版已產生。'; }
    } catch (err) {
      console.error(err);
      if (status) { status.className='status warn'; status.textContent=`產生主管審閱版失敗：${err.message||err}`; }
    }
  }

  async function refreshPublicMetrics() {
    try {
      const parsed=await parseCurrentFile();
      const {counts}=buildWorkbook(parsed);
      const mText=$('mText'),mApproval=$('mApproval'),mMgmt=$('mMgmt'),mHighlight=$('mHighlight'),mTotal=$('mTotal');
      if (mTotal) mTotal.textContent=counts.total;
      if (mText) mText.textContent=counts.review;
      if (mApproval) mApproval.textContent=counts.approval;
      if (mMgmt) mMgmt.textContent=counts.mgmt;
      if (mHighlight) mHighlight.textContent=counts.highlight;
      const textLabel=mText?.parentElement?.querySelector('span');
      if (textLabel) textLabel.textContent='語句／內容案件';
    } catch (_) {}
  }

  supervisorBtn.addEventListener('click',e=>{
    e.preventDefault(); e.stopImmediatePropagation(); makeSupervisor();
  },true);

  fileInput.addEventListener('change',()=>{
    setTimeout(refreshPublicMetrics,250);
    setTimeout(refreshPublicMetrics,900);
  });
})();
