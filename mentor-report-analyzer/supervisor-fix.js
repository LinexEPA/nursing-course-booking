(() => {
  'use strict';

  const btn = document.getElementById('supervisorBtn');
  const input = document.getElementById('fileInput');
  const status = document.getElementById('status');
  if (!btn || !input) return;

  const setStatus = (msg, kind='') => {
    status.className = `status ${kind}`.trim();
    status.textContent = msg;
  };
  const text = v => {
    if (v == null) return '';
    if (v instanceof Date) return v.toISOString().slice(0,10);
    if (typeof v === 'object') {
      if (Array.isArray(v.richText)) return v.richText.map(x=>x.text||'').join('');
      if ('result' in v) return text(v.result);
      if ('text' in v) return String(v.text ?? '');
    }
    return String(v).trim();
  };
  const h = v => text(v).replace(/\s+/g,'').replace(/／/g,'/').trim();
  const safe = v => String(v||'').replace(/[\\/:*?"<>|]/g,'_').replace(/\s+/g,'').slice(0,50);
  const fill = (argb) => ({type:'pattern', pattern:'solid', fgColor:{argb}});
  const border = {
    top:{style:'thin',color:{argb:'FFD8E2E8'}}, bottom:{style:'thin',color:{argb:'FFD8E2E8'}},
    left:{style:'thin',color:{argb:'FFD8E2E8'}}, right:{style:'thin',color:{argb:'FFD8E2E8'}}
  };
  const setColWidths = (ws, widths) => widths.forEach((w,i)=>ws.getColumn(i+1).width=w);
  const styleRow = (row, opts={}) => {
    row.eachCell({includeEmpty:true}, c => {
      c.alignment = {vertical:'top', wrapText:true, ...(opts.align||{})};
      c.border = border;
      if (opts.fill) c.fill = fill(opts.fill);
      if (opts.font) c.font = {...opts.font};
    });
    if (opts.height) row.height = opts.height;
  };
  const mergeTitle = (ws, range, value, color='FF203B63') => {
    ws.mergeCells(range); const c=ws.getCell(range.split(':')[0]); c.value=value;
    c.fill=fill(color); c.font={bold:true,color:{argb:'FFFFFFFF'},size:14};
    c.alignment={vertical:'middle',horizontal:'left',wrapText:true};
  };
  const addHeader = (ws, headers) => {
    const r=ws.addRow(headers); styleRow(r,{fill:'FF0D8F98',font:{bold:true,color:{argb:'FFFFFFFF'}},align:{horizontal:'center'},height:36});
  };
  async function download(wb, filename){
    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    const url=URL.createObjectURL(blob); const a=document.createElement('a');
    a.href=url; a.download=filename; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(()=>URL.revokeObjectURL(url),1500);
  }

  async function parseCurrentFile(){
    const file=input.files?.[0];
    if(!file) throw new Error('請先選擇原始 Excel。');
    const wb=new ExcelJS.Workbook(); await wb.xlsx.load(await file.arrayBuffer());
    const ws=wb.worksheets[0]; if(!ws) throw new Error('找不到來源工作表。');
    const headers=[]; ws.getRow(1).eachCell({includeEmpty:true},(c,col)=>headers[col]=h(c.value));
    const rows=[];
    for(let r=2;r<=ws.rowCount;r++){
      const obj={}; headers.forEach((name,c)=>{if(name)obj[name]=text(ws.getRow(r).getCell(c).value)});
      if(!(obj['學員姓名']||obj['導師姓名'])) continue;
      rows.push({sourceRow:r,data:obj});
    }
    if(!rows.length) throw new Error('沒有讀到可分析的輔導紀錄。');
    const analyses=rows.map(x=>MentorRules.analyzeRecord(x.data,x.sourceRow));
    return {file,rows,analyses,year:rows[0].data['年度']||'',month:rows[0].data['月份']||rows[0].data['輔導月份']||''};
  }

  function monthLabel(year,month){
    let m=String(month||'').replace(/月份$/,'').replace(/月$/,'');
    return `${year?year+'年':''}${m?m+'月':''}` || '本月';
  }

  function buildSupervisor(data){
    const {analyses,year,month}=data;
    const total=analyses.length;
    const textCases=analyses.filter(x=>x.textIssue);
    const approval=analyses.filter(x=>x.needsApproval);
    const mgmtCases=analyses.filter(x=>x.management);
    const highlights=analyses.filter(x=>x.highlight);
    const label=monthLabel(year,month);

    const wb=new ExcelJS.Workbook(); wb.creator='Mentor Report Analyzer';

    // 主管月報
    const m=wb.addWorksheet('主管月報'); setColWidths(m,[7,32,9,4,18,4,20,4,18,4,14]);
    mergeTitle(m,'A1:K1',`${label}｜新手導師計畫 主管審閱版`);
    m.mergeCells('A2:K2'); m.getCell('A2').value='主管審閱原則：退件審批與管理確認完全分流。單純明顯錯字不作退件理由；主管只需審批真正需要教師確認／補充的案件。';
    m.getCell('A2').alignment={wrapText:true,vertical:'middle'}; m.getRow(2).height=34;
    m.addRow([]);
    m.addRow(['總紀錄','', '文字／內容案件','', '可直接處理不退件','', '主管需審批','', '管理確認','', '教學亮點']);
    styleRow(m.getRow(4),{fill:'FFE9F3F4',font:{bold:true,color:{argb:'FF203B63'}},align:{horizontal:'center'},height:30});
    m.addRow([total,'',textCases.length,'',textCases.length-approval.length,'',approval.length,'',mgmtCases.length,'',highlights.length]);
    styleRow(m.getRow(5),{font:{bold:true,size:14,color:{argb:'FF203B63'}},align:{horizontal:'center'},height:28});
    m.addRow([]); m.addRow([]);
    m.mergeCells('A8:D8'); m.getCell('A8').value='退件原因摘要（主管真正需要看的差異）';
    m.getCell('A8').font={bold:true,color:{argb:'FF203B63'},size:12};
    m.mergeCells('E8:K8'); m.getCell('E8').value='兩條路徑分開處理'; m.getCell('E8').font={bold:true,color:{argb:'FF203B63'},size:12};
    m.addRow(['處理類型','件數','主管判讀','','退件審批','','','','管理確認','','']); styleRow(m.getRow(9),{fill:'FF0D8F98',font:{bold:true,color:{argb:'FFFFFFFF'}},height:30});
    const countType=t=>textCases.filter(x=>x.textType===t).length;
    const summaryRows=[
      ['明顯錯字／輕微語句，可直接處理',textCases.filter(x=>!x.needsApproval).length,'不退件；不讓教師因作文式校正承受不必要壓力。','','目的','','判斷正式紀錄是否需要教師修正／補充','','目的','','確認主管知悉、制度處理與後續追蹤'],
      ['姓名／原意不明，需教師確認',countType('語意不明／可能相反'),'由規則標記後交人工確認。','','主管動作','','只審批「需主管審批＝是」的案件','','主管動作','','確認處理狀態與是否需下月追蹤'],
      ['紀錄表述風險',countType('紀錄表述風險'),'負向標籤或責任歸屬等內容需人工確認。','','不代表','','管理風險高','','不代表','','文字寫得不好'],
      ['關鍵內容不足',countType('關鍵內容不足'),'若不足以完成重要病安／管理判讀，需補充。','','交集','','同一案件可同時出現在兩頁','','原則','','兩頁各自判斷，不互相取代'],
      ['前後敘述待判斷',countType('前後敘述待確認'),'由主管決定是否需退件確認。','','','','','','','','']
    ];
    summaryRows.forEach(v=>{const r=m.addRow(v); styleRow(r,{height:38});});
    m.addRow([]); m.mergeCells('A16:K16'); m.getCell('A16').value='主管本月只需完成 3 件事'; m.getCell('A16').font={bold:true,color:{argb:'FF203B63'},size:12};
    [[1,`到「退件審批」篩選是否需主管審批＝是，共${approval.length}筆；主管決定退件或不退件。`],[2,`到「管理確認」查看${mgmtCases.length}件案件；確認主管知悉、制度處理及是否需要下月追蹤。`],[3,`到「教學亮點附件」查看${highlights.length}件候選；主管選擇適合的正向回饋方式。`]].forEach(v=>{const r=m.addRow(v); m.mergeCells(`B${r.number}:K${r.number}`); styleRow(r,{height:30});});
    m.views=[{state:'frozen',ySplit:2}];

    // 退件審批
    const rj=wb.addWorksheet('退件審批'); setColWidths(rj,[10,9,12,12,18,42,32,28,15,14,16,24]);
    mergeTitle(rj,'A1:L1','A｜退件審批：文字／內容是否需要退回教師');
    rj.mergeCells('A2:L2'); rj.getCell('A2').value=`主管只需處理「是否需主管審批＝是」的${approval.length}筆。單純明顯錯字原則上直接修正，不作退件理由。`; rj.getCell('A2').alignment={wrapText:true}; rj.getRow(2).height=32;
    rj.addRow([]); addHeader(rj,['單位','案例','學員','導師','問題類型','問題原句／摘要','為什麼需要處理','護理部初審建議','是否需主管審批','涉及管理確認','主管審批結果','主管備註']);
    textCases.forEach(a=>{const r=rj.addRow([a.row['單位']||'待補',a.caseId,a.row['學員姓名']||'',a.row['導師姓名']||'',a.textType,a.excerpt,a.textReason,a.textSuggestion,a.needsApproval?'是':'否',a.management?'是':'否','','']);styleRow(r,{height:70});});
    rj.autoFilter={from:'A4',to:'L4'}; rj.views=[{state:'frozen',ySplit:4}];

    // 管理確認
    const mg=wb.addWorksheet('管理確認'); setColWidths(mg,[10,9,12,12,24,32,36,18,14,24]);
    mergeTitle(mg,'A1:J1','B｜管理確認：主管知悉、病安制度與後續追蹤');
    mg.mergeCells('A2:J2'); mg.getCell('A2').value='此頁與退件審批無關：管理確認不代表紀錄文字不合格。主管只需確認是否知悉、是否已處理，以及是否需要後續追蹤。'; mg.getCell('A2').alignment={wrapText:true}; mg.getRow(2).height=32;
    mg.addRow([]); addHeader(mg,['單位','案例','學員','導師','管理議題','現有紀錄狀態','護理部待確認事項','主管確認結果','下月追蹤','主管備註']);
    mgmtCases.forEach(a=>{const r=mg.addRow([a.row['單位']||'待補',a.caseId,a.row['學員姓名']||'',a.row['導師姓名']||'',a.managementIssue,a.managementState,a.managementQuestion,'','','']);styleRow(r,{height:62});});
    mg.autoFilter={from:'A4',to:'J4'}; mg.views=[{state:'frozen',ySplit:4}];

    // 教學亮點附件
    const hi=wb.addWorksheet('教學亮點附件'); setColWidths(hi,[10,9,12,12,26,42,34,26,20,24]);
    mergeTitle(hi,'A1:J1','C｜教學亮點附件：供主管具體正向回饋');
    hi.mergeCells('A2:J2'); hi.getCell('A2').value='這一頁不是教師排名；只列出能從紀錄看見具體教學行為的代表案例。主管可直接選擇回饋方式。'; hi.getCell('A2').alignment={wrapText:true}; hi.getRow(2).height=32;
    hi.addRow([]); addHeader(hi,['單位','案例','學員','導師','教學亮點類型','具體教學內容摘要','為什麼值得肯定','護理部建議回饋方式','主管確認回饋方式','主管備註']);
    highlights.forEach(a=>{const r=hi.addRow([a.row['單位']||'待補',a.caseId,a.row['學員姓名']||'',a.row['導師姓名']||'',a.highlightType,a.excerpt,a.highlightReason,a.highlightFeedback,'','']);styleRow(r,{height:62});});
    hi.autoFilter={from:'A4',to:'J4'}; hi.views=[{state:'frozen',ySplit:4}];

    return {wb,label};
  }

  btn.addEventListener('click', async (e) => {
    e.preventDefault(); e.stopImmediatePropagation();
    try{
      setStatus('正在產生主管審閱版…');
      const data=await parseCurrentFile();
      const {wb,label}=buildSupervisor(data);
      await download(wb,`${safe(label)}_新手導師計畫_主管審閱.xlsx`);
      setStatus('主管審閱版已產生。','ok');
    }catch(err){
      console.error(err); setStatus(`產生主管審閱版失敗：${err.message||err}`,'warn');
    }
  }, true);
})();
