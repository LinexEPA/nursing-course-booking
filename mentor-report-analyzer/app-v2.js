(() => {
  'use strict';

  const $ = id => document.getElementById(id);
  const els = {
    input:$('fileInput'), drop:$('dropZone'), meta:$('fileMeta'), status:$('status'),
    detail:$('detailBtn'), supervisor:$('supervisorBtn'),
    total:$('mTotal'), text:$('mText'), approval:$('mApproval'), mgmt:$('mMgmt'), highlight:$('mHighlight')
  };
  const state={fileName:'',rows:[],analyses:[],year:'',month:''};
  const REQUIRED=['學員姓名','導師姓名','陪伴訓練紀錄'];
  const setStatus=(msg,kind='')=>{els.status.className=`status ${kind}`.trim();els.status.textContent=msg;};
  const cellText=v=>{
    if(v==null)return''; if(v instanceof Date)return v.toISOString().slice(0,10);
    if(typeof v==='object'){
      if(Array.isArray(v.richText))return v.richText.map(x=>x.text||'').join('');
      if('result' in v)return cellText(v.result); if('text' in v)return String(v.text??'');
      if('hyperlink' in v)return String(v.text??v.hyperlink??'');
    }
    return String(v).trim();
  };
  const h=v=>cellText(v).replace(/\s+/g,'').replace(/／/g,'/').trim();
  const safe=v=>String(v||'').replace(/[\\/:*?"<>|]/g,'_').replace(/\s+/g,'').slice(0,50);
  const monthLabel=()=>{
    let y=String(state.year||'').replace(/年$/,'');
    let m=String(state.month||'').replace(/月份$/,'').replace(/月$/,'');
    return `${y?y+'年':''}${m?m+'月':''}`||'本月';
  };

  async function loadFile(file){
    if(!file)return;
    if(!/\.xlsx?$/i.test(file.name)){setStatus('請選擇 Excel（.xlsx 或 .xls）檔案。','warn');return;}
    try{
      setStatus('正在讀取 Excel…');
      const wb=new ExcelJS.Workbook(); await wb.xlsx.load(await file.arrayBuffer());
      const ws=wb.worksheets[0]; if(!ws)throw new Error('找不到工作表。');
      const headers=[]; ws.getRow(1).eachCell({includeEmpty:true},(c,col)=>headers[col]=h(c.value));
      const headerSet=new Set(headers.filter(Boolean));
      const missing=REQUIRED.filter(x=>!headerSet.has(h(x))); if(missing.length)throw new Error(`來源格式不符，缺少欄位：${missing.join('、')}`);
      const rows=[];
      for(let r=2;r<=ws.rowCount;r++){
        const obj={}; headers.forEach((name,c)=>{if(name)obj[name]=cellText(ws.getRow(r).getCell(c).value)});
        if(!(obj['學員姓名']||obj['導師姓名']))continue;
        rows.push({sourceRow:r,data:obj});
      }
      if(!rows.length)throw new Error('沒有讀到可分析的輔導紀錄。');
      state.fileName=file.name; state.rows=rows; state.year=rows[0].data['年度']||''; state.month=rows[0].data['月份']||rows[0].data['輔導月份']||'';
      state.analyses=rows.map(x=>MentorRules.analyzeRecord(x.data,x.sourceRow));
      updateSummary(); els.meta.textContent=`${file.name}｜${rows.length} 筆紀錄`; els.detail.disabled=false; els.supervisor.disabled=false;
      setStatus(`分析完成：固定規則 ${MentorRules.RULE_VERSION} 已完成第一輪分類。`,'ok');
    }catch(err){console.error(err);state.rows=[];state.analyses=[];els.detail.disabled=true;els.supervisor.disabled=true;setStatus(`讀取失敗：${err.message||err}`,'warn');}
  }

  function updateSummary(){
    const a=state.analyses, counts={total:a.length,text:a.filter(x=>x.textIssue).length,approval:a.filter(x=>x.needsApproval).length,mgmt:a.filter(x=>x.management).length,highlight:a.filter(x=>x.highlight).length};
    els.total.textContent=counts.total;els.text.textContent=counts.text;els.approval.textContent=counts.approval;els.mgmt.textContent=counts.mgmt;els.highlight.textContent=counts.highlight;return counts;
  }

  async function download(wb,filename){
    const buffer=await wb.xlsx.writeBuffer();
    const blob=new Blob([buffer],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=filename;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1500);
  }

  const border={top:{style:'thin',color:{argb:'FFD8E2E8'}},bottom:{style:'thin',color:{argb:'FFD8E2E8'}},left:{style:'thin',color:{argb:'FFD8E2E8'}},right:{style:'thin',color:{argb:'FFD8E2E8'}}};
  const fill=argb=>({type:'pattern',pattern:'solid',fgColor:{argb}});
  function styleRow(row,opts={}){
    row.eachCell({includeEmpty:true},c=>{c.alignment={vertical:'top',wrapText:true,...(opts.align||{})};c.border=border;if(opts.fill)c.fill=fill(opts.fill);if(opts.font)c.font={...opts.font};}); if(opts.height)row.height=opts.height;
  }
  const setWidths=(ws,widths)=>widths.forEach((w,i)=>ws.getColumn(i+1).width=w);
  function mergeTitle(ws,range,value){ws.mergeCells(range);const c=ws.getCell(range.split(':')[0]);c.value=value;c.fill=fill('FF203B63');c.font={bold:true,color:{argb:'FFFFFFFF'},size:14};c.alignment={vertical:'middle',horizontal:'left',wrapText:true};}
  function addHeader(ws,headers){const r=ws.addRow(headers);styleRow(r,{fill:'FF0D8F98',font:{bold:true,color:{argb:'FFFFFFFF'}},align:{horizontal:'center'},height:36});}

  async function makeDetailed(){
    try{
      setStatus('正在產生細緻分析版…');
      const wb=new ExcelJS.Workbook(); wb.creator='Mentor Report Analyzer';
      const ws=wb.addWorksheet('細緻分析');
      const cols=[
        ['原始列號',10],['案例',10],['年度',8],['月份',10],['學員姓名',12],['導師姓名',12],['原始陪伴訓練紀錄',50],
        ['明顯錯字提醒',26],['問題類型',24],['是否建議退回',14],['為什麼需要處理',34],['建議教師修改方向',30],
        ['管理確認',12],['管理議題',24],['管理待確認事項',36],['教學亮點候選',14],['教學亮點類型',26],
        ['規則命中',16],['規則信心',10],['Codex判定',18],['最終人工判定',18],['備註',24]
      ];
      ws.columns=cols.map(([header,width])=>({header,key:header,width})); ws.getRow(1).height=34; ws.getRow(1).eachCell(c=>{c.font={bold:true,color:{argb:'FFFFFFFF'}};c.fill=fill('FF0D8F98');c.alignment={vertical:'middle',horizontal:'center',wrapText:true};c.border=border;});
      ws.views=[{state:'frozen',ySplit:1}]; ws.autoFilter={from:'A1',to:'V1'};
      state.analyses.forEach(a=>{
        const r=a.row; const out=ws.addRow({
          '原始列號':a.sourceRow,'案例':a.caseId,'年度':r['年度']||'','月份':r['月份']||r['輔導月份']||'','學員姓名':r['學員姓名']||'','導師姓名':r['導師姓名']||'',
          '原始陪伴訓練紀錄':r['陪伴訓練紀錄']||'','明顯錯字提醒':a.typoAlert||'','問題類型':a.textType,'是否建議退回':a.needsApproval?'是':'否',
          '為什麼需要處理':a.textReason,'建議教師修改方向':a.textSuggestion,'管理確認':a.management?'是':'否','管理議題':a.managementIssue,'管理待確認事項':a.managementQuestion,
          '教學亮點候選':a.highlight?'是':'否','教學亮點類型':a.highlightType,'規則命中':a.ruleHits,'規則信心':a.confidence,'Codex判定':'','最終人工判定':'','備註':''
        }); styleRow(out,{height:72});
      });
      const info=wb.addWorksheet('驗證說明');setWidths(info,[24,90]);
      info.addRows([['規則版本',MentorRules.RULE_VERSION],['退件原則','明顯錯字、姓名／原意錯誤、語意不明、重要內容不足等需退件；標點符號、空格與純格式問題不退件。'],['標籤化判斷','不只看負向字眼；若有具體可觀察行為、結果與改善措施，不應僅因「心不在焉／注意力不集中」等詞直接判為標籤。'],['驗證目的','保留 Codex 判定與最終人工判定，用於後續逐筆比較。'],['來源檔',state.fileName]]);
      info.getColumn(1).font={bold:true,color:{argb:'FF203B63'}};info.eachRow(r=>{r.alignment={vertical:'top',wrapText:true};r.height=30;});
      await download(wb,`${safe(monthLabel())}_新手導師計畫_細緻分析.xlsx`);setStatus('細緻分析版已產生。','ok');
    }catch(err){console.error(err);setStatus(`產生細緻分析版失敗：${err.message||err}`,'warn');}
  }

  function buildSupervisorWorkbook(){
    const analyses=state.analyses,total=analyses.length,textCases=analyses.filter(x=>x.textIssue),approval=analyses.filter(x=>x.needsApproval),mgmtCases=analyses.filter(x=>x.management),highlights=analyses.filter(x=>x.highlight),label=monthLabel();
    const wb=new ExcelJS.Workbook();wb.creator='Mentor Report Analyzer';

    const m=wb.addWorksheet('主管月報');setWidths(m,[7,32,9,4,18,4,20,4,18,4,14]);mergeTitle(m,'A1:K1',`${label}｜新手導師計畫 主管審閱版`);
    m.mergeCells('A2:K2');m.getCell('A2').value='主管審閱原則：退件審批與管理確認完全分流。明顯錯字需退回修正；標點、空格與純格式問題不作退件理由。';m.getCell('A2').alignment={wrapText:true,vertical:'middle'};m.getRow(2).height=34;
    m.addRow([]);m.addRow(['總紀錄','', '文字／內容案件','', '可直接處理不退件','', '主管需審批','', '管理確認','', '教學亮點']);styleRow(m.getRow(4),{fill:'FFE9F3F4',font:{bold:true,color:{argb:'FF203B63'}},align:{horizontal:'center'},height:30});
    m.addRow([total,'',textCases.length,'',textCases.length-approval.length,'',approval.length,'',mgmtCases.length,'',highlights.length]);styleRow(m.getRow(5),{font:{bold:true,size:14,color:{argb:'FF203B63'}},align:{horizontal:'center'},height:28});m.addRow([]);m.addRow([]);
    m.mergeCells('A8:D8');m.getCell('A8').value='退件原因摘要（主管真正需要看的差異）';m.getCell('A8').font={bold:true,color:{argb:'FF203B63'},size:12};m.mergeCells('E8:K8');m.getCell('E8').value='兩條路徑分開處理';m.getCell('E8').font={bold:true,color:{argb:'FF203B63'},size:12};
    m.addRow(['處理類型','件數','主管判讀','','退件審批','','','','管理確認','','']);styleRow(m.getRow(9),{fill:'FF0D8F98',font:{bold:true,color:{argb:'FFFFFFFF'}},height:30});
    const countType=t=>textCases.filter(x=>x.textType===t).length;
    [
      ['明顯錯字',countType('明顯錯字'),'正式紀錄需修正後再送出。','','目的','','判斷正式紀錄是否需教師修正／補充','','目的','','確認主管知悉、制度處理與後續追蹤'],
      ['語意不明／可能相反',countType('語意不明／可能相反'),'不能由審閱者自行猜測原意。','','主管動作','','只審批「需主管審批＝是」的案件','','主管動作','','確認處理狀態與是否需下月追蹤'],
      ['疑似不適當評價／標籤化描述',countType('疑似不適當評價／標籤化描述'),'需看是否有具體可觀察行為支持，再決定是否退件。','','不代表','','只要出現負向字眼就退件','','不代表','','文字寫得不好'],
      ['關鍵內容不足',countType('關鍵內容不足'),'重要病安／管理資訊不足時需補充。','','交集','','同一案件可同時出現在兩頁','','原則','','兩頁各自判斷，不互相取代'],
      ['前後敘述待確認',countType('前後敘述待確認'),'由主管確認是否需退件。','','','','','','','','']
    ].forEach(v=>{const r=m.addRow(v);styleRow(r,{height:40});});
    m.addRow([]);m.mergeCells('A16:K16');m.getCell('A16').value='主管本月只需完成 3 件事';m.getCell('A16').font={bold:true,color:{argb:'FF203B63'},size:12};
    [[1,`到「退件審批」查看${approval.length}筆需主管審批案件；主管決定退件或不退件。`],[2,`到「管理確認」查看${mgmtCases.length}件案件；確認主管知悉、制度處理及是否需要下月追蹤。`],[3,`到「教學亮點附件」查看${highlights.length}件候選；主管選擇適合的正向回饋方式。`]].forEach(v=>{const r=m.addRow(v);m.mergeCells(`B${r.number}:K${r.number}`);styleRow(r,{height:30});});m.views=[{state:'frozen',ySplit:2}];

    const rj=wb.addWorksheet('退件審批');setWidths(rj,[10,9,12,12,22,42,30,28,15,14,16,24]);mergeTitle(rj,'A1:L1','A｜退件審批：文字／內容是否需要退回教師');
    rj.mergeCells('A2:L2');rj.getCell('A2').value=`主管只需處理「是否需主管審批＝是」的${approval.length}筆。明顯錯字需修正；標點、空格與純格式問題不列入退件。`;rj.getCell('A2').alignment={wrapText:true};rj.getRow(2).height=32;rj.addRow([]);addHeader(rj,['單位','案例','學員','導師','問題類型','問題原句／摘要','為什麼需要處理','護理部初審建議','是否需主管審批','涉及管理確認','主管審批結果','主管備註']);
    textCases.forEach(a=>{const r=rj.addRow([a.row['單位']||'待補',a.caseId,a.row['學員姓名']||'',a.row['導師姓名']||'',a.textType,a.excerpt,a.textReason,a.typoAlert?`${a.textSuggestion}｜${a.typoAlert}`:a.textSuggestion,a.needsApproval?'是':'否',a.management?'是':'否','','']);styleRow(r,{height:72});});rj.autoFilter={from:'A4',to:'L4'};rj.views=[{state:'frozen',ySplit:4}];

    const mg=wb.addWorksheet('管理確認');setWidths(mg,[10,9,12,12,24,32,36,18,14,24]);mergeTitle(mg,'A1:J1','B｜管理確認：主管知悉、病安制度與後續追蹤');
    mg.mergeCells('A2:J2');mg.getCell('A2').value='此頁與退件審批無關：管理確認不代表紀錄文字不合格。主管只需確認是否知悉、是否已處理，以及是否需要後續追蹤。';mg.getCell('A2').alignment={wrapText:true};mg.getRow(2).height=32;mg.addRow([]);addHeader(mg,['單位','案例','學員','導師','管理議題','現有紀錄狀態','護理部待確認事項','主管確認結果','下月追蹤','主管備註']);
    mgmtCases.forEach(a=>{const r=mg.addRow([a.row['單位']||'待補',a.caseId,a.row['學員姓名']||'',a.row['導師姓名']||'',a.managementIssue,a.managementState,a.managementQuestion,'','','']);styleRow(r,{height:64});});mg.autoFilter={from:'A4',to:'J4'};mg.views=[{state:'frozen',ySplit:4}];

    const hi=wb.addWorksheet('教學亮點附件');setWidths(hi,[10,9,12,12,26,42,34,26,20,24]);mergeTitle(hi,'A1:J1','C｜教學亮點附件：供主管具體正向回饋');
    hi.mergeCells('A2:J2');hi.getCell('A2').value='這一頁不是教師排名；只列出能從紀錄看見具體教學行為的代表案例。主管可直接選擇回饋方式。';hi.getCell('A2').alignment={wrapText:true};hi.getRow(2).height=32;hi.addRow([]);addHeader(hi,['單位','案例','學員','導師','教學亮點類型','具體教學內容摘要','為什麼值得肯定','護理部建議回饋方式','主管確認回饋方式','主管備註']);
    highlights.forEach(a=>{const r=hi.addRow([a.row['單位']||'待補',a.caseId,a.row['學員姓名']||'',a.row['導師姓名']||'',a.highlightType,a.excerpt,a.highlightReason,a.highlightFeedback,'','']);styleRow(r,{height:64});});hi.autoFilter={from:'A4',to:'J4'};hi.views=[{state:'frozen',ySplit:4}];
    return wb;
  }

  async function makeSupervisor(){
    try{setStatus('正在產生主管審閱版…');const wb=buildSupervisorWorkbook();await download(wb,`${safe(monthLabel())}_新手導師計畫_主管審閱.xlsx`);setStatus('主管審閱版已產生。','ok');}
    catch(err){console.error(err);setStatus(`產生主管審閱版失敗：${err.message||err}`,'warn');}
  }

  els.input.addEventListener('change',e=>loadFile(e.target.files?.[0]));els.detail.addEventListener('click',makeDetailed);els.supervisor.addEventListener('click',makeSupervisor);
  ['dragenter','dragover'].forEach(evt=>els.drop.addEventListener(evt,e=>{e.preventDefault();els.drop.classList.add('drag')}));['dragleave','drop'].forEach(evt=>els.drop.addEventListener(evt,e=>{e.preventDefault();els.drop.classList.remove('drag')}));els.drop.addEventListener('drop',e=>loadFile(e.dataTransfer?.files?.[0]));
})();
