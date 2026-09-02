(() => {
  'use strict';

  const $ = id => document.getElementById(id);
  const status = $('status');
  const detailBtn = $('detailBtn');
  const fileInput = $('fileInput');

  // 正式使用介面：不顯示開發／驗證術語。
  const heroText = document.querySelector('.hero p');
  if (heroText) heroText.textContent = '上傳每月輔導紀錄清冊，系統先行整理可能需要關注的文字、管理事項與教學亮點，供人工審閱與確認，並產生審閱用 Excel。';

  const note = document.querySelector('.note');
  if (note) note.textContent = '資料處理說明：上傳檔案僅於目前瀏覽器中處理，本系統不儲存上傳檔案；產出的 Excel 會直接下載至您的裝置。';

  const version = document.querySelector('.version');
  if (version) version.style.display = 'none';

  const textMetric = $('mText')?.parentElement?.querySelector('span');
  if (textMetric) textMetric.textContent = '需文字檢視';

  function publicStatusText(text) {
    if (/分析完成：/.test(text)) return '分析完成，可下載細緻分析版或主管審閱版。';
    return text;
  }

  if (status) {
    const cleanStatus = () => {
      const next = publicStatusText(status.textContent || '');
      if (next !== status.textContent) status.textContent = next;
    };
    cleanStatus();
    new MutationObserver(cleanStatus).observe(status, { childList: true, characterData: true, subtree: true });
  }

  if (!detailBtn || !fileInput || typeof ExcelJS === 'undefined') return;

  const cellText = v => {
    if (v == null) return '';
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    if (typeof v === 'object') {
      if (Array.isArray(v.richText)) return v.richText.map(x => x.text || '').join('');
      if ('result' in v) return cellText(v.result);
      if ('text' in v) return String(v.text ?? '');
      if ('hyperlink' in v) return String(v.text ?? v.hyperlink ?? '');
    }
    return String(v).trim();
  };
  const h = v => cellText(v).replace(/\s+/g, '').replace(/／/g, '/').trim();
  const safe = v => String(v || '').replace(/[\\/:*?"<>|]/g, '_').replace(/\s+/g, '').slice(0, 50);
  const border = {
    top:{style:'thin',color:{argb:'FFD8E2E8'}}, bottom:{style:'thin',color:{argb:'FFD8E2E8'}},
    left:{style:'thin',color:{argb:'FFD8E2E8'}}, right:{style:'thin',color:{argb:'FFD8E2E8'}}
  };
  const fill = argb => ({type:'pattern',pattern:'solid',fgColor:{argb}});

  async function download(wb, filename) {
    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  }

  async function parseCurrentFile() {
    const file = fileInput.files?.[0];
    if (!file) throw new Error('請先選擇 Excel。');
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(await file.arrayBuffer());
    const ws = wb.worksheets[0];
    if (!ws) throw new Error('找不到工作表。');
    const headers = [];
    ws.getRow(1).eachCell({includeEmpty:true}, (c,col) => headers[col] = h(c.value));
    const rows = [];
    for (let r=2; r<=ws.rowCount; r++) {
      const obj = {};
      headers.forEach((name,c) => { if (name) obj[name] = cellText(ws.getRow(r).getCell(c).value); });
      if (!(obj['學員姓名'] || obj['導師姓名'])) continue;
      rows.push({sourceRow:r, data:obj});
    }
    if (!rows.length) throw new Error('沒有讀到可分析的輔導紀錄。');
    return {file, rows, analyses:rows.map(x => MentorRules.analyzeRecord(x.data, x.sourceRow))};
  }

  async function makePublicDetailed() {
    try {
      if (status) { status.className='status'; status.textContent='正在產生細緻分析版…'; }
      const parsed = await parseCurrentFile();
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('細緻分析');
      const cols = [
        ['原始列號',10],['案例',10],['年度',8],['月份',10],['學員姓名',12],['導師姓名',12],['原始陪伴訓練紀錄',50],
        ['明顯錯字提醒',26],['問題類型',25],['是否建議退回',14],['為什麼需要處理',34],['建議教師修改方向',30],
        ['管理確認',12],['管理議題',24],['管理待確認事項',36],['教學亮點候選',14],['教學亮點類型',26],['備註',24]
      ];
      ws.columns = cols.map(([header,width]) => ({header,key:header,width}));
      ws.getRow(1).height = 34;
      ws.getRow(1).eachCell(c => {
        c.font={bold:true,color:{argb:'FFFFFFFF'}}; c.fill=fill('FF0D8F98');
        c.alignment={vertical:'middle',horizontal:'center',wrapText:true}; c.border=border;
      });
      ws.views=[{state:'frozen',ySplit:1}];
      ws.autoFilter={from:'A1',to:'R1'};

      parsed.analyses.forEach(a => {
        const r=a.row;
        const out=ws.addRow({
          '原始列號':a.sourceRow,'案例':a.caseId,'年度':r['年度']||'','月份':r['月份']||r['輔導月份']||'',
          '學員姓名':r['學員姓名']||'','導師姓名':r['導師姓名']||'','原始陪伴訓練紀錄':r['陪伴訓練紀錄']||'',
          '明顯錯字提醒':a.typoAlert||'','問題類型':a.textType||'','是否建議退回':a.needsApproval?'是':'否',
          '為什麼需要處理':a.textReason||'','建議教師修改方向':a.textSuggestion||'',
          '管理確認':a.management?'是':'否','管理議題':a.managementIssue||'','管理待確認事項':a.managementQuestion||'',
          '教學亮點候選':a.highlight?'是':'否','教學亮點類型':a.highlightType||'','備註':''
        });
        out.eachCell({includeEmpty:true}, c => { c.alignment={vertical:'top',wrapText:true}; c.border=border; });
        out.height=72;
      });

      const first=parsed.rows[0].data;
      const y=String(first['年度']||'').replace(/年$/,'');
      const m=String(first['月份']||first['輔導月份']||'').replace(/月份$/,'').replace(/月$/,'');
      const label=`${y?y+'年':''}${m?m+'月':''}` || '本月';
      await download(wb, `${safe(label)}_新手導師計畫_細緻分析.xlsx`);
      if (status) { status.className='status ok'; status.textContent='細緻分析版已產生。'; }
    } catch (err) {
      console.error(err);
      if (status) { status.className='status warn'; status.textContent=`產生細緻分析版失敗：${err.message||err}`; }
    }
  }

  // 以正式版細緻輸出取代原本含驗證欄位的下載。
  detailBtn.addEventListener('click', e => {
    e.preventDefault();
    e.stopImmediatePropagation();
    makePublicDetailed();
  }, true);
})();
