(() => {
  'use strict';

  const $ = (id) => document.getElementById(id);
  const els = {
    input: $('fileInput'), drop: $('dropZone'), meta: $('fileMeta'), status: $('status'),
    detail: $('detailBtn'), supervisor: $('supervisorBtn'),
    total: $('mTotal'), text: $('mText'), approval: $('mApproval'), mgmt: $('mMgmt'), highlight: $('mHighlight')
  };

  const state = { fileName: '', rows: [], analyses: [], year: '', month: '' };
  const REQUIRED = ['學員姓名', '導師姓名', '陪伴訓練紀錄'];

  function setStatus(message, kind = '') {
    els.status.className = `status ${kind}`.trim();
    els.status.textContent = message;
  }

  function cellText(v) {
    if (v == null) return '';
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    if (typeof v === 'object') {
      if (Array.isArray(v.richText)) return v.richText.map(x => x.text || '').join('');
      if ('result' in v) return cellText(v.result);
      if ('text' in v) return String(v.text ?? '');
      if ('hyperlink' in v) return String(v.text ?? v.hyperlink ?? '');
    }
    return String(v).trim();
  }

  function normalizeHeader(v) {
    return cellText(v).replace(/\s+/g, '').replace(/／/g, '/').trim();
  }

  function sourceLabel() {
    const y = state.year || '';
    const m = String(state.month || '').replace(/月份$/, '月');
    return `${y ? `${y}年` : ''}${m}` || '本月';
  }

  async function loadFile(file) {
    if (!file) return;
    if (!/\.xlsx?$/i.test(file.name)) {
      setStatus('請選擇 Excel（.xlsx 或 .xls）檔案。', 'warn');
      return;
    }
    try {
      setStatus('正在讀取 Excel…');
      const buf = await file.arrayBuffer();
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buf);
      const ws = wb.worksheets[0];
      if (!ws) throw new Error('找不到工作表。');

      const headerRow = ws.getRow(1);
      const headers = [];
      headerRow.eachCell({ includeEmpty: true }, (cell, col) => { headers[col] = normalizeHeader(cell.value); });
      const headerSet = new Set(headers.filter(Boolean));
      const missing = REQUIRED.filter(x => !headerSet.has(normalizeHeader(x)));
      if (missing.length) throw new Error(`來源格式不符，缺少欄位：${missing.join('、')}`);

      const rows = [];
      for (let r = 2; r <= ws.rowCount; r++) {
        const rowObj = {};
        const row = ws.getRow(r);
        headers.forEach((h, c) => { if (h) rowObj[h] = cellText(row.getCell(c).value); });
        const learner = rowObj[normalizeHeader('學員姓名')] || '';
        const tutor = rowObj[normalizeHeader('導師姓名')] || '';
        if (!learner && !tutor) continue;
        rows.push({ sourceRow: r, data: rowObj });
      }
      if (!rows.length) throw new Error('沒有讀到可分析的輔導紀錄。');

      state.fileName = file.name;
      state.rows = rows;
      state.year = rows[0].data[normalizeHeader('年度')] || '';
      state.month = rows[0].data[normalizeHeader('月份')] || rows[0].data[normalizeHeader('輔導月份')] || '';
      state.analyses = rows.map(x => MentorRules.analyzeRecord(x.data, x.sourceRow));

      updateSummary();
      els.meta.textContent = `${file.name}｜${rows.length} 筆紀錄`;
      els.detail.disabled = false;
      els.supervisor.disabled = false;
      setStatus(`分析完成：已用固定規則 ${MentorRules.RULE_VERSION} 完成第一輪分類。語意不確定案件保留給人工／Codex 後續比對。`, 'ok');
    } catch (err) {
      console.error(err);
      state.rows = []; state.analyses = [];
      els.detail.disabled = true; els.supervisor.disabled = true;
      setStatus(`讀取失敗：${err.message || err}`, 'warn');
    }
  }

  function updateSummary() {
    const a = state.analyses;
    const counts = {
      total: a.length,
      text: a.filter(x => x.textIssue).length,
      approval: a.filter(x => x.needsApproval).length,
      mgmt: a.filter(x => x.management).length,
      highlight: a.filter(x => x.highlight).length
    };
    els.total.textContent = counts.total;
    els.text.textContent = counts.text;
    els.approval.textContent = counts.approval;
    els.mgmt.textContent = counts.mgmt;
    els.highlight.textContent = counts.highlight;
    return counts;
  }

  function safeFilePart(v) {
    return String(v || '').replace(/[\\/:*?"<>|]/g, '_').replace(/\s+/g, '').slice(0, 50);
  }

  async function saveWorkbook(wb, filename) {
    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  }

  function headerStyle(cell) {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0D8F98' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFD5E0E6' } }, bottom: { style: 'thin', color: { argb: 'FFD5E0E6' } },
      left: { style: 'thin', color: { argb: 'FFD5E0E6' } }, right: { style: 'thin', color: { argb: 'FFD5E0E6' } }
    };
  }

  function bodyStyle(row) {
    row.alignment = { vertical: 'top', wrapText: true };
    row.eachCell({ includeEmpty: true }, cell => {
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFE2E8EC' } }, bottom: { style: 'thin', color: { argb: 'FFE2E8EC' } },
        left: { style: 'thin', color: { argb: 'FFE2E8EC' } }, right: { style: 'thin', color: { argb: 'FFE2E8EC' } }
      };
    });
  }

  async function makeDetailed() {
    try {
      setStatus('正在產生細緻分析版…');
      const wb = new ExcelJS.Workbook();
      wb.creator = 'Mentor Report Analyzer';
      const ws = wb.addWorksheet('細緻分析');
      const cols = [
        ['原始列號', 10], ['案例', 10], ['年度', 8], ['月份', 10], ['學員姓名', 12], ['導師姓名', 12],
        ['原始陪伴訓練紀錄', 46], ['清理後紀錄', 46], ['文字/內容案件', 12], ['問題類型', 18],
        ['是否需主管審批', 14], ['為什麼需要處理', 32], ['護理部初審建議', 28], ['管理確認', 12],
        ['管理議題', 24], ['管理待確認事項', 34], ['教學亮點候選', 14], ['教學亮點類型', 26],
        ['規則命中', 16], ['規則信心', 10], ['Codex判定', 18], ['最終人工判定', 18], ['備註', 24]
      ];
      ws.columns = cols.map(([header, width]) => ({ header, key: header, width }));
      ws.getRow(1).height = 32;
      ws.getRow(1).eachCell(headerStyle);
      ws.views = [{ state: 'frozen', ySplit: 1 }];
      ws.autoFilter = { from: 'A1', to: 'W1' };

      for (const a of state.analyses) {
        const r = a.row;
        const out = ws.addRow({
          '原始列號': a.sourceRow, '案例': a.caseId, '年度': r['年度'] || '', '月份': r['月份'] || r['輔導月份'] || '',
          '學員姓名': r['學員姓名'] || '', '導師姓名': r['導師姓名'] || '',
          '原始陪伴訓練紀錄': r['陪伴訓練紀錄'] || '', '清理後紀錄': a.cleanedText,
          '文字/內容案件': a.textIssue ? '是' : '否', '問題類型': a.textType,
          '是否需主管審批': a.needsApproval ? '是' : '否', '為什麼需要處理': a.textReason,
          '護理部初審建議': a.textSuggestion, '管理確認': a.management ? '是' : '否',
          '管理議題': a.managementIssue, '管理待確認事項': a.managementQuestion,
          '教學亮點候選': a.highlight ? '是' : '否', '教學亮點類型': a.highlightType,
          '規則命中': a.ruleHits, '規則信心': a.confidence, 'Codex判定': '', '最終人工判定': '', '備註': ''
        });
        bodyStyle(out); out.height = 72;
      }

      const info = wb.addWorksheet('驗證說明');
      info.columns = [{ width: 24 }, { width: 90 }];
      info.addRows([
        ['規則版本', MentorRules.RULE_VERSION],
        ['分析方式', '第一版僅使用固定 JavaScript 規則，不呼叫 AI。'],
        ['驗證目的', '後續將「規則判定、Codex判定、最終人工判定」逐筆比較，以評估一致率、誤抓與漏抓。'],
        ['使用原則', '規則高信心不代表自動退件；最終退件／管理決策仍由人工確認。'],
        ['來源檔', state.fileName]
      ]);
      info.getColumn(1).font = { bold: true, color: { argb: 'FF203B63' } };
      info.eachRow(r => { r.alignment = { vertical: 'top', wrapText: true }; r.height = 28; });

      const prefix = safeFilePart(sourceLabel()) || '本月';
      await saveWorkbook(wb, `${prefix}_新手導師計畫_細緻分析.xlsx`);
      setStatus('細緻分析版已產生。', 'ok');
    } catch (err) {
      console.error(err); setStatus(`產生細緻分析版失敗：${err.message || err}`, 'warn');
    }
  }

  function cloneStyle(src, dst) {
    dst.style = JSON.parse(JSON.stringify(src.style || {}));
    dst.numFmt = src.numFmt;
  }

  function writeTemplateRows(ws, startRow, rows, colCount) {
    const styleRow = ws.getRow(startRow);
    const rowHeight = styleRow.height || 45;
    const styleCells = [];
    for (let c = 1; c <= colCount; c++) styleCells[c] = styleRow.getCell(c);

    const clearTo = Math.max(ws.rowCount, startRow + rows.length + 3);
    for (let r = startRow; r <= clearTo; r++) {
      const target = ws.getRow(r);
      for (let c = 1; c <= colCount; c++) target.getCell(c).value = null;
    }

    rows.forEach((vals, i) => {
      const row = ws.getRow(startRow + i);
      row.height = rowHeight;
      for (let c = 1; c <= colCount; c++) {
        const cell = row.getCell(c);
        cell.value = vals[c - 1] ?? null;
        cloneStyle(styleCells[c], cell);
        cell.alignment = { ...(cell.alignment || {}), vertical: 'top', wrapText: true };
      }
    });
  }

  function categoryCounts(analyses) {
    const count = type => analyses.filter(x => x.textType === type).length;
    return {
      low: analyses.filter(x => x.textIssue && !x.needsApproval).length,
      ambiguous: count('語意不明／可能相反'),
      risk: count('紀錄表述風險'),
      gap: count('關鍵內容不足'),
      contradiction: count('前後敘述待確認')
    };
  }

  async function makeSupervisor() {
    try {
      setStatus('正在套用主管審閱範本…');
      const response = await fetch('./templates/supervisor-template.xlsx', { cache: 'no-store' });
      if (!response.ok) throw new Error('找不到主管審閱範本。');
      const templateBuffer = await response.arrayBuffer();
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(templateBuffer);

      const monthly = wb.getWorksheet('主管月報');
      const reject = wb.getWorksheet('退件審批');
      const mgmt = wb.getWorksheet('管理確認');
      const highlight = wb.getWorksheet('教學亮點附件');
      if (!monthly || !reject || !mgmt || !highlight) throw new Error('主管範本工作表不完整。');

      const counts = updateSummary();
      const textCases = state.analyses.filter(x => x.textIssue);
      const approvalCases = state.analyses.filter(x => x.needsApproval);
      const mgmtCases = state.analyses.filter(x => x.management);
      const highlightCases = state.analyses.filter(x => x.highlight);
      const cats = categoryCounts(state.analyses);

      monthly.getCell('A1').value = `${sourceLabel()}｜新手導師計畫 主管審閱版`;
      monthly.getCell('A5').value = counts.total;
      monthly.getCell('C5').value = counts.text;
      monthly.getCell('E5').value = counts.text - counts.approval;
      monthly.getCell('G5').value = counts.approval;
      monthly.getCell('I5').value = counts.mgmt;
      monthly.getCell('K5').value = counts.highlight;
      monthly.getCell('B10').value = cats.low;
      monthly.getCell('B11').value = cats.ambiguous;
      monthly.getCell('B12').value = cats.risk;
      monthly.getCell('B13').value = cats.gap;
      monthly.getCell('B14').value = cats.contradiction;
      monthly.getCell('C11').value = approvalCases.filter(x => x.textType === '語意不明／可能相反').map(x => x.caseId).join('、') || '本月無';
      monthly.getCell('C12').value = approvalCases.filter(x => x.textType === '紀錄表述風險').map(x => x.caseId).join('、') || '本月無';
      monthly.getCell('C13').value = approvalCases.filter(x => x.textType === '關鍵內容不足').map(x => x.caseId).join('、') || '本月無';
      monthly.getCell('C14').value = approvalCases.filter(x => x.textType === '前後敘述待確認').map(x => x.caseId).join('、') || '本月無';
      monthly.getCell('G11').value = `只審批「需主管審批＝是」的${counts.approval}筆`;
      monthly.getCell('B17').value = `到「退件審批」篩選是否需主管審批＝是，共${counts.approval}筆；主管決定退件或不退件。`;
      monthly.getCell('B18').value = `到「管理確認」查看${counts.mgmt}件案件；確認主管知悉、制度處理及是否需要下月追蹤。`;
      monthly.getCell('B19').value = `到「教學亮點附件」查看${counts.highlight}件候選；主管選擇適合的正向回饋方式。`;

      reject.getCell('A2').value = `主管只需處理「是否需主管審批＝是」的${counts.approval}筆。單純明顯錯字原則上直接修正，不作退件理由。`;
      const rejectRows = textCases.map(a => [
        a.row['單位'] || '待補', a.caseId, a.row['學員姓名'] || '', a.row['導師姓名'] || '', a.textType,
        a.excerpt, a.textReason, a.textSuggestion, a.needsApproval ? '是' : '否', a.management ? '是' : '否', '', ''
      ]);
      writeTemplateRows(reject, 5, rejectRows, 12);

      const mgmtRows = mgmtCases.map(a => [
        a.row['單位'] || '待補', a.caseId, a.row['學員姓名'] || '', a.row['導師姓名'] || '', a.managementIssue,
        a.managementState, a.managementQuestion, '', '', ''
      ]);
      writeTemplateRows(mgmt, 5, mgmtRows, 10);

      const highlightRows = highlightCases.map(a => [
        a.row['單位'] || '待補', a.caseId, a.row['學員姓名'] || '', a.row['導師姓名'] || '', a.highlightType,
        a.excerpt, a.highlightReason, a.highlightFeedback, '', ''
      ]);
      writeTemplateRows(highlight, 5, highlightRows, 10);

      const prefix = safeFilePart(sourceLabel()) || '本月';
      await saveWorkbook(wb, `${prefix}_新手導師計畫_主管審閱.xlsx`);
      setStatus('主管審閱版已依固定範本產生。', 'ok');
    } catch (err) {
      console.error(err); setStatus(`產生主管審閱版失敗：${err.message || err}`, 'warn');
    }
  }

  els.input.addEventListener('change', e => loadFile(e.target.files?.[0]));
  els.detail.addEventListener('click', makeDetailed);
  els.supervisor.addEventListener('click', makeSupervisor);

  ['dragenter', 'dragover'].forEach(evt => els.drop.addEventListener(evt, e => { e.preventDefault(); els.drop.classList.add('drag'); }));
  ['dragleave', 'drop'].forEach(evt => els.drop.addEventListener(evt, e => { e.preventDefault(); els.drop.classList.remove('drag'); }));
  els.drop.addEventListener('drop', e => loadFile(e.dataTransfer?.files?.[0]));
})();
