(() => {
  'use strict';

  const answerLike = /答案|作答|選項|正確答案|標準答案|測驗|題目|分數|得分|評分|分數欄|是否正確|正解/i;
  const metaLike = /時間戳記|提交時間|開始時間|完成時間|電子郵件|email|帳號|流水號|編號|id$/i;

  const aliases = {
    student: [
      /^學員姓名$/i, /^學員$/i, /新進.*姓名/i, /新人.*姓名/i, /受輔導.*姓名/i,
      /受訓.*姓名/i, /學習者.*姓名/i, /姓名.*學員/i, /學員.*姓名/i
    ],
    mentor: [
      /^導師姓名$/i, /^導師$/i, /臨床教師.*姓名/i, /^臨床教師$/i, /輔導教師.*姓名/i,
      /輔導者.*姓名/i, /教師姓名/i, /指導者.*姓名/i
    ],
    record: [
      /^陪伴訓練紀錄$/i, /陪伴.*紀錄/i, /輔導.*紀錄/i, /輔導.*內容/i, /本月.*輔導/i,
      /導師.*紀錄/i, /紀錄.*內容/i, /輔導.*摘要/i, /訓練.*紀錄/i, /輔導.*情形/i,
      /本月.*紀錄/i, /學習.*狀況.*紀錄/i
    ],
    unit: [/^單位$/i, /單位名稱/i, /服務單位/i, /所屬單位/i, /部門/i, /科別/i, /病房/i],
    year: [/^年度$/i, /輔導年度/i, /年份/i],
    month: [/^月份$/i, /輔導月份/i, /月份.*輔導/i, /年月/i]
  };

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

  const headerText = v => cellText(v).replace(/\s+/g, '').replace(/／/g, '/').trim();

  function matchHeader(headers, patterns) {
    for (let i = 1; i < headers.length; i++) {
      const value = headers[i] || '';
      if (patterns.some(rx => rx.test(value))) return i;
    }
    return -1;
  }

  function columnSamples(ws, col, headerRow, limit = 80) {
    const out = [];
    const end = Math.min(ws.rowCount, headerRow + limit);
    for (let r = headerRow + 1; r <= end; r++) {
      const value = cellText(ws.getRow(r).getCell(col).value);
      if (value) out.push(value);
    }
    return out;
  }

  function avgTextLength(values) {
    if (!values.length) return 0;
    return values.reduce((sum, value) => sum + value.length, 0) / values.length;
  }

  function detectRecordColumn(ws, headers, headerRow) {
    const exact = matchHeader(headers, aliases.record);
    if (exact > 0 && !answerLike.test(headers[exact])) {
      return { col: exact, reason: '欄位名稱辨識' };
    }

    const candidates = [];
    for (let c = 1; c < headers.length; c++) {
      const header = headers[c] || '';
      if (!header || answerLike.test(header) || metaLike.test(header)) continue;
      if ([...aliases.student, ...aliases.mentor, ...aliases.unit, ...aliases.year, ...aliases.month].some(rx => rx.test(header))) continue;

      const values = columnSamples(ws, c, headerRow);
      if (values.length < 1) continue;
      const avg = avgTextLength(values);
      const longRatio = values.filter(v => v.length >= 30).length / values.length;
      if (avg >= 35 || longRatio >= 0.45) {
        candidates.push({ col: c, header, avg, longRatio });
      }
    }

    candidates.sort((a, b) => (b.avg + b.longRatio * 40) - (a.avg + a.longRatio * 40));
    if (!candidates.length) return { col: -1, reason: '' };

    // 若前兩欄非常接近，仍優先選含「紀錄／輔導／陪伴／訓練」語意者。
    const preferred = candidates.find(x => /紀錄|輔導|陪伴|訓練|學習狀況/i.test(x.header));
    if (preferred) return { col: preferred.col, reason: `依內容辨識：${preferred.header}` };

    return { col: candidates[0].col, reason: `依長文字內容辨識：${candidates[0].header}` };
  }

  function findHeaderRow(ws) {
    const max = Math.min(ws.rowCount, 15);
    let best = null;
    for (let r = 1; r <= max; r++) {
      const headers = [];
      ws.getRow(r).eachCell({ includeEmpty: true }, (cell, col) => { headers[col] = headerText(cell.value); });
      const filled = headers.filter(Boolean).length;
      if (filled < 3) continue;

      const student = matchHeader(headers, aliases.student);
      const mentor = matchHeader(headers, aliases.mentor);
      const record = matchHeader(headers, aliases.record);
      let score = 0;
      if (student > 0) score += 3;
      if (mentor > 0) score += 3;
      if (record > 0) score += 5;
      if (matchHeader(headers, aliases.unit) > 0) score += 1;
      if (filled >= 5) score += 1;

      if (!best || score > best.score) best = { row: r, headers, score };
    }
    return best;
  }

  function detectSheet(workbook) {
    let best = null;
    for (const ws of workbook.worksheets || []) {
      const head = findHeaderRow(ws);
      if (!head) continue;
      const recordInfo = detectRecordColumn(ws, head.headers, head.row);
      const student = matchHeader(head.headers, aliases.student);
      const mentor = matchHeader(head.headers, aliases.mentor);

      let score = head.score;
      if (recordInfo.col > 0) score += 5;
      if (student > 0 || mentor > 0) score += 2;
      if (!best || score > best.score) best = { ws, ...head, recordInfo, student, mentor, score };
    }
    return best;
  }

  function parseWorkbook(workbook) {
    const detected = detectSheet(workbook);
    if (!detected) throw new Error('找不到可辨識的資料表。');

    const { ws, row: headerRow, headers, recordInfo } = detected;
    const studentCol = matchHeader(headers, aliases.student);
    const mentorCol = matchHeader(headers, aliases.mentor);
    const unitCol = matchHeader(headers, aliases.unit);
    const yearCol = matchHeader(headers, aliases.year);
    const monthCol = matchHeader(headers, aliases.month);
    const recordCol = recordInfo.col;

    if (recordCol < 1) {
      throw new Error('找不到「輔導紀錄」內容欄。請確認檔案中有一欄是每位學員的長文字輔導紀錄。');
    }

    // 至少要能辨識學員或導師其中一項；若都沒有，仍以有輔導紀錄的列為資料列。
    const rows = [];
    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
      const record = cellText(ws.getRow(r).getCell(recordCol).value);
      const student = studentCol > 0 ? cellText(ws.getRow(r).getCell(studentCol).value) : '';
      const mentor = mentorCol > 0 ? cellText(ws.getRow(r).getCell(mentorCol).value) : '';
      if (!record && !student && !mentor) continue;
      if (!record && !(student || mentor)) continue;

      const data = {};
      // 保留原始欄位，讓既有管理規則仍可讀取「停止輔導原因／連續請假天數」等欄位。
      headers.forEach((name, col) => {
        if (!name) return;
        data[name] = cellText(ws.getRow(r).getCell(col).value);
      });

      data['學員姓名'] = student;
      data['導師姓名'] = mentor;
      data['陪伴訓練紀錄'] = record;
      if (unitCol > 0) data['單位'] = cellText(ws.getRow(r).getCell(unitCol).value);
      if (yearCol > 0) data['年度'] = cellText(ws.getRow(r).getCell(yearCol).value);
      if (monthCol > 0) data['月份'] = cellText(ws.getRow(r).getCell(monthCol).value);

      rows.push({ sourceRow: r, data });
    }

    if (!rows.length) throw new Error('沒有讀到可分析的輔導紀錄。');

    const mapping = {
      worksheet: ws.name,
      headerRow,
      student: studentCol > 0 ? headers[studentCol] : '',
      mentor: mentorCol > 0 ? headers[mentorCol] : '',
      record: headers[recordCol] || '',
      unit: unitCol > 0 ? headers[unitCol] : '',
      year: yearCol > 0 ? headers[yearCol] : '',
      month: monthCol > 0 ? headers[monthCol] : '',
      recordReason: recordInfo.reason || ''
    };

    return { rows, mapping };
  }

  window.MentorInputAdapter = { parseWorkbook, cellText, headerText };
})();
