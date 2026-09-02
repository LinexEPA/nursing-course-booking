(() => {
  'use strict';
  const M = window.MentorRules;
  if (!M || typeof M.analyzeRecord !== 'function') return;

  const baseAnalyze = M.analyzeRecord;
  const text = v => (v == null ? '' : String(v));
  const unique = xs => [...new Set(xs.filter(Boolean))];

  function addRuleHit(out, id) {
    out.ruleHits = unique([...(out.ruleHits ? out.ruleHits.split('、') : []), id]).join('、');
  }

  function addTextIssue(out, type, reason, suggestion, ruleId) {
    if (!out.textIssue) {
      out.textIssue = true;
      out.textType = type;
      out.textReason = reason;
      out.textSuggestion = suggestion;
    } else if (!String(out.textType || '').includes(type)) {
      out.textType = `${out.textType}＋${type}`;
      out.textReason = [out.textReason, reason].filter(Boolean).join('；');
      out.textSuggestion = [out.textSuggestion, suggestion].filter(Boolean).join('；');
    }
    out.needsApproval = true;
    addRuleHit(out, ruleId);
  }

  function actualManagerIntervention(t) {
    return /已.{0,4}(與|向)(護理長|阿長|主管).{0,8}(會談|討論|回報|說明)|與(護理長|阿長|主管).{0,8}(會談|討論)後|(護理長|阿長|主管).{0,6}(已|有).{0,8}(介入|處理|調整|知悉)|經(護理長|阿長|主管).{0,8}(處理|調整|介入)|調整.{0,8}班別|異動班別/.test(t);
  }

  function negativeAdaptation(t) {
    return /無法適應|適應不良|適應困難|難以適應|跟不上|適應.{0,8}(困難|不佳|不穩)|無法順利完成班內工作/.test(t);
  }

  function emotionAffectsWork(t) {
    return /影響工作|影響臨床|無法完成|無法執行|無法上班|臨床照護.{0,10}(混亂|忘記)|工作.{0,8}(混亂|延誤|中斷)|忘記流程/.test(t);
  }

  function frequentTardiness(t) {
    return /遲到.{0,10}(頻繁|多次|反覆|經常|常常)|(?:頻繁|多次|反覆|經常|常常).{0,10}遲到/.test(t);
  }

  // 115/06 人工基準鎖定：這些已確認為正式紀錄中的明顯錯字，後續改規則不可遺失。
  function extraTypos(t) {
    const out = [];
    if (t.includes('給於')) out.push('給於 → 給予');
    if (t.includes('不錯得方法')) out.push('不錯得方法 → 不錯的方法');
    if (/白斑.{0,8}(適應|工作|訓練|班)/.test(t)) out.push('白斑 → 白班（依上下文）');
    if (t.includes('資深學會')) out.push('資深學會 → 請確認是否為「資深學姊／資深同仁」');
    if (/較容易謊|容易謊/.test(t)) out.push('容易謊 → 容易慌（依上下文）');
    if (t.includes('在過作過程中')) out.push('在過作過程中 → 在工作過程中');
    return out;
  }

  M.analyzeRecord = function(row, sourceRow) {
    const out = baseAnalyze(row, sourceRow);
    const t = text(row?.['陪伴訓練紀錄']).trim();
    let issues = out.managementIssue ? out.managementIssue.split('、').filter(Boolean) : [];
    const managerActual = actualManagerIntervention(t);

    // 單純提到「阿長／主管可以協助」不等於主管已介入。
    if (!managerActual) issues = issues.filter(x => x !== '主管／管理介入');

    // 「工作適應穩定／良好」不應因出現「工作適應」字樣而被列為管理問題。
    if (!negativeAdaptation(t)) issues = issues.filter(x => x !== '工作適應／教學支持');

    // 情緒或壓力只有在已影響工作，或確有主管介入時，才獨立列為管理確認。
    if (!emotionAffectsWork(t) && !managerActual) issues = issues.filter(x => x !== '情緒／壓力影響工作');

    // 頻繁或反覆遲到屬出勤管理，不只是一般輔導內容。
    if (frequentTardiness(t)) issues.push('出勤／遲到');
    issues = unique(issues);

    out.management = issues.length > 0;
    out.managementIssue = issues.join('、');
    if (!out.management) {
      out.managementState = '';
      out.managementQuestion = '';
    } else if (issues.length === 1 && issues[0] === '出勤／遲到') {
      out.managementState = '紀錄顯示反覆出勤問題';
      out.managementQuestion = '確認遲到改善情形、是否影響單位運作及後續追蹤。';
    } else if (!managerActual && out.managementState === '已見主管／制度處理訊號' && !/事件檢討|PDCA|已通報|事件通報|已處理|已調整|已介入/.test(t)) {
      out.managementState = '現有紀錄無法確認主管／制度是否已處理';
    }

    // 六月人工基準已確認的明顯錯字：正式紀錄需退件修正。
    const typos = extraTypos(t);
    if (typos.length) {
      out.typoAlert = unique([...(out.typoAlert ? out.typoAlert.split('；') : []), ...typos]).join('；');
      addTextIssue(
        out,
        '明顯錯字',
        `偵測到正式紀錄中的明顯錯字或疑似誤植：${typos.join('；')}`,
        '建議退件｜修正明顯錯字後再送出',
        'T01B'
      );
    }

    // 有具體困難描述可以保留，但貶抑性比喻本身不適合作為正式紀錄。
    if (/無頭蒼蠅|漫無章法/.test(t)) {
      addTextIssue(
        out,
        '正式紀錄表述風險',
        '出現貶抑性或形象化比喻；即使前文已有具體事實，正式紀錄仍宜改為可觀察的工作排序、延遲或遺漏描述。',
        '建議退件｜保留具體事實，移除貶抑性比喻',
        'T07'
      );
    }

    // 正向親暱或口語肯定（如「寶貝學妹、棒棒的學妹」）若已有具體事實支持，不因語氣本身退件。

    out.confidence = out.textIssue || out.management ? (out.needsApproval ? '高' : '中') : '高';
    return out;
  };

  M.RULE_VERSION = 'v1.4';
})();
