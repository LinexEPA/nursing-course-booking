(() => {
  'use strict';
  const M = window.MentorRules;
  if (!M || typeof M.analyzeRecord !== 'function') return;

  const baseAnalyze = M.analyzeRecord;
  const txt = v => (v == null ? '' : String(v));
  const unique = xs => [...new Set(xs.filter(Boolean))];

  function confirmedSafetyEvent(t) {
    return /(?:已|曾|近期|當下|此次|本次).{0,10}(?:發生|出現).{0,8}(?:給藥錯誤|病安事件|異常事件|跌倒|針扎|針刺)|(?:發生|出現).{0,8}(?:給藥錯誤|病安事件|異常事件|跌倒|針扎|針刺)|(?:給藥錯誤|病安事件|異常事件).{0,12}(?:已通報|通報)|已通報|事件通報|near\s*miss|給錯|誤給|錯給|漏給|誤拔|放錯.{0,8}(?:藥|床|病人)|造成.{0,12}(?:傷害|影響)|病人.{0,8}(?:受傷|傷害)/i.test(t);
  }

  function removeSafetyManagement(out) {
    let issues = out.managementIssue ? out.managementIssue.split('、').filter(Boolean) : [];
    issues = issues.filter(x => !/給藥／病安事件|病人安全事件|侵入性處置事件/.test(x));
    out.managementIssue = unique(issues).join('、');
    out.management = issues.length > 0;
    if (!out.management) {
      out.managementState = '';
      out.managementQuestion = '';
    }
  }

  function removeGapOnly(out) {
    const hits = (out.ruleHits || '').split('、').filter(Boolean).filter(x => x !== 'T06');
    out.ruleHits = hits.join('、');
    if (out.textType === '關鍵內容不足') {
      out.textIssue = false;
      out.textType = '';
      out.textReason = '';
      out.textSuggestion = '';
      out.needsApproval = false;
    }
  }

  M.analyzeRecord = function(row, sourceRow) {
    const out = baseAnalyze(row, sourceRow);
    const t = txt(row?.['陪伴訓練紀錄']).trim();
    const confirmed = confirmedSafetyEvent(t);

    // 「防止給藥錯誤」「避免影響病人安全」「如同針扎」等預防性教學，不代表事件已發生。
    if (!confirmed) {
      removeSafetyManagement(out);
      removeGapOnly(out);
    }

    let issues = out.managementIssue ? out.managementIssue.split('、').filter(Boolean) : [];

    // 重要流程反覆疏漏：可管理追蹤，但不是病安事件。
    if (/偶有.{0,8}疏漏.{0,12}藥物對點|反覆.{0,10}(?:漏|遺漏).{0,12}(?:藥物對點|交班|查核)|藥物對點.{0,12}(?:疏漏|遺漏).{0,20}(?:下一班|交接班)/.test(t)) {
      issues.push('重要工作流程／交班品質需追蹤');
      out.managementState = '紀錄顯示重要流程有反覆或偶發疏漏';
      out.managementQuestion = '確認藥物對點／交班流程改善情形，以及是否仍需持續追蹤。';
    }

    // 臨床評估警覺性不足且已有具體漏察與持續關注訊號：管理追蹤，但不直接稱病安事件。
    if (/尿液顏色異常未及時發現|異常.{0,12}未及時發現/.test(t) && /需多加關注|進步幅度不大|持續關注|持續指導/.test(t)) {
      issues.push('臨床評估／警覺性需追蹤');
      out.managementState = '已有具體臨床評估漏察，且紀錄顯示仍需持續關注';
      out.managementQuestion = '確認臨床評估與異常辨識能力的改善情形，以及下月是否仍需追蹤。';
    }

    issues = unique(issues);
    out.managementIssue = issues.join('、');
    out.management = issues.length > 0;
    if (!out.management) {
      out.managementState = '';
      out.managementQuestion = '';
    }

    out.confidence = out.textIssue || out.management ? (out.needsApproval ? '高' : '中') : '高';
    return out;
  };

  M.RULE_VERSION = 'v1.5';
})();
