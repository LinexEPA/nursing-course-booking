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

  function addLearningRecordSuggestions(out, text) {
    const suggestions = [];
    const vagueLearning = text.match(/偶爾還是會漏(?=[，。；、\s]|$)|偶爾仍會漏(?=[，。；、\s]|$)|(?:提醒後|指導後)[，、\s]*(?:目前可以|已有改善)(?=[，。；、\s]|$)/);
    if (vagueLearning) {
      suggestions.push(`學習歷程｜「${vagueLearning[0]}」請補充學員在哪一步遇到困難，以及提醒後能做到哪一步。若尚未再觀察，可寫下次要確認的項目。`);
    }

    const vagueGuidance = text.match(/(?:^|[，。；、\s])(?:導師|老師|教師)?(?:已|有再|再次)(?:提醒|指導|衛教)(?=[，。；、\s]|$)/);
    const concreteGuidance = /(?:導師|老師|教師)[^，。；\n]{0,30}(?:示範|提問|陪同|引導|(?:用|依照|對照).{1,12}(?:提示|提醒))/.test(text);
    if (vagueGuidance && !concreteGuidance) {
      const phrase = vagueGuidance[0].replace(/^[，。；、\s]+/, '');
      suggestions.push(`教師輔導｜「${phrase}」請寫明提醒的內容，或補充你如何示範、提問或引導學員練習。`);
    }

    if (!suggestions.length) return;
    out.textIssue = true;
    out.textType ||= '輔導紀錄建議補充';
    out.textReason ||= '紀錄中的學員表現或教師輔導方式可再寫具體。';
    out.textSuggestion = [out.textSuggestion, ...suggestions].filter(Boolean).join('\n');
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

    const continueMentoring = txt(row?.['下個月是否持續輔導']).trim();
    const hasStopReason = Boolean(txt(row?.['停止輔導原因']).trim());
    let followUpWarning = '';
    if (continueMentoring === '是' && hasStopReason) {
      followUpWarning = '已選擇持續輔導，但有填寫停止輔導原因，請確認下月是否續輔。';
    } else if (continueMentoring === '否' && !hasStopReason) {
      followUpWarning = '已選擇不再持續輔導，請補填停止輔導原因。';
    }
    if (followUpWarning) {
      issues.push('續輔與停止原因不一致');
      out.managementState ||= '續輔資料需確認';
      out.managementQuestion = [out.managementQuestion, followUpWarning].filter(Boolean).join(' ');
    }

    issues = unique(issues);
    out.managementIssue = issues.join('、');
    out.management = issues.length > 0;
    if (!out.management) {
      out.managementState = '';
      out.managementQuestion = '';
    }

    addLearningRecordSuggestions(out, t);
    out.confidence = out.textIssue || out.management ? (out.needsApproval ? '高' : '中') : '高';
    return out;
  };

  M.RULE_VERSION = 'v1.8';
})();
