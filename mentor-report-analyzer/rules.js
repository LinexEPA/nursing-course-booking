(() => {
  'use strict';

  const RULE_VERSION = 'v1.1';
  const textField = '陪伴訓練紀錄';
  const rx = (source) => new RegExp(source, 'i');
  const has = (text, source) => rx(source).test(text || '');
  const norm = (v) => (v == null ? '' : String(v)).trim();
  const clean = (v) => norm(v).replace(/_x000D_/gi, '').replace(/\r\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
  const clip = (s, n = 180) => { s = clean(s); return s.length > n ? `${s.slice(0, n)}…` : s; };

  const TYPO_PAIRS = [
    ['到時侯','到時候'], ['學院','學員'], ['提除','提出'], ['的的','的'],
    ['一起病人安全事件','一起病人安全事件（請確認原意）'], ['在多點細心','再多點細心'], ['在跟下一班','再跟下一班']
  ];

  function typoFindings(text){
    const out=[];
    for(const [wrong,right] of TYPO_PAIRS){
      if((text||'').includes(wrong)) out.push(`${wrong} → ${right}`);
    }
    return out;
  }

  function suspiciousLabel(text){
    const strong = has(text,'耍小聰明|懶惰|不用心|散漫|態度很差|能力很差|沒有責任感|不負責任');
    const soft = has(text,'心不在焉|注意力不夠集中|注意力不集中|學習態度被動');
    if(!strong && !soft) return false;
    const behaviorEvidence = has(text,'導致|因此|例如|曾|發生|遺漏|漏做|漏寫|錯誤|未完成|需.{0,8}提醒|經.{0,12}提醒|才能.{0,8}(發現|修正)|具體|工作事項');
    const improvementPlan = has(text,'建議|加強|工作清單|查核|自我檢核|後續|改善|提醒|規劃|重點確認');
    if(soft && behaviorEvidence && improvementPlan) return false;
    if(strong && behaviorEvidence && improvementPlan && !has(text,'懶惰|耍小聰明|沒有責任感|不負責任')) return false;
    return true;
  }

  const TEXT_RULES = [
    {
      id:'T01', type:'明顯錯字', level:'high', approval:true,
      test:t => typoFindings(t).length>0,
      reason:'偵測到正式紀錄中的明顯錯字或誤植；正式送出紀錄應修正。',
      suggestion:'建議退件｜修正明顯錯字後再送出'
    },
    {
      id:'T03', type:'語意不明／可能相反', level:'high', approval:true,
      test:t => has(t,'降低病患安全|降低病人安全|沒尚無法|即時向觀察|無法及時處理.*可以|可以.*無法'),
      reason:'句意可能相反或存在關鍵語意缺口，不能由審閱者自行猜測原意。',
      suggestion:'建議退件｜請教師確認原意'
    },
    {
      id:'T04', type:'疑似不適當評價／標籤化描述', level:'medium', approval:true,
      test:t => suspiciousLabel(t),
      reason:'紀錄出現較主觀或人格化評價，但缺少足夠的可觀察行為與脈絡支持，建議人工確認。',
      suggestion:'待人工確認｜若屬人格化定性，改為具體可觀察行為描述'
    },
    {
      id:'T05', type:'前後敘述待確認', level:'medium', approval:true,
      test:t => has(t,'(已|逐漸|明顯).{0,12}(改善|進步).{0,28}(時常|仍常|持續).{0,12}(遺漏|錯誤|無法|需要提醒)'),
      reason:'同一紀錄同時出現改善與持續性問題，需人工確認目前實際程度。',
      suggestion:'待人工確認｜確認目前狀況與紀錄一致性'
    }
  ];

  function managementSignals(row,t){
    const issues=[];
    const leaveDays=Number(norm(row['學員連續請假天數']).replace(/[^0-9.-]/g,''))||0;
    const stopReason=[row['停止輔導原因'],row['其他結案說明'],row['離職類別'],row['其他離職說明']].map(norm).join(' ');
    const med=has(t,'給藥|藥物|藥盒|三讀五對|點滴|化療|藥品')&&has(t,'錯|誤|放置別床|遺漏|漏|未|near\s*miss|異常|事件');
    const safety=has(t,'病人安全|病患安全|病安|near\s*miss|跌倒|針扎|事件檢討|異常事件');
    const invasive=has(t,'侵入性|導管|鼻胃管|尿管|引流')&&has(t,'錯|未|權責|事件|異常|自行|不確定');
    const manager=has(t,'護理長|阿長|主管|與.*討論後.*班別|調整.*班別|異動班別');
    const conflict=has(t,'口角|爭執|衝突|跨單位|交班事件|抱怨');
    const adaptation=has(t,'無法適應|工作適應')&&has(t,'主管|護理長|班別|持續|反覆|影響');
    const emotionWork=has(t,'哭泣|情緒|壓力|憂鬱')&&has(t,'完成臨床工作|影響工作|主管|護理長|調整班別');
    const turnover=has(stopReason,'離職|轉調|調職|生涯規劃')||has(t,'離職|轉調|調職');
    if(med) issues.push('給藥／病安事件'); else if(safety) issues.push('病人安全');
    if(invasive) issues.push('侵入性處置');
    if(manager) issues.push('主管／管理介入');
    if(conflict) issues.push('溝通／跨單位事件');
    if(adaptation) issues.push('工作適應／教學支持');
    if(emotionWork) issues.push('情緒／壓力影響工作');
    if(leaveDays>0) issues.push(`連續請假 ${leaveDays} 天`);
    if(turnover) issues.push('離職／轉調');
    return [...new Set(issues)];
  }

  function managementDetail(t,issues){
    const managerKnown=has(t,'護理長|阿長|主管|事件檢討|PDCA|通報|晨會|已處理|已調整|已介入');
    let state=managerKnown?'已見主管／制度處理訊號':'現有紀錄無法確認主管／制度是否已處理';
    let question='確認主管是否知悉、目前處理狀態及是否需要後續追蹤。';
    if(issues.some(x=>/給藥|病人安全/.test(x))){
      const hasOutcome=has(t,'未造成|無傷害|病人無|病患無|已給入|未給入|立即停止|處置|通報|事件檢討|PDCA');
      if(!hasOutcome){ state='事件訊號明確，但影響／處理資訊不足'; question='確認事件是否到達病人、影響、即時處置、制度處理及後續追蹤。'; }
      else question='確認事件釐清結果、制度處理與改善措施是否持續落實。';
    } else if(issues.some(x=>/離職|轉調/.test(x))) question='確認人員異動狀態、輔導結案與必要交接。';
    else if(issues.some(x=>/請假/.test(x))) question='確認連續請假是否影響輔導進度，以及是否需要調整後續安排。';
    else if(issues.some(x=>/工作適應|情緒/.test(x))) question='確認目前支持措施、工作安排及下月是否仍需追蹤。';
    return {state,question};
  }

  function keyContentGap(t,managementIssues){
    if(!managementIssues.some(x=>/給藥|病人安全/.test(x))) return false;
    const eventWords=has(t,'錯誤|給錯|放錯|near\s*miss|病安|病人安全|事件');
    const followWords=has(t,'通報|處置|未造成|無傷害|事件檢討|PDCA|主管|護理長|改善|追蹤|已處理');
    return eventWords&&!followWords;
  }

  function highlightAnalysis(t){
    const signals=[];
    if(has(t,'示範|示教|逐步帶領|床邊陪同|陪同完成|依SOP')) signals.push('示範／實作帶領');
    if(has(t,'引導.*思考|個案討論|案例討論|共同討論|提問')) signals.push('引導臨床思考');
    if(has(t,'立即回饋|即時回饋|給予回饋|共同檢視|修正建議')) signals.push('即時回饋');
    if(has(t,'學員回饋|學員表示|能說出|能逐漸|信心.*提升|願意調整|內化')) signals.push('看見學員反應／改變');
    if(has(t,'反思|省思|PDCA|晨會導讀|知識分享')) signals.push('反思／團隊學習');
    const unique=[...new Set(signals)];
    const score=unique.length+(t.length>=180?1:0);
    const yes=score>=3&&unique.some(x=>/學員反應|即時回饋|反思/.test(x));
    if(!yes) return {yes:false,type:'',reason:'',feedback:''};
    const type=unique.slice(0,2).join('＋');
    const reason=`紀錄可看見具體教學方法，且包含${unique.includes('看見學員反應／改變')?'學員反應或改變':'回饋／反思'}，具可回溯的教學價值。`;
    const feedback=unique.includes('反思／團隊學習')?'可作為教學案例分享':(unique.includes('即時回饋')?'建議教學亮點感謝函':'主管口頭肯定');
    return {yes:true,type,reason,feedback};
  }

  function analyzeRecord(row,sourceRow){
    const raw=norm(row[textField]);
    const t=clean(raw);
    const typos=typoFindings(raw);
    const ruleHits=[];
    let primary=null;
    for(const rule of TEXT_RULES){
      if(rule.test(raw)){
        ruleHits.push(rule.id);
        if(!primary||({low:1,medium:2,high:3}[rule.level]>{low:1,medium:2,high:3}[primary.level])) primary=rule;
      }
    }
    const managementIssues=managementSignals(row,t);
    if(keyContentGap(t,managementIssues)){
      const gapRule={id:'T06',type:'關鍵內容不足',level:'high',approval:true,reason:'紀錄已出現重要病安／事件訊號，但缺少足以完成管理判讀的影響或處理資訊。',suggestion:'建議退件｜補充必要事件與後續資訊'};
      ruleHits.push(gapRule.id); if(!primary||primary.level!=='high') primary=gapRule;
    }
    const detail=managementDetail(t,managementIssues);
    const highlight=highlightAnalysis(t);
    const textIssue=!!primary;
    const needsApproval=!!(primary&&primary.approval);
    const management=managementIssues.length>0;
    const confidence=primary?.level==='high'||managementIssues.some(x=>/離職|請假/.test(x))?'高':(textIssue||management?'中':'高');
    return {
      sourceRow, caseId:`C-${String(sourceRow).padStart(2,'0')}`, row,
      textIssue, textType:primary?.type||'', textReason:primary?.reason||'', textSuggestion:primary?.suggestion||'',
      typoAlert:typos.join('；'), needsApproval,
      management, managementIssue:managementIssues.join('、'), managementState:management?detail.state:'', managementQuestion:management?detail.question:'',
      highlight:highlight.yes, highlightType:highlight.type, highlightReason:highlight.reason, highlightFeedback:highlight.feedback,
      confidence, ruleHits:[...new Set(ruleHits)].join('、'), excerpt:clip(t,220)
    };
  }

  window.MentorRules={RULE_VERSION,analyzeRecord,cleanText:clean,clip};
})();
