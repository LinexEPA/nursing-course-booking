(() => {
  'use strict';

  if (document.getElementById('recordCoachSection')) return;

  const style = document.createElement('style');
  style.textContent = `
    .coach-section{margin:0 0 30px;padding:0 2px}
    .coach-panel{border:1px solid #eadfe1;border-radius:24px;padding:24px;background:rgba(255,253,252,.9);box-shadow:0 13px 35px rgba(108,88,94,.045)}
    .coach-head{display:flex;align-items:flex-start;justify-content:space-between;gap:18px;margin-bottom:16px}
    .coach-step{display:inline-flex;align-items:center;gap:7px;color:#a97482;font-size:13px;font-weight:700;margin-bottom:7px}
    .coach-head h2{margin:0 0 7px;color:#6b6067;font-size:22px;line-height:1.4}
    .coach-head p{margin:0;color:#8a8186;font-size:14px;line-height:1.75;max-width:760px}
    .coach-rule{flex:0 0 auto;padding:6px 10px;border-radius:999px;background:#f8eef1;color:#9c7480;font-size:12px;font-weight:700;white-space:nowrap}
    .coach-textarea{width:100%;min-height:150px;resize:vertical;border:1px solid #e7dadd;border-radius:17px;padding:15px 16px;background:#fff;color:#5f585d;font:inherit;font-size:14px;line-height:1.7;outline:none}
    .coach-textarea:focus{border-color:#d3adb7;box-shadow:0 0 0 3px rgba(201,149,162,.09)}
    .coach-under{display:flex;align-items:center;justify-content:space-between;gap:14px;margin-top:10px;flex-wrap:wrap}
    .coach-note{color:#968b90;font-size:12.5px;line-height:1.6}
    .coach-actions{display:flex;gap:8px;flex-wrap:wrap}
    .coach-btn{border:0;border-radius:12px;padding:9px 14px;font-size:13px;font-weight:750;cursor:pointer}
    .coach-check{background:#bd8996;color:#fff;box-shadow:0 6px 15px rgba(169,116,130,.12)}
    .coach-copy{background:#f4ecee;color:#8d6d76}
    .coach-result{margin-top:16px;display:none}
    .coach-result.show{display:block}
    .coach-summary{padding:12px 14px;border-radius:15px;background:#faf5f6;border:1px solid #eee1e4;color:#74696f;font-size:13.5px;line-height:1.65}
    .coach-summary.good{background:#f3f7f4;border-color:#e2ece5;color:#5e7469}
    .coach-chips{display:flex;gap:7px;flex-wrap:wrap;margin:10px 0 0}
    .coach-chip{padding:4px 9px;border-radius:999px;background:#f4eeee;color:#8a777e;font-size:11.5px}
    .coach-chip.on{background:#f0f6f2;color:#647b6e}
    .coach-guides{display:grid;gap:10px;margin-top:12px}
    .coach-guide{border:1px solid #eadfe1;border-radius:16px;padding:13px 14px;background:#fff}
    .coach-guide strong{display:block;color:#7a6870;font-size:13.5px;margin-bottom:5px}
    .coach-guide p{margin:0;color:#7f767a;font-size:13px;line-height:1.65}
    .coach-example{margin-top:7px;padding:8px 10px;border-radius:11px;background:#fbf5f6;color:#907e84;font-size:12.5px;line-height:1.6}
    .coach-caution{margin-top:12px;color:#9a8e93;font-size:12px;line-height:1.65}
    @media(max-width:720px){.coach-panel{padding:19px 17px}.coach-head{flex-direction:column}.coach-rule{align-self:flex-start}.coach-textarea{min-height:180px}}
  `;
  document.head.appendChild(style);

  const host = document.querySelector('.card');
  const hero = document.querySelector('.hero');
  if (!host || !hero) return;

  const section = document.createElement('section');
  section.id = 'recordCoachSection';
  section.className = 'coach-section';
  section.innerHTML = `
    <div class="coach-panel">
      <div class="coach-head">
        <div>
          <div class="coach-step">♡ 送出前先看一下</div>
          <h2>輔導紀錄小助手</h2>
          <p>貼上您原本寫的紀錄。這裡不會幫您重寫，也不會替您補上沒有發生的內容；只會提醒哪些地方若能再具體一點，正式紀錄會更容易被理解。</p>
        </div>
        <span class="coach-rule">保留您的語句</span>
      </div>
      <textarea id="coachText" class="coach-textarea" placeholder="把您原本寫的輔導紀錄貼在這裡，也可以直接在這裡補充或修改…"></textarea>
      <div class="coach-under">
        <div class="coach-note">♡ 範例只示範「可以補哪一類資訊」，請以您實際看到、實際做過的內容為準。</div>
        <div class="coach-actions">
          <button id="coachCheck" class="coach-btn coach-check" type="button">幫我看看還缺什麼</button>
          <button id="coachCopy" class="coach-btn coach-copy" type="button">複製我現在的文字</button>
        </div>
      </div>
      <div id="coachResult" class="coach-result" aria-live="polite"></div>
    </div>
  `;
  host.parentNode.insertBefore(section, host);

  const textEl = document.getElementById('coachText');
  const checkBtn = document.getElementById('coachCheck');
  const copyBtn = document.getElementById('coachCopy');
  const resultEl = document.getElementById('coachResult');

  const esc = s => String(s || '').replace(/[&<>\"]/g, c => c === '&' ? '&amp;' : c === '<' ? '&lt;' : c === '>' ? '&gt;' : '&quot;');
  const clean = s => String(s || '').replace(/\s+/g, ' ').trim();

  function firstMatch(text, regex) {
    const m = text.match(regex);
    return m ? m[0] : '';
  }

  function analyze(text) {
    const t = clean(text);
    const guides = [];

    const hasContext = /(當|在|於|面對|遇到|執行|進行|近期|本月|白班|小夜|大夜|夜班|交班|給藥|照護|病人|工作量|突發|操作|班後|術前|術中|術後|第一次|首次|轉班|輪班)/.test(t);
    const hasTeacherAction = /(提醒|指導|示範|教導|說明|陪同|討論|分享|回顧|詢問|關心|協助|引導|鼓勵|傾聽|回饋|修正|練習|模擬|帶著|陪伴|建議|提供支持|確認|共同檢視|追問)/.test(t);
    const hasLearnerResponse = /(經[^。；，]{0,30}後|學員表示|學員回饋|已能|逐漸|願意|開始|能夠|較能|改善|提升|理解|掌握|主動|完成|減少|穩定|進步|有明顯|已逐步)/.test(t);
    const hasFollowup = /(後續|持續追蹤|下個月|再觀察|繼續|將持續|後續將|下一步|再加強|持續加強)/.test(t);

    const vagueRegex = /(個性.{0,6}(內向|外向|大而化之)|散漫|懶惰|不用心|沒責任感|沒有責任感|漫無章法|無頭蒼蠅|表現不好|表現不佳|態度不好|不積極|較被動|很被動|容易慌亂|較容易慌|心不在焉|注意力不集中|容易忘記|常忘記|能力差|能力不足|不夠細心|粗心|常出錯|偶有錯誤|錯誤或疏忽|有疏失)/;
    const vague = firstMatch(t, vagueRegex);
    const hasBehaviorEvidence = /(遺漏|漏掉|未完成|延遲|提早|遲到|未.{0,8}更新|未.{0,8}紀錄|未.{0,8}交班|忘記|動作|語氣|主動詢問|主動求助|提問|完成|核對|紀錄|通報|告知|回報|交班|排序|待辦|清單|處置|處理多件|同時)/.test(t);

    const actualEvent = /(發生.{0,12}(錯誤|疏失|事件)|給錯|漏給|誤給|跌倒|針扎|near\s*miss|跡近|已通報|造成.{0,12}(傷害|影響)|用藥錯誤)/i.test(t);
    const hasEventHandling = /(立即|當下|已通報|通知|回報|處理|處置|檢討|PDCA|修正|確認|追蹤|主管|護理長|醫師)/.test(t);
    const hasEventOutcome = /(無傷害|受傷|未造成|有無影響|影響|未給入|已給入|生命徵象|無異常|結果|後續狀況)/.test(t);

    if (t.length < 35) {
      guides.push({
        title:'可以再多一點實際情境',
        body:'目前內容比較短。請想想：這是在什麼工作情境下發生？您實際看到學員做了什麼？只要補一個真正發生過的片段就好。',
        example:'思考框架：在＿＿＿＿時，我觀察到學員＿＿＿＿。'
      });
    } else if (!hasContext) {
      guides.push({
        title:'如果方便，可以補上發生的情境',
        body:'目前可以看出您的判斷，但較難知道是在什麼情況下觀察到。補一個情境，會比把句子寫得更漂亮更有幫助。',
        example:'思考框架：在＿＿＿＿（例如交班／給藥／同時處理多項工作）時，觀察到＿＿＿＿。'
      });
    }

    if (!hasTeacherAction && t.length >= 35) {
      guides.push({
        title:'可以想想：當時您實際做了什麼？',
        body:'這段目前較多是在描述學員。如果當時有提醒、示範、討論、陪同或回饋，可以補上那個真實行動；如果當下只是觀察、沒有介入，也不需要硬補。',
        example:'思考框架：當時我有＿＿＿＿，並和學員＿＿＿＿。'
      });
    }

    if (!hasLearnerResponse && hasTeacherAction) {
      guides.push({
        title:'教學後如果有觀察到變化，可以再補一句',
        body:'目前已看得到您怎麼帶。如果後續確實有看到學員的反應、理解或改變，可以補上；如果尚未看到結果，就維持原文，不需要預測。',
        example:'思考框架：提醒／討論後，學員＿＿＿＿。若尚未觀察到，就不用填。'
      });
    }

    if (vague && (!hasBehaviorEvidence || t.length < 85)) {
      guides.push({
        title:`「${vague}」如果是重要判斷，可以再回到可觀察的事實`,
        body:'不是不能寫困難，而是可以想想：您究竟看到了什麼行為，才會有這個感受？把「評價」換成一個實際觀察，主管會更容易理解。',
        example:'句型示意：不要只寫「較被動」，可以回想是否曾出現「遇到不熟悉事項時較少主動提問」這類可觀察行為。這只是示意，請勿套用不存在的情況。'
      });
    }

    if (actualEvent && (!hasEventHandling || !hasEventOutcome)) {
      guides.push({
        title:'如果這裡描述的是實際事件，可以確認關鍵事實是否齊全',
        body:'只補您確定知道的內容：實際發生什麼、有沒有影響、當下怎麼處理、後續是否需要追蹤。不知道的部分不要推測。',
        example:'思考順序：發生什麼 → 有無影響 → 當下處理 → 後續追蹤。不是每一項都一定要有，只寫已知事實。'
      });
    }

    if (guides.length > 3) guides.length = 3;

    return {t,guides,hasContext,hasTeacherAction,hasLearnerResponse,hasFollowup};
  }

  function render() {
    const text = textEl.value;
    if (!clean(text)) {
      resultEl.className = 'coach-result show';
      resultEl.innerHTML = '<div class="coach-summary">先貼上您的原始紀錄，我才知道要提醒哪一部分。這裡不會自動改寫您的文字。</div>';
      return;
    }

    const a = analyze(text);
    const chips = [
      ['情境', a.hasContext],
      ['老師怎麼帶', a.hasTeacherAction],
      ['學員反應／改變', a.hasLearnerResponse],
      ['後續追蹤', a.hasFollowup]
    ].map(([label,on]) => `<span class="coach-chip ${on?'on':''}">${on?'✓ ':'○ '}${label}</span>`).join('');

    if (!a.guides.length) {
      resultEl.className = 'coach-result show';
      resultEl.innerHTML = `
        <div class="coach-summary good"><strong>這段資訊已經相當清楚 ♡</strong><br>看得到具體內容，不需要為了讓文字更正式或更漂亮而重寫。保留您原本的語氣就可以。</div>
        <div class="coach-chips">${chips}</div>
        <div class="coach-caution">系統只看紀錄是否容易被理解，不判定老師教得好不好，也不替您補上沒有發生的內容。</div>
      `;
      return;
    }

    resultEl.className = 'coach-result show';
    resultEl.innerHTML = `
      <div class="coach-summary"><strong>再多一點點就會更清楚 ♡</strong><br>不用重寫整段，只挑下面對您有用的提醒補充即可。</div>
      <div class="coach-chips">${chips}</div>
      <div class="coach-guides">${a.guides.map(g => `
        <div class="coach-guide">
          <strong>${esc(g.title)}</strong>
          <p>${esc(g.body)}</p>
          <div class="coach-example">${esc(g.example)}</div>
        </div>
      `).join('')}</div>
      <div class="coach-caution">提醒：範例不是建議答案，也不會自動加入您的紀錄。請只補充您實際觀察、實際教學或確定知道的內容。</div>
    `;
  }

  checkBtn.addEventListener('click', render);
  copyBtn.addEventListener('click', async () => {
    if (!textEl.value.trim()) return;
    try {
      await navigator.clipboard.writeText(textEl.value);
      const old = copyBtn.textContent;
      copyBtn.textContent = '已複製 ✓';
      setTimeout(() => copyBtn.textContent = old, 1500);
    } catch (_) {
      textEl.focus(); textEl.select(); document.execCommand('copy');
      copyBtn.textContent = '已複製 ✓';
      setTimeout(() => copyBtn.textContent = '複製我現在的文字', 1500);
    }
  });
})();
