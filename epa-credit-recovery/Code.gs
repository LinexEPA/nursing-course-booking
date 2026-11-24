/**
 * 115年度護理部 EPAs 系列課程｜學分補查
 * Google Apps Script（綁定在全新的私人 Google 試算表）
 *
 * 部署前：
 * 1. 在試算表「擴充功能 → Apps Script」貼上本檔及 Index.html。
 * 2. 從 Apps Script 編輯器執行 initialize()，完成授權。
 * 3. 先以虛構資料確認回寫、重複登記與逾期阻擋。
 * 4. 經院內個資及平台使用核准後，才部署正式 Web App。
 *
 * 安全設計：
 * - 不在 GitHub、URL、郵件、試算表或稽核紀錄儲存完整身分證字號。
 * - 完整身分證字號僅在伺服端以秘密金鑰產生 HMAC 比對值。
 * - 不蒐集電子信箱；使用 Apps Script 私有指令碼屬性中的共用碼限制登記。
 * - 共用碼僅為基本門檻，不能證明填寫者身分，也不得寫入公開 GitHub。
 * - 所有登記僅供承辦人核對紙本證據，絕不自動認列學分。
 * - 請勿與既有課程預約 Apps Script 或 Sheets 共用。
 */

const EPAS = Object.freeze({
  TIMEZONE: 'Asia/Taipei',
  DEADLINE: '20261205235959',
  MAIN_SHEET: '補查資料',
  AUDIT_SHEET: '異動紀錄',
  COURSE_AB: ['已填答但未入學分', '未收到通知', '已收到尚未填答', '無須補查'],
  COURSE_C: ['已完成並繳交紙本', '已完成尚未繳交', '尚未完成評核', '無須補查']
});

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('EPAs 系列課程｜學分補查')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/** 僅供腳本擁有者從 Apps Script 編輯器執行。重複執行不清除紀錄。 */
function initialize() {
  const actor = Session.getActiveUser().getEmail();
  const owner = Session.getEffectiveUser().getEmail();
  if (!actor || !owner || actor.toLowerCase() !== owner.toLowerCase()) {
    throw new Error('僅限腳本擁有者在編輯器執行初始化。');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('請從新建立的 Google 試算表開啟 Apps Script。');
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('EPA_CREDIT_PEPPER')) {
    props.setProperty('EPA_CREDIT_PEPPER', Utilities.getUuid() + Utilities.getUuid());
  }
  props.setProperty('EPA_CREDIT_SPREADSHEET_ID', ss.getId());

  ensureSheet_(ss, EPAS.MAIN_SHEET, [
    '建檔時間（台灣）', '更新時間（台灣）', '單位', '護理長姓名', '登記來源',
    '臨床教師全名', '教師身分比對碼（HMAC）', '身分證遮蔽顯示',
    '第一堂｜甚麼學員可以被信賴', '第二堂｜你以為的，是我以為的？',
    '第三堂｜老師～換你上場！', '受評 PGY 學員全名', '處理狀態'
  ]);
  ensureSheet_(ss, EPAS.AUDIT_SHEET, ['時間（台灣）', '填寫護理長', '事件', '教師身分比對碼（HMAC）']);
  // 舊版欄位沒有蒐集資料，可以安全改名；保留既有欄序及資料。
  ss.getSheetByName(EPAS.MAIN_SHEET).getRange('E1').setValue('登記來源');
  ss.getSheetByName(EPAS.AUDIT_SHEET).getRange('B1').setValue('填寫護理長');
  Logger.log('初始化完成。請確認表單回寫；所有紀錄須人工核對後才可認列。');
}

function getFormStatus() {
  return {open: !isClosed_(), deadline: '115/12/05', title: '115年度護理部 EPAs 系列課程'};
}

/** 初次進入表單時核對共用碼；不將共用碼寫入試算表。 */
function verifySharedCode(accessCode) {
  if (isClosed_()) throw new Error('本次補查已截止，逾期申請請循護理部裁示辦理。');
  requireSharedCode_(accessCode);
  return {ok: true};
}

/** 登記只代表請求補查。保留每次提交供人工比對，不覆寫前次證據。 */
function submitCreditRecord(rawPayload) {
  if (isClosed_()) throw new Error('本次補查已截止，逾期申請請循護理部裁示辦理。');
  requireSharedCode_(rawPayload && rawPayload.accessCode);
  const payload = validateRecord_(rawPayload);
  const idHash = hmac_(payload.nationalId);
  const maskedId = payload.nationalId.slice(0, 1) + '＊＊＊＊＊' + payload.nationalId.slice(-4);
  const now = Utilities.formatDate(new Date(), EPAS.TIMEZONE, 'yyyy/MM/dd HH:mm:ss');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (isClosed_()) throw new Error('本次補查已截止。');
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(EPAS.MAIN_SHEET);
    if (!sheet) throw new Error('系統尚未完成初始化。');
    const record = [
      now, now, safeCell_(payload.unit), safeCell_(payload.managerName), '公開網頁（待人工核對）',
      safeCell_(payload.teacherName), idHash, maskedId,
      payload.course1, payload.course2, payload.course3, safeCell_(payload.pgyName), '待查核'
    ];
    sheet.appendRow(record);
    ss.getSheetByName(EPAS.AUDIT_SHEET).appendRow([
      now, safeCell_(payload.unit + '｜' + payload.managerName), '新增補查登記', idHash
    ]);
    SpreadsheetApp.flush();
    return {ok: true, message: '補查登記已送出，將依簽到單及紙本評核表查核。'};
  } finally {
    lock.releaseLock();
  }
}

function validateRecord_(payload) {
  if (!payload || typeof payload !== 'object') throw new Error('表單資料不完整。');
  const unit = String(payload.unit || '').trim();
  const managerName = String(payload.managerName || '').trim();
  const teacherName = String(payload.teacherName || '').trim();
  const nationalId = String(payload.nationalId || '').trim().toUpperCase().replace(/\s|-/g, '');
  const course1 = String(payload.course1 || '');
  const course2 = String(payload.course2 || '');
  const course3 = String(payload.course3 || '');
  const pgyName = String(payload.pgyName || '').trim();

  if (unit.length < 1 || unit.length > 40 || managerName.length < 2 || managerName.length > 40) {
    throw new Error('請填寫單位及護理長全名。');
  }
  if (teacherName.length < 2 || teacherName.length > 40) throw new Error('請填寫教師全名。');
  if (!/^[A-Z][A-Z0-9]{9}$/.test(nationalId)) {
    throw new Error('請確認身分證字號格式（英文字母與數字共 10 碼）。');
  }
  if (!EPAS.COURSE_AB.includes(course1) || !EPAS.COURSE_AB.includes(course2) ||
      !EPAS.COURSE_C.includes(course3)) throw new Error('請確認三堂課程的填寫狀態。');
  const completed = course3 === EPAS.COURSE_C[0] || course3 === EPAS.COURSE_C[1];
  if (completed && (pgyName.length < 2 || pgyName.length > 40)) {
    throw new Error('第三堂已完成者，請填寫接受評核的 PGY 學員全名。');
  }
  return {unit, managerName, teacherName, nationalId, course1, course2, course3, pgyName: completed ? pgyName : ''};
}

function isClosed_() {
  const current = Utilities.formatDate(new Date(), EPAS.TIMEZONE, 'yyyyMMddHHmmss');
  return current > EPAS.DEADLINE;
}

function requireSharedCode_(accessCode) {
  const expected = PropertiesService.getScriptProperties().getProperty('EPA_CREDIT_SHARED_CODE');
  if (!expected) throw new Error('表單尚未完成共用碼設定，請聯繫承辦人。');
  const supplied = String(accessCode || '').trim().toUpperCase();
  if (supplied.length !== expected.length) throw new Error('共用碼不正確，請向護理長確認。');
  let mismatch = 0;
  for (let i = 0; i < expected.length; i++) mismatch |= supplied.charCodeAt(i) ^ expected.charCodeAt(i);
  if (mismatch) throw new Error('共用碼不正確，請向護理長確認。');
}

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('EPA_CREDIT_SPREADSHEET_ID');
  if (!id) throw new Error('系統尚未完成初始化。');
  return SpreadsheetApp.openById(id);
}

function ensureSheet_(ss, name, headers) {
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#DFEBF7');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function hmac_(message) {
  const secret = PropertiesService.getScriptProperties().getProperty('EPA_CREDIT_PEPPER');
  if (!secret) throw new Error('系統尚未完成初始化。');
  const bytes = Utilities.computeHmacSha256Signature(String(message), secret);
  return bytes.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}

function safeCell_(text) {
  const value = String(text || '').trim();
  return /^[=+\-@]/.test(value) ? "'" + value : value;
}
