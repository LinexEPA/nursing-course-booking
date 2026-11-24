/**
 * 115年度護理部 EPAs 系列課程｜學分補查
 * Google Apps Script（綁定在全新的私人 Google 試算表）
 *
 * 部署前：
 * 1. 在試算表「擴充功能 → Apps Script」貼上本檔及 Index.html。
 * 2. 從 Apps Script 編輯器執行 initialize()，完成授權。
 * 3. 在「護理長權限」工作表填妥單位、姓名、院內信箱與啟用狀態。
 * 4. 先以測試帳號驗證信箱、回寫與逾期阻擋。
 * 5. 經院內個資及平台使用核准後，才部署正式 Web App。
 *
 * 安全設計：
 * - 不在 GitHub、URL、郵件、試算表或稽核紀錄儲存完整身分證字號。
 * - 完整身分證字號僅在伺服端以秘密金鑰產生 HMAC 比對值。
 * - 僅接受「護理長權限」中啟用的院內信箱完成 OTP 驗證後送件。
 * - 請勿與既有課程預約 Apps Script 或 Sheets 共用。
 */

const EPAS = Object.freeze({
  TIMEZONE: 'Asia/Taipei',
  DEADLINE: '20261205235959',
  EMAIL_DOMAIN: '@tzuchi.com.tw',
  OTP_TTL_SECONDS: 600,
  OTP_COOLDOWN_SECONDS: 75,
  OTP_WINDOW_SECONDS: 21600,
  MAX_OTP_REQUESTS: 5,
  SESSION_TTL_SECONDS: 1800,
  MAX_OTP_ATTEMPTS: 5,
  MAIN_SHEET: '補查資料',
  ACCESS_SHEET: '護理長權限',
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
    '建檔時間（台灣）', '更新時間（台灣）', '單位', '護理長姓名', '護理長院內信箱',
    '臨床教師全名', '教師身分比對碼（HMAC）', '身分證遮蔽顯示',
    '第一堂｜甚麼學員可以被信賴', '第二堂｜你以為的，是我以為的？',
    '第三堂｜老師～換你上場！', '受評 PGY 學員全名', '處理狀態'
  ]);
  ensureSheet_(ss, EPAS.ACCESS_SHEET, ['單位', '護理長姓名', '院內信箱', '啟用（是／否）']);
  ensureSheet_(ss, EPAS.AUDIT_SHEET, ['時間（台灣）', '護理長院內信箱', '事件', '教師身分比對碼（HMAC）']);
  Logger.log('初始化完成。請先維護「護理長權限」工作表，再執行測試。');
}

function getFormStatus() {
  return {open: !isClosed_(), deadline: '115/12/05', title: '115年度護理部 EPAs 系列課程'};
}

/** 申請一次性驗證碼；對未知信箱回傳相同文案，避免查探人員名單。 */
function requestOtp(rawEmail) {
  if (isClosed_()) throw new Error('本次補查已截止，逾期申請請循護理部裁示辦理。');
  const email = normalizeEmail_(rawEmail);
  if (!email) throw new Error('請輸入正確的院內信箱。');

  const cache = CacheService.getScriptCache();
  const emailKey = hmac_(email).slice(0, 36);
  const throttleKey = 'cooldown_' + emailKey;
  const limitKey = 'request_limit_' + emailKey;
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    if (cache.get(throttleKey)) {
      throw new Error('驗證碼已寄送，請稍候再試。');
    }
    const requests = Number(cache.get(limitKey) || 0);
    if (requests >= EPAS.MAX_OTP_REQUESTS) {
      throw new Error('驗證碼申請過於頻繁，請稍後聯絡承辦人。');
    }
    cache.put(throttleKey, '1', EPAS.OTP_COOLDOWN_SECONDS);
    cache.put(limitKey, String(requests + 1), EPAS.OTP_WINDOW_SECONDS);
  } finally {
    lock.releaseLock();
  }

  const manager = findManager_(email);
  if (manager) {
    const random = Utilities.getUuid().replace(/-/g, '');
    const code = String(parseInt(random.slice(0, 12), 16) % 1000000).padStart(6, '0');
    const nonce = Utilities.getUuid();
    const record = {hash: hmac_(nonce + ':' + code), nonce, attempts: 0};
    cache.put('otp_' + emailKey, JSON.stringify(record), EPAS.OTP_TTL_SECONDS);
    MailApp.sendEmail({
      to: email,
      subject: 'EPAs 學分補查｜一次性驗證碼',
      body: '您的學分補查驗證碼為：' + code + '\n\n10 分鐘內有效，請勿轉傳。\n本信不包含教師個人資料。'
    });
  }
  return {ok: true, message: '若此信箱已列入護理長名單，驗證碼將寄至院內信箱。'};
}

/** 驗證成功後建立短效工作階段。 */
function verifyOtp(rawEmail, rawCode) {
  if (isClosed_()) throw new Error('本次補查已截止。');
  const email = normalizeEmail_(rawEmail);
  const code = String(rawCode || '').trim();
  if (!email || !/^\d{6}$/.test(code)) throw new Error('請輸入 6 位數驗證碼。');

  const cache = CacheService.getScriptCache();
  const emailKey = hmac_(email).slice(0, 36);
  const key = 'otp_' + emailKey;
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const stored = cache.get(key);
    if (!stored) throw new Error('驗證碼已失效，請重新申請。');
    const record = JSON.parse(stored);
    record.attempts += 1;
    if (record.attempts > EPAS.MAX_OTP_ATTEMPTS) {
      cache.remove(key);
      throw new Error('嘗試次數過多，請重新申請驗證碼。');
    }
    if (!sameText_(record.hash, hmac_(record.nonce + ':' + code))) {
      if (record.attempts >= EPAS.MAX_OTP_ATTEMPTS) cache.remove(key);
      else cache.put(key, JSON.stringify(record), EPAS.OTP_TTL_SECONDS);
      throw new Error('驗證碼不正確。');
    }
    const manager = findManager_(email);
    if (!manager) throw new Error('目前無法確認填寫權限，請聯絡承辦人。');
    cache.remove(key);

    const token = Utilities.getUuid() + Utilities.getUuid();
    cache.put('session_' + hmac_(token).slice(0, 48), JSON.stringify(manager), EPAS.SESSION_TTL_SECONDS);
    return {ok: true, token, unit: manager.unit, manager: manager.name, expiresMinutes: 30};
  } finally {
    lock.releaseLock();
  }
}

/** 同一位教師在同一單位重複回填，更新原紀錄，不產生重複學分案件。 */
function submitCreditRecord(token, rawPayload) {
  if (isClosed_()) throw new Error('本次補查已截止，逾期申請請循護理部裁示辦理。');
  const manager = getSession_(token);
  if (!manager) throw new Error('驗證已逾時，請重新登入。');
  const payload = validateRecord_(rawPayload);
  const idHash = hmac_(payload.nationalId);
  const maskedId = payload.nationalId.slice(0, 1) + '＊＊＊＊＊＊＊' + payload.nationalId.slice(-2);
  const now = Utilities.formatDate(new Date(), EPAS.TIMEZONE, 'yyyy/MM/dd HH:mm:ss');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (isClosed_()) throw new Error('本次補查已截止。');
    const ss = getSpreadsheet_();
    const sheet = ss.getSheetByName(EPAS.MAIN_SHEET);
    if (!sheet) throw new Error('系統尚未完成初始化。');
    const rows = sheet.getDataRange().getValues();
    let matchingRow = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][6]) === idHash) {
        if (String(rows[i][2]) !== manager.unit) {
          throw new Error('此教師已有其他單位紀錄，請洽承辦人確認。');
        }
        matchingRow = i + 1;
        break;
      }
    }
    const createdAt = matchingRow > 0 ? rows[matchingRow - 1][0] : now;
    const record = [
      createdAt, now, safeCell_(manager.unit), safeCell_(manager.name), manager.email,
      safeCell_(payload.teacherName), idHash, maskedId,
      payload.course1, payload.course2, payload.course3, safeCell_(payload.pgyName), '待查核'
    ];
    if (matchingRow > 0) {
      sheet.getRange(matchingRow, 1, 1, record.length).setValues([record]);
    } else {
      sheet.appendRow(record);
    }
    ss.getSheetByName(EPAS.AUDIT_SHEET).appendRow([
      now, manager.email, matchingRow > 0 ? '更新' : '新增', idHash
    ]);
    SpreadsheetApp.flush();
    return {ok: true, message: matchingRow > 0 ? '已更新原補查紀錄。' : '補查資料已送出。'};
  } finally {
    lock.releaseLock();
  }
}

function validateRecord_(payload) {
  if (!payload || typeof payload !== 'object') throw new Error('表單資料不完整。');
  const teacherName = String(payload.teacherName || '').trim();
  const nationalId = String(payload.nationalId || '').trim().toUpperCase().replace(/\s|-/g, '');
  const course1 = String(payload.course1 || '');
  const course2 = String(payload.course2 || '');
  const course3 = String(payload.course3 || '');
  const pgyName = String(payload.pgyName || '').trim();

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
  return {teacherName, nationalId, course1, course2, course3, pgyName: completed ? pgyName : ''};
}

function findManager_(email) {
  const sheet = getSpreadsheet_().getSheetByName(EPAS.ACCESS_SHEET);
  if (!sheet) throw new Error('系統尚未完成初始化。');
  const values = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < values.length; i++) {
    const [unit, name, allowedEmail, enabled] = values[i].map(x => String(x || '').trim());
    if (allowedEmail.toLowerCase() === email && ['是', 'Y', 'YES', 'TRUE', '1'].includes(enabled.toUpperCase())) {
      if (!unit || !name) continue;
      return {unit, name, email};
    }
  }
  return null;
}

function getSession_(token) {
  const text = String(token || '');
  if (!/^[0-9a-f-]{72}$/i.test(text)) return null;
  const stored = CacheService.getScriptCache().get('session_' + hmac_(text).slice(0, 48));
  if (!stored) return null;
  const manager = JSON.parse(stored);
  const live = findManager_(manager.email);
  if (!live || live.unit !== manager.unit) return null;
  return live;
}

function normalizeEmail_(rawEmail) {
  const email = String(rawEmail || '').trim().toLowerCase();
  return /^[a-z0-9._%+-]+@tzuchi\.com\.tw$/.test(email) ? email : '';
}

function isClosed_() {
  const current = Utilities.formatDate(new Date(), EPAS.TIMEZONE, 'yyyyMMddHHmmss');
  return current > EPAS.DEADLINE;
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

function sameText_(a, b) {
  const aa = String(a), bb = String(b);
  if (aa.length !== bb.length) return false;
  let diff = 0;
  for (let i = 0; i < aa.length; i++) diff |= aa.charCodeAt(i) ^ bb.charCodeAt(i);
  return diff === 0;
}

function safeCell_(text) {
  const value = String(text || '').trim();
  return /^[=+\-@]/.test(value) ? "'" + value : value;
}
