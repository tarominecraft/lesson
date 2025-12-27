const SPREADSHEET_ID = '1TdXzJX9KUxuTBDu9cEobX0TmlabMSz-STSQ0aGYPE_8';
const TZ = 'Asia/Tokyo';

// 外部POST API用トークン（必ず変更してください）
const API_TOKEN = 'gWMBAMDzh6';

/**
 * 先生は1名想定（resources シートの先頭行を使用）
 * - 初期作成時はこのID/名前で resources を作ります
 */
const SINGLE_TEACHER_ID = 't01';
const SINGLE_TEACHER_NAME = '先生';

// 教材PDFを保存するフォルダID（空なら初回アップロード時に自動作成して ScriptProperties に保存）
const PDF_FOLDER_ID = '';

// ================================
// Webアプリ（iframe埋め込み対応）
// ================================
function doGet(e) {
  ensureSheets_();
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('授業スケジュール')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

// ================================
// 設定（ここを書き換えるだけでUIも追従）
// ================================
function getConfig_() {
  return { start: '08:00', end: '20:00', slotMinutes: 15, weekStart: 'mon' }; // weekStart: 'mon' or 'sun'
}

function getAdminPassword_() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD') || '';
}

function adminTokenCacheKey_(token) {
  return `adminToken:${token}`;
}

function isAdminToken_(adminToken) {
  if (!adminToken) return false;
  const cache = CacheService.getScriptCache();
  const key = adminTokenCacheKey_(adminToken);
  const hit = cache.get(key);
  if (hit) {
    cache.put(key, hit, 60 * 30); // extend TTL on use
  }
  return !!hit;
}

function assertAdminToken_(adminToken) {
  if (!isAdminToken_(adminToken)) throw new Error('権限がありません（管理者トークン）');
}


// ================================
// GAS公開関数（フロントから google.script.run で呼ばれる）
// ================================
function adminLogin(password) {
  try {
    const expected = getAdminPassword_();
    if (!expected) return { ok: false, message: '管理者パスワードが設定されていません' };

    if (String(password || '') !== expected) {
      return { ok: false, message: 'パスワードが違います' };
    }

    const adminToken = Utilities.getUuid();
    CacheService.getScriptCache().put(adminTokenCacheKey_(adminToken), '1', 60 * 30); // 30min TTL
    return { ok: true, adminToken, ttlSec: 60 * 30 };
  } catch (err) {
    return ng_(err);
  }
}

function getInit(adminKey) {
  try {
    ensureSheets_();
    const config = getConfig_();
    validateConfig_(config);

    const resources = listResources_();
    if (!resources.length) throw new Error('resources（先生一覧）が空です');

    const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    const teacherId = resources[0].resourceId;

    const weekStart = getWeekStartYmd_(today, config.weekStart);
    const days = buildWeekDays_(weekStart);

    const end = addDaysYmd_(weekStart, 6);
    const bookings = listBookingsRange_(teacherId, weekStart, end);

    const isAdmin = isAdminToken_(adminKey);

    return ok_({ resources, config, today, teacherId, weekStart, days, bookings, isAdmin });
  } catch (err) {
    return ng_(err);
  }
}


// teacherId と anchorDate（週のどこかの日付）を渡すと、その週の予約を返す
function getWeek(anchorDateStr, clientId, adminKey) {
  try {
    ensureSheets_();
    const config = getConfig_();
    validateConfig_(config);

    assertDateStr_(anchorDateStr);

    const resources = listResources_();
    if (!resources.length) throw new Error('resources（先生一覧）が空です');
    const teacherId = resources[0].resourceId;

    const weekStart = getWeekStartYmd_(anchorDateStr, config.weekStart);
    const days = buildWeekDays_(weekStart);
    const end = addDaysYmd_(weekStart, 6);

    const bookings = listBookingsRange_(teacherId, weekStart, end);

    const isAdmin = isAdminToken_(adminKey);
    const cid = String(clientId || '').trim();
    const requests = listRequestsRange_(teacherId, weekStart, end, isAdmin ? '' : cid);

    return ok_({ weekStart, days, bookings, requests, isAdmin });
  } catch (err) {
    return ng_(err);
  }
}


function addBookingInternal_(data) {
  const lock = LockService.getScriptLock();
  try {
    ensureSheets_();
    lock.waitLock(30000);

    const config = getConfig_();
    validateConfig_(config);
    const normalized = normalizeBookingInput_(data, config);

    // 重複チェック（同じ日・同じ先生のみ）
    const existing = listBookings_(normalized.date)
      .filter(b => (b.resourceId === normalized.resourceId) || (String(b.userName || '').trim() === normalized.userName));
    assertNoOverlap_(normalized, existing, null);

    const ss = getSs_();
    const sh = ss.getSheetByName('bookings');

    const now = nowJst_();
    const id = Utilities.getUuid();

    sh.appendRow([
      id,
      normalized.resourceId,
      normalized.date,
      normalized.startTime,
      normalized.endTime,
      normalized.userName,
      normalized.title,
      normalized.isVisitor ? 'TRUE' : 'FALSE',
      now,
      now,
      '',
      ''
    ]);

    return ok_({ id });
  } catch (err) {
    return ng_(err);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function updateBookingInternal_(id, data) {
  const lock = LockService.getScriptLock();
  try {
    ensureSheets_();
    lock.waitLock(30000);

    id = String(id || '').trim();
    if (!id) throw new Error('id が不正です');

    const config = getConfig_();
    validateConfig_(config);
    const normalized = normalizeBookingInput_(data, config);

    const ss = getSs_();
    const sh = ss.getSheetByName('bookings');
    const values = sh.getDataRange().getValues();
    if (values.length < 2) throw new Error('予約が存在しません');

    let rowIndex = -1;
    let oldRow = null;
    for (let r = 2; r <= values.length; r++) {
      const row = values[r - 1];
      if (String(row[0] || '').trim() === id) {
        rowIndex = r;
        oldRow = row;
        break;
      }
    }
    if (rowIndex === -1) throw new Error('対象の予約が見つかりません');

    // 重複チェック（更新対象を除外）
    const existing = listBookings_(normalized.date)
      .filter(b => (b.resourceId === normalized.resourceId) || (String(b.userName || '').trim() === normalized.userName));
    assertNoOverlap_(normalized, existing, id);

    const createdAt = formatDateTimeCellAsJst_(oldRow[8]) || nowJst_();
    const updatedAt = nowJst_();

    // PDFはこの関数では触らない（既存を保持）
    const pdfFileId = String((oldRow.length >= 11 ? (oldRow[10] || '') : '')).trim();
    const pdfFileName = String((oldRow.length >= 12 ? (oldRow[11] || '') : '')).trim();

    sh.getRange(rowIndex, 1, 1, 12).setValues([[
      id,
      normalized.resourceId,
      normalized.date,
      normalized.startTime,
      normalized.endTime,
      normalized.userName,
      normalized.title,
      normalized.isVisitor ? 'TRUE' : 'FALSE',
      createdAt,
      updatedAt,
      pdfFileId,
      pdfFileName
    ]]);

    return ok_({ id });
  } catch (err) {
    return ng_(err);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function deleteBookingInternal_(id) {
  const lock = LockService.getScriptLock();
  try {
    ensureSheets_();
    lock.waitLock(30000);

    id = String(id || '').trim();
    if (!id) throw new Error('id が不正です');

    const ss = getSs_();
    const sh = ss.getSheetByName('bookings');
    const values = sh.getDataRange().getValues();
    if (values.length < 2) throw new Error('予約が存在しません');

    let rowIndex = -1;
    for (let r = 2; r <= values.length; r++) {
      const row = values[r - 1];
      if (String(row[0] || '').trim() === id) {
        rowIndex = r;
        break;
      }
    }
    if (rowIndex === -1) throw new Error('対象の予約が見つかりません');

    sh.deleteRow(rowIndex);
    return ok_({ id });

  } catch (err) {
    return ng_(err);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


function uploadBookingPdfInternal_(id, fileName, base64) {
  const lock = LockService.getScriptLock();
  try {
    ensureSheets_();
    lock.waitLock(30000);

    id = String(id || '').trim();
    if (!id) throw new Error('id が不正です');

    fileName = String(fileName || '').trim();
    if (!fileName) fileName = 'lesson.pdf';

    base64 = String(base64 || '').trim();
    if (!base64) throw new Error('base64 が空です');

    const folder = getOrCreatePdfFolder_();
    const bytes = Utilities.base64Decode(base64);
    const blob = Utilities.newBlob(bytes, 'application/pdf', fileName);

    const file = folder.createFile(blob);

    // 共有設定：リンクを知っている全員が閲覧可（必要ならドメイン内に変更してください）
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {
      // ドメイン制約などで失敗する場合があるため握りつぶす
    }

    // bookings シート更新（K=pdfFileId, L=pdfFileName）
    const ss = getSs_();
    const sh = ss.getSheetByName('bookings');
    const values = sh.getDataRange().getValues();
    if (values.length < 2) throw new Error('予約が存在しません');

    let rowIndex = -1;
    for (let r = 2; r <= values.length; r++) {
      const row = values[r - 1];
      if (String(row[0] || '').trim() === id) {
        rowIndex = r;
        break;
      }
    }
    if (rowIndex === -1) throw new Error('対象の予約が見つかりません');

    sh.getRange(rowIndex, 11, 1, 2).setValues([[file.getId(), file.getName()]]);
    sh.getRange(rowIndex, 10).setValue(nowJst_()); // updatedAt も更新

    return ok_({
      id,
      pdfFileId: file.getId(),
      pdfFileName: file.getName(),
      pdfUrl: getDriveFileUrl_(file.getId())
    });
  } catch (err) {
    return ng_(err);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


function addBooking(adminKey, data) {
  try {
    assertAdminToken_(adminKey);
    return addBookingInternal_(data);
  } catch (err) {
    return ng_(err);
  }
}

function updateBooking(adminKey, id, data) {
  try {
    assertAdminToken_(adminKey);
    return updateBookingInternal_(id, data);
  } catch (err) {
    return ng_(err);
  }
}

function deleteBooking(adminKey, id) {
  try {
    assertAdminToken_(adminKey);
    return deleteBookingInternal_(id);
  } catch (err) {
    return ng_(err);
  }
}

function uploadBookingPdf(adminKey, bookingId, fileName, base64) {
  try {
    assertAdminToken_(adminKey);
    return uploadBookingPdfInternal_(bookingId, fileName, base64);
  } catch (err) {
    return ng_(err);
  }
}





// ================================
// 生徒の希望（requests）
// ================================
function createRequest(clientId, data) {
  const lock = LockService.getScriptLock();
  try {
    ensureSheets_();
    lock.waitLock(30000);

    const config = getConfig_();
    validateConfig_(config);

    const cid = String(clientId || '').trim();
    if (!cid) throw new Error('clientId が必要です');

    const resources = listResources_();
    if (!resources.length) throw new Error('resources（先生一覧）が空です');
    const teacherId = resources[0].resourceId;

    const normalized = normalizeRequestInput_(data, config);
    normalized.resourceId = teacherId;

    const ss = getSs_();
    const sh = ss.getSheetByName('requests');

    const now = nowJst_();
    const id = Utilities.getUuid();

    sh.appendRow([
      id,
      normalized.resourceId,
      normalized.date,
      normalized.startTime,
      normalized.endTime,
      normalized.title,
      normalized.studentName,
      normalized.studentEmail,
      normalized.note,
      cid,
      'pending',
      now,
      now,
      ''
    ]);

    return ok_({ id });
  } catch (err) {
    return ng_(err);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function updateRequest(clientId, id, data, adminKey) {
  const lock = LockService.getScriptLock();
  try {
    ensureSheets_();
    lock.waitLock(30000);

    const cid = String(clientId || '').trim();
    if (!cid) throw new Error('clientId が必要です');

    const isAdmin = isAdminToken_(adminKey);

    id = String(id || '').trim();
    if (!id) throw new Error('id が不正です');

    const config = getConfig_();
    validateConfig_(config);

    const ss = getSs_();
    const sh = ss.getSheetByName('requests');
    const values = sh.getDataRange().getValues();
    if (values.length < 2) throw new Error('希望が存在しません');

    let rowIndex = -1;
    let rowClientId = '';
    let status = '';
    for (let r = 2; r <= values.length; r++) {
      const row = values[r - 1];
      if (String(row[0] || '').trim() === id) {
        rowIndex = r;
        rowClientId = String(row[9] || '').trim();
        status = String(row[10] || '').trim();
        break;
      }
    }
    if (rowIndex === -1) throw new Error('対象の希望が見つかりません');

    if (!isAdmin && rowClientId !== cid) throw new Error('権限がありません');
    if (!isAdmin && status !== 'pending') throw new Error('確定済みの希望は編集できません');

    const normalized = normalizeRequestInput_(data, config);

    const now = nowJst_();
    // columns: date(3) start(4) end(5) title(6) name(7) email(8) note(9) updatedAt(13)
    sh.getRange(rowIndex, 3, 1, 7).setValues([[
      normalized.date,
      normalized.startTime,
      normalized.endTime,
      normalized.title,
      normalized.studentName,
      normalized.studentEmail,
      normalized.note
    ]]);
    sh.getRange(rowIndex, 13).setValue(now);

    return ok_({ id });
  } catch (err) {
    return ng_(err);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function deleteRequest(clientId, id, adminKey) {
  const lock = LockService.getScriptLock();
  try {
    ensureSheets_();
    lock.waitLock(30000);

    const cid = String(clientId || '').trim();
    if (!cid) throw new Error('clientId が必要です');

    const isAdmin = isAdminToken_(adminKey);

    id = String(id || '').trim();
    if (!id) throw new Error('id が不正です');

    const ss = getSs_();
    const sh = ss.getSheetByName('requests');
    const values = sh.getDataRange().getValues();
    if (values.length < 2) throw new Error('希望が存在しません');

    let rowIndex = -1;
    let rowClientId = '';
    let status = '';
    for (let r = 2; r <= values.length; r++) {
      const row = values[r - 1];
      if (String(row[0] || '').trim() === id) {
        rowIndex = r;
        rowClientId = String(row[9] || '').trim();
        status = String(row[10] || '').trim();
        break;
      }
    }
    if (rowIndex === -1) throw new Error('対象の希望が見つかりません');

    if (!isAdmin && rowClientId !== cid) throw new Error('権限がありません');
    if (!isAdmin && status !== 'pending') throw new Error('確定済みの希望は削除できません');

    sh.deleteRow(rowIndex);
    return ok_({ id });
  } catch (err) {
    return ng_(err);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function approveRequest(adminKey, requestId) {
  const lock = LockService.getScriptLock();
  try {
    assertAdminToken_(adminKey);
    ensureSheets_();
    lock.waitLock(30000);

    const config = getConfig_();
    validateConfig_(config);

    requestId = String(requestId || '').trim();
    if (!requestId) throw new Error('requestId が不正です');

    const ss = getSs_();
    const sh = ss.getSheetByName('requests');
    const values = sh.getDataRange().getValues();
    if (values.length < 2) throw new Error('希望が存在しません');

    let rowIndex = -1;
    let req = null;
    for (let r = 2; r <= values.length; r++) {
      const row = values[r - 1];
      if (String(row[0] || '').trim() === requestId) {
        rowIndex = r;
        req = {
          id: requestId,
          resourceId: String(row[1] || '').trim(),
          date: String(row[2] || '').trim(),
          startTime: String(row[3] || '').trim(),
          endTime: String(row[4] || '').trim(),
          title: String(row[5] || '').trim(),
          studentName: String(row[6] || '').trim(),
          studentEmail: String(row[7] || '').trim(),
          note: String(row[8] || '').trim(),
          status: String(row[10] || '').trim(),
        };
        break;
      }
    }
    if (rowIndex === -1 || !req) throw new Error('対象の希望が見つかりません');
    if (req.status !== 'pending') throw new Error('この希望は既に処理済みです');

    // 予約作成（重複チェック込み）
    const res = addBookingInternal_({
      resourceId: req.resourceId,
      date: req.date,
      startTime: req.startTime,
      endTime: req.endTime,
      userName: req.studentName,
      title: req.title || '',
      isVisitor: false
    });
    if (!res || !res.ok) throw new Error(res && res.message ? res.message : '授業の作成に失敗しました');

    const bookingId = res.data && res.data.id ? res.data.id : '';

    const now = nowJst_();
    sh.getRange(rowIndex, 11).setValue('approved'); // status
    sh.getRange(rowIndex, 13).setValue(now);        // updatedAt
    sh.getRange(rowIndex, 14).setValue(bookingId);  // bookingId

    return ok_({ requestId, bookingId });
  } catch (err) {
    return ng_(err);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function rejectRequest(adminKey, requestId) {
  const lock = LockService.getScriptLock();
  try {
    assertAdminToken_(adminKey);
    ensureSheets_();
    lock.waitLock(30000);

    requestId = String(requestId || '').trim();
    if (!requestId) throw new Error('requestId が不正です');

    const ss = getSs_();
    const sh = ss.getSheetByName('requests');
    const values = sh.getDataRange().getValues();
    if (values.length < 2) throw new Error('希望が存在しません');

    let rowIndex = -1;
    let status = '';
    for (let r = 2; r <= values.length; r++) {
      const row = values[r - 1];
      if (String(row[0] || '').trim() === requestId) {
        rowIndex = r;
        status = String(row[10] || '').trim();
        break;
      }
    }
    if (rowIndex === -1) throw new Error('対象の希望が見つかりません');
    if (status !== 'pending') throw new Error('この希望は既に処理済みです');

    const now = nowJst_();
    sh.getRange(rowIndex, 11).setValue('rejected');
    sh.getRange(rowIndex, 13).setValue(now);

    return ok_({ requestId });
  } catch (err) {
    return ng_(err);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function normalizeRequestInput_(data, config) {
  if (!data || typeof data !== 'object') throw new Error('data が不正です');

  const date = String(data.date || '').trim();
  const startTime = String(data.startTime || '').trim();
  const endTime = String(data.endTime || '').trim();

  const title = String(data.title || '').trim();
  const studentName = String(data.studentName || data.userName || '').trim();
  const studentEmail = String(data.studentEmail || data.email || '').trim();
  const note = String(data.note || '').trim();

  assertDateStr_(date);
  assertTimeStr_(startTime);
  assertTimeStr_(endTime);
  if (!studentName) throw new Error('studentName が必要です');

  const slot = Number(config.slotMinutes || 15);
  const startMin = timeToMinutes_(startTime);
  const endMin = timeToMinutes_(endTime);
  const openMin = timeToMinutes_(config.start);
  const closeMin = timeToMinutes_(config.end);

  if (endMin <= startMin) throw new Error('endTime は startTime より後である必要があります');
  if ((endMin - startMin) > 24 * 60) throw new Error('時間範囲が不正です');
  if (startMin % slot !== 0 || endMin % slot !== 0) throw new Error('start/end は slotMinutes 単位である必要があります');
  if (startMin < openMin || endMin > closeMin) throw new Error('営業時間外です');

  // 今日の「過去時間」は希望禁止
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  if (date === today) {
    const nowHm = Utilities.formatDate(new Date(), TZ, 'HH:mm');
    const nowMin = timeToMinutes_(nowHm);
    const minStartAllowed = Math.ceil(nowMin / slot) * slot;
    if (startMin < minStartAllowed) {
      throw new Error('過去の時間は指定できません（現在時刻より前）');
    }
  }

  return { date, startTime, endTime, title, studentName, studentEmail, note, resourceId: '' };
}

function listRequestsRange_(teacherId, startYmd, endYmd, filterClientId) {
  teacherId = String(teacherId || '').trim();
  assertDateStr_(startYmd);
  assertDateStr_(endYmd);

  const cid = String(filterClientId || '').trim();

  const ss = getSs_();
  const sh = ss.getSheetByName('requests');
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const out = [];
  for (let r = 2; r <= values.length; r++) {
    const row = values[r - 1];
    const id = String(row[0] || '').trim();
    if (!id) continue;

    const resourceId = String(row[1] || '').trim();
    const date = normalizeDateCellToYmd_(row[2]);
    if (!date) continue;

    if (resourceId !== teacherId) continue;
    if (date < startYmd || date > endYmd) continue;

    const status = String(row[10] || '').trim();
    if (status && status !== 'pending') continue;

    const rowClientId = String(row[9] || '').trim();
    if (cid && rowClientId !== cid) continue;

    out.push({
      id,
      resourceId,
      date,
      startTime: normalizeTimeCellToHm_(row[3]) || String(row[3] || '').trim(),
      endTime: normalizeTimeCellToHm_(row[4]) || String(row[4] || '').trim(),
      title: String(row[5] || '').trim(),
      studentName: String(row[6] || '').trim(),
      studentEmail: String(row[7] || '').trim(),
      note: String(row[8] || '').trim(),
      clientId: rowClientId,
      status: status || 'pending',
      createdAt: formatDateTimeCellAsJst_(row[11]) || String(row[11] || '').trim(),
      updatedAt: formatDateTimeCellAsJst_(row[12]) || String(row[12] || '').trim(),
      bookingId: String(row[13] || '').trim()
    });
  }

  out.sort((a, b) => {
    if (a.date !== b.date) return a.date.localeCompare(b.date);
    return timeToMinutes_(a.startTime) - timeToMinutes_(b.startTime);
  });
  return out;
}


function getDriveFileUrl_(fileId) {
  const id = String(fileId || '').trim();
  if (!id) return '';
  return 'https://drive.google.com/file/d/' + encodeURIComponent(id) + '/view?usp=sharing';
}

function getOrCreatePdfFolder_() {
  const props = PropertiesService.getScriptProperties();

  let folderId = String(PDF_FOLDER_ID || '').trim();
  if (!folderId) folderId = String(props.getProperty('PDF_FOLDER_ID') || '').trim();

  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (e) {
      // fallthrough
    }
  }

  // create
  const folder = DriveApp.createFolder('授業PDF');
  props.setProperty('PDF_FOLDER_ID', folder.getId());
  return folder;
}

// ================================
// 外部POST API（任意）
// ================================
function doPost(e) {
  try {
    ensureSheets_();

    const payload = parsePostPayload_(e);
    const token = String(payload.token || '').trim();
    if (!token || token !== API_TOKEN) {
      return jsonOut_({ ok: false, message: 'Invalid token' });
    }

    const action = String(payload.action || '').trim();

    if (action === 'add') {
      const res = addBookingInternal_({
        resourceId: payload.resourceId,
        date: payload.date,
        startTime: payload.startTime,
        endTime: payload.endTime,
        userName: payload.userName,
        title: payload.title,
        isVisitor: payload.isVisitor
      });
      return jsonOut_(res);
    }

    if (action === 'delete') {
      const res = deleteBookingInternal_(String(payload.id || ''));
      return jsonOut_(res);
    }

    if (action === 'list') {
      const date = String(payload.date || '').trim();
      assertDateStr_(date);
      return jsonOut_(ok_(listBookings_(date)));
    }

    return jsonOut_({ ok: false, message: 'Unknown action' });
  } catch (err) {
    return jsonOut_(ng_(err));
  }
}

// ================================
// 内部処理（スプレッドシート）
// ================================
function getSs_() {
  if (!SPREADSHEET_ID || SPREADSHEET_ID === 'YOUR_ID_HERE') {
    throw new Error('SPREADSHEET_ID を実際のスプレッドシートIDに変更してください');
  }
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    throw new Error('スプレッドシートを開けません。SPREADSHEET_ID と権限を確認してください');
  }
}

function getDefaultResources_() {
  return [
    { resourceId: SINGLE_TEACHER_ID, resourceName: SINGLE_TEACHER_NAME, capacity: 0, sortOrder: 1, isActive: true },
  ];
}

function ensureSheets_() {
  const ss = getSs_();

  // resources
  let rs = ss.getSheetByName('resources');
  if (!rs) {
    rs = ss.insertSheet('resources');
    rs.getRange(1, 1, 1, 5).setValues([['resourceId', 'resourceName', 'capacity', 'sortOrder', 'isActive']]);

    const init = getDefaultResources_().map(r => ([
      r.resourceId, r.resourceName, r.capacity, r.sortOrder, r.isActive ? 'TRUE' : 'FALSE'
    ]));
    rs.getRange(2, 1, init.length, 5).setValues(init);
    rs.setFrozenRows(1);
  } else {
    const expected = ['resourceId', 'resourceName', 'capacity', 'sortOrder', 'isActive'];
    const header = rs.getRange(1, 1, 1, 5).getValues()[0];
    const ok = expected.every((v, i) => String(header[i] || '') === v);
    if (!ok) rs.getRange(1, 1, 1, 5).setValues([expected]);

    syncResourcesToDefault_(rs, getDefaultResources_());
    rs.setFrozenRows(1);
  }

  // bookings
  let bs = ss.getSheetByName('bookings');
  const expectedB = ['id','resourceId','date','startTime','endTime','userName','title','isVisitor','createdAt','updatedAt','pdfFileId','pdfFileName'];
  if (!bs) {
    bs = ss.insertSheet('bookings');
    bs.getRange(1, 1, 1, expectedB.length).setValues([expectedB]);
    bs.setFrozenRows(1);
  } else {
    // 足りない列があれば追加
    const lastCol = bs.getLastColumn();
    if (lastCol < expectedB.length) {
      bs.insertColumnsAfter(lastCol, expectedB.length - lastCol);
    }
    const header = bs.getRange(1, 1, 1, expectedB.length).getValues()[0];
    const ok = expectedB.every((v, i) => String(header[i] || '') === v);
    if (!ok) bs.getRange(1, 1, 1, expectedB.length).setValues([expectedB]);
    bs.setFrozenRows(1);
  }


  // requests（生徒の希望）
  let qs = ss.getSheetByName('requests');
  const expectedQ = [
    'id','resourceId','date','startTime','endTime',
    'title','studentName','studentEmail','note',
    'clientId','status', // status: pending/approved/rejected
    'createdAt','updatedAt',
    'bookingId'
  ];
  if (!qs) {
    qs = ss.insertSheet('requests');
    qs.getRange(1, 1, 1, expectedQ.length).setValues([expectedQ]);
    qs.setFrozenRows(1);
  } else {
    const lastCol = qs.getLastColumn();
    if (lastCol < expectedQ.length) {
      qs.insertColumnsAfter(lastCol, expectedQ.length - lastCol);
    }
    const header = qs.getRange(1, 1, 1, expectedQ.length).getValues()[0];
    const ok = expectedQ.every((v, i) => String(header[i] || '') === v);
    if (!ok) qs.getRange(1, 1, 1, expectedQ.length).setValues([expectedQ]);
    qs.setFrozenRows(1);
  }

  // 文字列扱いに寄せる（事故回避）
  try {
    bs.getRange('A:A').setNumberFormat('@');
    bs.getRange('B:B').setNumberFormat('@');
    bs.getRange('C:C').setNumberFormat('@');
    bs.getRange('D:E').setNumberFormat('@');
    bs.getRange('F:G').setNumberFormat('@');
    bs.getRange('H:H').setNumberFormat('@');
    bs.getRange('I:J').setNumberFormat('@');
    bs.getRange('K:L').setNumberFormat('@');
  } catch (e) {}
}

// resources をデフォルト一覧に同期（不足分追加・既存更新）
function syncResourcesToDefault_(sheet, defaults) {
  const values = sheet.getDataRange().getValues();
  const existingMap = new Map(); // resourceId -> rowIndex
  for (let r = 2; r <= values.length; r++) {
    const row = values[r - 1];
    const id = String(row[0] || '').trim();
    if (!id) continue;
    existingMap.set(id, r);
  }

  const appendRows = [];
  for (const d of defaults) {
    const rid = String(d.resourceId).trim();
    const newRow = [
      rid,
      String(d.resourceName || '').trim(),
      Number(d.capacity || 0),
      Number(d.sortOrder || 0),
      d.isActive ? 'TRUE' : 'FALSE'
    ];

    const rowIndex = existingMap.get(rid);
    if (rowIndex) {
      sheet.getRange(rowIndex, 1, 1, 5).setValues([newRow]);
    } else {
      appendRows.push(newRow);
    }
  }

  if (appendRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, appendRows.length, 5).setValues(appendRows);
  }
}

function listResources_() {
  const ss = getSs_();
  const sh = ss.getSheetByName('resources');
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const out = [];
  for (let r = 2; r <= values.length; r++) {
    const row = values[r - 1];
    const resourceId = String(row[0] || '').trim();
    if (!resourceId) continue;

    const resourceName = String(row[1] || '').trim();
    const capacity = Number(row[2] || 0) || 0;
    const sortOrder = Number(row[3] || 0) || 0;
    const isActive = boolFromCell_(row[4]);
    if (!isActive) continue;

    out.push({ resourceId, resourceName, capacity, sortOrder, isActive });
  }

  out.sort((a, b) => (a.sortOrder - b.sortOrder) || a.resourceName.localeCompare(b.resourceName));
  return out;
}

function listBookings_(dateStr) {
  assertDateStr_(dateStr);

  const ss = getSs_();
  const sh = ss.getSheetByName('bookings');
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  // リソースの並び順を反映
  const resources = listResources_();
  const orderMap = {};
  resources.forEach((r, i) => orderMap[r.resourceId] = (Number(r.sortOrder) || (i + 1)));

  const out = [];
  for (let r = 2; r <= values.length; r++) {
    const row = values[r - 1];
    const id = String(row[0] || '').trim();
    if (!id) continue;

    const date = normalizeDateCellToYmd_(row[2]);
    if (date !== dateStr) continue;

    const startTime = normalizeTimeCellToHm_(row[3]);
    const endTime = normalizeTimeCellToHm_(row[4]);

    const pdfFileId = String((row.length >= 11 ? (row[10] || '') : '')).trim();
    const pdfFileName = String((row.length >= 12 ? (row[11] || '') : '')).trim();

    out.push({
      id,
      resourceId: String(row[1] || '').trim(),
      date,
      startTime,
      endTime,
      userName: String(row[5] || '').trim(),
      title: String(row[6] || '').trim(),
      isVisitor: boolFromCell_(row[7]),
      createdAt: formatDateTimeCellAsJst_(row[8]) || String(row[8] || '').trim(),
      updatedAt: formatDateTimeCellAsJst_(row[9]) || String(row[9] || '').trim(),
      pdfFileId,
      pdfFileName,
      pdfUrl: pdfFileId ? getDriveFileUrl_(pdfFileId) : ''
    });
  }

  out.sort((a, b) => {
    const oa = orderMap[a.resourceId] || 999999;
    const ob = orderMap[b.resourceId] || 999999;
    if (oa !== ob) return oa - ob;
    const ra = a.resourceId.localeCompare(b.resourceId);
    if (ra !== 0) return ra;
    return timeToMinutes_(a.startTime) - timeToMinutes_(b.startTime);
  });

  return out;
}

// teacherId + date range
function listBookingsRange_(teacherId, startYmd, endYmd) {
  teacherId = String(teacherId || '').trim();
  assertDateStr_(startYmd);
  assertDateStr_(endYmd);

  const ss = getSs_();
  const sh = ss.getSheetByName('bookings');
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const out = [];
  for (let r = 2; r <= values.length; r++) {
    const row = values[r - 1];
    const id = String(row[0] || '').trim();
    if (!id) continue;

    const rid = String(row[1] || '').trim();
    if (rid !== teacherId) continue;

    const date = normalizeDateCellToYmd_(row[2]);
    if (!date) continue;
    if (date < startYmd || date > endYmd) continue;

    const startTime = normalizeTimeCellToHm_(row[3]);
    const endTime = normalizeTimeCellToHm_(row[4]);

    const pdfFileId = String((row.length >= 11 ? (row[10] || '') : '')).trim();
    const pdfFileName = String((row.length >= 12 ? (row[11] || '') : '')).trim();

    out.push({
      id,
      resourceId: rid,
      date,
      startTime,
      endTime,
      userName: String(row[5] || '').trim(),
      title: String(row[6] || '').trim(),
      isVisitor: boolFromCell_(row[7]),
      createdAt: formatDateTimeCellAsJst_(row[8]) || String(row[8] || '').trim(),
      updatedAt: formatDateTimeCellAsJst_(row[9]) || String(row[9] || '').trim(),
      pdfFileId,
      pdfFileName,
      pdfUrl: pdfFileId ? getDriveFileUrl_(pdfFileId) : ''
    });
  }

  out.sort((a, b) => {
    if (a.date !== b.date) return a.date.localeCompare(b.date);
    return timeToMinutes_(a.startTime) - timeToMinutes_(b.startTime);
  });

  return out;
}

// ================================
// 週関連
// ================================
function parseYmd_(ymd) {
  const m = String(ymd || '').trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  const y = Number(m[1]);
  const mo = Number(m[2]);
  const d = Number(m[3]);
  return new Date(y, mo - 1, d);
}

function formatYmd_(d) {
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}

function addDaysYmd_(ymd, delta) {
  const dt = parseYmd_(ymd);
  if (!dt) throw new Error('date が不正です');
  dt.setDate(dt.getDate() + Number(delta || 0));
  return formatYmd_(dt);
}

function getWeekStartYmd_(anchorYmd, weekStart) {
  const dt = parseYmd_(anchorYmd);
  if (!dt) throw new Error('date が不正です');

  const ws = String(weekStart || 'mon').toLowerCase();
  const startDow = (ws === 'sun') ? 0 : 1; // 0=Sun, 1=Mon

  const dow = dt.getDay();
  let diff = dow - startDow;
  if (diff < 0) diff += 7;
  dt.setDate(dt.getDate() - diff);
  return formatYmd_(dt);
}

function buildWeekDays_(weekStartYmd) {
  const dowJa = ['日','月','火','水','木','金','土'];
  const out = [];
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');

  let cur = parseYmd_(weekStartYmd);
  for (let i = 0; i < 7; i++) {
    const ymd = formatYmd_(cur);
    const md = Utilities.formatDate(cur, TZ, 'M/d');
    out.push({
      date: ymd,
      dow: dowJa[cur.getDay()],
      label: md + '(' + dowJa[cur.getDay()] + ')',
      isToday: (ymd === today)
    });
    cur.setDate(cur.getDate() + 1);
  }
  return out;
}

// ================================
// 正規化（既存データ救済）
// ================================
function pad2_(n) { return String(n).padStart(2, '0'); }

function normalizeDateCellToYmd_(cell) {
  if (cell === null || typeof cell === 'undefined') return '';

  if (Object.prototype.toString.call(cell) === '[object Date]' && !isNaN(cell.getTime())) {
    return Utilities.formatDate(cell, TZ, 'yyyy-MM-dd');
  }

  const s = String(cell).trim();
  if (!s) return '';

  const m1 = s.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})/);
  if (m1) return `${m1[1]}-${pad2_(m1[2])}-${pad2_(m1[3])}`;

  const m2 = s.match(/^(\d{4})年(\d{1,2})月(\d{1,2})日/);
  if (m2) return `${m2[1]}-${pad2_(m2[2])}-${pad2_(m2[3])}`;

  return s;
}

function normalizeTimeCellToHm_(cell) {
  if (cell === null || typeof cell === 'undefined') return '';

  if (Object.prototype.toString.call(cell) === '[object Date]' && !isNaN(cell.getTime())) {
    return Utilities.formatDate(cell, TZ, 'HH:mm');
  }

  if (typeof cell === 'number' && isFinite(cell)) {
    const totalMin = Math.round(cell * 24 * 60);
    const hh = Math.floor(totalMin / 60) % 24;
    const mm = totalMin % 60;
    return `${pad2_(hh)}:${pad2_(mm)}`;
  }

  const s = String(cell).trim();
  if (!s) return '';

  let m = s.match(/^(\d{1,2}):(\d{2})$/);
  if (m) return `${pad2_(m[1])}:${m[2]}`;

  m = s.match(/^(\d{1,2}):(\d{2}):(\d{2})$/);
  if (m) return `${pad2_(m[1])}:${m[2]}`;

  m = s.match(/(\d{1,2}):(\d{2})/);
  if (m) return `${pad2_(m[1])}:${m[2]}`;

  return s;
}

// ================================
// 妥当性・重複チェック
// ================================
function validateConfig_(config) {
  if (!config || typeof config !== 'object') throw new Error('config が不正です');
  if (!config.start || !config.end) throw new Error('config.start/end が必要です');
  assertTimeStr_(config.start);
  assertTimeStr_(config.end);

  const slot = Number(config.slotMinutes);
  if (!Number.isFinite(slot) || slot <= 0) throw new Error('slotMinutes が不正です');

  const openMin = timeToMinutes_(config.start);
  const closeMin = timeToMinutes_(config.end);
  if (!(openMin < closeMin)) throw new Error('config.start < config.end である必要があります');
  if (openMin % slot !== 0 || closeMin % slot !== 0) {
    throw new Error('config.start/end は slotMinutes 単位である必要があります');
  }
}

function normalizeBookingInput_(data, config) {
  if (!data || typeof data !== 'object') throw new Error('data が不正です');

  const resourceId = String(data.resourceId || '').trim();
  const date = String(data.date || '').trim();
  const startTime = String(data.startTime || '').trim();
  const endTime = String(data.endTime || '').trim();
  const userName = String(data.userName || '').trim();
  const title = String(data.title || '').trim();
  const isVisitor = boolFromCell_(data.isVisitor);

  if (!resourceId) throw new Error('resourceId が必要です');
  assertDateStr_(date);
  assertTimeStr_(startTime);
  assertTimeStr_(endTime);
  if (!userName) throw new Error('userName が必要です');
  if (!title) throw new Error('title が必要です');

  const resources = listResources_();
  const exists = resources.some(r => r.resourceId === resourceId);
  if (!exists) throw new Error('resourceId が無効です（resourcesを確認してください）');

  const slot = Number(config.slotMinutes);
  const startMin = timeToMinutes_(startTime);
  const endMin = timeToMinutes_(endTime);
  const openMin = timeToMinutes_(config.start);
  const closeMin = timeToMinutes_(config.end);

  if (!(startMin < endMin)) throw new Error('startTime < endTime である必要があります');
  if (startMin % slot !== 0 || endMin % slot !== 0) throw new Error('start/end は slotMinutes 単位である必要があります');
  if (startMin < openMin || endMin > closeMin) throw new Error('営業時間外です');

  // 今日の「過去時間」は予約禁止（次スロット以降のみ許可）
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  if (date === today) {
    const nowHm = Utilities.formatDate(new Date(), TZ, 'HH:mm');
    const nowMin = timeToMinutes_(nowHm);
    const minStartAllowed = Math.ceil(nowMin / slot) * slot;
    if (startMin < minStartAllowed) {
      throw new Error('過去の時間は予約できません（現在時刻より前）');
    }
  }

  return { resourceId, date, startTime, endTime, userName, title, isVisitor };
}

function assertNoOverlap_(candidate, existingBookings, excludeId) {
  const ns = timeToMinutes_(candidate.startTime);
  const ne = timeToMinutes_(candidate.endTime);

  for (const b of existingBookings) {
    if (excludeId && b.id === excludeId) continue;
    const es = timeToMinutes_(b.startTime);
    const ee = timeToMinutes_(b.endTime);
    if (ns < ee && ne > es) throw new Error('同じ先生または同じ生徒で時間が重なっています');
  }
}

function assertDateStr_(s) {
  const v = String(s || '').trim();
  const m = v.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) throw new Error('date は "YYYY-MM-DD" 形式である必要があります');

  const y = Number(m[1]);
  const mo = Number(m[2]);
  const d = Number(m[3]);

  if (mo < 1 || mo > 12) throw new Error('date の月が不正です');
  if (d < 1 || d > 31) throw new Error('date の日が不正です');

  const dt = new Date(y, mo - 1, d);
  if (dt.getFullYear() !== y || (dt.getMonth() + 1) !== mo || dt.getDate() !== d) {
    throw new Error('date が不正です（実在しない日付）');
  }
}

function assertTimeStr_(s) {
  const v = String(s || '').trim();
  if (!/^\d{2}:\d{2}$/.test(v)) throw new Error('time は "HH:mm" 形式である必要があります');
  const parts = v.split(':');
  const hh = Number(parts[0]);
  const mm = Number(parts[1]);
  if (!Number.isFinite(hh) || !Number.isFinite(mm) || hh < 0 || hh > 23 || mm < 0 || mm > 59) {
    throw new Error('time が不正です');
  }
}

function timeToMinutes_(hhmm) {
  const parts = String(hhmm).trim().split(':');
  const hh = parseInt(parts[0], 10);
  const mm = parseInt(parts[1], 10);
  return (hh * 60) + mm;
}

function nowJst_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
}

function formatDateTimeCellAsJst_(cell) {
  if (!cell) return '';
  if (Object.prototype.toString.call(cell) === '[object Date]' && !isNaN(cell.getTime())) {
    return Utilities.formatDate(cell, TZ, 'yyyy-MM-dd HH:mm:ss');
  }
  return String(cell).trim();
}

function boolFromCell_(v) {
  if (v === true) return true;
  if (v === false) return false;
  const s = String(v || '').trim().toLowerCase();
  return (s === 'true' || s === '1' || s === 'yes' || s === 'y');
}

// ================================
// doPostヘルパ（JSON / form 両対応）
// ================================
function parsePostPayload_(e) {
  const p = {};

  if (e && e.parameter) {
    Object.keys(e.parameter).forEach(k => p[k] = e.parameter[k]);
  }

  if (e && e.postData && e.postData.contents) {
    const ct = String(e.postData.type || '').toLowerCase();
    if (ct.indexOf('application/json') !== -1) {
      try {
        const obj = JSON.parse(e.postData.contents);
        if (obj && typeof obj === 'object') {
          Object.keys(obj).forEach(k => p[k] = obj[k]);
        }
      } catch (err) {}
    }
  }

  if (typeof p.isVisitor === 'string') {
    const v = p.isVisitor.toLowerCase();
    p.isVisitor = (v === 'true' || v === '1' || v === 'yes');
  }
  return p;
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ================================
// レスポンス統一
// ================================
function ok_(data) {
  return { ok: true, data: (typeof data === 'undefined' ? null : data) };
}

function ng_(err) {
  const msg = (err && err.message) ? err.message : String(err);
  return { ok: false, message: msg };
}
