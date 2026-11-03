/**
 * 學生檢核系統 — Google Apps Script 後端
 *
 * 這個檔案貼到「試算表 > 擴充功能 > Apps Script」的 Code.gs。
 * 第一次使用請先執行 setup()，再到「專案設定 > 指令碼屬性」填入：
 *   CLIENT_ID      Google Cloud 建立的 OAuth 用戶端 ID（xxxx.apps.googleusercontent.com）
 *   TEACHER_EMAILS 教師的 Gmail，可用逗號分隔多個
 *   STUDENT_DOMAIN 學生帳號網域，預設 student.ksvs.kh.edu.tw
 *
 * 部署：部署 > 新增部署作業 > 網頁應用程式
 *   執行身分：我
 *   誰可以存取：所有人
 */

var SHEETS = {
  classes: { name: '班級', headers: ['班級'] },
  students: { name: '學生名單', headers: ['Email', '班級', '座號', '姓名', '登記時間'] },
  assignments: {
    name: '作業',
    headers: ['作業編號', '上課日期', '標題', '說明', '適用班級', '檢核項目', '上傳檔案', '文字回覆', '開放中', '建立時間']
  },
  submissions: {
    name: '繳交狀態',
    headers: ['作業編號', '上課日期', '班級', '座號', '姓名', 'Email', '已勾選項目', '已完成', '文字內容', '檔案連結', '首次提交', '最後更新']
  },
  log: { name: '提交紀錄', headers: ['時間', '作業編號', '班級', '座號', '姓名', 'Email', '動作', '摘要'] }
};

// 上傳檔案 / 文字回覆 的三種設定
var REQ = { NONE: '不需要', OPTIONAL: '選填', REQUIRED: '必填' };
var MAX_FILE_BYTES = 10 * 1024 * 1024;
var MAX_FILES = 5;

/* ======================= 初始化 ======================= */

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var props = PropertiesService.getScriptProperties();
  props.setProperty('SPREADSHEET_ID', ss.getId());
  if (!props.getProperty('STUDENT_DOMAIN')) props.setProperty('STUDENT_DOMAIN', 'student.ksvs.kh.edu.tw');
  if (!props.getProperty('CLIENT_ID')) props.setProperty('CLIENT_ID', '請填入OAuth用戶端ID');
  if (!props.getProperty('TEACHER_EMAILS')) props.setProperty('TEACHER_EMAILS', Session.getEffectiveUser().getEmail());

  Object.keys(SHEETS).forEach(function (key) {
    var def = SHEETS[key];
    var sh = ss.getSheetByName(def.name) || ss.insertSheet(def.name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold').setBackground('#e8f0fe');
      sh.setFrozenRows(1);
    }
  });

  var classSheet = ss.getSheetByName(SHEETS.classes.name);
  if (classSheet.getLastRow() < 2) {
    var sample = ['一年1班', '一年2班', '一年3班', '一年4班', '一年5班', '一年6班', '一年7班'];
    classSheet.getRange(2, 1, sample.length, 1).setValues(sample.map(function (c) { return [c]; }));
  }
  // 讓座號、作業編號等欄位保持文字格式，避免 01 變成 1
  ss.getSheetByName(SHEETS.students.name).getRange('A:D').setNumberFormat('@');
  ss.getSheetByName(SHEETS.assignments.name).getRange('A:A').setNumberFormat('@');
  ss.getSheetByName(SHEETS.submissions.name).getRange('A:F').setNumberFormat('@');

  var def1 = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
  if (def1 && ss.getSheets().length > 1 && def1.getLastRow() === 0) ss.deleteSheet(def1);

  getUploadRoot_();
  Logger.log('初始化完成。請到「專案設定 > 指令碼屬性」確認 CLIENT_ID 與 TEACHER_EMAILS。');
}

/* ======================= HTTP 入口 ======================= */

function doGet() {
  return json_({ ok: true, message: '學生檢核系統 API 運作中' });
}

function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    var user = verifyIdToken_(req.idToken);
    var result = route_(req.action, req.data || {}, user);
    return json_({ ok: true, data: result });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function route_(action, data, user) {
  var studentActions = {
    me: apiMe_,
    register: apiRegister_,
    submit: apiSubmit_
  };
  var teacherActions = {
    adminLoad: apiAdminLoad_,
    saveAssignment: apiSaveAssignment_,
    deleteAssignment: apiDeleteAssignment_,
    saveClasses: apiSaveClasses_,
    saveStudents: apiSaveStudents_
  };
  if (action === 'whoami') return { email: user.email, name: user.name, isTeacher: user.isTeacher, isStudent: user.isStudent };
  if (studentActions[action]) {
    if (!user.isStudent) throw new Error('請使用學校帳號（@' + getProp_('STUDENT_DOMAIN') + '）登入');
    return studentActions[action](data, user);
  }
  if (teacherActions[action]) {
    if (!user.isTeacher) throw new Error('此帳號沒有教師權限');
    return teacherActions[action](data, user);
  }
  throw new Error('未知的操作：' + action);
}

/* ======================= 身分驗證 ======================= */

function verifyIdToken_(idToken) {
  if (!idToken) throw new Error('請先登入');
  var cache = CacheService.getScriptCache();
  var key = 'tok_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, idToken));
  var cached = cache.get(key);
  var info;
  if (cached) {
    info = JSON.parse(cached);
  } else {
    var res = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken), { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw new Error('登入已過期，請重新登入');
    info = JSON.parse(res.getContentText());
    if (info.aud !== getProp_('CLIENT_ID')) throw new Error('登入憑證不符（CLIENT_ID 設定錯誤）');
    if (String(info.email_verified) !== 'true') throw new Error('Email 尚未驗證');
    var ttl = Math.min(600, Number(info.exp) - Math.floor(Date.now() / 1000));
    if (ttl > 30) cache.put(key, JSON.stringify(info), ttl);
  }
  if (Number(info.exp) * 1000 < Date.now()) throw new Error('登入已過期，請重新登入');

  var email = String(info.email).toLowerCase();
  var domain = String(getProp_('STUDENT_DOMAIN') || '').toLowerCase();
  var teachers = String(getProp_('TEACHER_EMAILS') || '').toLowerCase().split(/[,\s]+/).filter(String);
  return {
    email: email,
    name: info.name || email,
    isTeacher: teachers.indexOf(email) >= 0,
    isStudent: !!domain && email.slice(-(domain.length + 1)) === '@' + domain
  };
}

/* ======================= 學生 API ======================= */

function apiMe_(data, user) {
  var student = findStudent_(user.email);
  var res = {
    email: user.email,
    googleName: user.name,
    classes: readClasses_(),
    student: student,
    assignments: []
  };
  if (!student) return res;

  var subs = readRows_(SHEETS.submissions.name).filter(function (r) { return r['Email'] === user.email; });
  var subMap = {};
  subs.forEach(function (s) { subMap[s['作業編號']] = s; });

  res.assignments = readAssignments_()
    .filter(function (a) { return a.open && appliesTo_(a, student.className); })
    .map(function (a) {
      var s = subMap[a.id];
      a.submission = s ? {
        checked: parseJson_(s['已勾選項目'], []),
        done: s['已完成'] === '是',
        text: s['文字內容'] || '',
        files: parseFiles_(s['檔案連結']).map(function (f) { return { name: f.name }; }),
        updatedAt: fmtDate_(s['最後更新'], 'yyyy/MM/dd HH:mm')
      } : null;
      return a;
    })
    .sort(function (a, b) { return b.date.localeCompare(a.date) || b.id.localeCompare(a.id); });
  return res;
}

function apiRegister_(data, user) {
  var className = String(data.className || '').trim();
  var seat = String(data.seat || '').trim();
  var name = String(data.name || '').trim();
  if (readClasses_().indexOf(className) < 0) throw new Error('請選擇班級');
  if (!/^\d{1,3}$/.test(seat)) throw new Error('座號請填數字');
  if (!name) throw new Error('請填姓名');
  seat = ('0' + seat).slice(-2);

  return withLock_(function () {
    var sh = sheet_(SHEETS.students.name);
    var rows = readRows_(SHEETS.students.name);
    var existing = rows.filter(function (r) { return r['Email'] === user.email; })[0];
    var taken = rows.filter(function (r) { return r['班級'] === className && pad2_(r['座號']) === seat && r['Email'] !== user.email; })[0];
    if (taken) throw new Error(className + ' ' + seat + ' 號已被登記，如有錯誤請告訴老師');
    var values = [user.email, className, seat, name, new Date()];
    if (existing) {
      if (existing['班級'] && existing['姓名']) throw new Error('你已經登記過了，如需修改請告訴老師');
      sh.getRange(existing._row, 1, 1, values.length).setValues([values]);
    } else {
      sh.appendRow(values);
    }
    return { ok: true };
  });
}

function apiSubmit_(data, user) {
  var student = findStudent_(user.email);
  if (!student) throw new Error('請先登記班級與座號');
  var a = readAssignments_().filter(function (x) { return x.id === String(data.assignmentId); })[0];
  if (!a) throw new Error('找不到這個作業');
  if (!a.open) throw new Error('這個作業已經關閉');
  if (!appliesTo_(a, student.className)) throw new Error('這個作業不屬於你的班級');

  var checked = (data.checked || []).map(Number).filter(function (i) { return i >= 0 && i < a.items.length; });
  var text = a.textMode === REQ.NONE ? '' : String(data.text || '').slice(0, 5000);
  var newFiles = a.fileMode === REQ.NONE ? [] : (data.files || []);
  if (newFiles.length > MAX_FILES) throw new Error('最多上傳 ' + MAX_FILES + ' 個檔案');

  return withLock_(function () {
    var sh = sheet_(SHEETS.submissions.name);
    var rows = readRows_(SHEETS.submissions.name);
    var existing = rows.filter(function (r) { return r['Email'] === user.email && r['作業編號'] === a.id; })[0];

    var files = existing ? parseFiles_(existing['檔案連結']) : [];
    var keep = data.keepFiles;
    if (Array.isArray(keep)) files = files.filter(function (f) { return keep.indexOf(f.name) >= 0; });
    newFiles.forEach(function (f) {
      var bytes = Utilities.base64Decode(f.base64);
      if (bytes.length > MAX_FILE_BYTES) throw new Error('檔案 ' + f.name + ' 超過 10MB');
      var folder = getFolder_(a.id + '_' + a.title, student.className);
      var fileName = student.className + '_' + student.seat + '_' + student.name + '_' + f.name;
      var file = folder.createFile(Utilities.newBlob(bytes, f.type || 'application/octet-stream', fileName));
      files.push({ name: f.name, url: file.getUrl() });
    });
    if (files.length > MAX_FILES) files = files.slice(-MAX_FILES);

    var done = checked.length === a.items.length &&
      (a.fileMode !== REQ.REQUIRED || files.length > 0) &&
      (a.textMode !== REQ.REQUIRED || text.trim().length > 0);

    var now = new Date();
    var row = [
      a.id, a.date, student.className, student.seat, student.name, user.email,
      JSON.stringify(checked), done ? '是' : '否', text,
      files.map(function (f) { return f.name + ' | ' + f.url; }).join('\n'),
      existing ? existing['首次提交'] : now, now
    ];
    if (existing) sh.getRange(existing._row, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);

    sheet_(SHEETS.log.name).appendRow([
      now, a.id, student.className, student.seat, student.name, user.email,
      existing ? '更新' : '首次提交',
      '勾選 ' + checked.length + '/' + a.items.length + (files.length ? '，檔案 ' + files.length : '') + (text ? '，有文字' : '') + (done ? '，已完成' : '')
    ]);
    return { done: done };
  });
}

/* ======================= 教師 API ======================= */

function apiAdminLoad_() {
  var subs = readRows_(SHEETS.submissions.name).map(function (s) {
    return {
      id: s['作業編號'],
      email: s['Email'],
      className: s['班級'],
      seat: pad2_(s['座號']),
      name: s['姓名'],
      checked: parseJson_(s['已勾選項目'], []),
      done: s['已完成'] === '是',
      text: s['文字內容'] || '',
      files: parseFiles_(s['檔案連結']),
      firstAt: fmtDate_(s['首次提交'], 'yyyy/MM/dd HH:mm'),
      updatedAt: fmtDate_(s['最後更新'], 'yyyy/MM/dd HH:mm')
    };
  });
  var students = readRows_(SHEETS.students.name).map(function (r) {
    return { email: r['Email'], className: r['班級'], seat: pad2_(r['座號']), name: r['姓名'] };
  }).filter(function (s) { return s.email; });
  return {
    classes: readClasses_(),
    students: students,
    assignments: readAssignments_(),
    submissions: subs,
    spreadsheetUrl: SpreadsheetApp.openById(getProp_('SPREADSHEET_ID')).getUrl(),
    serverTime: fmtDate_(new Date(), 'yyyy/MM/dd HH:mm:ss')
  };
}

function apiSaveAssignment_(data) {
  var id = String(data.id || '').trim();
  if (!id) throw new Error('請填作業編號');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(data.date || ''))) throw new Error('請選擇上課日期');
  var items = (data.items || []).map(function (s) { return String(s).trim(); }).filter(String);
  if (!items.length) throw new Error('至少要有一個檢核項目');
  var modes = [REQ.NONE, REQ.OPTIONAL, REQ.REQUIRED];
  var fileMode = modes.indexOf(data.fileMode) >= 0 ? data.fileMode : REQ.NONE;
  var textMode = modes.indexOf(data.textMode) >= 0 ? data.textMode : REQ.NONE;
  var classes = (data.classes || []).filter(function (c) { return readClasses_().indexOf(c) >= 0; });
  if (!classes.length) throw new Error('請至少選擇一個班級');

  return withLock_(function () {
    var sh = sheet_(SHEETS.assignments.name);
    var rows = readRows_(SHEETS.assignments.name);
    var origId = String(data.originalId || '');
    var existing = rows.filter(function (r) { return String(r['作業編號']) === (origId || id); })[0];
    if (!origId && existing) throw new Error('作業編號 ' + id + ' 已存在');
    if (origId && origId !== id && rows.some(function (r) { return String(r['作業編號']) === id; })) throw new Error('作業編號 ' + id + ' 已存在');
    var row = [id, data.date, String(data.title || ''), String(data.description || ''), classes.join(','),
      items.join('\n'), fileMode, textMode, data.open ? '是' : '否', existing ? existing['建立時間'] : new Date()];
    if (existing) sh.getRange(existing._row, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
    if (existing && origId !== id) {
      // 作業編號改名時，同步更新繳交狀態
      var subSh = sheet_(SHEETS.submissions.name);
      readRows_(SHEETS.submissions.name).forEach(function (s) {
        if (String(s['作業編號']) === origId) subSh.getRange(s._row, 1).setValue(id);
      });
    }
    if (existing && fmtDate_(existing['上課日期'], 'yyyy-MM-dd') !== data.date) {
      var subSh2 = sheet_(SHEETS.submissions.name);
      readRows_(SHEETS.submissions.name).forEach(function (s) {
        if (String(s['作業編號']) === id) subSh2.getRange(s._row, 2).setValue(data.date);
      });
    }
    return { ok: true };
  });
}

function apiDeleteAssignment_(data) {
  return withLock_(function () {
    var rows = readRows_(SHEETS.assignments.name);
    var r = rows.filter(function (x) { return String(x['作業編號']) === String(data.id); })[0];
    if (!r) throw new Error('找不到作業');
    sheet_(SHEETS.assignments.name).deleteRow(r._row);
    return { ok: true };
  });
}

function apiSaveClasses_(data) {
  var list = (data.classes || []).map(function (c) { return String(c).trim(); }).filter(String);
  if (!list.length) throw new Error('至少要有一個班級');
  return withLock_(function () {
    var sh = sheet_(SHEETS.classes.name);
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).clearContent();
    sh.getRange(2, 1, list.length, 1).setValues(list.map(function (c) { return [c]; }));
    return { ok: true };
  });
}

/** 匯入或修改學生名單：data.students = [{email, className, seat, name}]，data.replaceClass 選填 */
function apiSaveStudents_(data) {
  var classes = readClasses_();
  var incoming = (data.students || []).map(function (s) {
    return {
      email: String(s.email || '').trim().toLowerCase(),
      className: String(s.className || '').trim(),
      seat: pad2_(s.seat),
      name: String(s.name || '').trim()
    };
  }).filter(function (s) { return s.email; });
  incoming.forEach(function (s) {
    if (classes.indexOf(s.className) < 0) throw new Error(s.email + ' 的班級「' + s.className + '」不在班級清單中');
  });
  var remove = (data.remove || []).map(function (e) { return String(e).toLowerCase(); });

  return withLock_(function () {
    var sh = sheet_(SHEETS.students.name);
    var rows = readRows_(SHEETS.students.name);
    var map = {};
    rows.forEach(function (r) { map[r['Email']] = r; });
    incoming.forEach(function (s) {
      var vals = [s.email, s.className, s.seat, s.name, (map[s.email] && map[s.email]['登記時間']) || new Date()];
      if (map[s.email]) sh.getRange(map[s.email]._row, 1, 1, 5).setValues([vals]);
      else { sh.appendRow(vals); map[s.email] = { _row: sh.getLastRow(), '登記時間': vals[4] }; }
    });
    rows.filter(function (r) { return remove.indexOf(r['Email']) >= 0; })
      .sort(function (a, b) { return b._row - a._row; })
      .forEach(function (r) { sh.deleteRow(r._row); });
    return { ok: true, count: incoming.length };
  });
}

/* ======================= 資料存取 ======================= */

function ss_() {
  var id = getProp_('SPREADSHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('找不到工作表「' + name + '」，請先執行 setup()');
  return sh;
}

function readRows_(name) {
  var sh = sheet_(name);
  var values = sh.getDataRange().getValues();
  var headers = values.shift() || [];
  return values.map(function (row, i) {
    var o = { _row: i + 2 };
    headers.forEach(function (h, j) {
      var v = row[j];
      o[h] = (v instanceof Date) ? v : (h === 'Email' ? String(v).trim().toLowerCase() : (typeof v === 'string' ? v.trim() : v));
    });
    return o;
  }).filter(function (o) { return headers.some(function (h) { return o[h] !== '' && o[h] !== null; }); });
}

function readClasses_() {
  return readRows_(SHEETS.classes.name).map(function (r) { return String(r['班級']); }).filter(String);
}

function readAssignments_() {
  return readRows_(SHEETS.assignments.name).map(function (r) {
    return {
      id: String(r['作業編號']),
      date: fmtDate_(r['上課日期'], 'yyyy-MM-dd'),
      title: String(r['標題'] || ''),
      description: String(r['說明'] || ''),
      classes: String(r['適用班級'] || '').split(',').map(function (s) { return s.trim(); }).filter(String),
      items: String(r['檢核項目'] || '').split('\n').map(function (s) { return s.trim(); }).filter(String),
      fileMode: r['上傳檔案'] || REQ.NONE,
      textMode: r['文字回覆'] || REQ.NONE,
      open: r['開放中'] === '是' || r['開放中'] === true
    };
  });
}

function findStudent_(email) {
  var r = readRows_(SHEETS.students.name).filter(function (x) { return x['Email'] === email; })[0];
  if (!r || !r['班級'] || !r['姓名']) return null;
  return { email: email, className: String(r['班級']), seat: pad2_(r['座號']), name: String(r['姓名']) };
}

function appliesTo_(a, className) {
  return a.classes.indexOf(className) >= 0;
}

/* ======================= 雲端硬碟 ======================= */

function getUploadRoot_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('UPLOAD_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* 資料夾被刪掉就重建 */ }
  }
  var folder = DriveApp.createFolder('學生檢核_上傳檔案');
  props.setProperty('UPLOAD_FOLDER_ID', folder.getId());
  return folder;
}

function getFolder_(assignmentName, className) {
  var cache = CacheService.getScriptCache();
  var key = 'fld_' + Utilities.base64EncodeWebSafe(assignmentName + '/' + className, Utilities.Charset.UTF_8);
  var id = cache.get(key);
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* ignore */ } }
  var a = childFolder_(getUploadRoot_(), assignmentName);
  var c = childFolder_(a, className);
  cache.put(key, c.getId(), 21600);
  return c;
}

function childFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/* ======================= 小工具 ======================= */

function getProp_(k) { return PropertiesService.getScriptProperties().getProperty(k); }

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('系統忙碌中，請稍後再試');
  try { return fn(); } finally { lock.releaseLock(); }
}

function fmtDate_(v, pattern) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), pattern);
  return String(v);
}

function pad2_(v) {
  var s = String(v === undefined || v === null ? '' : v).trim();
  return /^\d$/.test(s) ? '0' + s : s;
}

function parseJson_(s, fallback) {
  try { return s ? JSON.parse(s) : fallback; } catch (e) { return fallback; }
}

function parseFiles_(s) {
  return String(s || '').split('\n').filter(String).map(function (line) {
    var i = line.lastIndexOf(' | ');
    return i >= 0 ? { name: line.slice(0, i), url: line.slice(i + 3) } : { name: line, url: line };
  });
}

/** 讓 Apps Script 一次要求所有需要的權限（setup 後可手動執行一次） */
function authorize() {
  UrlFetchApp.fetch('https://www.google.com');
  DriveApp.getRootFolder();
  SpreadsheetApp.getActiveSpreadsheet();
}
