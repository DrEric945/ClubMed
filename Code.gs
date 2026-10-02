/**
 * 學海計劃參與意願登記 — 後端程式（貼到 Apps Script 的 Code.gs）
 * 資料寫入本試算表的「登記資料」工作表，成績截圖存到雲端硬碟資料夾。
 * 只有試算表擁有者（方老師）看得到資料與檔案。
 */

const SHEET_NAME = '登記資料';
const FOLDER_NAME = '學海計劃_成績截圖';
const MAX_FILE_MB = 5;

const FIELDS = [
  ['ts', '最後更新時間'],
  ['nameZh', '中文姓名'],
  ['nameEn', '英文姓名'],
  ['sid', '學號'],
  ['birthday', '出生年月日（西元）'],
  ['phone', '聯絡電話'],
  ['mobile', '手機'],
  ['email', 'Email'],
  ['dept', '就讀系所'],
  ['grade', '年級'],
  ['engScore', '大一英文成績'],
  ['engFile', '大一英文成績截圖'],
  ['gpa', '在校平均成績'],
  ['gpaFile', '在校平均成績截圖'],
  ['newRes', '是否為新住民'],
  ['fatherNat', '父親國籍'],
  ['motherNat', '母親國籍'],
  ['indig', '是否為原住民'],
  ['tribe', '族別'],
  ['pref1', '第一志願'],
  ['pref2', '第二志願'],
  ['pref3', '第三志願'],
  ['pref4', '第四志願'],
  ['pref5', '第五志願'],
  ['created', '首次登記時間'],
];
const KEYS = FIELDS.map(f => f[0]);
const HEADERS = FIELDS.map(f => f[1]);
const REQUIRED = ['nameZh', 'nameEn', 'sid', 'birthday', 'phone', 'mobile', 'email',
  'dept', 'grade', 'newRes', 'indig', 'pref1'];

/** 網頁進入點 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('學海計劃參與意願登記')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 接收 GitHub Pages 網頁送來的資料（查詢與送出） */
function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    if (req.action === 'lookup') {
      out = { ok: true, result: lookup(req.sid, req.birthday, req.mobile) };
    } else if (req.action === 'submit') {
      out = { ok: true, result: submit(req.data || {}) };
    } else {
      throw new Error('未知的操作。');
    }
  } catch (err) {
    out = { ok: false, error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 讓 Index.html 載入 Photos.html（照片檔） */
function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/** 第一次請手動執行這個函式，完成授權並建立工作表與資料夾 */
function setup() {
  getSheet_();
  getFolder_();
}

/** 學生用 學號＋生日＋手機 載入自己的資料 */
function lookup(sid, birthday, mobile) {
  const sh = getSheet_();
  const row = findRow_(sh, normSid_(sid));
  if (!row) throw new Error('查無這個學號的登記資料，請直接填寫下方表單。');
  const o = readRow_(sh, row);
  if (o.birthday !== String(birthday) || digits_(o.mobile) !== digits_(mobile)) {
    throw new Error('學號、出生年月日或手機不正確，請再確認一次。');
  }
  // 不把雲端檔案連結傳回前端，只告訴學生「已上傳過」
  o.engFile = !!o.engFile;
  o.gpaFile = !!o.gpaFile;
  return o;
}

/** 送出或更新登記資料 */
function submit(d) {
  REQUIRED.forEach(k => {
    if (!d[k] || !String(d[k]).trim()) throw new Error('有必填欄位沒有填寫（' + HEADERS[KEYS.indexOf(k)] + '）。');
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_();
    const sid = normSid_(d.sid);
    const rowIdx = findRow_(sh, sid);
    let old = null;

    if (rowIdx) {
      old = readRow_(sh, rowIdx);
      const auth = d.auth || {};
      if (old.birthday !== String(auth.birthday || '') || digits_(old.mobile) !== digits_(auth.mobile || '')) {
        throw new Error('這個學號已經登記過了。若要修改，請先在上方「修改已登記的資料」輸入學號、生日與手機載入資料。');
      }
    }

    const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
    const rec = {};
    KEYS.forEach(k => rec[k] = clean_(d[k]));
    rec.sid = sid;
    rec.ts = now;
    rec.created = old ? old.created : now;

    rec.engFile = rec.engScore
      ? (d.engFileObj ? saveFile_(d.engFileObj, sid, rec.nameZh, '大一英文成績') : (old ? old.engFile : ''))
      : '';
    rec.gpaFile = rec.gpa
      ? (d.gpaFileObj ? saveFile_(d.gpaFileObj, sid, rec.nameZh, '在校平均成績') : (old ? old.gpaFile : ''))
      : '';

    if (rec.newRes !== '是') { rec.fatherNat = ''; rec.motherNat = ''; }
    if (rec.indig !== '是') rec.tribe = '';

    const target = rowIdx || sh.getLastRow() + 1;
    const range = sh.getRange(target, 1, 1, KEYS.length);
    range.setNumberFormat('@'); // 全部存成純文字，避免電話開頭的 0 或日期被改格式
    range.setValues([KEYS.map(k => rec[k] || '')]);
    return { updated: !!rowIdx };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 內部工具 ---------- */

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function getFolder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function findRow_(sh, sid) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const col = KEYS.indexOf('sid') + 1;
  const vals = sh.getRange(2, col, last - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (normSid_(vals[i][0]) === sid) return i + 2;
  }
  return 0;
}

function readRow_(sh, row) {
  const v = sh.getRange(row, 1, 1, KEYS.length).getDisplayValues()[0];
  const o = {};
  KEYS.forEach((k, i) => o[k] = String(v[i]).replace(/^'/, ''));
  return o;
}

function saveFile_(f, sid, name, label) {
  if (!/^image\/|^application\/pdf$/.test(f.mimeType || '')) {
    throw new Error('截圖只接受圖片或 PDF 檔。');
  }
  const bytes = Utilities.base64Decode(f.data);
  if (bytes.length > MAX_FILE_MB * 1024 * 1024) {
    throw new Error('檔案太大了，請小於 ' + MAX_FILE_MB + ' MB。');
  }
  const ext = ((f.name || '').match(/\.[^.]+$/) || [''])[0];
  const blob = Utilities.newBlob(bytes, f.mimeType, sid + '_' + name + '_' + label + ext);
  return getFolder_().createFile(blob).getUrl(); // 檔案預設只有擁有者看得到
}

function normSid_(s) { return String(s || '').trim().toUpperCase(); }
function digits_(s) { return String(s || '').replace(/\D/g, ''); }
function clean_(s) {
  s = String(s == null ? '' : s).trim();
  return /^[=@]/.test(s) ? "'" + s : s; // 防止被試算表當成公式
}
