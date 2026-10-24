/**
 * MoneyPro — Data layer
 * --------------------------------------------------------------
 * แต่ละตาราง = 1 ชีต แถวแรกเป็นหัวคอลัมน์
 * ฟังก์ชันที่ลงท้ายด้วย _ เป็นฟังก์ชันภายใน หน้าเว็บเรียกตรงไม่ได้
 */

const APP_NAME = 'MoneyPro';

const TABLES = {
  Transactions: ['ID', 'Date', 'Type', 'Category', 'Amount', 'Note', 'CreatedAt', 'UpdatedAt', 'Account', 'RecurringId', 'Attachment', 'ToAccount'],
  Accounts:     ['ID', 'Name', 'Type', 'OpeningBalance', 'Color', 'Archived', 'SortOrder'],
  Categories:   ['ID', 'Name', 'Type', 'Icon', 'Color', 'MonthlyBudget', 'Group', 'Archived', 'SortOrder'],
  Recurring:    ['ID', 'Type', 'Category', 'Account', 'Amount', 'DayOfMonth', 'Note', 'Active', 'StartDate', 'LastRun', 'CreatedAt', 'ToAccount'],
  Goals:        ['ID', 'Name', 'TargetAmount', 'Account', 'TargetDate', 'Color', 'Archived', 'CreatedAt'],
  Settings:     ['Key', 'Value'],
};

/** คอลัมน์ที่เก็บเป็นข้อความ (กัน Sheets แปลงเป็นวันที่) */
const TEXT_COLUMNS = { Transactions: ['Date'], Recurring: ['StartDate', 'LastRun'], Goals: ['TargetDate'] };
const NUMBER_COLUMNS = {
  Transactions: ['Amount'], Accounts: ['OpeningBalance', 'SortOrder'],
  Categories: ['MonthlyBudget', 'SortOrder'], Recurring: ['Amount', 'DayOfMonth'], Goals: ['TargetAmount'],
};
const BOOL_COLUMNS = { Accounts: ['Archived'], Categories: ['Archived'], Recurring: ['Active'], Goals: ['Archived'] };

const SCHEMA_VERSION = 2;
const CACHE_TTL = 600; // วินาที

const DEFAULT_ACCOUNTS = [
  { id: 'acc_cash', name: 'เงินสด', type: 'cash', openingBalance: 0, color: '#10B981', sortOrder: 1 },
  { id: 'acc_bank', name: 'บัญชีธนาคาร', type: 'bank', openingBalance: 0, color: '#4F46E5', sortOrder: 2 },
];

/**
 * Group ใช้กับการวิเคราะห์แบบ 50/30/20
 *   need = ค่าใช้จ่ายจำเป็น, want = ค่าใช้จ่ายตามความต้องการ, income = รายรับ
 */
const DEFAULT_CATEGORIES = [
  { id: 'food',          name: 'อาหารและเครื่องดื่ม',  type: 'expense', icon: 'utensils',        color: '#F97316', group: 'need' },
  { id: 'housing',       name: 'ที่พักและสาธารณูปโภค', type: 'expense', icon: 'home',            color: '#4F46E5', group: 'need' },
  { id: 'transport',     name: 'เดินทาง',              type: 'expense', icon: 'car',             color: '#0EA5E9', group: 'need' },
  { id: 'bills',         name: 'บิลและค่าสมาชิก',       type: 'expense', icon: 'receipt',         color: '#14B8A6', group: 'need' },
  { id: 'health',        name: 'สุขภาพ',               type: 'expense', icon: 'heart-pulse',     color: '#EF4444', group: 'need' },
  { id: 'shopping',      name: 'ช้อปปิ้ง',             type: 'expense', icon: 'shopping-bag',    color: '#EC4899', group: 'want' },
  { id: 'entertainment', name: 'บันเทิงและท่องเที่ยว',  type: 'expense', icon: 'clapperboard',    color: '#A855F7', group: 'want' },
  { id: 'education',     name: 'การศึกษาและพัฒนาตนเอง', type: 'expense', icon: 'graduation-cap',  color: '#EAB308', group: 'want' },
  { id: 'other_exp',     name: 'รายจ่ายอื่นๆ',          type: 'expense', icon: 'circle-ellipsis', color: '#64748B', group: 'want' },
  { id: 'salary',        name: 'เงินเดือน',            type: 'income',  icon: 'briefcase',       color: '#10B981', group: 'income' },
  { id: 'bonus',         name: 'โบนัส',                type: 'income',  icon: 'award',           color: '#22C55E', group: 'income' },
  { id: 'side',          name: 'รายได้เสริม',           type: 'income',  icon: 'laptop',          color: '#84CC16', group: 'income' },
  { id: 'invest_inc',    name: 'ดอกเบี้ย/เงินปันผล',    type: 'income',  icon: 'trending-up',     color: '#06B6D4', group: 'income' },
  { id: 'other_inc',     name: 'รายรับอื่นๆ',           type: 'income',  icon: 'plus-circle',     color: '#94A3B8', group: 'income' },
];

/* ==============================================================
 * Sheet access
 * ============================================================ */

function ss_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('สคริปต์นี้ต้องผูกกับ Google Sheet (สร้างจาก Extensions → Apps Script ในชีต)');
  return ss;
}

const SHEET_MEMO_ = {};

function sheet_(name) {
  if (!SHEET_MEMO_[name]) SHEET_MEMO_[name] = openSheet_(name);
  return SHEET_MEMO_[name];
}

function openSheet_(name) {
  const headers = TABLES[name];
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    formatSheet_(sh, name);
    return sh;
  }
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const current = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  if (current.length < headers.length || headers.some((h, i) => current[i] !== h)) formatSheet_(sh, name);
  return sh;
}

function formatSheet_(sh, name) {
  const headers = TABLES[name];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#0F172A').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  (TEXT_COLUMNS[name] || []).forEach((h) => {
    sh.getRange(1, headers.indexOf(h) + 1, sh.getMaxRows(), 1).setNumberFormat('@');
  });
  (NUMBER_COLUMNS[name] || []).forEach((h) => {
    sh.getRange(2, headers.indexOf(h) + 1, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('#,##0.00');
  });
}

/** 'OpeningBalance' → 'openingBalance', 'ID' → 'id' */
function key_(header) {
  return header === 'ID' ? 'id' : header.charAt(0).toLowerCase() + header.slice(1);
}

function rowToObj_(name, row) {
  const o = {};
  TABLES[name].forEach((h, i) => {
    let v = row[i];
    if (v instanceof Date) v = (TEXT_COLUMNS[name] || []).indexOf(h) > -1 ? fmtDate_(v) : v.toISOString();
    if ((NUMBER_COLUMNS[name] || []).indexOf(h) > -1) v = round2_(Number(v) || 0);
    else if ((BOOL_COLUMNS[name] || []).indexOf(h) > -1) v = v === true || String(v).toUpperCase() === 'TRUE';
    else v = v === null || v === undefined ? '' : String(v);
    o[key_(h)] = v;
  });
  return o;
}

function objToRow_(name, o) {
  return TABLES[name].map((h) => {
    const v = o[key_(h)];
    if (v === undefined || v === null) return '';
    return typeof v === 'string' ? safeCell_(v) : v;
  });
}

/* ==============================================================
 * Cache (CacheService แบ่งเป็นชิ้น ชิ้นละไม่เกิน ~100KB)
 * ============================================================ */

function cacheGet_(k) {
  try {
    const c = CacheService.getScriptCache();
    const n = Number(c.get(k + ':n'));
    if (!n) return null;
    const keys = [];
    for (let i = 0; i < n; i++) keys.push(k + ':' + i);
    const parts = c.getAll(keys);
    let s = '';
    for (let i = 0; i < n; i++) {
      if (parts[keys[i]] === undefined) return null;
      s += parts[keys[i]];
    }
    return JSON.parse(s);
  } catch (e) {
    return null;
  }
}

function cachePut_(k, value) {
  try {
    const s = JSON.stringify(value);
    const size = 90000;
    const map = {};
    let n = 0;
    for (let i = 0; i < s.length; i += size) map[k + ':' + n++] = s.slice(i, i + size);
    if (n > 80) return;
    map[k + ':n'] = String(n);
    CacheService.getScriptCache().putAll(map, CACHE_TTL);
  } catch (e) { /* ไม่มี cache ก็ทำงานได้ */ }
}

function cacheDrop_(name) {
  try { CacheService.getScriptCache().remove('tbl:' + name + ':n'); } catch (e) {}
}

/* ==============================================================
 * Table operations
 * ============================================================ */

function tRead_(name, noCache) {
  const ck = 'tbl:' + name;
  if (!noCache) {
    const hit = cacheGet_(ck);
    if (hit) return hit;
  }
  const sh = sheet_(name);
  const last = sh.getLastRow();
  let list = [];
  if (last >= 2) {
    list = sh.getRange(2, 1, last - 1, TABLES[name].length).getValues()
      .filter((r) => r[0] !== '' && r[0] !== null)
      .map((r) => rowToObj_(name, r));
  }
  cachePut_(ck, list);
  return list;
}

function tFindRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function tInsert_(name, obj) {
  sheet_(name).appendRow(objToRow_(name, obj));
  cacheDrop_(name);
  return obj;
}

function tInsertMany_(name, objs) {
  if (!objs.length) return;
  const sh = sheet_(name);
  sh.getRange(sh.getLastRow() + 1, 1, objs.length, TABLES[name].length).setValues(objs.map((o) => objToRow_(name, o)));
  cacheDrop_(name);
}

function tGet_(name, id) {
  const sh = sheet_(name);
  const r = tFindRow_(sh, id);
  if (r < 0) return null;
  return rowToObj_(name, sh.getRange(r, 1, 1, TABLES[name].length).getValues()[0]);
}

function tUpdate_(name, id, patch) {
  const sh = sheet_(name);
  const r = tFindRow_(sh, id);
  if (r < 0) throw new Error('ไม่พบรายการ (อาจถูกลบไปแล้ว)');
  const range = sh.getRange(r, 1, 1, TABLES[name].length);
  const next = Object.assign(rowToObj_(name, range.getValues()[0]), patch);
  next.id = String(id);
  range.setValues([objToRow_(name, next)]);
  cacheDrop_(name);
  return next;
}

function tDelete_(name, id) {
  const sh = sheet_(name);
  const r = tFindRow_(sh, id);
  if (r < 0) throw new Error('ไม่พบรายการ (อาจถูกลบไปแล้ว)');
  sh.deleteRow(r);
  cacheDrop_(name);
}

/* ==============================================================
 * Settings
 * ============================================================ */

function getSettingsMap_() {
  const m = {};
  tRead_('Settings').forEach((r) => { m[r.key] = r.value; });
  return m;
}

function setSetting_(k, v) {
  const sh = sheet_('Settings');
  const r = tFindRow_(sh, k);
  if (r < 0) sh.appendRow([k, String(v)]);
  else sh.getRange(r, 2).setValue(String(v));
  cacheDrop_('Settings');
}

/* ==============================================================
 * Setup
 * ============================================================ */

function ensureSetup_() {
  if (Number(getSettingsMap_().schemaVersion) >= SCHEMA_VERSION) return;
  withLock_(() => {
    Object.keys(TABLES).forEach((t) => sheet_(t));
    if (!tRead_('Accounts', true).length) {
      tInsertMany_('Accounts', DEFAULT_ACCOUNTS.map((a) => Object.assign({ archived: false }, a)));
    }
    const cats = tRead_('Categories', true);
    const have = {};
    cats.forEach((c) => { have[c.id] = true; });
    tInsertMany_('Categories', DEFAULT_CATEGORIES.filter((c) => !have[c.id])
      .map((c, i) => Object.assign({ monthlyBudget: 0, archived: false, sortOrder: cats.length + i + 1 }, c)));
    const s = getSettingsMap_();
    if (s.backupEnabled === undefined) setSetting_('backupEnabled', 'false');
    setSetting_('schemaVersion', SCHEMA_VERSION);
  });
}

/* ==============================================================
 * Utils
 * ============================================================ */

function tz_() {
  return Session.getScriptTimeZone() || 'Asia/Bangkok';
}

function today_() {
  return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
}

function fmtDate_(d) {
  return Utilities.formatDate(d, tz_(), 'yyyy-MM-dd');
}

function newId_(prefix) {
  return (prefix || 'TX') + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase();
}

function round2_(n) {
  return Math.round(n * 100) / 100;
}

function safeCell_(s) {
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function isDate_(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(s || ''));
}

function pad2_(n) {
  return String(n).padStart(2, '0');
}

function daysInMonth_(y, m) {
  return new Date(y, m, 0).getDate();
}

/** วันถัดไปที่ตรงกับวันที่ day (1–31) หลัง afterDate โดยปัดวันที่เกินให้เป็นสิ้นเดือน */
function nextOccurrence_(afterDate, day) {
  let y = Number(afterDate.slice(0, 4));
  let m = Number(afterDate.slice(5, 7));
  for (let i = 0; i < 3; i++) {
    const cand = y + '-' + pad2_(m) + '-' + pad2_(Math.min(day, daysInMonth_(y, m)));
    if (cand > afterDate) return cand;
    m++;
    if (m > 12) { m = 1; y++; }
  }
  return null;
}

function prevDay_(date) {
  const d = new Date(Number(date.slice(0, 4)), Number(date.slice(5, 7)) - 1, Number(date.slice(8, 10)) - 1);
  return d.getFullYear() + '-' + pad2_(d.getMonth() + 1) + '-' + pad2_(d.getDate());
}
