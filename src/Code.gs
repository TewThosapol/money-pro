/**
 * MoneyPro — Public API
 * --------------------------------------------------------------
 * ฟังก์ชันในไฟล์นี้เรียกจากหน้าเว็บผ่าน google.script.run
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** รันครั้งเดียวจาก Editor: สร้างชีตทั้งหมดและติดตั้งงานอัตโนมัติ */
function setup() {
  ensureSetup_();
  installTriggers();
  return 'ตั้งค่าเรียบร้อย';
}

/* ==============================================================
 * Bootstrap & data
 * ============================================================ */

/** ข้อมูลทั้งหมดที่หน้าเว็บต้องใช้ตอนเปิด */
function bootstrap() {
  ensureSetup_();
  let generated = 0;
  try { generated = processRecurring_(); } catch (e) { console.error(e); }
  return Object.assign(referenceData_(), getAppData(), { recurringGenerated: generated });
}

function referenceData_() {
  const s = getSettingsMap_();
  return {
    today: today_(),
    timezone: tz_(),
    accounts: tRead_('Accounts'),
    categories: tRead_('Categories'),
    recurring: tRead_('Recurring'),
    goals: tRead_('Goals'),
    settings: {
      backupEnabled: s.backupEnabled === 'true',
      lastBackup: s.lastBackup || '',
      lastBackupUrl: s.lastBackupUrl || '',
      receiptFolderUrl: s.receiptFolderId ? 'https://drive.google.com/drive/folders/' + s.receiptFolderId : '',
    },
  };
}

/** รายการทั้งหมด + ยอดคงเหลือแต่ละบัญชี (หน้าเว็บคำนวณการวิเคราะห์เอง) */
function getAppData() {
  const all = allTransactions_();
  return { transactions: all, balances: accountBalances_(all) };
}

/**
 * ดึงรายการพร้อมตัวกรอง
 * @param {{from?, to?, type?, account?, category?, q?}} f
 */
function getTransactions(f) {
  f = f || {};
  const q = String(f.q || '').trim().toLowerCase();
  const names = {};
  tRead_('Categories').forEach((c) => { names[c.id] = c.name.toLowerCase(); });
  return allTransactions_().filter((t) =>
    (!f.from || t.date >= f.from) && (!f.to || t.date <= f.to) &&
    (!f.type || f.type === 'all' || t.type === f.type) &&
    (!f.account || t.account === f.account) &&
    (!f.category || t.category === f.category) &&
    (!q || t.note.toLowerCase().indexOf(q) > -1 || (names[t.category] || '').indexOf(q) > -1 || String(t.amount).indexOf(q) > -1)
  ).sort(byDateDesc_);
}

/** สรุปตัวเลขในช่วงวันที่ (สำหรับเรียกใช้ภายนอก หน้าเว็บคำนวณเองจาก getAppData) */
function getDashboardSummary(range) {
  range = range || {};
  const list = getTransactions({ from: range.from, to: range.to });
  let income = 0, expense = 0;
  const byCategory = {};
  list.forEach((t) => {
    if (t.type === 'transfer') return;
    if (t.type === 'income') income += t.amount;
    else {
      expense += t.amount;
      byCategory[t.category] = round2_((byCategory[t.category] || 0) + t.amount);
    }
  });
  return {
    income: round2_(income), expense: round2_(expense), net: round2_(income - expense),
    savingsRate: income > 0 ? round2_(((income - expense) / income) * 100) : null,
    byCategory: byCategory, balances: accountBalances_(allTransactions_()), count: list.length,
  };
}

/* ==============================================================
 * Transactions CRUD
 * ============================================================ */

function addTransaction(data) {
  const c = validateTx_(data);
  return withLock_(() => {
    const now = new Date().toISOString();
    return tInsert_('Transactions', Object.assign(c, { id: newId_('TX'), createdAt: now, updatedAt: now, recurringId: '' }));
  });
}

function updateTransaction(data) {
  if (!data || !data.id) throw new Error('ไม่พบรหัสรายการ');
  const c = validateTx_(data);
  return withLock_(() => {
    const before = tGet_('Transactions', data.id);
    const saved = tUpdate_('Transactions', data.id, Object.assign(c, { updatedAt: new Date().toISOString() }));
    if (before && before.attachment && before.attachment !== saved.attachment) trashFile_(before.attachment);
    return saved;
  });
}

function deleteTransaction(id) {
  if (!id) throw new Error('ไม่พบรหัสรายการ');
  withLock_(() => {
    const before = tGet_('Transactions', id);
    tDelete_('Transactions', id);
    if (before && before.attachment) trashFile_(before.attachment);
  });
  return { ok: true, id: id };
}

/* ==============================================================
 * Accounts
 * ============================================================ */

const ACCOUNT_TYPES = ['cash', 'bank', 'credit', 'ewallet', 'saving', 'invest'];

function saveAccount(data) {
  data = data || {};
  const name = String(data.name || '').trim().slice(0, 60);
  if (!name) throw new Error('กรุณาตั้งชื่อบัญชี');
  const clean = {
    name: name,
    type: ACCOUNT_TYPES.indexOf(data.type) > -1 ? data.type : 'bank',
    openingBalance: round2_(Number(data.openingBalance) || 0),
    color: hexColor_(data.color, '#4F46E5'),
    archived: data.archived === true,
  };
  return withLock_(() => {
    if (data.id) return tUpdate_('Accounts', data.id, clean);
    const n = tRead_('Accounts', true).length;
    return tInsert_('Accounts', Object.assign(clean, { id: newId_('ACC'), sortOrder: n + 1 }));
  });
}

/** ลบได้เฉพาะบัญชีที่ยังไม่มีรายการ ถ้ามีแล้วจะเก็บเข้าคลัง */
function deleteAccount(id) {
  return withLock_(() => {
    if (!tRead_('Accounts', true).some((a) => !a.archived && a.id !== id)) {
      throw new Error('ต้องมีบัญชีที่ใช้งานอยู่อย่างน้อย 1 บัญชี');
    }
    const used = allTransactions_().some((t) => t.account === id || t.toAccount === id) ||
      tRead_('Recurring', true).some((r) => r.account === id || r.toAccount === id) ||
      tRead_('Goals', true).some((g) => g.account === id && !g.archived);
    if (used) {
      tUpdate_('Accounts', id, { archived: true });
      return { archived: true };
    }
    tDelete_('Accounts', id);
    return { deleted: true };
  });
}

/* ==============================================================
 * Categories & budgets
 * ============================================================ */

function saveCategory(data) {
  data = data || {};
  const name = String(data.name || '').trim().slice(0, 60);
  if (!name) throw new Error('กรุณาตั้งชื่อหมวดหมู่');
  const type = data.type === 'income' ? 'income' : 'expense';
  const clean = {
    name: name, type: type,
    icon: /^[a-z0-9-]{2,40}$/.test(String(data.icon || '')) ? data.icon : 'tag',
    color: hexColor_(data.color, '#64748B'),
    monthlyBudget: Math.max(0, round2_(Number(data.monthlyBudget) || 0)),
    group: type === 'income' ? 'income' : (data.group === 'need' ? 'need' : 'want'),
    archived: data.archived === true,
  };
  return withLock_(() => {
    if (data.id) {
      const cur = tGet_('Categories', data.id);
      if (cur && cur.type !== type && allTransactions_().some((t) => t.category === data.id)) {
        throw new Error('เปลี่ยนประเภทของหมวดที่มีรายการแล้วไม่ได้');
      }
      return tUpdate_('Categories', data.id, clean);
    }
    const n = tRead_('Categories', true).length;
    return tInsert_('Categories', Object.assign(clean, { id: newId_('CAT'), sortOrder: n + 1 }));
  });
}

/** ตั้งงบรายเดือนของหมวด (0 = ไม่ตั้งงบ) */
function setBudget(categoryId, amount) {
  const v = Math.max(0, round2_(Number(amount) || 0));
  return withLock_(() => tUpdate_('Categories', categoryId, { monthlyBudget: v }));
}

function deleteCategory(id) {
  return withLock_(() => {
    const used = allTransactions_().some((t) => t.category === id) || tRead_('Recurring', true).some((r) => r.category === id);
    if (used) {
      tUpdate_('Categories', id, { archived: true });
      return { archived: true };
    }
    tDelete_('Categories', id);
    return { deleted: true };
  });
}

/* ==============================================================
 * Savings goals
 * ============================================================ */

/**
 * เป้าหมายการออม ความคืบหน้าคิดจากยอดคงเหลือของบัญชีที่ผูกไว้
 * account ว่าง = ใช้ยอดรวมของบัญชีออมทรัพย์และการลงทุนทั้งหมด
 */
function saveGoal(data) {
  data = data || {};
  const name = String(data.name || '').trim().slice(0, 60);
  if (!name) throw new Error('กรุณาตั้งชื่อเป้าหมาย');
  const target = round2_(Number(data.targetAmount));
  if (!(target > 0)) throw new Error('เป้าหมายต้องมากกว่า 0');
  const account = String(data.account || '');
  if (account && !tRead_('Accounts').some((a) => a.id === account)) throw new Error('ไม่พบบัญชีที่เลือก');
  const targetDate = isDate_(data.targetDate) ? data.targetDate : '';
  const clean = { name: name, targetAmount: target, account: account, targetDate: targetDate,
    color: hexColor_(data.color, '#0EA5E9'), archived: data.archived === true };
  return withLock_(() => {
    if (data.id) return tUpdate_('Goals', data.id, clean);
    return tInsert_('Goals', Object.assign(clean, { id: newId_('GOAL'), createdAt: new Date().toISOString() }));
  });
}

function deleteGoal(id) {
  withLock_(() => tDelete_('Goals', id));
  return { ok: true };
}

/* ==============================================================
 * Recurring
 * ============================================================ */

function saveRecurring(data) {
  data = data || {};
  const tx = validateTx_({
    date: data.startDate, type: data.type, category: data.category,
    account: data.account, toAccount: data.toAccount, amount: data.amount, note: data.note,
  });
  const day = Math.round(Number(data.dayOfMonth));
  if (!(day >= 1 && day <= 31)) throw new Error('วันที่ของเดือนต้องอยู่ระหว่าง 1–31');
  const clean = {
    type: tx.type, category: tx.category, account: tx.account, toAccount: tx.toAccount, amount: tx.amount, note: tx.note,
    dayOfMonth: day, active: data.active !== false, startDate: tx.date,
  };
  const saved = withLock_(() => {
    if (data.id) return tUpdate_('Recurring', data.id, clean);
    return tInsert_('Recurring', Object.assign(clean, { id: newId_('REC'), lastRun: '', createdAt: new Date().toISOString() }));
  });
  processRecurring_();
  return saved;
}

function deleteRecurring(id) {
  withLock_(() => tDelete_('Recurring', id));
  return { ok: true };
}

/** สร้างรายการประจำที่ถึงกำหนดทันที คืนค่าจำนวนรายการที่สร้าง */
function runRecurringNow() {
  return processRecurring_();
}

/* ==============================================================
 * Settings & backup
 * ============================================================ */

function saveSettings(s) {
  s = s || {};
  if (s.backupEnabled !== undefined) {
    setSetting_('backupEnabled', s.backupEnabled ? 'true' : 'false');
    installTriggers();
  }
  return referenceData_().settings;
}

function backupNow() {
  return backup_();
}

function listBackups() {
  return listBackups_();
}

/* ==============================================================
 * Internals
 * ============================================================ */

/** รายการทั้งหมดที่ปรับหมวด/บัญชีที่หายไปให้เป็นค่าสำรองแล้ว */
function allTransactions_() {
  const catIds = {};
  tRead_('Categories').forEach((c) => { catIds[c.id] = c.type; });
  const accounts = tRead_('Accounts');
  const accIds = {};
  accounts.forEach((a) => { accIds[a.id] = true; });
  const fallbackAcc = (accounts.filter((a) => !a.archived)[0] || accounts[0] || { id: '' }).id;
  return tRead_('Transactions').map((t) => {
    t.type = t.type === 'income' || t.type === 'transfer' ? t.type : 'expense';
    t.date = String(t.date).slice(0, 10);
    if (t.type === 'transfer') {
      t.category = '';
      t.sub = '';
      if (!accIds[t.toAccount]) t.toAccount = fallbackAcc;
    } else {
      if (!catIds[t.category]) t.category = t.type === 'income' ? 'other_inc' : 'other_exp';
      t.toAccount = '';
    }
    if (!accIds[t.account]) t.account = fallbackAcc;
    return t;
  });
}

function accountBalances_(all) {
  const bal = {};
  tRead_('Accounts').forEach((a) => { bal[a.id] = a.openingBalance; });
  const today = today_();
  all.forEach((t) => {
    if (t.date > today) return; // รายการล่วงหน้ายังไม่นับ
    if (bal[t.account] === undefined) bal[t.account] = 0;
    if (t.type === 'transfer') {
      if (bal[t.toAccount] === undefined) bal[t.toAccount] = 0;
      bal[t.account] -= t.amount;
      bal[t.toAccount] += t.amount;
      return;
    }
    bal[t.account] += t.type === 'income' ? t.amount : -t.amount;
  });
  Object.keys(bal).forEach((k) => { bal[k] = round2_(bal[k]); });
  return bal;
}

function validateTx_(d) {
  if (!d) throw new Error('ไม่มีข้อมูล');
  const date = String(d.date || '').trim();
  if (!isDate_(date)) throw new Error('รูปแบบวันที่ไม่ถูกต้อง (yyyy-MM-dd)');
  const type = ['income', 'expense', 'transfer'].indexOf(d.type) > -1 ? d.type : null;
  if (!type) throw new Error('ประเภทต้องเป็นรายรับ รายจ่าย หรือโอน');
  const accounts = tRead_('Accounts');
  const acc = accounts.filter((a) => a.id === d.account)[0];
  if (!acc) throw new Error('ไม่พบบัญชีที่เลือก');
  let catId = '', toAccount = '';
  if (type === 'transfer') {
    const to = accounts.filter((a) => a.id === d.toAccount)[0];
    if (!to) throw new Error('เลือกบัญชีปลายทาง');
    if (to.id === acc.id) throw new Error('บัญชีต้นทางและปลายทางต้องไม่ใช่บัญชีเดียวกัน');
    toAccount = to.id;
  } else {
    const cat = tRead_('Categories').filter((c) => c.id === d.category)[0];
    if (!cat) throw new Error('ไม่พบหมวดหมู่ที่เลือก');
    if (cat.type !== type) throw new Error('หมวดหมู่ไม่ตรงกับประเภทรายการ');
    catId = cat.id;
  }
  const amount = round2_(Number(d.amount));
  if (!isFinite(amount) || amount <= 0) throw new Error('จำนวนเงินต้องมากกว่า 0');
  if (amount > 1e12) throw new Error('จำนวนเงินมากเกินไป');
  const attachment = /^[A-Za-z0-9_-]{10,100}$/.test(String(d.attachment || '')) ? d.attachment : '';
  return {
    date: date, type: type, category: catId, account: acc.id, toAccount: toAccount, amount: amount,
    note: String(d.note || '').trim().slice(0, 500), attachment: attachment,
    sub: type === 'transfer' ? '' : String(d.sub || '').trim().slice(0, 40),
  };
}

function hexColor_(v, fallback) {
  return /^#[0-9a-fA-F]{6}$/.test(String(v || '')) ? v : fallback;
}

function byDateDesc_(a, b) {
  return (b.date + b.createdAt).localeCompare(a.date + a.createdAt);
}
