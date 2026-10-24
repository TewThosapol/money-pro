// ทดสอบ Backend ด้วย Google services จำลอง:  node tests/backend.test.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

/* ---------- mocks ---------- */
function makeSheet() {
  const rows = [];
  let fmtCalls = 0;
  const range = (r, c, nr = 1, nc = 1) => ({
    setValues(v) { v.forEach((row, i) => { rows[r - 1 + i] = rows[r - 1 + i] || []; row.forEach((x, j) => { rows[r - 1 + i][c - 1 + j] = x; }); }); return this; },
    setValue(x) { rows[r - 1] = rows[r - 1] || []; rows[r - 1][c - 1] = x; return this; },
    getValues() { const o = []; for (let i = 0; i < nr; i++) { const row = rows[r - 1 + i] || []; const out = []; for (let j = 0; j < nc; j++) out.push(row[c - 1 + j] === undefined ? '' : row[c - 1 + j]); o.push(out); } return o; },
    setFontWeight() { return this; }, setBackground() { return this; }, setFontColor() { return this; }, setNumberFormat() { fmtCalls++; return this; },
  });
  return {
    rows,
    getRange: range,
    getLastRow: () => rows.length,
    getLastColumn: () => Math.max(0, ...rows.map((r) => r.length)),
    getMaxRows: () => 1000,
    setFrozenRows() {},
    appendRow(r) { rows.push(r.slice()); },
    deleteRow(i) { rows.splice(i - 1, 1); },
  };
}
const sheets = {};
const cache = {};
const ctx = {
  console,
  SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: (n) => sheets[n] || null, insertSheet: (n) => (sheets[n] = makeSheet()), getId: () => 'SS', getName: () => 'Money' }) },
  CacheService: { getScriptCache: () => ({
    get: (k) => cache[k] ?? null,
    getAll: (ks) => Object.fromEntries(ks.filter((k) => k in cache).map((k) => [k, cache[k]])),
    putAll: (m) => Object.assign(cache, m),
    remove: (k) => { delete cache[k]; },
  }) },
  LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
  Session: { getScriptTimeZone: () => 'Asia/Bangkok' },
  Utilities: {
    getUuid: () => Math.random().toString(16).slice(2) + Math.random().toString(16).slice(2),
    formatDate: (d, tz, f) => {
      const p = (n) => String(n).padStart(2, '0');
      const s = d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
      return f === 'yyyy-MM-dd' ? s : s + ' 0000';
    },
  },
  ScriptApp: { getProjectTriggers: () => [], newTrigger: () => ({ timeBased: () => ({ everyDays: () => ({ atHour: () => ({ create() {} }) }) }) }) },
  DriveApp: { getFileById: () => ({ setTrashed() { ctx.__trashed = (ctx.__trashed || 0) + 1; } }) },
};
vm.createContext(ctx);
for (const f of ['Data.gs', 'Jobs.gs', 'Receipt.gs', 'Code.gs']) {
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'src', f), 'utf8'), ctx, { filename: f });
}
const run = (code) => vm.runInContext(code, ctx);

/* ---------- tests ---------- */
let passed = 0;
const t = (name, fn) => { fn(); passed++; console.log('✓', name); };

const boot = run('bootstrap()');
t('bootstrap creates sheets and defaults', () => {
  for (const n of ['Transactions', 'Accounts', 'Categories', 'Recurring', 'Settings']) assert(sheets[n], n);
  assert.strictEqual(boot.accounts.length, 2);
  assert(boot.categories.length >= 14);
  assert(boot.categories.every((c) => c.group));
});

const today = run('today_()');
let tx;
t('add transaction', () => {
  tx = run(`addTransaction({ date: '${today}', type: 'expense', category: 'food', account: 'acc_cash', amount: '150.555', note: '=SUM(A1)' })`);
  assert.strictEqual(tx.amount, 150.56);
  assert.strictEqual(sheets.Transactions.rows[1][5], "'=SUM(A1)", 'formula injection guarded');
});
t('validation rejects wrong category type / bad amount / bad account', () => {
  assert.throws(() => run(`addTransaction({ date: '${today}', type: 'income', category: 'food', account: 'acc_cash', amount: 1 })`), /ไม่ตรง/);
  assert.throws(() => run(`addTransaction({ date: '${today}', type: 'expense', category: 'food', account: 'acc_cash', amount: 0 })`), /มากกว่า 0/);
  assert.throws(() => run(`addTransaction({ date: '${today}', type: 'expense', category: 'food', account: 'nope', amount: 5 })`), /บัญชี/);
  assert.throws(() => run(`addTransaction({ date: '2026/10/01', type: 'expense', category: 'food', account: 'acc_cash', amount: 5 })`), /วันที่/);
});
t('update keeps id and replaces attachment (old file trashed)', () => {
  run(`updateTransaction({ id: '${tx.id}', date: '${today}', type: 'expense', category: 'food', account: 'acc_cash', amount: 200, attachment: 'AAAAAAAAAAAA1' })`);
  const u = run(`updateTransaction({ id: '${tx.id}', date: '${today}', type: 'expense', category: 'shopping', account: 'acc_bank', amount: 300, attachment: 'BBBBBBBBBBBB2' })`);
  assert.strictEqual(u.id, tx.id);
  assert.strictEqual(u.category, 'shopping');
  assert.strictEqual(ctx.__trashed, 1);
});
t('balances and summary', () => {
  run(`addTransaction({ date: '${today}', type: 'income', category: 'salary', account: 'acc_bank', amount: 1000 })`);
  const d = run('getAppData()');
  assert.strictEqual(d.balances.acc_bank, 700);
  const s = run(`getDashboardSummary({ from: '${today}', to: '${today}' })`);
  assert.strictEqual(s.net, 700);
  assert.strictEqual(s.savingsRate, 70);
});
t('cache is invalidated after writes', () => {
  const before = run('getAppData()').transactions.length;
  run(`addTransaction({ date: '${today}', type: 'expense', category: 'food', account: 'acc_cash', amount: 1 })`);
  assert.strictEqual(run('getAppData()').transactions.length, before + 1);
});
t('delete transaction trashes attachment', () => {
  run(`deleteTransaction('${tx.id}')`);
  assert(!run('getAppData()').transactions.some((x) => x.id === tx.id));
  assert.strictEqual(ctx.__trashed, 2);
});
t('recurring backfills monthly and clamps day 31', () => {
  const n0 = run('getAppData()').transactions.length;
  run(`saveRecurring({ type: 'expense', category: 'housing', account: 'acc_bank', amount: 9000, dayOfMonth: 31, startDate: '2026-01-15', note: 'rent' })`);
  const recs = run('getAppData()').transactions.filter((x) => x.recurringId);
  assert(recs.length >= 9, 'backfilled ' + recs.length);
  assert(recs.some((x) => x.date === '2026-02-28'), 'Feb clamped');
  assert.strictEqual(run('getAppData()').transactions.length, n0 + recs.length);
  assert.strictEqual(run('runRecurringNow()'), 0, 'no duplicates on second run');
});
t('nextOccurrence helper', () => {
  assert.strictEqual(run(`nextOccurrence_('2026-01-31', 31)`), '2026-02-28');
  assert.strictEqual(run(`nextOccurrence_('2026-02-28', 31)`), '2026-03-31');
  assert.strictEqual(run(`nextOccurrence_('2026-12-15', 10)`), '2027-01-10');
});
t('category save, budget, archive when used', () => {
  const c = run(`saveCategory({ name: 'Pets', type: 'expense', group: 'need', icon: 'shield', color: '#123456', monthlyBudget: 800 })`);
  assert.strictEqual(c.group, 'need');
  run(`setBudget('food', 5000)`);
  assert.strictEqual(run('tRead_("Categories")').find((x) => x.id === 'food').monthlyBudget, 5000);
  assert.strictEqual(JSON.stringify(run(`deleteCategory('food')`)), JSON.stringify({ archived: true }));
  assert.strictEqual(JSON.stringify(run(`deleteCategory('${c.id}')`)), JSON.stringify({ deleted: true }));
});
t('account delete rules', () => {
  const a = run(`saveAccount({ name: 'KTC', type: 'credit', openingBalance: -500, color: '#EF4444' })`);
  assert.strictEqual(JSON.stringify(run(`deleteAccount('${a.id}')`)), JSON.stringify({ deleted: true }));
  assert.strictEqual(JSON.stringify(run(`deleteAccount('acc_cash')`)), JSON.stringify({ archived: true }));
  assert.throws(() => run(`deleteAccount('acc_bank')`), /อย่างน้อย 1/);
});
t('missing category/account fall back instead of crashing', () => {
  sheets.Transactions.appendRow(['TX-OLD', '2026-01-01', 'expense', 'ไม่มีหมวดนี้', 50, '', '', '', '']);
  run('cacheDrop_("Transactions")');
  const old = run('getAppData()').transactions.find((x) => x.id === 'TX-OLD');
  assert.strictEqual(old.category, 'other_exp');
  assert(old.account);
});

t('transfer moves money between accounts and is not income/expense', () => {
  const save = run(`saveAccount({ name: 'ออมทรัพย์', type: 'saving', openingBalance: 0, color: '#0EA5E9' })`);
  const before = run('getAppData()').balances;
  const tr = run(`addTransaction({ date: '${today}', type: 'transfer', account: 'acc_bank', toAccount: '${save.id}', amount: 500, note: 'ออม' })`);
  assert.strictEqual(tr.category, '');
  const after = run('getAppData()').balances;
  assert.strictEqual(after.acc_bank, before.acc_bank - 500);
  assert.strictEqual(after[save.id], 500);
  const s = run(`getDashboardSummary({ from: '${today}', to: '${today}' })`);
  assert(!JSON.stringify(s.byCategory).includes('""'), 'no empty category');
  assert.throws(() => run(`addTransaction({ date: '${today}', type: 'transfer', account: 'acc_bank', toAccount: 'acc_bank', amount: 5 })`), /ไม่ใช่บัญชีเดียวกัน/);
  assert.strictEqual(JSON.stringify(run(`deleteAccount('${save.id}')`)), JSON.stringify({ archived: true }), 'used by transfer → archived');
});
t('recurring transfer creates transfers', () => {
  const save = run(`saveAccount({ name: 'ออมอัตโนมัติ', type: 'saving', openingBalance: 0, color: '#0EA5E9' })`);
  run(`saveRecurring({ type: 'transfer', account: 'acc_bank', toAccount: '${save.id}', amount: 1000, dayOfMonth: 1, startDate: '2026-08-01', note: 'ออมทุกเดือน' })`);
  const list = run('getAppData()').transactions.filter((x) => x.toAccount === save.id);
  assert(list.length >= 3 && list.every((x) => x.type === 'transfer'), 'got ' + list.length);
  assert.strictEqual(run('getAppData()').balances[save.id], list.length * 1000);
});
t('goals CRUD', () => {
  const g = run(`saveGoal({ name: 'เงินสำรองฉุกเฉิน', targetAmount: 150000, account: '', targetDate: '2027-06-30' })`);
  assert(g.id && g.targetAmount === 150000);
  run(`saveGoal({ id: '${g.id}', name: 'เงินสำรอง', targetAmount: 120000, account: 'acc_bank' })`);
  const goals = run('bootstrap()').goals;
  assert.strictEqual(goals.length, 1);
  assert.strictEqual(goals[0].account, 'acc_bank');
  assert.throws(() => run(`saveGoal({ name: 'x', targetAmount: 0 })`), /มากกว่า 0/);
  run(`deleteGoal('${g.id}')`);
  assert.strictEqual(run('bootstrap()').goals.length, 0);
});

console.log(`\n${passed} passed`);
