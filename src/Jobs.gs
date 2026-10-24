/**
 * MoneyPro — Background jobs
 * --------------------------------------------------------------
 * รายการประจำ, สำรองข้อมูล, ตัวตั้งเวลา (trigger) และไฟล์ใน Drive
 */

const BACKUP_KEEP = 14;

/**
 * สร้างรายการจากกฎรายการประจำที่ถึงกำหนดแล้ว (ย้อนหลังได้สูงสุด 24 งวดต่อกฎ)
 * @return {number} จำนวนรายการที่สร้าง
 */
function processRecurring_() {
  const rules = tRead_('Recurring').filter((r) => r.active);
  if (!rules.length) return 0;
  const today = today_();
  // ตรวจแบบเร็วก่อนล็อก: มีกฎไหนถึงกำหนดไหม
  const due = rules.some((r) => {
    const next = nextOccurrence_(r.lastRun || prevDay_(r.startDate), r.dayOfMonth);
    return next && next <= today;
  });
  if (!due) return 0;

  return withLock_(() => {
    const fresh = tRead_('Recurring', true).filter((r) => r.active);
    const now = new Date().toISOString();
    const newTx = [];
    fresh.forEach((r) => {
      let last = r.lastRun || prevDay_(r.startDate);
      let next = nextOccurrence_(last, r.dayOfMonth);
      let guard = 0;
      while (next && next <= today && guard++ < 24) {
        newTx.push({
          id: newId_('TX'), date: next, type: r.type, category: r.category, amount: r.amount,
          note: r.note || 'รายการประจำ', createdAt: now, updatedAt: now,
          account: r.account, recurringId: r.id, attachment: '', toAccount: r.type === 'transfer' ? r.toAccount : '',
        });
        last = next;
        next = nextOccurrence_(last, r.dayOfMonth);
      }
      if (last !== (r.lastRun || prevDay_(r.startDate))) tUpdate_('Recurring', r.id, { lastRun: last });
    });
    tInsertMany_('Transactions', newTx);
    return newTx.length;
  });
}

/** งานรายวัน: ตัวตั้งเวลาเรียกฟังก์ชันนี้ */
function dailyJob() {
  processRecurring_();
  if (getSettingsMap_().backupEnabled === 'true') backup_();
}

/** ติดตั้งตัวตั้งเวลารายวัน (ไม่สร้างซ้ำ) */
function installTriggers() {
  const exists = ScriptApp.getProjectTriggers().some((t) => t.getHandlerFunction() === 'dailyJob');
  if (!exists) ScriptApp.newTrigger('dailyJob').timeBased().everyDays(1).atHour(2).create();
  return 'ok';
}

/* ==============================================================
 * Backup
 * ============================================================ */

function backup_() {
  const ss = ss_();
  const folder = folder_('backupFolderId', APP_NAME + ' Backups');
  const stamp = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HHmm');
  const copy = DriveApp.getFileById(ss.getId()).makeCopy(ss.getName() + ' — backup ' + stamp, folder);

  // เก็บไว้เฉพาะ BACKUP_KEEP ไฟล์ล่าสุด
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) files.push(it.next());
  files.sort((a, b) => b.getDateCreated() - a.getDateCreated());
  files.slice(BACKUP_KEEP).forEach((f) => f.setTrashed(true));

  setSetting_('lastBackup', new Date().toISOString());
  setSetting_('lastBackupUrl', copy.getUrl());
  return { name: copy.getName(), url: copy.getUrl(), folderUrl: folder.getUrl() };
}

function listBackups_() {
  const id = getSettingsMap_().backupFolderId;
  if (!id) return [];
  const out = [];
  try {
    const it = DriveApp.getFolderById(id).getFiles();
    while (it.hasNext()) {
      const f = it.next();
      out.push({ name: f.getName(), url: f.getUrl(), created: f.getDateCreated().toISOString() });
    }
  } catch (e) {
    return [];
  }
  return out.sort((a, b) => b.created.localeCompare(a.created));
}

/* ==============================================================
 * Drive helpers
 * ============================================================ */

/** โฟลเดอร์ใน Drive ที่จำ ID ไว้ใน Settings (สร้างใหม่ถ้ายังไม่มีหรือถูกลบ) */
function folder_(settingKey, name) {
  const id = getSettingsMap_()[settingKey];
  if (id) {
    try {
      const f = DriveApp.getFolderById(id);
      if (!f.isTrashed()) return f;
    } catch (e) { /* สร้างใหม่ */ }
  }
  const f = DriveApp.createFolder(name);
  setSetting_(settingKey, f.getId());
  return f;
}

function trashFile_(fileId) {
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) { /* ไม่เป็นไร */ }
}
