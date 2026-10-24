/**
 * MoneyPro — Receipt / slip reader
 * --------------------------------------------------------------
 * 1) บันทึกรูปลงโฟลเดอร์ "MoneyPro Receipts" ใน Google Drive
 * 2) อ่านตัวอักษรด้วย OCR ของ Google Drive (ภาษาไทย + อังกฤษ)
 * 3) แยกยอดเงิน วันที่ เวลา ร้านค้า/ผู้รับ และเดาหมวดหมู่
 *
 * ต้องเปิด Advanced Service "Drive API" (Editor → Services → + → Drive API)
 * ถ้าใช้ไฟล์ appsscript.json ใน repo นี้ จะเปิดให้อัตโนมัติ
 */

const RECEIPT_MAX_BYTES = 8 * 1024 * 1024;
const RECEIPT_TYPES = ['image/jpeg', 'image/png', 'image/gif', 'image/webp', 'application/pdf'];

/**
 * อัปโหลดรูปสลิป/ใบเสร็จไปเก็บใน Drive
 * หน้าเว็บอ่านข้อความเองด้วย Tesseract แล้วเรียก parseReceipt() ดังนั้นปกติไม่ต้องใช้ OCR ของ Drive
 * (OCR ของ Drive มีโควตาต่ำ เมื่อเกินโควตาจะรอนานหลายนาที) ส่ง ocr: true เพื่อบังคับใช้
 * @param {{data: string, mimeType: string, name?: string, ocr?: boolean}} p  data = base64 (ไม่มี prefix data:)
 */
function uploadReceipt(p) {
  if (!p || !p.data) throw new Error('ไม่พบไฟล์รูป');
  const mime = String(p.mimeType || '').toLowerCase();
  if (RECEIPT_TYPES.indexOf(mime) < 0) throw new Error('รองรับเฉพาะ JPG, PNG, GIF, WEBP หรือ PDF');
  const bytes = Utilities.base64Decode(p.data);
  if (bytes.length > RECEIPT_MAX_BYTES) throw new Error('ไฟล์ใหญ่เกิน 8MB');

  const stamp = Utilities.formatDate(new Date(), tz_(), 'yyyyMMdd-HHmmss');
  const ext = mime === 'application/pdf' ? 'pdf' : mime.split('/')[1].replace('jpeg', 'jpg');
  const blob = Utilities.newBlob(bytes, mime, 'receipt-' + stamp + '.' + ext);

  const folder = folder_('receiptFolderId', APP_NAME + ' Receipts');
  const file = folder.createFile(blob);

  let text = '';
  let ocrError = '';
  if (p.ocr === true) {
    try {
      text = ocr_(blob);
    } catch (e) {
      ocrError = String(e && e.message || e);
    }
  }
  const cats = tRead_('Categories').filter((c) => !c.archived);
  return {
    fileId: file.getId(),
    url: file.getUrl(),
    text: text,
    ocrError: ocrError,
    parsed: parseReceiptText_(text, cats),
  };
}

/** แยกยอดเงิน/วันที่/ร้านค้า/หมวด จากข้อความที่หน้าเว็บอ่านได้ */
function parseReceipt(text) {
  return parseReceiptText_(String(text || '').slice(0, 20000), tRead_('Categories').filter((c) => !c.archived));
}

/** ลบรูปที่อัปโหลดแล้วแต่ไม่ได้บันทึกรายการ */
function discardReceipt(fileId) {
  if (!/^[A-Za-z0-9_-]{10,100}$/.test(String(fileId || ''))) return { ok: false };
  const used = tRead_('Transactions').some((t) => t.attachment === fileId);
  if (!used) trashFile_(fileId);
  return { ok: !used };
}

/** ทดสอบตัวแยกข้อความจาก Editor ได้: parseReceiptText('ข้อความ') */
function parseReceiptText(text) {
  return parseReceiptText_(String(text || ''), DEFAULT_CATEGORIES);
}

/* ==============================================================
 * OCR (Drive API v2 หรือ v3)
 * ============================================================ */

function ocr_(blob) {
  if (typeof Drive === 'undefined') {
    throw new Error('ยังไม่ได้เปิด Drive API — ไปที่ Editor → Services → + → Drive API → Add');
  }
  const t0 = Date.now();
  let docId;
  if (Drive.Files.insert) {
    docId = Drive.Files.insert({ title: 'ocr-temp', mimeType: blob.getContentType() }, blob,
      { ocr: true, ocrLanguage: 'th' }).id;
  } else {
    docId = Drive.Files.create({ name: 'ocr-temp', mimeType: 'application/vnd.google-apps.document' }, blob,
      { ocrLanguage: 'th' }).id;
  }
  const t1 = Date.now();
  try {
    const text = DocumentApp.openById(docId).getBody().getText();
    console.log('OCR timing: convert ' + (t1 - t0) + 'ms, read ' + (Date.now() - t1) + 'ms');
    return text;
  } finally {
    const t2 = Date.now();
    try { Drive.Files.remove(docId); } catch (e) { trashFile_(docId); }
    console.log('OCR timing: cleanup ' + (Date.now() - t2) + 'ms');
  }
}

/** ใช้ทดสอบจาก Editor: อ่านรูปล่าสุดในโฟลเดอร์ Receipts แล้วแสดงข้อความและผลที่แยกได้ */
function debugLatestReceipt() {
  const t0 = Date.now();
  const id = getSettingsMap_().receiptFolderId;
  if (!id) return 'ยังไม่มีโฟลเดอร์ Receipts';
  const it = DriveApp.getFolderById(id).getFiles();
  let latest = null;
  while (it.hasNext()) {
    const f = it.next();
    if (!latest || f.getDateCreated() > latest.getDateCreated()) latest = f;
  }
  if (!latest) return 'ไม่มีไฟล์';
  const t1 = Date.now();
  const text = ocr_(latest.getBlob());
  console.log('find file ' + (t1 - t0) + 'ms, total ' + (Date.now() - t0) + 'ms, size ' + latest.getSize());
  console.log(text);
  console.log(JSON.stringify(parseReceiptText_(text, DEFAULT_CATEGORIES)));
  return 'ok';
}

/* ==============================================================
 * Parser (ฟังก์ชันล้วน ไม่เรียกบริการของ Google จึงทดสอบนอก Apps Script ได้)
 * ============================================================ */

const TH_MONTHS_SHORT = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
const TH_MONTHS_LONG = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
const EN_MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];

const CATEGORY_KEYWORDS = {
  food: ['อาหาร', 'ร้านอาหาร', 'ข้าว', 'ก๋วยเตี๋ยว', 'กาแฟ', 'คาเฟ่', 'ชานม', 'ขนม', 'เบเกอรี่', 'restaurant', 'cafe', 'coffee',
    'starbucks', 'amazon', 'kfc', 'mcdonald', 'burger', 'pizza', 'sushi', 'mk ', 'bonchon', 'swensen', 'after you', 'grabfood',
    'grab food', 'lineman', 'line man', 'foodpanda', 'robinhood', 'shabu', 'bbq', 'บุฟเฟ่ต์', 'โภชนา', 'food'],
  transport: ['grab', 'bolt', 'taxi', 'แท็กซี่', 'bts', 'mrt', 'รถไฟฟ้า', 'airport rail', 'น้ำมัน', 'ปตท', 'ptt', 'shell',
    'bangchak', 'บางจาก', 'caltex', 'esso', 'ทางด่วน', 'easy pass', 'm-flow', 'ที่จอดรถ', 'parking', 'วินมอเตอร์ไซค์', 'rabbit'],
  housing: ['ค่าเช่า', 'คอนโด', 'หอพัก', 'ค่าส่วนกลาง', 'การไฟฟ้า', 'ค่าไฟ', 'ไฟฟ้า', 'การประปา', 'ค่าน้ำ', 'ประปา', 'mea', 'pea',
    'นิติบุคคล', 'rent'],
  bills: ['ais', 'true', 'dtac', '3bb', 'ais fibre', 'internet', 'อินเทอร์เน็ต', 'โทรศัพท์', 'netflix', 'spotify', 'youtube',
    'disney', 'icloud', 'google one', 'apple.com', 'chatgpt', 'openai', 'claude', 'ประกัน', 'insurance', 'บัตรเครดิต'],
  health: ['โรงพยาบาล', 'hospital', 'คลินิก', 'clinic', 'ร้านยา', 'เภสัช', 'pharmacy', 'boots', 'watsons', 'ทันตกรรม', 'dental',
    'fitness', 'ฟิตเนส', 'gym'],
  shopping: ['shopee', 'lazada', 'central', 'เซ็นทรัล', 'robinson', 'the mall', 'emporium', 'siam paragon', 'uniqlo', 'h&m',
    'zara', 'ikea', 'homepro', 'big c', 'บิ๊กซี', 'lotus', 'โลตัส', 'tops', 'makro', 'แม็คโคร', '7-eleven', 'เซเว่น',
    'familymart', 'lawson', 'tiktok shop', 'jd central', 'power buy', 'banana', 'apple store'],
  entertainment: ['major', 'sf cinema', 'cinema', 'โรงภาพยนตร์', 'ภาพยนตร์', 'concert', 'คอนเสิร์ต', 'ticket', 'ตั๋ว', 'steam',
    'playstation', 'nintendo', 'agoda', 'booking.com', 'airasia', 'thai airways', 'nok air', 'โรงแรม', 'hotel', 'resort'],
  education: ['course', 'คอร์ส', 'udemy', 'coursera', 'หนังสือ', 'book', 'se-ed', 'ซีเอ็ด', 'นายอินทร์', 'kinokuniya',
    'มหาวิทยาลัย', 'ค่าเทอม', 'tuition', 'สัมมนา'],
  salary: ['เงินเดือน', 'salary', 'payroll'],
  invest_inc: ['ดอกเบี้ย', 'interest', 'เงินปันผล', 'dividend'],
};

function parseReceiptText_(raw, cats, now) {
  const text = normalizeOcr_(raw || '');
  const lines = text.split('\n').map((l) => l.trim()).filter(Boolean);
  const type = /(รับเงิน|เงินเข้า|โอนเข้า|ได้รับเงิน|received|deposit|incoming)/i.test(text) &&
    !/(โอนเงินสำเร็จ|ชำระเงินสำเร็จ|จ่ายบิลสำเร็จ|payment successful)/i.test(text) ? 'income' : 'expense';
  const amount = findAmount_(lines);
  const date = findDate_(text, now);
  const time = findTime_(text);
  const merchant = findMerchant_(lines);
  const category = guessCategory_(text, type, cats || []);
  return {
    amount: amount,
    date: date,
    time: time,
    merchant: merchant,
    type: type,
    category: category,
    found: !!(amount || date || merchant),
  };
}

function normalizeOcr_(s) {
  return String(s)
    .replace(/\u0E4D\u0E32/g, '\u0E33') // นิคหิต + สระอา (OCR มักแยก) → สระอำ
    .replace(/([\u0E01-\u0E4E])[ \t]+(?=[\u0E01-\u0E4E])/g, '$1') // OCR บางรุ่นเว้นวรรคระหว่างตัวอักษรไทย
    .replace(/[๐-๙]/g, (d) => String('๐๑๒๓๔๕๖๗๘๙'.indexOf(d)))
    .replace(/\r/g, '')
    .replace(/[ \t ]+/g, ' ')
    .replace(/(\d)\s*,\s*(\d{3})/g, '$1,$2')
    .replace(/(\d)\s+\.\s*(\d{2})\b/g, '$1.$2');
}

function findAmount_(lines) {
  const strong = /(ยอดสุทธิ|รวมทั้งสิ้น|ยอดรวมสุทธิ|ยอดชำระ|ยอดที่ต้องชำระ|ยอดเงินที่ชำระ|grand\s*total|net\s*total|total\s*amount|amount\s*paid|total\s*due|^total\b|\btotal\s*[:฿])/i;
  const medium = /(จำนวนเงิน|จำนวน|ยอดเงิน|ยอดโอน|ยอดรวม|รวมเงิน|^รวม|amount|total|ชำระ|paid)/i;
  const exclude = /(ค่าธรรมเนียม|fee|vat|ภาษี|tax\b|เงินทอน|ทอน|change|ส่วนลด|discount|คงเหลือ|balance|เลขที่|รหัสอ้างอิง|ref|no\.|โทร|tel|เบอร์|points?|แต้ม|cash\s*received|รับเงินสด|subtotal|sub\s*total|ก่อนภาษี)/i;
  const numRe = /(฿|thb)?\s*(\d{1,3}(?:,\d{3})+(?:\.\d{1,2})?|\d+(?:\.\d{1,2})?)\s*(บาท|thb|฿|baht)?/gi;

  const candidates = [];
  const pushFrom = (line, weight) => {
    // ตัดวันที่/เวลา/เลขบัญชีออกก่อน กันไปจับผิดตัว
    const clean = line
      .replace(/\d{1,4}[\/\-.]\d{1,2}[\/\-.]\d{2,4}/g, ' ')
      .replace(/\b\d{1,2}:\d{2}(:\d{2})?\b/g, ' ')
      .replace(/\b[x\d]{3}-[x\d]-[x\d]{4,5}-[x\d]\b/gi, ' ')
      .replace(/\bx{2,}[\dx-]*\b/gi, ' ');
    let m;
    numRe.lastIndex = 0;
    while ((m = numRe.exec(clean))) {
      const txt = m[2];
      const v = Number(txt.replace(/,/g, ''));
      if (!isFinite(v) || v <= 0 || v >= 1e8) continue;
      const hasDec = /\.\d{2}$/.test(txt);
      const hasCur = !!(m[1] || m[3]);
      const digits = txt.replace(/\D/g, '').length;
      if (!hasDec && !hasCur && digits >= 7) continue; // น่าจะเป็นเลขอ้างอิง
      if (!hasDec && !hasCur && weight === 0) continue;
      candidates.push({ v: Math.round(v * 100) / 100, score: weight * 10 + (hasDec ? 3 : 0) + (hasCur ? 4 : 0) });
    }
  };

  lines.forEach((line, i) => {
    if (exclude.test(line) && !strong.test(line)) return;
    const w = strong.test(line) ? 3 : medium.test(line) ? 2 : 0;
    pushFrom(line, w);
    // ตัวเลขอยู่บรรทัดถัดไปจากคำว่า "จำนวนเงิน"
    if (w > 0 && !/\d/.test(line.replace(/\d{1,2}:\d{2}/g, '')) && lines[i + 1] && !exclude.test(lines[i + 1])) {
      pushFrom(lines[i + 1], w);
    }
  });
  if (!candidates.length) return null;
  candidates.sort((a, b) => b.score - a.score || b.v - a.v);
  return candidates[0].v;
}

function findDate_(text, nowArg) {
  const now = nowArg ? new Date(nowArg) : new Date();
  const maxYear = now.getFullYear() + 1;
  const toCE = (y) => {
    y = Number(y);
    if (y < 100) y = y >= 40 ? 2500 + y - 543 : 2000 + y;
    else if (y > 2400) y -= 543;
    return y;
  };
  const ok = (y, m, d) => y >= 2000 && y <= maxYear && m >= 1 && m <= 12 && d >= 1 && d <= daysInMonthPure_(y, m);
  const fmt = (y, m, d) => y + '-' + String(m).padStart(2, '0') + '-' + String(d).padStart(2, '0');

  const loose = (s) => s.replace(/\./g, '').split('').map((ch) => ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('\\.?\\s?') + '\\.?';
  const thAlts = TH_MONTHS_LONG.map((m, i) => ({ re: m, i: i }))
    .concat(TH_MONTHS_SHORT.map((m, i) => ({ re: loose(m), i: i })));
  for (let k = 0; k < thAlts.length; k++) {
    const re = new RegExp('(\\d{1,2})\\s*' + thAlts[k].re + '[\\s:.,]*(\\d{2,4})');
    const m = text.match(re);
    if (m) {
      const y = toCE(m[2]), d = Number(m[1]), mo = thAlts[k].i + 1;
      if (ok(y, mo, d)) return fmt(y, mo, d);
    }
  }
  const en = text.match(/(\d{1,2})[\s\-]*(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?[\s,\-]*(\d{2,4})/i);
  if (en) {
    const y = toCE(en[3]), mo = EN_MONTHS.indexOf(en[2].toLowerCase()) + 1, d = Number(en[1]);
    if (ok(y, mo, d)) return fmt(y, mo, d);
  }
  const en2 = text.match(/(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+(\d{1,2}),?\s*(\d{4})/i);
  if (en2) {
    const y = toCE(en2[3]), mo = EN_MONTHS.indexOf(en2[1].toLowerCase()) + 1, d = Number(en2[2]);
    if (ok(y, mo, d)) return fmt(y, mo, d);
  }
  const iso = text.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (iso) {
    const y = toCE(iso[1]), mo = Number(iso[2]), d = Number(iso[3]);
    if (ok(y, mo, d)) return fmt(y, mo, d);
  }
  const re = /(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/g;
  let m;
  while ((m = re.exec(text))) {
    const y = toCE(m[3]), mo = Number(m[2]), d = Number(m[1]);
    if (ok(y, mo, d)) return fmt(y, mo, d);
  }
  // OCR อ่านชื่อเดือนเพี้ยน (เช่น "ต.ค." เป็น "0.n.") แต่ยังเห็น วัน ปี และเวลา:
  // ใช้เดือนล่าสุดที่วันนั้นยังไม่เกินวันนี้ (สลิปส่วนใหญ่บันทึกไม่นานหลังโอน)
  const g = text.match(/(?:^|\s)(\d{1,2})\s+\S{1,3}\.\s?\S{1,3}\.?\s*(\d{2}|25\d{2}|20\d{2})\s*[,\s]\s*\d{1,2}[:.]\d{2}/);
  if (g) {
    const d = Number(g[1]), y = toCE(g[2]);
    let mo = now.getMonth() + 1, yy = now.getFullYear();
    if (d > now.getDate()) { mo -= 1; if (mo < 1) { mo = 12; yy -= 1; } }
    if (y === yy && ok(yy, mo, d)) return fmt(yy, mo, d);
  }
  return null;
}

function findTime_(text) {
  const m = text.match(/\b([01]?\d|2[0-3])[:.]([0-5]\d)(?:[:.][0-5]\d)?\s*(?:น\.|น|hrs?|am|pm)?/i);
  if (!m) return null;
  // กันไปจับยอดเงิน เช่น 12.50 → ต้องมี ":" หรือคำว่า น.
  if (m[0].indexOf(':') < 0 && !/(น|hr|am|pm)/i.test(m[0])) return null;
  return String(m[1]).padStart(2, '0') + ':' + m[2];
}

function findMerchant_(lines) {
  const marker = /^(ไปยัง|ไปที่|ถึง|ผู้รับ(?:เงิน)?|โอนให้|ชำระให้|to|merchant|ร้านค้า|ชื่อร้าน|payee|biller)(?![A-Za-z])\s*[:：]?\s*(.*)$/i;
  const junk = /(สำเร็จ|ใบเสร็จ|ใบกำกับ|โอนเงิน|เลขที่|รหัส|วันที่|เวลา|จำนวน|บาท|ค่าธรรมเนียม|ธนาคาร|สาขา|โทร|พร้อมเพย์|www\.|http|@|x{3,}|\d{3}-\d|\b(successful|receipt|tax\s*invoice|transfer|ref|date|time|amount|thb|fee|bank|branch|tel|scan|qr|promptpay|total|order)\b)/i;
  const good = (s) => s && s.length >= 2 && s.length <= 60 && /[A-Za-zก-๙]/.test(s) && !junk.test(s) &&
    (s.replace(/[^A-Za-zก-๙]/g, '').length >= 3);
  // ตัดสัญลักษณ์/ไอคอนที่ OCR อ่านเป็นตัวอักษรเดี่ยวหน้าชื่อ เช่น "5 646 WATSONS" → "646 WATSONS"
  const clean = (s) => String(s || '').replace(/^(\S\s+)+(?=\S{2,})/, '').replace(/\s{2,}/g, ' ').trim();
  for (let i = 0; i < lines.length; i++) {
    const m = lines[i].match(marker);
    if (m) {
      if (good(clean(m[2]))) return clean(m[2]);
      if (good(clean(lines[i + 1]))) return clean(lines[i + 1]);
    }
  }
  for (let i = 0; i < Math.min(lines.length, 6); i++) {
    if (good(clean(lines[i]))) return clean(lines[i]);
  }
  return null;
}

function guessCategory_(text, type, cats) {
  const t = ' ' + text.toLowerCase() + ' ';
  const allowed = {};
  cats.forEach((c) => { if (c.type === type) allowed[c.id] = true; });
  let best = null, bestLen = 0;
  Object.keys(CATEGORY_KEYWORDS).forEach((id) => {
    if (!allowed[id]) return;
    CATEGORY_KEYWORDS[id].forEach((kw) => {
      if (kw.length > bestLen && hasKeyword_(t, kw.toLowerCase())) { best = id; bestLen = kw.length; }
    });
  });
  // คำค้นที่ตรงกับชื่อหมวดที่ผู้ใช้ตั้งเอง
  cats.forEach((c) => {
    if (c.type !== type || !c.name) return;
    const n = c.name.toLowerCase();
    if (n.length > bestLen && t.indexOf(n) > -1) { best = c.id; bestLen = n.length; }
  });
  return best;
}

/** คำภาษาอังกฤษต้องตรงทั้งคำ (กัน "pea" ไปตรงกับ "repeat") ส่วนภาษาไทยไม่มีช่องว่างจึงใช้ค้นแบบมีอยู่ในข้อความ */
function hasKeyword_(t, kw) {
  if (!/^[\x20-\x7e]+$/.test(kw)) return t.indexOf(kw) > -1;
  const esc = kw.trim().replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return new RegExp('(^|[^a-z0-9])' + esc + '([^a-z0-9]|$)').test(t);
}

function daysInMonthPure_(y, m) {
  return new Date(y, m, 0).getDate();
}
