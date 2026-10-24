// ทดสอบตัวแยกข้อความสลิป/ใบเสร็จนอก Apps Script:  node tests/receipt.test.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ctx = { console };
vm.createContext(ctx);
for (const f of ['Data.gs', 'Receipt.gs']) {
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'src', f), 'utf8'), ctx, { filename: f });
}
const parse = (text) => vm.runInContext('parseReceiptText_', ctx)(text, vm.runInContext('DEFAULT_CATEGORIES', ctx));

const cases = [
  {
    name: 'KBank transfer slip (Thai, BE short year, amount on next line)',
    text: `โอนเงินสำเร็จ
24 ต.ค. 69 14:32 น.
นาย ทศพล ท.
ธ.กสิกรไทย
xxx-x-x1234-x
ไปยัง
ร้านกาแฟบ้านสวน
xxx-x-x5678-x
เลขที่รายการ:
016297143258AQR04512
จำนวน:
1,250.00 บาท
ค่าธรรมเนียม:
0.00 บาท`,
    expect: { amount: 1250, date: '2026-10-24', time: '14:32', merchant: 'ร้านกาแฟบ้านสวน', type: 'expense', category: 'food' },
  },
  {
    name: 'SCB slip with Thai digits and full month',
    text: `จ่ายบิลสำเร็จ
๓ ตุลาคม ๒๕๖๙ - ๐๙:๑๕
จาก นาย ทศพล
ไปยัง การไฟฟ้านครหลวง
จำนวนเงิน ๒,๔๓๑.๕๐`,
    expect: { amount: 2431.5, date: '2026-10-03', time: '09:15', merchant: 'การไฟฟ้านครหลวง', category: 'housing' },
  },
  {
    name: 'Cafe receipt (English, total vs subtotal/vat/cash/change)',
    text: `STARBUCKS COFFEE
Central World 3F
Date: 21/10/2026  Time: 08:47
Caffe Latte (G)      145.00
Croissant             95.00
Subtotal             224.30
VAT 7%                15.70
TOTAL                240.00
Cash                 500.00
Change               260.00
Tel. 02-123-4567`,
    expect: { amount: 240, date: '2026-10-21', time: '08:47', merchant: 'STARBUCKS COFFEE', category: 'food' },
  },
  {
    name: 'Fuel receipt with BE 4-digit year',
    text: `บริษัท ปตท. น้ำมันและการค้าปลีก จำกัด
สาขา วิภาวดี
วันที่ 15/09/2569 18:02
แก๊สโซฮอล์ 95  32.45 ลิตร
ยอดสุทธิ 1,200.00 บาท`,
    expect: { amount: 1200, date: '2026-09-15', category: 'transport' },
  },
  {
    name: 'Incoming transfer → income',
    text: `ได้รับเงิน
1 Oct 2026, 10:00
จาก บริษัท ตัวอย่าง จำกัด
จำนวนเงิน 35,000.00 THB
หมายเหตุ เงินเดือน`,
    expect: { amount: 35000, date: '2026-10-01', type: 'income', category: 'salary' },
  },
  {
    name: 'Online shop order',
    text: `Shopee
Order Total ฿ 589
Payment date Oct 18, 2026`,
    expect: { amount: 589, date: '2026-10-18', category: 'shopping' },
  },
  {
    name: 'Garbage text → nothing found',
    text: `hello world`,
    expect: { amount: null, date: null },
  },
];

let fail = 0;
for (const c of cases) {
  const got = parse(c.text);
  const bad = Object.keys(c.expect).filter((k) => got[k] !== c.expect[k]);
  if (bad.length) {
    fail++;
    console.log('✗', c.name);
    bad.forEach((k) => console.log('   ', k, 'expected', JSON.stringify(c.expect[k]), 'got', JSON.stringify(got[k])));
  } else {
    console.log('✓', c.name);
  }
}
console.log(fail ? `\n${fail} failed` : '\nall passed');
process.exit(fail ? 1 : 0);
