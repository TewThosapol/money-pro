// Apps Script HtmlService ตัดข้อความแบบคอมเมนต์ (// ถึงท้ายบรรทัด และ /* ... */) ในสคริปต์ทิ้ง
// แม้อยู่ในสตริง ทดสอบว่าโค้ดยังถูกต้องหลังโดนตัด:  node tests/html-gas.test.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const html = fs.readFileSync(path.join(__dirname, '..', 'src', 'Index.html'), 'utf8');
const scripts = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].map((m) => m[1]);
let fail = 0;
scripts.forEach((code, i) => {
  const stripped = code.split('\n').map((l) => l.split('//')[0]).join('\n').replace(/\/\*[\s\S]*?\*\//g, '');
  try {
    new vm.Script(stripped, { filename: `inline-script-${i}.js` });
    console.log(`✓ inline script ${i} survives Apps Script comment stripping`);
  } catch (e) {
    fail++;
    console.log(`✗ inline script ${i}: ${e.message}`);
  }
});
if (/href="#/.test(html)) {
  fail++;
  console.log('✗ found href="#..." link (navigates out of the Apps Script iframe)');
} else {
  console.log('✓ no #hash links');
}
process.exit(fail ? 1 : 0);
