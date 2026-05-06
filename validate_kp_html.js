
const fs = require('fs');
const lines = fs.readFileSync('index.html', 'utf8').split('\n');
const kpLine = lines[1776]; // 1777行目(0-indexed)
console.log('Line 1777 length:', kpLine.length);
console.log('Starts with window.KP_HTML:', kpLine.startsWith('window.KP_HTML = {'));
console.log('Ends with ;', kpLine.trimEnd().endsWith('};'));

// JSON部分を抽出してパース
const jsonPart = kpLine.replace(/^window\.KP_HTML = /, '').replace(/;\s*$/, '');
try {
  const parsed = JSON.parse(jsonPart);
  console.log('JSON parse: OK');
  console.log('Value type:', typeof parsed.value);
  console.log('Value length:', parsed.value ? parsed.value.length : 0);
  if (parsed.value && parsed.value.includes('kpRequestInitData')) {
    console.log('OK: kpRequestInitData found in value');
  } else {
    console.log('MISSING: kpRequestInitData not in value');
  }
} catch(e) {
  console.log('JSON parse ERROR:', e.message.substring(0, 200));
}
