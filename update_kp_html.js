
const fs = require('fs');

// KPindex.htmlを読み込み
const kpContent = fs.readFileSync('KPindex.html', 'utf8');

// JSON.stringifyで正確にエスケープ（ダブルクォートを含む）
const escaped = JSON.stringify(kpContent);
// JSON.stringifyは "..." の形式なので、外のクォートを除いた値部分のみ使う
// window.KP_HTML = {"value": "..."}; の形式にする
const newKpHtmlLine = 'window.KP_HTML = {"value":' + escaped + '};';

console.log('New KP_HTML line length:', newKpHtmlLine.length);
console.log('Last 50 chars:', newKpHtmlLine.slice(-50));

// index.htmlを読み込み
const indexContent = fs.readFileSync('index.html', 'utf8');

// 1777行目（0-indexedで1776）のwindow.KP_HTML行を特定
const lines = indexContent.split('\n');
let kpLineIdx = -1;
for (let i = 0; i < lines.length; i++) {
  if (lines[i].trimStart().startsWith('window.KP_HTML = {')) {
    kpLineIdx = i;
    break;
  }
}

if (kpLineIdx === -1) {
  console.error('ERROR: window.KP_HTML line not found!');
  process.exit(1);
}

console.log('Found KP_HTML at line:', kpLineIdx + 1);
console.log('Old line length:', lines[kpLineIdx].length);

// 置換
lines[kpLineIdx] = newKpHtmlLine;

// 書き込み
const newContent = lines.join('\n');
fs.writeFileSync('index.html', newContent, 'utf8');
console.log('SUCCESS: index.html updated. Total length:', newContent.length);
