
const fs = require('fs');
let content = fs.readFileSync('index.html', 'utf8');
const newVer = 'kp_' + Date.now().toString(36);
content = content.replace(/window\.KP_HTML_VERSION = "[^"]*";/, 'window.KP_HTML_VERSION = "' + newVer + '";');
fs.writeFileSync('index.html', content, 'utf8');
console.log('Updated KP_HTML_VERSION to:', newVer);
