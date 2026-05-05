import fs from "node:fs";
import path from "node:path";
import crypto from "node:crypto";

const repoRoot = process.cwd();
const indexPath = path.join(repoRoot, "index.html");
const kpSourcePathArg = process.argv[2];
const kpSourcePath = kpSourcePathArg
  ? path.resolve(repoRoot, kpSourcePathArg)
  : path.resolve(repoRoot, "..", "KPindex.html");

if (!fs.existsSync(indexPath)) {
  throw new Error(`index.html が見つかりません: ${indexPath}`);
}
if (!fs.existsSync(kpSourcePath)) {
  throw new Error(`KPindex.html が見つかりません: ${kpSourcePath}`);
}

const indexRaw = fs.readFileSync(indexPath, "utf8");
const kpRaw = fs.readFileSync(kpSourcePath, "utf8");
// HTML内 <script> 直下の文字列として埋め込むため、閉じタグを無害化する
const kpRawForEmbed = kpRaw.replace(/<\/script/gi, "<\\/script");

const startToken = "window.KP_HTML = ";
const endToken = "\n    const LS_APP_CACHED_MATERIALS";
const start = indexRaw.indexOf(startToken);
if (start < 0) {
  throw new Error("index.html 内に window.KP_HTML が見つかりません。");
}
const end = indexRaw.indexOf(endToken, start);
if (end < 0) {
  throw new Error("index.html 内で KP_HTML 終端を特定できません。");
}

const hash = crypto.createHash("sha1").update(kpRaw).digest("hex").slice(0, 12);
const replacement =
  `window.KP_HTML = ${JSON.stringify({ value: kpRawForEmbed })};\n` +
  `window.KP_HTML_VERSION = "${hash}";`;

const next =
  indexRaw.slice(0, start) +
  replacement +
  indexRaw.slice(end);

if (next === indexRaw) {
  console.log("変更なし（既に最新）。");
  process.exit(0);
}

fs.writeFileSync(indexPath, next, "utf8");
console.log(`KP_HTML を更新しました: version=${hash}`);
console.log(`source: ${kpSourcePath}`);
