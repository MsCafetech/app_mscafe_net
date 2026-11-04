import * as acorn from 'acorn';
import fs from 'fs';
const [src, dst] = process.argv.slice(2);
// Shift_JIS (CP932) で表せない文字（簡体字・ハングル・é・© など）は、サーバーで貼り付けると化けて構文エラーになるため、
// スクリプト内は \uXXXX、HTML 部分は &#xXXXX; に置き換える（2026-11-04）。
const sj = new TextDecoder('shift_jis'), OK = new Set();
for (let a = 0x81; a <= 0xfc; a++) for (let b = 0x40; b <= 0xfc; b++) { const c = sj.decode(Buffer.from([a, b])); if (c.length === 1 && c !== '\ufffd') OK.add(c); }
for (let a = 0xa1; a <= 0xdf; a++) OK.add(sj.decode(Buffer.from([a])));
const safe = c => c.codePointAt(0) < 128 || OK.has(c);
const escJs = t => [...t].map(c => safe(c) ? c : c.split('').map(u => '\\u' + u.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0')).join('')).join('');
const escHtml = t => [...t].map(c => safe(c) ? c : '&#x' + c.codePointAt(0).toString(16).toUpperCase() + ';').join('');
const html = fs.readFileSync(src, 'utf8');
let out = '', last = 0; const re = /<script>([\s\S]*?)<\/script>/g; let m;
while ((m = re.exec(html))) {
  out += escHtml(html.slice(last, m.index));
  let js = m[1]; const cs = [];
  acorn.parse(js, { ecmaVersion: 'latest', sourceType: 'script', onComment: (block, text, start, end) => cs.push([start, end]) });
  for (const [s, e] of cs.reverse()) js = js.slice(0, s) + js.slice(e);
  js = js.replace(/(const CSS = `)([\s\S]*?)(`;)/, (_, a, b, c) => a + b.replace(/\/\*[\s\S]*?\*\//g, '') + c);
  js = js.split('\n').map(l => l.replace(/[ \t]+$/, '')).filter(l => l.trim() !== '').join('\n');
  acorn.parse(js, { ecmaVersion: 'latest', sourceType: 'script' });
  js = escJs(js); acorn.parse(js, { ecmaVersion: 'latest', sourceType: 'script' });
  out += '<script>\n' + js + '\n</script>'; last = re.lastIndex;
}
out += escHtml(html.slice(last));
const mm = out.match(/正しいサイズ: ([\d,]+) バイト/);
const base = Buffer.byteLength(out) - Buffer.byteLength(mm[1]); let size = base;
for (let i = 0; i < 5; i++) size = base + Buffer.byteLength(size.toLocaleString('en-US'));
out = out.replace(mm[0], `正しいサイズ: ${size.toLocaleString('en-US')} バイト`);
fs.writeFileSync(dst, out); console.log('bytes', Buffer.byteLength(out), 'lines', out.split('\n').length, 'maxline', Math.max(...out.split('\n').map(l => l.length)));
