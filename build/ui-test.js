/* =============================================================
 *  ui-test.js - 화면이 조용히 안 도는 자리 검증
 *
 *  실행:  node build\ui-test.js
 *
 *  배경: 지난 미완료 배너의 [펼쳐보기]가 눌러도 아무 일이 없었다.
 *        코드는 멀쩡했다 — el.hidden = true 를 제대로 넣고 있었는데
 *        .carry-list{display:flex} 가 그것을 덮고 있었다.
 *        hidden 의 display:none 은 브라우저 기본 규칙이라 우선순위가 가장 낮아서
 *        클래스 규칙 하나에 진다. 에러도 안 나고 검사도 안 걸려서
 *        "왜 안 되지" 하고 한참 들여다보게 되는 종류다.
 *
 *        같은 이유로 자동 백업을 꺼도 헤더의 카운트다운이 안 사라졌다.
 *        [hidden]{display:none !important} 를 전역으로 못 박아 고쳤고,
 *        여기서 지킨다.
 *
 *  두 가지를 본다.
 *    1) hidden 을 켜도 CSS display 에 덮여 안 숨는 곳
 *    2) JS 가 부르는데 문서에 없는 id — $(...) 가 null 이라 그 줄에서 멈춘다
 * ============================================================= */
'use strict';
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const htmlPath = path.join(root, 'dist', '배포', 'sched.html');

let pass = 0, fail = 0;
function chk(name, cond, detail){
  if(cond){ pass++; console.log('  [통과] ' + name); }
  else { fail++; console.log('  [실패] ' + name + '  -> ' + (detail === undefined ? '' : detail)); }
}

const src = fs.readFileSync(htmlPath, 'utf8');
const s0 = src.indexOf('<style>'), s1 = src.indexOf('</style>');
if(s0 < 0 || s1 < 0){ console.log('  [실패] style 블록을 찾지 못했습니다'); process.exit(1); }
const css = src.slice(s0 + 7, s1).replace(/\/\*[\s\S]*?\*\//g, ' ');

/* ---------------------------------------------------------------
   1. hidden 이 덮이는가
   --------------------------------------------------------------- */
console.log('1. hidden 을 켜면 정말 숨는가');

// 전역으로 못 박아 둔 규칙이 있으면 !important 없는 규칙은 전부 진다
const guard = /(^|[,}\s])\[hidden\]\s*\{[^}]*display\s*:\s*none\s*!important/.test(css);
chk('[hidden]{display:none !important} 를 전역으로 두었다', guard);

// hidden 으로 여닫는 요소 — id 와 그 요소가 실제로 단 class
const targets = new Map();
for(const m of src.matchAll(/<[a-z]+\b([^>]*)>/g)){
  const a = m[1];
  if(!/\bhidden(?=[\s>=]|$)/.test(a)) continue;
  const id = (a.match(/\bid="([^"]+)"/) || [])[1];
  if(!id) continue;
  const cls = (a.match(/\bclass="([^"]+)"/) || [, ''])[1].split(/\s+/).filter(Boolean);
  targets.set(id, new Set(cls));
}
for(const m of src.matchAll(/\$\('#([a-zA-Z0-9_]+)'\)\.hidden/g)){
  if(targets.has(m[1])) continue;
  const tag = src.match(new RegExp('<[a-z]+\\b[^>]*\\bid="' + m[1] + '"[^>]*>'));
  const cls = tag ? (tag[0].match(/\bclass="([^"]+)"/) || [, ''])[1].split(/\s+/).filter(Boolean) : [];
  targets.set(m[1], new Set(cls));
}

// display 를 none 아닌 값으로 주는 규칙만 추린다
const rules = [];
for(const m of css.matchAll(/([^{}]+)\{([^{}]*)\}/g)){
  const sel = m[1].trim();
  const d = m[2].match(/(?:^|;)\s*display\s*:\s*([^;}]+)/);
  if(!d) continue;
  const val = d[1].trim();
  if(/^none\b/.test(val) || sel.startsWith('@') || sel.includes('[hidden]')) continue;
  if(guard && !/!important/.test(val)) continue;     // 전역 규칙이 이긴다
  rules.push({ sel, val });
}

// 마지막 조각(그 요소 자신에게 걸리는 부분)만 정확히 대조한다
const hits = [];
for(const [id, classes] of targets){
  for(const r of rules) for(const one of r.sel.split(',')){
    const last = one.trim().split(/[\s>+~]+/).pop();
    if(!last || last.includes(':')) continue;
    const ids = [...last.matchAll(/#([\w-]+)/g)].map(x => x[1]);
    const cls = [...last.matchAll(/\.([\w-]+)/g)].map(x => x[1]);
    if(!ids.length && !cls.length) continue;                 // 태그 선택자만은 건너뛴다
    if(ids.length && !ids.every(x => x === id)) continue;
    if(!cls.every(c => classes.has(c))) continue;
    hits.push('#' + id + ' <- ' + one.trim() + ' {display:' + r.val + '}');
  }
}
chk('hidden 으로 여닫는 요소를 ' + targets.size + '개 찾았다', targets.size > 0);
chk('그중 CSS display 에 덮이는 것이 없다', hits.length === 0, hits.join(' · '));

/* ---------------------------------------------------------------
   2. 없는 id 를 부르는가
   --------------------------------------------------------------- */
console.log('');
console.log('2. JS 가 부르는 id 가 문서에 다 있는가');

const have = new Set([...src.matchAll(/\bid="([^"]+)"/g)].map(m => m[1]));
const asked = new Set([...src.matchAll(/\$\('#([a-zA-Z0-9_]+)'\)/g)].map(m => m[1]));
const missing = [...asked].filter(id => !have.has(id));
chk('부르는 id ' + asked.size + '개가 모두 문서에 있다', missing.length === 0, missing.join(', '));

console.log('');
if(fail === 0){ console.log('===== 전부 통과 (' + pass + '건) ====='); process.exit(0); }
else { console.log('===== 성공 ' + pass + ' / 실패 ' + fail + ' ====='); process.exit(1); }
