/* =============================================================
 *  undef-test.js - 선언되지 않은 이름을 쓰고 있는 곳 찾기
 *  실행:  node build\undef-test.js
 *
 *  왜 만들었나
 *    가져오기 창이 통째로 안 열린 적이 있다. 원인은 drawSteps() 안에서
 *    없어진 변수(inboxTotal)를 읽는 줄 하나였다. 거기서 ReferenceError 가
 *    나면서 그 아래 이벤트 연결이 전부 실행되지 않았다.
 *    문법 검사(new Function)로는 못 잡는다. 실행해 봐야 알기 때문이다.
 *    그래서 "읽기만 하고 이 파일 어디서도 선언한 적 없는 이름" 을 글자로 찾는다.
 *
 *  한계
 *    정식 파서가 아니라 글자를 훑는 방식이다. 놓치는 것이 있을 수 있다.
 *    잘못 잡는 것(오탐)은 아래 KNOWN 에 적어서 없앤다.
 * ============================================================= */
'use strict';
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const src = fs.readFileSync(path.join(root, 'web', '스케줄러.html'), 'utf8');

// 본문 스크립트만 떼어 온다 (데이터 블록은 JSON 이라 제외)
const marks = [];
const reScript = /<script(?![^>]*type="application\/json")[^>]*>([\s\S]*?)<\/script>/g;
let mm;
while((mm = reScript.exec(src))) marks.push({ code: mm[1] });
if(!marks.length) throw new Error('스크립트를 찾지 못했습니다.');

const NL = String.fromCharCode(10);
const ID_SRC = '[A-Za-z_$][\\w$]*';
const isId = s => new RegExp('^' + ID_SRC + '$').test(s);

/* 주석과 따옴표 안의 글은 코드가 아니다. 자리는 그대로 두고 공백으로 지운다.
   (자리를 지켜야 몇 번째 줄인지 그대로 셀 수 있다) */
function blankOut(code){
  const out = code.split('');
  const n = code.length;
  let i = 0;
  const isRegexPos = k => {
    // 정규식인지 나눗셈인지 - 바로 앞의 뜻있는 글자로 가른다
    for(let j = k - 1; j >= 0; j--){
      const c = code[j];
      if(c === ' ' || c === '\t' || c === NL || c === '\r') continue;
      return !/[\w$)\]]/.test(c);
    }
    return true;
  };
  while(i < n){
    const c = code[i];
    if(c === '/' && code[i+1] === '/'){
      while(i < n && code[i] !== NL){ out[i] = ' '; i++; }
      continue;
    }
    if(c === '/' && code[i+1] === '*'){
      const end = code.indexOf('*/', i + 2);
      const stop = end < 0 ? n : end + 2;
      for(; i < stop; i++) if(code[i] !== NL) out[i] = ' ';
      continue;
    }
    if(c === '"' || c === "'" || c === '`'){
      const q = c;
      out[i] = ' '; i++;
      while(i < n){
        if(code[i] === '\\'){ out[i] = ' '; out[i+1] = ' '; i += 2; continue; }
        if(code[i] === q){ out[i] = ' '; i++; break; }
        if(code[i] !== NL) out[i] = ' ';
        i++;
      }
      continue;
    }
    if(c === '/' && isRegexPos(i)){
      out[i] = ' '; i++;
      let cls = false;
      while(i < n && code[i] !== NL){
        if(code[i] === '\\'){ out[i] = ' '; out[i+1] = ' '; i += 2; continue; }
        if(code[i] === '[') cls = true;
        else if(code[i] === ']') cls = false;
        else if(code[i] === '/' && !cls){ out[i] = ' '; i++; break; }
        out[i] = ' '; i++;
      }
      while(i < n && /[gimsuy]/.test(code[i])){ out[i] = ' '; i++; }
      continue;
    }
    i++;
  }
  return out.join('');
}

/* 어디선가 선언한 적이 있는 이름을 모두 모은다.
   범위(스코프)는 따지지 않는다. 여기서 잡으려는 것은 스코프 실수가 아니라
   "이 파일 어디에도 없는 이름" 이기 때문이다. */
function declaredNames(code){
  const set = new Set();
  const add = s => { if(s && isId(s)) set.add(s); };
  const argNames = list => list.split(',').forEach(p => {
    add(p.split('=')[0].replace(/\.\.\./g, '').trim());
  });
  let r, x;

  // function / class 이름
  r = new RegExp('\\b(?:function|class)\\s*\\*?\\s*(' + ID_SRC + ')', 'g');
  while((x = r.exec(code))) add(x[1]);

  /* const/let/var 는 한 줄에 여러 개를 붙여 쓴다.
       let a = 1, b = 2, {c} = o;
     첫 이름만 보면 b, c 를 놓쳐서 멀쩡한 변수를 없는 이름이라고 잡는다.
     그래서 선언 목록 끝까지 읽는다. */
  r = /\b(?:const|let|var)\s+/g;
  while((x = r.exec(code))){
    let depth = 0, buf = '';
    const pieces = [];
    for(let i = x.index + x[0].length; i < code.length; i++){
      const c = code[i];
      if('([{'.indexOf(c) >= 0){ depth++; }
      else if(')]}'.indexOf(c) >= 0){ if(depth === 0) break; depth--; }
      else if(c === ';' && depth === 0) break;
      else if(c === ',' && depth === 0){ pieces.push(buf); buf = ''; continue; }
      else if(c === NL && depth === 0 && !/[=,+\-*/&|?:(]\s*$/.test(buf)) break;
      buf += c;
    }
    pieces.push(buf);
    for(const piece of pieces){
      const head = piece.split('=')[0].trim();
      if(isId(head)){ add(head); continue; }
      if(/^[{[]/.test(head)){          // 구조 분해 - 이름만 훑어 담는다
        head.replace(/[{}[\]]/g, ' ').split(',').forEach(t => {
          add(t.split(':').pop().replace(/\.\.\./g, '').trim());
        });
      }
    }
  }

  // function f(a, b) / catch(e)
  r = new RegExp('(?:\\bfunction\\s*\\*?\\s*(?:' + ID_SRC + ')?\\s*|\\bcatch\\s*)\\(([^)]*)\\)', 'g');
  while((x = r.exec(code))) argNames(x[1]);

  // (a, b) => ...
  r = /\(([^()]*)\)\s*=>/g;
  while((x = r.exec(code))) argNames(x[1]);

  // a => ...
  r = new RegExp('(?:^|[^\\w$.])(' + ID_SRC + ')\\s*=>', 'g');
  while((x = r.exec(code))) add(x[1]);

  // { foo(a, b){ ... } } 축약 메서드의 인자
  r = new RegExp('(?:^|[,{' + NL + '])\\s*(?:' + ID_SRC + ')\\s*\\(([^)]*)\\)\\s*\\{', 'g');
  while((x = r.exec(code))) argNames(x[1]);

  // for(const x of ...) / for(let i = 0; ...)
  r = new RegExp('\\bfor\\s*\\(\\s*(?:const|let|var)\\s+(' + ID_SRC + ')', 'g');
  while((x = r.exec(code))) add(x[1]);

  // for(const [a, b] of ...) / for(const {a} of ...) - 구조 분해도 선언이다
  r = /\bfor\s*\(\s*(?:const|let|var)\s*([{[][^)]*?[}\]])\s+(?:of|in)\b/g;
  while((x = r.exec(code))){
    x[1].replace(/[{}[\]]/g, ' ').split(',').forEach(t => {
      add(t.split(':').pop().replace(/\.\.\./g, '').trim());
    });
  }

  return set;
}

// 브라우저와 언어가 원래 주는 이름들
const KNOWN = new Set(`
window document navigator location history screen console alert confirm prompt
setTimeout clearTimeout setInterval clearInterval requestAnimationFrame
localStorage sessionStorage getComputedStyle matchMedia
Object Array String Number Boolean Math JSON Date RegExp Error TypeError RangeError
Map Set WeakMap WeakSet Promise Symbol Proxy Reflect Intl BigInt
Function parseInt parseFloat isNaN isFinite encodeURIComponent decodeURIComponent
encodeURI decodeURI escape unescape structuredClone queueMicrotask
Blob File FileReader FormData URL URLSearchParams Image Audio
TextEncoder TextDecoder DecompressionStream CompressionStream Response Request
DOMParser XMLSerializer Node Element HTMLElement Event CustomEvent
Uint8Array Uint16Array Uint32Array Int8Array Int16Array Int32Array
Float32Array Float64Array ArrayBuffer DataView
arguments globalThis
`.trim().split(/\s+/));

// 예약어는 이름이 아니다
const RESERVED = new Set(`
if else for while do switch case default break continue return function class
const let var new typeof instanceof in of delete void try catch finally throw
this super extends static get set async await yield import export
true false null undefined NaN Infinity
`.trim().split(/\s+/));

let fail = 0;
const found = [];

for(const part of marks){
  const clean = blankOut(part.code);
  const declared = declaredNames(clean);
  const lines = part.code.split(NL);
  const scan = /[A-Za-z_$][\w$]*/g;
  let x;
  while((x = scan.exec(clean))){
    const name = x[0];
    if(RESERVED.has(name) || KNOWN.has(name) || declared.has(name)) continue;

    const prev = clean.slice(Math.max(0, x.index - 40), x.index);
    const next = clean.slice(x.index + name.length, x.index + name.length + 24);
    if(/[.?]\s*$/.test(prev)) continue;               // obj.name / obj?.name
    if(/^\s*:/.test(next)) continue;                  // { name: ... }
    if(/[\w$]/.test(clean[x.index - 1] || '')) continue;

    const line = clean.slice(0, x.index).split(NL).length;
    found.push({ name, line, near: (lines[line - 1] || '').trim().slice(0, 90) });
  }
}

console.log('선언 없이 쓰는 이름 찾기');
console.log('');
if(found.length){
  const seen = new Set();
  for(const b of found){
    if(seen.has(b.name)) continue;
    seen.add(b.name);
    fail++;
    console.log('  [실패] ' + b.name + '  (' + b.line + '번째 줄)');
    console.log('         ' + b.near);
  }
}else{
  console.log('  [통과] 어디서도 선언하지 않은 이름을 쓰는 곳이 없습니다');
}

console.log('');
if(fail === 0){ console.log('===== 이상 없음 ====='); process.exit(0); }
else { console.log('===== ' + fail + '개 찾음 ====='); process.exit(1); }
