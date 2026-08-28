/* =============================================================
 *  batch-test.js - 여러 사람이 보낸 조각이 섞일 때의 동작 검증
 *
 *  검토 문서 1-1 의 두 가지 재현 시나리오를 실제로 돌려 본다.
 *  실행:  node build\batch-test.js
 * ============================================================= */
'use strict';
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const htmlPath = path.join(root, 'dist', '배포', 'sched.html');
const src = fs.readFileSync(htmlPath, 'utf8');

function slice(a, b){
  const i = src.indexOf(a), j = src.indexOf(b, i);
  if(i < 0 || j < 0) throw new Error('앵커 없음: ' + a);
  return src.slice(i, j);
}
const code =
  'let DB = { share: [], cfg: {}, log: [] };\n' +
  'const alive = l => l.filter(r => !r.del);\n' +
  'const myName = () => DB.cfg.name || "사용자";\n' +
  'const markDirty = () => {};\n' +
  slice('const pad = (n, w)', 'const LS_KEY') + '\n' +
  slice("const PROTO = 'WSCH1'", 'function closeModal') + '\n' +
  'module.exports = { buildParts, feed, inboxRecords, syncPreview, syncApply, resetInbox, FD,' +
  ' setDB: d => { DB = d; }, getDB: () => DB };\n';

const modPath = path.join(__dirname, '.fixture', '_batch.js');
fs.mkdirSync(path.dirname(modPath), { recursive: true });
fs.writeFileSync(modPath, code, 'utf8');
const M = require(modPath);

let pass = 0, fail = 0;
function chk(name, cond, detail){
  if(cond){ pass++; console.log('  [통과] ' + name); }
  else { fail++; console.log('  [실패] ' + name + '  -> ' + (detail === undefined ? '' : detail)); }
}

function rec(id, doc, title, mtime){
  return { id, docNo:doc, title, content:'내용 ' + title, regDate:'2026-08-01',
           dueDate:'2026-09-01', author:id.split('-')[0], mtime, del:false };
}
function partsOf(list, sender){
  M.setDB({ share:list, cfg:{ name:sender }, log:[] });
  return M.buildParts(list);
}

console.log('');
console.log('===== 여러 사람 조각 섞임 검증 =====');

/* ---- 재현 1: 1조각짜리 두 개가 연달아 들어옴 ---- */
console.log('');
console.log('1. 서로 다른 사람이 보낸 1조각짜리를 이어서 붙여넣기');

const kim  = [rec('김대리-20260827-100000-1111', 'DOC-K', '김대리 건', '2026-08-27 10:00:00')];
const park = [rec('박과장-20260827-110000-2222', 'DOC-P', '박과장 건', '2026-08-27 11:00:00')];
const pk = partsOf(kim, '김대리');
const pp = partsOf(park, '박과장');
chk('둘 다 1조각', pk.length === 1 && pp.length === 1, pk.length + ',' + pp.length);

M.setDB({ share:[], cfg:{ name:'수신자' }, log:[] });
M.resetInbox();
const r1 = M.feed(pk[0]);
const r2 = M.feed(pp[0]);
const got = M.inboxRecords();
const ids = Object.keys(got);

console.log('     1번째 투입: ' + r1.msg);
console.log('     2번째 투입: ' + r2.msg);
console.log('     버퍼에 남은 건: ' + (ids.length ? ids.join(', ') : '(없음)'));

chk('두 사람 것이 모두 살아 있어야 한다', ids.length === 2,
    '실제 ' + ids.length + '건 — 한쪽이 조용히 버려짐');

M.syncApply();
chk('반영 후에도 2건', M.getDB().share.length === 2,
    '실제 ' + M.getDB().share.length + '건: ' + M.getDB().share.map(r => r.docNo).join(', '));

/* ---- 재현 2: 3조각 중 2개 받은 상태에서 다른 배치가 끼어듦 ---- */
console.log('');
console.log('2. 3조각 받는 중에 2조각짜리가 끼어듦');

const big = [];
for(let i = 0; i < 3; i++) big.push(rec('김대리-2026082710000' + i + '-' + i, 'BIG-' + i, '큰 건 ' + i, '2026-08-27 10:0' + i + ':00'));
M.setDB({ share:big, cfg:{ name:'김대리', partMax:250 }, log:[] });
const pbig = M.buildParts(big);

const mid = [];
for(let i = 0; i < 2; i++) mid.push(rec('박과장-2026082711000' + i + '-' + i, 'MID-' + i, '중간 건 ' + i, '2026-08-27 11:0' + i + ':00'));
M.setDB({ share:mid, cfg:{ name:'박과장', partMax:250 }, log:[] });
const pmid = M.buildParts(mid);

console.log('     김대리 ' + pbig.length + '조각 / 박과장 ' + pmid.length + '조각');

M.setDB({ share:[], cfg:{ name:'수신자', partMax:250 }, log:[] });
M.resetInbox();
M.feed(pbig[0]);
if(pbig[1]) M.feed(pbig[1]);
const beforeCount = Object.keys(M.inboxRecords()).length;
const r3 = M.feed(pmid[0]);
const afterCount = Object.keys(M.inboxRecords()).length;

console.log('     끼어들기 전 버퍼: ' + beforeCount + '건');
console.log('     끼어든 뒤 버퍼:   ' + afterCount + '건');
console.log('     상태 문구: ' + r3.msg);

chk('앞서 받은 조각이 경고 없이 사라지면 안 된다', afterCount >= beforeCount,
    beforeCount + '건 -> ' + afterCount + '건으로 줄어듦');

/* ---- 3. 같은 사람이 같은 배치를 두 번 넣는 것은 여전히 안전해야 ---- */
console.log('');
console.log('3. 같은 배치 재투입은 여전히 멱등이어야 한다');
M.setDB({ share:[], cfg:{ name:'수신자' }, log:[] });
M.resetInbox();
M.feed(pk[0]);
M.feed(pk[0]);
M.syncApply();
chk('같은 것을 두 번 넣어도 1건', M.getDB().share.length === 1, M.getDB().share.length + '건');

console.log('');
if(fail === 0){ console.log('===== 전부 통과 (' + pass + '건) ====='); process.exit(0); }
else { console.log('===== 성공 ' + pass + ' / 실패 ' + fail + ' ====='); process.exit(1); }
