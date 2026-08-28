/* =============================================================
 *  interop.js - HTML 버전이 엑셀 버전과 호환되는지 검증
 *
 *  실행:  node build\interop.js
 *         (먼저 build\simulate.ps1 -FixtureDir <dir> 로 기준 파일을 만든다.
 *          build\interop.ps1 이 두 단계를 순서대로 돌려 준다)
 *
 *  방법
 *    1) web\스케줄러.html 에서 유틸·동기화 부분만 떼어내 그대로 실행한다.
 *       (사본을 따로 두지 않는다. 실제 배포되는 코드를 검사해야 의미가 있다)
 *    2) PowerShell 기준 구현이 만든 파트 텍스트와 바이트 단위로 비교한다.
 *    3) 기준 파트를 JS 파서에 먹여 원본 레코드가 그대로 복원되는지 본다.
 *
 *  PowerShell 기준 구현은 VBA 와 같은 규칙임이 simulate.ps1 37건으로 확인됐다.
 *  따라서 여기서 일치하면 HTML 버전도 엑셀 버전과 텍스트를 주고받을 수 있다.
 * ============================================================= */
'use strict';
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const htmlPath = path.join(root, 'web', '스케줄러.html');
const fixDir = process.argv[2] || path.join(root, 'build', '.fixture');

let pass = 0, fail = 0;
function chk(name, cond, detail){
  if(cond){ pass++; console.log('  [통과] ' + name); }
  else { fail++; console.log('  [실패] ' + name + '  -> ' + (detail === undefined ? '' : detail)); }
}

/* ---- 1. 실제 HTML 에서 검사 대상 코드 떼어내기 ---- */
const src = fs.readFileSync(htmlPath, 'utf8');

function slice(startAnchor, endAnchor, label){
  const a = src.indexOf(startAnchor);
  if(a < 0) throw new Error('앵커를 찾지 못했습니다(' + label + ' 시작): ' + startAnchor);
  const b = src.indexOf(endAnchor, a);
  if(b < 0) throw new Error('앵커를 찾지 못했습니다(' + label + ' 끝): ' + endAnchor);
  return src.slice(a, b);
}

const utilsCode = slice('const pad = (n, w)', 'const LS_KEY', '유틸');
const syncCode  = slice("const PROTO = 'WSCH1'", 'function closeModal', '동기화');

/* DB 등 화면 쪽 의존성은 최소한으로 흉내낸다 */
const prelude = `
let DB = { share: [], cfg: {}, log: [] };
const alive = list => list.filter(r => !r.del);
const myName = () => (DB.cfg.name || '').trim() || '사용자';
const markDirty = () => {};
`;   // addLog / shareSummary / docNoOwner 는 HTML 원본에 들어 있는 것을 그대로 쓴다
const tail = `
module.exports = {
  esc, unesc, checksum4, normTime, parseDate, dateStr,
  shareToLine, lineToShare, buildParts, feed, inboxRecords,
  syncPreview, syncApply, resetInbox,
  setDB: d => { DB = d; }, getDB: () => DB,
  FD
};
`;

const modPath = path.join(fixDir, '_extracted.js');
fs.mkdirSync(fixDir, { recursive: true });
fs.writeFileSync(modPath, prelude + utilsCode + '\n' + syncCode + '\n' + tail, 'utf8');

let M;
try {
  M = require(modPath);
} catch(e){
  console.log('  [실패] HTML 에서 떼어낸 코드가 실행되지 않습니다 -> ' + e.message);
  process.exit(1);
}

console.log('');
console.log('===== HTML ↔ 엑셀 호환 검증 =====');
console.log('');
console.log('1. HTML 코드 적재');
chk('유틸·동기화 코드를 떼어내 실행', true);

/* ---- 2. 기준 파일과 대조 ---- */
console.log('');
console.log('2. 기준 구현(PowerShell = VBA 규칙)과 대조');

const psPartsPath = path.join(fixDir, 'ps_parts.json');
const psRecsPath  = path.join(fixDir, 'ps_records.json');
if(!fs.existsSync(psPartsPath)){
  console.log('  [실패] 기준 파일이 없습니다: ' + psPartsPath);
  console.log('         build\\simulate.ps1 -FixtureDir "' + fixDir + '" 를 먼저 실행하세요.');
  process.exit(1);
}
const psParts = JSON.parse(fs.readFileSync(psPartsPath, 'utf8'));
const psRecs  = JSON.parse(fs.readFileSync(psRecsPath, 'utf8'));
if(!Array.isArray(psParts) || !psParts.length || typeof psParts[0] !== 'string'){
  console.log('  [실패] 기준 파트 파일이 문자열 배열이 아닙니다: ' + JSON.stringify(psParts).slice(0, 120));
  process.exit(1);
}
if(!Array.isArray(psRecs) || !psRecs.length){
  console.log('  [실패] 기준 레코드 파일이 배열이 아닙니다.');
  process.exit(1);
}

// PowerShell 쪽 레코드(한글 키)를 JS 레코드 모양으로
const toJs = r => ({
  id: r['ID'], docNo: r['문서번호'], title: r['제목'], content: r['내용'],
  regDate: r['등록일'], dueDate: r['회신기한'], author: r['작성자'],
  mtime: r['수정시각'], del: r['삭제여부'] === '1'
});
const recs = psRecs.map(toJs);

// (1) 옛 형식(B= 배치ID 없음)도 그대로 읽히는가
//     기준 파일은 배치ID 도입 전 형식이다. 이미 나눠준 파일이 보내오는 텍스트를
//     새 버전이 계속 받을 수 있어야 한다.
chk('기준 파트가 옛 형식(B= 없음)인지 확인',
    psParts.every(p => p.indexOf('|B=') < 0), '기준 파일에 B= 가 들어 있음');

// (2) 기준 파트를 JS 파서에 먹여 원본이 복원되는가
M.resetInbox();
let st = null;
for(let i = psParts.length - 1; i >= 0; i--) st = M.feed(psParts[i]);   // 일부러 역순
chk('엑셀이 만든 텍스트를 역순으로 넣어도 완료', st && st.st === M.FD.COMPLETE,
    st ? ('상태 ' + st.st + ' / ' + st.msg) : '결과 없음');

const back = M.inboxRecords();
chk('복원된 건수 일치 (' + Object.keys(back).length + '건)',
    Object.keys(back).length === recs.length,
    Object.keys(back).length + ' vs ' + recs.length);

let fieldOk = true;
for(const r of recs){
  const b = back[r.id];
  if(!b){ fieldOk = false; console.log('     누락: ' + r.id); continue; }
  for(const k of ['docNo','title','content','regDate','dueDate','author','mtime','del']){
    if(String(b[k]) !== String(r[k])){
      fieldOk = false;
      console.log('     불일치 ' + r.id + '.' + k);
      console.log('       기대: ' + JSON.stringify(r[k]));
      console.log('       실제: ' + JSON.stringify(b[k]));
    }
  }
}
chk('모든 필드가 원본과 일치 (파이프·백슬래시·줄바꿈·탭 포함)', fieldOk, '위 목록 참조');

/* ---- 3. JS 단독 동작 ---- */
console.log('');
console.log('3. HTML 버전 자체 동작');

// 이스케이프 왕복
const nasty = ['\\\\p', '\\p', '\\\\', '\\', 'a\\\\\\\\pb', '\\n', '|\\|', '끝에역슬래시\\',
               '여러줄\n\n연속', '탭\t\t연속', '||||', '\\\\\\\\\\\\'];
let escOk = true;
for(const s of nasty){ if(M.unesc(M.esc(s)) !== s){ escOk = false; console.log('     깨짐: ' + JSON.stringify(s)); } }
chk('악성 입력 ' + nasty.length + '종 이스케이프 왕복', escOk);

// 체크섬
const line = 'S|홍길동-20260827-143052-8471|기획-2026-0451|상반기 실적 보고|본문|2026-08-27|2026-08-30|홍길동|2026-08-27 14:30:52|0';
chk('체크섬이 한 글자 변형을 잡아냄',
    M.checksum4(line) !== M.checksum4(line.replace('0451','0452')));
chk('체크섬이 4자리 16진수', /^[0-9A-F]{4}$/.test(M.checksum4(line)), M.checksum4(line));

// 잘림 / 변형 / 머리말 없음
M.resetInbox();
chk('본문 변형을 잡아냄', M.feed(psParts[0].replace('평범한', '평범핝')).st === M.FD.BADSUM);
M.resetInbox();
chk('머리말이 없으면 형식 불일치', M.feed('그냥 아무 말이나').st === M.FD.NOHEADER);
M.resetInbox();
chk('줄바꿈이 LF 로 바뀌어도 통과',
    (() => { for(const p of psParts) st = M.feed(p.replace(/\r\n/g, '\n')); return st.st === M.FD.COMPLETE; })());
M.resetInbox();
chk('줄 끝 공백이 붙어도 통과',
    (() => { for(const p of psParts) st = M.feed(p.split('\r\n').map(l => l + '   ').join('\r\n')); return st.st === M.FD.COMPLETE; })());

// 멱등
M.resetInbox();
M.feed(psParts[0]);
const dup = M.feed(psParts[0]);
chk('같은 조각 재투입은 중복으로 쌓이지 않음',
    dup.st === M.FD.DUPLICATE || dup.st === M.FD.COMPLETE, '상태 ' + dup.st);

// 병합 규칙
function freshDB(share){ return { share: share, cfg:{name:'테스터'}, log:[] }; }

M.resetInbox();
M.setDB(freshDB([{id:'FX-2', docNo:'DOC-002', title:'내가고친제목', content:'', regDate:'2026-01-01',
                  dueDate:'2026-01-10', author:'나', mtime:'2099-01-01 00:00:00', del:false}]));
for(const p of psParts) M.feed(p);
M.syncApply();
chk('과거 수정시각이 최신을 덮어쓰지 않음',
    M.getDB().share.find(r => r.id === 'FX-2').title === '내가고친제목',
    M.getDB().share.find(r => r.id === 'FX-2').title);

// 내가 지운 뒤로 상대가 손대지 않았다면(= 상대 텍스트가 더 과거) 되살아나면 안 된다.
M.resetInbox();
M.setDB(freshDB([{id:'FX-2', docNo:'DOC-002', title:'내가지움', content:'', regDate:'2026-01-01',
                  dueDate:'2026-01-10', author:'나', mtime:'2099-01-01 00:00:00', del:true}]));
for(const p of psParts) M.feed(p);
M.syncApply();
chk('내가 지운 건은 옛 텍스트로 되살아나지 않음',
    M.getDB().share.find(r => r.id === 'FX-2').del === true,
    '삭제여부 ' + M.getDB().share.find(r => r.id === 'FX-2').del);

// 반대로 내가 지운 뒤 상대가 더 최신으로 손봤다면 되살아나야 한다.
M.resetInbox();
M.setDB(freshDB([{id:'FX-2', docNo:'DOC-002', title:'내가지움', content:'', regDate:'2026-01-01',
                  dueDate:'2026-01-10', author:'나', mtime:'2000-01-01 00:00:00', del:true}]));
for(const p of psParts) M.feed(p);
M.syncApply();
chk('상대가 더 최신으로 손본 건은 되살아남',
    M.getDB().share.find(r => r.id === 'FX-2').del === false,
    '삭제여부 ' + M.getDB().share.find(r => r.id === 'FX-2').del);

M.resetInbox();
M.setDB(freshDB([{id:'FX-3', docNo:'DOC-003', title:'살아있음', content:'', regDate:'2026-01-01',
                  dueDate:'2026-01-10', author:'나', mtime:'2099-01-01 00:00:00', del:false}]));
for(const p of psParts) M.feed(p);
M.syncApply();
chk('남이 지운 표식은 수정시각이 과거여도 전달됨',
    M.getDB().share.find(r => r.id === 'FX-3').del === true);

M.resetInbox();
M.setDB(freshDB([]));
for(const p of psParts) M.feed(p);
M.syncApply();
chk('빈 상태에 전부 신규로 들어감 (' + M.getDB().share.length + '건)',
    M.getDB().share.length === recs.length);

/* ---- 결과 ---- */
console.log('');
if(fail === 0){
  console.log('===== 전부 통과 (' + pass + '건) =====');
  process.exit(0);
}else{
  console.log('===== 성공 ' + pass + ' / 실패 ' + fail + ' =====');
  process.exit(1);
}
