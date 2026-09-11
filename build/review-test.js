/* =============================================================
 *  review-test.js - 검토 지적사항 수정분 검증
 *  실행:  node build\review-test.js
 * ============================================================= */
'use strict';
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const src = fs.readFileSync(path.join(root, 'dist', '배포', 'sched.html'), 'utf8');
function slice(a, b){
  const i = src.indexOf(a), j = src.indexOf(b, i);
  if(i < 0 || j < 0) throw new Error('앵커 없음: ' + a);
  return src.slice(i, j);
}

// empLookup ~ empExt 범위 (앵커가 애매해 직접 잡는다)
const empStart = src.indexOf('function empLookup(');
const empEnd = src.indexOf('function parseEmpCsv') > 0 ? src.indexOf('function parseEmpCsv') : src.indexOf('const EMP_HEAD');
const empCode = src.slice(empStart, empEnd);

const syncCode = slice("const PROTO = 'WSCH1'", '/* ------------------------------------------------------------\n   검색');
const searchCode = slice('const KIND_LABEL', 'function searchModal');
// carryToToday 는 renderCarry 뒤에 있으므로 renderToday 앞까지 잘라 온다
// QA_WEEK ~ carryToToday 까지 (renderToday 앞에서 끊는다)
const carryCode = slice('const QA_WEEK', 'function renderToday()');
// 업무 정렬·순서 바꾸기
const sortCode = slice('/* 회신기한 관련 규칙 두 가지.', 'function taskDensity');
const replCode = slice('function myReplied(id)', String.fromCharCode(10));
// 제목/팀 이름, 메일 양식, 출장 공유 글
const titleCode = slice('const DEFAULT_TITLE', 'function paintTitle()');
const mailCode = slice('function teamName()', 'function exportModal()');

// 떼어낸 코드에는 브라우저에서만 도는 줄이 섞여 있다(이벤트 등록 등).
// 여기서는 순수 로직만 검사하므로 최소한의 껍데기만 세워 준다.
const mod =
  'const document = { addEventListener(){}, createElement: () => ({ style:{}, appendChild(){}, remove(){} }),' +
  ' body:{ appendChild(){} } };\n' +
  'const window = { getSelection: () => ({ removeAllRanges(){}, addRange(){} }) };\n' +
  'const $ = () => null; const $$ = () => [];\n' +
  'let DB = { task:[], contact:[], share:[], lunch:[], emp:[], fav:[], cfg:{}, log:[] };\n' +
  'const alive = l => l.filter(r => !r.del);\n' +
  'const myName = () => DB.cfg.name || "사용자";\n' +
  'const markDirty = () => {};\n' +
  slice('const pad = (n, w)', 'const LS_KEY') + '\n' +
  empCode + '\n' + syncCode + '\n' + searchCode + '\n' + carryCode + '\n' + sortCode + '\n' +
  titleCode + '\n' + mailCode + '\n' + replCode + '\n' +
  'module.exports = { empLookup, empExt, buildParts, lastBuildInfo: () => lastBuildInfo,' +
  ' searchAll, overdueTasks, carryToToday, mergeVerdict, parseQuickAdd,' +
  ' sortTasks, dayTasks, materializeOrder, nextOrderFor, moveTask,' +
  ' feed, inboxRecords, resetInbox, FD,' +
  ' splitMulti, joinMulti, TRIP_TRANSPORT, TRIP_REGION, TRIP_MAX_MEMBER,' +
  ' teamName, mailText, tripShareText, sharePark, shareSpanDays,' +
  ' setDB: d => { DB = d;\n' +
  '   for(const k of ["task","contact","share","lunch","trip","emp","fav","log"]) if(!DB[k]) DB[k] = [];\n' +
  '   if(!DB.cfg) DB.cfg = {}; },\n' +
  ' getDB: () => DB, todayStr, dateStr, addDays, startOfDay, parseDate };\n';

const modPath = path.join(__dirname, '.fixture', '_review.js');
fs.mkdirSync(path.dirname(modPath), { recursive: true });
fs.writeFileSync(modPath, mod, 'utf8');
const M = require(modPath);

let pass = 0, fail = 0;
function chk(n, c, d){
  if(c){ pass++; console.log('  [통과] ' + n); }
  else { fail++; console.log('  [실패] ' + n + '  -> ' + (d === undefined ? '' : d)); }
}

console.log('');
console.log('===== 검토 수정분 검증 =====');

/* ---- 1-6 동명이인 오조회 ---- */
console.log('');
console.log('1-6. 동명이인이면 번호를 확정하지 않는다');
M.setDB({ task:[], contact:[], share:[], lunch:[], cfg:{}, log:[],
  emp:[ {name:'이지은', dept:'기획팀', ext:'1101'},
        {name:'이지은', dept:'총무팀', ext:'1102'},
        {name:'이지은', dept:'영업팀', ext:'1103'},
        {name:'박유일', dept:'인사팀', ext:'1200'} ] });

const amb = M.empLookup('이지은', '');
chk('부서 없이 동명이인 -> ambiguous', amb.state === 'ambiguous', amb.state);
chk('그때 번호를 주지 않는다', M.empExt('이지은', '') === '', '[' + M.empExt('이지은','') + ']');
chk('몇 명인지 알려준다', amb.count === 3, amb.count);
chk('부서를 주면 정확히 찾는다', M.empExt('이지은', '총무팀') === '1102', M.empExt('이지은','총무팀'));
chk('동명이인 아니면 부서 없이도 찾는다', M.empExt('박유일', '') === '1200', M.empExt('박유일',''));
chk('없는 이름은 빈 값', M.empExt('없는사람', '') === '');

/* ---- 1-2 미완료 이월 ---- */
console.log('');
console.log('1-2. 지난 미완료 이월');
const today = M.todayStr();
M.setDB({ emp:[], contact:[], share:[], lunch:[], cfg:{}, log:[], task:[
  {id:'t1', title:'어제 못 끝낸 일', date:'2026-01-05', done:false, del:false, mtime:'x'},
  {id:'t2', title:'그제 것', date:'2026-01-04', done:false, del:false, mtime:'x'},
  {id:'t3', title:'어제 끝낸 일', date:'2026-01-05', done:true,  del:false, mtime:'x'},
  {id:'t4', title:'지운 것',     date:'2026-01-05', done:false, del:true,  mtime:'x'},
  {id:'t5', title:'오늘 것',     date:today,        done:false, del:false, mtime:'x'}
]});
const od = M.overdueTasks();
chk('미완료 지난 건만 2건', od.length === 2, od.map(r => r.id).join(','));
chk('오래된 것부터 정렬', od[0].id === 't2', od[0].id);
chk('완료·삭제·오늘 것은 빠짐', !od.some(r => ['t3','t4','t5'].includes(r.id)));

M.carryToToday('t1');
const t1 = M.getDB().task.find(r => r.id === 't1');
chk('오늘로 옮기면 날짜가 오늘', t1.date === today, t1.date);
chk('원래 날짜를 origDate 에 남김', t1.origDate === '2026-01-05', t1.origDate);
chk('옮긴 뒤 이월 목록에서 빠짐', M.overdueTasks().length === 1);

/* ---- 2-1 검색 ---- */
console.log('');
console.log('2-1. 검색');
M.setDB({ emp:[], cfg:{}, log:[],
  task:[{id:'a1', title:'예산 조정 회의', content:'3층 회의실', date:'2026-03-02', done:false, del:false}],
  share:[{id:'b1', docNo:'기획-2026-0451', title:'상반기 실적', content:'본문', dueDate:'2026-03-10', author:'홍길동', del:false}],
  contact:[{id:'c1', name:'김도현', dept:'기술지원팀', date:'2026-03-05', done:false, del:false}],
  lunch:[{id:'d1', target:'박준영', place:'본관 한식당', date:'2026-03-06', del:false}] });

chk('업무 제목으로 찾음', M.searchAll('예산').length === 1);
chk('업무 내용으로도 찾음', M.searchAll('회의실').length === 1);
chk('문서번호로 찾음', M.searchAll('2026-0451').length === 1, M.searchAll('2026-0451').length);
chk('작성자로 찾음', M.searchAll('홍길동').length === 1);
chk('연락처 이름으로 찾음', M.searchAll('김도현').length === 1);
chk('점심 장소로 찾음', M.searchAll('한식당').length === 1);
chk('대소문자 무시', M.searchAll('기획').length === 1);
chk('없는 말은 0건', M.searchAll('없는말').length === 0);
chk('빈 검색어는 0건', M.searchAll('').length === 0);
chk('종류가 섞여 나옴', M.searchAll('2026').length >= 1);

/* ---- 1-7 조각 크기가 머리말·꼬리말을 포함하는가 ---- */
console.log('');
console.log('1-7. 조각 하나가 설정한 글자 수를 넘지 않는가');
const many = [];
for(let i = 0; i < 30; i++){
  many.push({ id:'홍길동-2026082714000' + (i % 10) + '-' + (1000 + i), docNo:'기획-2026-' + (400 + i),
    title:'협조 요청 건 ' + i, content:'본문 내용이 어느 정도 있는 건 ' + i + ' 입니다. 확인 부탁드립니다.',
    regDate:'2026-08-27', dueDate:'2026-09-02', author:'홍길동', mtime:'2026-08-27 14:00:0' + (i % 10), del:false });
}
for(const limit of [500, 800, 1500]){
  M.setDB({ emp:[], task:[], contact:[], lunch:[], log:[], cfg:{ name:'홍길동', partMax:limit }, share:many });
  const ps = M.buildParts(many);
  const over = ps.filter(p => p.length > limit);
  chk('조각당 ' + limit + '자 설정 -> 실제로 넘지 않음 (' + ps.length + '조각, 최대 ' + Math.max(...ps.map(p => p.length)) + '자)',
      over.length === 0, over.length + '개 초과');
}

/* ---- 1-8 긴 레코드 경고 ---- */
console.log('');
console.log('1-8. 조각 하나에 안 들어가는 건을 알려주는가');
const huge = [{ id:'홍길동-20260827-140000-1000', docNo:'기획-2026-0999', title:'아주 긴 건',
  content:'가'.repeat(2000), regDate:'2026-08-27', dueDate:'2026-09-02',
  author:'홍길동', mtime:'2026-08-27 14:00:00', del:false }];
M.setDB({ emp:[], task:[], contact:[], lunch:[], log:[], cfg:{ name:'홍길동', partMax:800 }, share:huge });
M.buildParts(huge);
const info = M.lastBuildInfo();
chk('너무 긴 건을 잡아냄', info.oversize.length === 1, info.oversize.length);
chk('어느 건인지 알려줌', info.oversize[0] && info.oversize[0].docNo === '기획-2026-0999',
    JSON.stringify(info.oversize[0] || {}));

/* ---- 1-3 삭제/되살리기 판정 ---- */
console.log('');
console.log('1-3. 삭제와 되살리기');
M.setDB({ emp:[], task:[], contact:[], lunch:[], log:[], cfg:{}, share:[
  { id:'x1', docNo:'D1', title:'내가지움', mtime:'2026-05-05 00:00:00', del:true }
]});
chk('옛 텍스트로는 안 되살아남',
    M.mergeVerdict({ id:'x1', docNo:'D1', title:'옛것', mtime:'2026-01-01 00:00:00', del:false }) === '변화없음');
chk('더 최신이면 되살아남',
    M.mergeVerdict({ id:'x1', docNo:'D1', title:'새것', mtime:'2026-09-09 00:00:00', del:false }) === '되살아남');

/* ---- 한 줄 빠른 입력 파서 ---- */
console.log('');
console.log('빠른 입력. 한 줄을 적으면 제목·날짜·시간으로 갈라내는가');
const base = M.startOfDay(new Date(2026, 7, 27));      // 2026-08-27
const T = M.todayStr();
const tomorrow = M.dateStr(M.addDays(M.startOfDay(new Date()), 1));
const dayAfter = M.dateStr(M.addDays(M.startOfDay(new Date()), 2));

function q(text){ return M.parseQuickAdd(text, base); }
// "2026-09-04" -> 요일 번호 (0=일)
function parseDateDow(s){ const d = new Date(+s.slice(0,4), +s.slice(5,7)-1, +s.slice(8,10)); return d.getDay(); }

let r;
r = q('보고서 초안 14:00');
chk('시간 14:00 을 떼어냄', r.title === '보고서 초안' && r.time === '14:00', JSON.stringify(r));
r = q('회의 준비 9시');
chk('9시 -> 09:00', r.title === '회의 준비' && r.time === '09:00', JSON.stringify(r));
r = q('점검 14시30분');
chk('14시30분 -> 14:30', r.title === '점검' && r.time === '14:30', JSON.stringify(r));
r = q('내일 팀 회의');
chk('내일 -> 다음날 날짜', r.title === '팀 회의' && r.date === tomorrow, JSON.stringify(r));
r = q('모레 출장 준비');
chk('모레 -> 이틀 뒤', r.title === '출장 준비' && r.date === dayAfter, JSON.stringify(r));
r = q('오늘 마감 확인');
chk('오늘 -> 오늘 날짜', r.title === '마감 확인' && r.date === T, JSON.stringify(r));
r = q('9/3 예산 회의');
chk('9/3 -> 9월 3일', r.title === '예산 회의' && r.date.endsWith('-09-03'), JSON.stringify(r));
r = q('9월 3일 예산 회의');
chk('9월 3일 형식도 인식', r.title === '예산 회의' && r.date.endsWith('-09-03'), JSON.stringify(r));
r = q('내일 보고서 검토 15:30');
chk('날짜와 시간을 같이', r.title === '보고서 검토' && r.date === tomorrow && r.time === '15:30', JSON.stringify(r));
r = q('그냥 할 일');
chk('아무것도 없으면 보고 있는 날짜', r.title === '그냥 할 일' && r.date === '2026-08-27' && r.time === '', JSON.stringify(r));
r = q('25:99 이상한 시간');
chk('말이 안 되는 시간은 제목으로 둔다', r.time === '' && r.title.includes('25:99'), JSON.stringify(r));
r = q('13월 40일 이상한 날짜');
chk('말이 안 되는 날짜도 제목으로', r.date === '2026-08-27' && r.title.includes('13월'), JSON.stringify(r));
r = q('   ');
chk('빈 입력은 제목 없음', r.title === '', JSON.stringify(r));
r = q('A/B 테스트 결과 정리');
chk('제목 속 슬래시를 날짜로 오인하지 않음', r.title === 'A/B 테스트 결과 정리', JSON.stringify(r));

console.log('');
console.log('  오전·오후');
r = q('보고 오후 2시');       chk('오후 2시 -> 14:00', r.time === '14:00' && r.title === '보고', JSON.stringify(r));
r = q('조회 오전 9시');       chk('오전 9시 -> 09:00', r.time === '09:00', JSON.stringify(r));
r = q('마감 오후 12시');      chk('오후 12시 -> 12:00', r.time === '12:00', JSON.stringify(r));
r = q('당직 오전 12시');      chk('오전 12시 -> 00:00', r.time === '00:00', JSON.stringify(r));
r = q('회식 저녁 7시');       chk('저녁 7시 -> 19:00', r.time === '19:00', JSON.stringify(r));
r = q('점검 오후 3:30');      chk('오후 3:30 -> 15:30', r.time === '15:30', JSON.stringify(r));
r = q('회의 2시반');          chk('2시반 -> 02:30', r.time === '02:30', JSON.stringify(r));
r = q('회의 오후 2시반');     chk('오후 2시반 -> 14:30', r.time === '14:30', JSON.stringify(r));
r = q('교육 14:00~16:00');    chk('시간 범위는 앞의 것', r.time === '14:00' && r.title === '교육', JSON.stringify(r));

console.log('');
console.log('  요일');
const dow = new Date().getDay();
/* 2026-09-11 개정 — 「까지」는 시작일이 아니라 마감일자로 넣는다.
   전에는 date 에 넣어서, 그날이 오기 전에는 목록에 안 보였다.
   마감인데 마감 당일에야 나타나는 셈이라 쓸모가 없었다.
   지금은 until 에 넣어 기준일부터 그날까지 매일 보인다. */
r = q('보고서 금요일까지');
chk('금요일까지 -> 금요일이 마감', parseDateDow(r.until) === 5 && r.usedDue === true && r.title === '보고서', JSON.stringify(r));
r = q('다음주 화요일 팀 회의');
chk('다음주 화요일 -> 화요일', parseDateDow(r.date) === 2 && r.title === '팀 회의', JSON.stringify(r));
r = q('이번주 목요일 점검');
chk('이번주 목요일 -> 목요일', parseDateDow(r.date) === 4 && r.title === '점검', JSON.stringify(r));
r = q('월요일 조회');
chk('요일만 적어도 인식', parseDateDow(r.date) === 1 && r.title === '조회', JSON.stringify(r));
r = q('다음주 화요일 회의');
const nextTue = r.date;
r = q('이번주 화요일 회의');
chk('다음주가 이번주보다 뒤', nextTue > r.date, nextTue + ' vs ' + r.date);

console.log('');
console.log('  상대 날짜 · 숫자');
r = q('3일 뒤 점검');        chk('3일 뒤', r.date === M.dateStr(M.addDays(M.startOfDay(new Date()),3)), JSON.stringify(r));
r = q('2주 후 보고');        chk('2주 후', r.date === M.dateStr(M.addDays(M.startOfDay(new Date()),14)), JSON.stringify(r));
r = q('글피 출장');          chk('글피 -> 3일 뒤', r.date === M.dateStr(M.addDays(M.startOfDay(new Date()),3)), JSON.stringify(r));
r = q('9.3 예산 회의');      chk('9.3 형식', r.date.endsWith('-09-03') && r.title === '예산 회의', JSON.stringify(r));
r = q('9-3 예산 회의');      chk('9-3 형식', r.date.endsWith('-09-03'), JSON.stringify(r));
r = q('27일 정산');          chk('27일 -> 이번달/다음달 27일', r.date.endsWith('-27') && r.title === '정산', JSON.stringify(r));
r = q('내일까지 보고서 제출'); chk('내일까지 -> 내일이 마감', r.until === tomorrow && r.usedDue === true && r.title === '보고서 제출', JSON.stringify(r));

console.log('');
console.log('  헷갈리면 안 되는 것');
r = q('2026년 계획 수립');   chk('연도는 날짜로 안 본다', r.title === '2026년 계획 수립', JSON.stringify(r));
/* 2026-09-11 개정 — 「N일 이내」를 이제 알아듣는다.
   전에는 30일만 날짜로 보고 '이내' 를 제목에 남겼다. */
r = q('30일 이내 처리');     chk('30일 이내 -> 30일 뒤가 마감', r.usedDue === true && r.title === '처리', JSON.stringify(r));
r = q('회의실 3층 예약');    chk('층수는 시간이 아니다', r.time === '' && r.title === '회의실 3층 예약', JSON.stringify(r));
r = q('자료 12부 인쇄');     chk('부수는 시간이 아니다', r.time === '' && r.title === '자료 12부 인쇄', JSON.stringify(r));

/* ---- 업무 정렬 · 끌어다 놓기 ---- */
console.log('');
console.log('업무 정렬 — 완료는 뒤로, 손으로 옮긴 순서는 유지');
const D = '2026-08-27';
function seed(){
  M.setDB({ emp:[], contact:[], share:[], lunch:[], cfg:{}, log:[], task:[
    { id:'A', title:'A', date:D, time:'11:00', done:false, del:false },
    { id:'B', title:'B', date:D, time:'09:00', done:false, del:false },
    { id:'C', title:'C', date:D, time:'',      done:false, del:false },
    { id:'D', title:'D', date:D, time:'08:00', done:true,  del:false },
    { id:'E', title:'E', date:D, time:'10:00', done:false, del:false }
  ]});
}
const ids = l => l.map(r => r.id).join('');

seed();
chk('처음엔 시간순 · 완료는 맨 뒤', ids(M.sortTasks(M.dayTasks(D))) === 'BEACD',
    ids(M.sortTasks(M.dayTasks(D))));

seed();
chk('완료 처리하면 뒤로 밀림', (() => {
  M.getDB().task.find(r => r.id === 'B').done = true;
  return ids(M.sortTasks(M.dayTasks(D))) === 'EACDB' || ids(M.sortTasks(M.dayTasks(D))) === 'EACBD';
})(), ids(M.sortTasks(M.dayTasks(D))));

seed();
M.moveTask(D, 'A', 'B', false);          // A 를 B 앞으로
chk('A 를 B 앞으로 끌면 순서가 바뀐다', ids(M.sortTasks(M.dayTasks(D))).startsWith('AB'),
    ids(M.sortTasks(M.dayTasks(D))));

seed();
M.moveTask(D, 'B', 'E', true);           // B 를 E 뒤로
chk('B 를 E 뒤로 끌기', (() => { const s = ids(M.sortTasks(M.dayTasks(D))); return s.indexOf('B') > s.indexOf('E'); })(),
    ids(M.sortTasks(M.dayTasks(D))));

seed();
M.moveTask(D, 'A', 'B', false);
chk('옮긴 뒤에도 완료 건은 계속 맨 뒤', ids(M.sortTasks(M.dayTasks(D))).endsWith('D'),
    ids(M.sortTasks(M.dayTasks(D))));

seed();
M.moveTask(D, 'A', 'B', false);
chk('옮기고 나면 그날 전부에 순서가 박힌다',
    M.dayTasks(D).every(r => typeof r.order === 'number'),
    JSON.stringify(M.dayTasks(D).map(r => [r.id, r.order])));

seed();
chk('아직 안 옮긴 날은 순서를 안 박는다', M.nextOrderFor(D) === undefined, M.nextOrderFor(D));
M.moveTask(D, 'A', 'B', false);
chk('옮긴 날에 새 업무는 맨 뒤 번호를 받는다', M.nextOrderFor(D) === 5, M.nextOrderFor(D));

seed();
M.moveTask(D, 'A', 'A', false);
chk('자기 자신에게 떨구면 아무 일 없음', ids(M.sortTasks(M.dayTasks(D))) === 'BEACD',
    ids(M.sortTasks(M.dayTasks(D))));

seed();
chk('없는 대상이면 아무 일 없음', M.moveTask(D, 'A', 'ZZZ', false) === false);

/* ---- 2-3 회신 표시가 동기화에 새지 않는가 ---- */
console.log('');
console.log('회신 표시는 내 컴퓨터에만 — 공유 텍스트에 절대 안 들어가야 한다');
const one = [{ id:'홍길동-20260827-140000-1000', docNo:'기획-2026-0451', title:'설비 교체 승인',
  content:'본문', regDate:'2026-08-27', dueDate:'2026-09-02',
  author:'홍길동', mtime:'2026-08-27 14:00:00', del:false }];
M.setDB({ emp:[], task:[], contact:[], lunch:[], log:[],
          cfg:{ name:'홍길동', shareDone:{ '홍길동-20260827-140000-1000':1 } }, share:one });
const outParts = M.buildParts(one).join('\n');
chk('회신 표시가 켜져 있어도 텍스트에 안 나온다',
    !/shareDone|회신함/.test(outParts), '텍스트에 회신 정보가 섞임');
// 칸 수는 기능이 늘면 바뀐다. 중요한 건 "회신 표시가 여기 끼지 않는 것".
chk('공유 줄은 정해진 칸 수만 (지금은 11칸)', (() => {
  const line = outParts.split('\n').find(l => l.startsWith('S|'));
  return line && line.split('|').length === 12;   // 'S' + 11칸
})(), outParts.split('\n').find(l => l.startsWith('S|')));

/* ---- 이름만 치면 부서까지 채워지는가 ---- */
console.log('');
console.log('연락처 - 풀네임을 치면 부서를 알아서 채우는가');
M.setDB({ task:[], contact:[], share:[], lunch:[], cfg:{}, log:[],
  emp:[ {name:'김도현', dept:'기술지원팀', ext:'1274'},
        {name:'이지은', dept:'기획팀', ext:'1101'},
        {name:'이지은', dept:'총무팀', ext:'1102'} ] });

// autoFillDept 와 같은 판단: 이름만으로 한 명이면 부서를 준다
const look1 = M.empLookup('김도현', '');
chk('한 명뿐이면 부서를 알려준다', look1.state === 'ok' && look1.dept === '기술지원팀',
    JSON.stringify(look1));
const look2 = M.empLookup('이지은', '');
chk('동명이인이면 부서를 함부로 정하지 않는다',
    look2.state === 'ambiguous' && !look2.dept, JSON.stringify(look2));
const look3 = M.empLookup('없는사람', '');
chk('명부에 없으면 아무것도 안 준다', look3.state === 'none' && !look3.dept);
chk('이름이 비면 아무것도 안 준다', M.empLookup('', '').state === 'none');

/* ---- 공유에 새로 붙인 칸(회신방식·담당자) ---- */
console.log('');
console.log('공유 - 회신방식·담당자가 오가는가');
const withNew = [{
  id:'RT-1', docNo:'기획-2026-0501', title:'설비 교체 승인', content:'본문',
  regDate:'2026-08-27', dueDate:'2026-09-02', author:'홍길동',
  mtime:'2026-08-27 14:00:00', del:false, replyType:'메일', owner:'김철수'
}];
M.setDB({ emp:[], task:[], contact:[], lunch:[], log:[], cfg:{name:'홍길동'}, share:withNew });
const rtParts = M.buildParts(withNew);
M.setDB({ emp:[], task:[], contact:[], lunch:[], log:[], cfg:{}, share:[] });
M.resetInbox();
for(const p of rtParts) M.feed(p);
const rtBack = M.inboxRecords()['RT-1'];
chk('회신방식이 그대로 도착', rtBack && rtBack.replyType === '메일', JSON.stringify(rtBack));
chk('담당자가 그대로 도착', rtBack && rtBack.owner === '김철수', rtBack && rtBack.owner);
chk('기존 칸도 그대로', rtBack && rtBack.docNo === '기획-2026-0501' && rtBack.author === '홍길동');

// 이미 나눠준 옛 버전은 9칸만 보낸다. 그것도 읽혀야 한다.
console.log('');
console.log('옛 버전(9칸)이 보낸 줄도 읽히는가');
const oldBody = 'S|OLD-1|기획-2026-0001|옛 방식 건|본문|2026-08-01|2026-08-10|박과장|2026-08-01 09:00:00|0';
const sum = (s => { let x = 0; for(let i = 0; i < s.length; i++){ x = (x * 31 + s.charCodeAt(i)) % 65536; }
                    return x.toString(16).toUpperCase().padStart(4, '0'); })(oldBody);
const oldPart = '###WSCH1|P1/1|N=1###\r\n' + oldBody + '\r\n###END|P1/1|C=' + sum + '###';
M.setDB({ emp:[], task:[], contact:[], lunch:[], log:[], cfg:{}, share:[] });
M.resetInbox();
const st9 = M.feed(oldPart);
chk('옛 형식 조각을 받아들임', st9.st === M.FD.COMPLETE, st9.msg);
const oldBack = M.inboxRecords()['OLD-1'];
chk('9칸 줄이 정상 복원됨', oldBack && oldBack.title === '옛 방식 건', JSON.stringify(oldBack));
chk('없는 칸은 빈 값으로', oldBack && oldBack.replyType === '' && oldBack.owner === '',
    oldBack && JSON.stringify([oldBack.replyType, oldBack.owner]));

/* ---- 출장: 복수 선택 ---- */
console.log('');
console.log('출장. 이동수단·지역을 여러 개 고를 수 있는가');
const TRP = M.TRIP_TRANSPORT;
let sp = M.splitMulti(TRP, 'KTX, 자차');
chk('고른 것 두 개를 그대로 알아봄', sp.sel.length === 2 && sp.sel.includes('KTX') && sp.sel.includes('자차'), JSON.stringify(sp));
chk('직접 적은 것은 없음', sp.etc === '', sp.etc);
sp = M.splitMulti(TRP, 'KTX, 셔틀버스');
chk('목록에 없는 것은 직접입력으로 감', sp.sel.join() === 'KTX' && sp.etc === '셔틀버스', JSON.stringify(sp));
sp = M.splitMulti(TRP, '');
chk('빈 값이면 아무것도 안 고른 상태', sp.sel.length === 0 && sp.etc === '');
sp = M.splitMulti(TRP, undefined);
chk('값이 아예 없어도 터지지 않음', sp.sel.length === 0 && sp.etc === '');
chk('고른 것과 적은 것을 한 줄로 합침', M.joinMulti(['자차', 'KTX'], '셔틀') === '자차, KTX, 셔틀', M.joinMulti(['자차', 'KTX'], '셔틀'));
chk('직접입력이 비면 쉼표가 남지 않음', M.joinMulti(['자차'], '') === '자차', M.joinMulti(['자차'], ''));
chk('아무것도 안 고르면 빈 값', M.joinMulti([], '') === '');
const rtv = M.splitMulti(TRP, M.joinMulti(['KTX', '무궁화호'], '셔틀, 도보'));
chk('저장했다 다시 열어도 그대로', rtv.sel.join() === 'KTX,무궁화호' && rtv.etc === '셔틀, 도보', JSON.stringify(rtv));
chk('지역도 같은 규칙', M.splitMulti(M.TRIP_REGION, '우정사업본부(세종), 서울청').etc === '서울청');
chk('출장자는 10명까지', M.TRIP_MAX_MEMBER === 10);

/* ---- 출장도 찾기에 걸리는가 ---- */
console.log('');
console.log('출장. 찾기에 함께 걸리는가');
M.setDB({ cfg:{}, trip:[
  {id:'p1', title:'전산망 점검 회의', content:'서버 이관 협의', date:'2026-03-04',
   time:'14:00', transport:'KTX, 자차', region:'우정정보관리원(나주)',
   members:['홍길동','김철수'], del:false},
  {id:'p2', title:'지운 출장', content:'', date:'2026-03-04', del:true} ] });
chk('출장 제목으로 찾음', M.searchAll('전산망').length === 1);
chk('출장 내용으로도 찾음', M.searchAll('이관').length === 1);
chk('지운 출장은 안 나옴', M.searchAll('지운').length === 0);


const NL = String.fromCharCode(10);

/* ---- 메일 양식 ---- */
console.log('');
console.log('메일 쓰기. 정해진 양식대로 나오는가');
M.setDB({ cfg:{ name:'홍길동', appTitle:'AX금융혁신팀 업무스케줄러' } });
chk('제목에서 팀 이름만 뽑아낸다', M.teamName() === 'AX금융혁신팀', M.teamName());
M.setDB({ cfg:{ appTitle:'디지털혁신부 업무 스케줄러' } });
chk('띄어쓴 "업무 스케줄러" 도 떼어낸다', M.teamName() === '디지털혁신부', M.teamName());
M.setDB({ cfg:{ appTitle:'기획팀' } });
chk('뒤에 붙은 말이 없으면 제목 그대로', M.teamName() === '기획팀', M.teamName());
M.setDB({ cfg:{} });
chk('제목을 안 바꿨으면 기본 제목에서 뽑는다', M.teamName() === 'AX금융혁신팀', M.teamName());

M.setDB({ cfg:{ name:'홍길동', appTitle:'AX금융혁신팀 업무스케줄러' } });
const mail = M.mailText({ docNo:'기획-2026-0451', title:'상반기 실적 보고',
                          content:'첨부 참고 부탁드립니다.', author:'김철수' });
const ML = mail.split(NL);
chk('첫 줄 인사에 팀 이름과 작성자', ML[0] === '안녕하십니까 AX금융혁신팀 김철수 입니다.', ML[0]);
chk('제목 줄', ML[2] === '상반기 실적 보고관련하여 공유드립니다.', ML[2]);
chk('문서번호 줄', ML[4] === '관련한 내용은 아래와 같습니다. (관련문서 : 기획-2026-0451)', ML[4]);
chk('내용 줄', ML[6] === '첨부 참고 부탁드립니다.', ML[6]);
chk('맺음 줄', ML[ML.length - 1] === '확인 후 회신부탁드립니다.', ML[ML.length - 1]);
chk('문단 사이가 한 줄씩 비어 있다', ML[1] === '' && ML[3] === '' && ML[5] === '' && ML[7] === '');
const noAuthor = M.mailText({ docNo:'', title:'급한 건', content:'', author:'' });
chk('작성자가 비면 내 이름으로', noAuthor.split(NL)[0] === '안녕하십니까 AX금융혁신팀 홍길동 입니다.', noAuthor.split(NL)[0]);
chk('문서번호가 비면 - 로', noAuthor.indexOf('(관련문서 : -)') > 0);
chk('양식에 ### 같은 부호가 안 섞인다', mail.indexOf('###') < 0 && mail.indexOf('|') < 0);

/* ---- 출장 공유 글 ---- */
console.log('');
console.log('출장 공유. 붙여넣을 글이 읽을 만한가');
const tShare = M.tripShareText([
  { title:'전산망 점검 회의', date:'2026-03-04', time:'14:00',
    region:'우정정보관리원(나주)', transport:'KTX, 자차',
    members:['홍길동','김철수'], content:'서버 이관 협의' },
  { title:'교육 참석', date:'2026-03-05', time:'',
    region:'우정인재개발원(천안)', transport:'자차', members:[], content:'' } ]);
chk('맨 위에 팀 이름', tShare.split(NL)[0] === '[AX금융혁신팀] 출장 공유', tShare.split(NL)[0]);
chk('날짜에 요일이 붙는다', tShare.indexOf('▶ 2026-03-04 (수)') > 0);
chk('날짜별로 묶인다', tShare.indexOf('▶ 2026-03-05 (목)') > 0);
chk('복수 선택한 이동수단이 그대로', tShare.indexOf('- 이동수단: KTX, 자차') > 0);
chk('출장자 수를 같이 적는다', tShare.indexOf('- 출장자(2명): 홍길동, 김철수') > 0);
chk('회의 시간이 없으면 그 줄을 빼 준다', tShare.indexOf('교육 참석') > 0 && tShare.split('회의 시작').length === 2);
chk('출장자가 없으면 - 로', tShare.indexOf('- 출장자(0명): -') > 0);
chk('내용이 없으면 빈 줄을 안 만든다', tShare.indexOf('- 내용: ' + NL) < 0);
chk('양식 부호가 안 섞인다', tShare.indexOf('###') < 0);
const tMulti = M.tripShareText([{ title:'출장', date:'2026-03-04', region:'세종', transport:'자차',
  members:['홍길동'], content:'첫 줄' + NL + '둘째 줄' }]);
chk('내용이 여러 줄이면 줄마다 들여쓴다',
    tMulti.indexOf('- 내용: 첫 줄' + NL + '        둘째 줄') > 0, JSON.stringify(tMulti));


/* ---- 회신기한 동안 계속 노출 ---- */
console.log('');
console.log('업무공유. 회신기한이 남아 있는 동안 자리를 지키는가');
const VD = M.parseDate('2026-03-05');
const mk = (id, due) => ({ id, dueDate:due, regDate:'2026-03-01' });
M.setDB({ cfg:{ shareDone:{ done1:1, done2:1 } } });
chk('회신 안 한 건은 뒤로 안 감', M.sharePark(mk('open1', '2026-03-10'), VD) === false);
chk('회신했어도 기한이 남았으면 자리 지킴', M.sharePark(mk('done1', '2026-03-10'), VD) === false);
chk('기한 당일도 아직 자리 지킴', M.sharePark(mk('done1', '2026-03-05'), VD) === false);
chk('회신했고 기한도 지났으면 뒤로', M.sharePark(mk('done1', '2026-03-04'), VD) === true);
chk('회신했는데 기한이 없으면 뒤로', M.sharePark(mk('done2', ''), VD) === true);
chk('기한 지났어도 회신 안 했으면 그대로', M.sharePark(mk('open1', '2026-02-01'), VD) === false);

console.log('');
console.log('달력. 등록일부터 기한 전날까지 남는가');
const span = M.shareSpanDays({ regDate:'2026-03-02', dueDate:'2026-03-05' });
chk('등록일부터 기한 전날까지', span.join(' ') === '2026-03-02 2026-03-03 2026-03-04', span.join(' '));
chk('기한 당일은 빠진다 (기한 칩이 따로 있으므로)', span.indexOf('2026-03-05') < 0);
chk('등록일과 기한이 같으면 없음', M.shareSpanDays({ regDate:'2026-03-05', dueDate:'2026-03-05' }).length === 0);
chk('기한이 등록일보다 앞이면 없음', M.shareSpanDays({ regDate:'2026-03-05', dueDate:'2026-03-01' }).length === 0);
chk('날짜가 비면 없음', M.shareSpanDays({ regDate:'', dueDate:'2026-03-05' }).length === 0);
chk('기간이 터무니없이 길면 120일에서 끊는다',
    M.shareSpanDays({ regDate:'2026-01-01', dueDate:'2030-01-01' }).length === 120,
    M.shareSpanDays({ regDate:'2026-01-01', dueDate:'2030-01-01' }).length);


console.log('');
if(fail === 0){ console.log('===== 전부 통과 (' + pass + '건) ====='); process.exit(0); }
else { console.log('===== 성공 ' + pass + ' / 실패 ' + fail + ' ====='); process.exit(1); }
