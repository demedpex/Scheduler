/* =============================================================
 *  sort-test.js - 업무 정렬 기준 검증
 *
 *  실행:  node build\sort-test.js
 *
 *  dist\배포\sched.html 의 정렬 함수를 그대로 떼어내 돌린다.
 *
 *  배경: 순서를 바꾸는 길이 드래그뿐이라 건수가 늘면 손이 많이 갔다.
 *        기준으로 세우는 길을 더했는데, 여기서 지켜야 할 것이 둘 있다.
 *
 *        1) 어느 기준이든 완료한 건은 맨 뒤다. 남은 일이 먼저 보여야 한다.
 *        2) 기준은 보는 순서만 바꾼다. order 를 건드리면 안 된다 —
 *           마감일 순으로 훑어보다 [직접 정한 순서]로 돌아왔을 때
 *           손으로 맞춰 둔 순서가 사라지면 다시는 안 쓰게 된다.
 * ============================================================= */
'use strict';
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const htmlPath = path.join(root, 'dist', '배포', 'sched.html');
const src = fs.readFileSync(htmlPath, 'utf8');

let pass = 0, fail = 0;
function chk(name, got, want){
  if(got === want){ pass++; console.log('  [통과] ' + name + '  ->  ' + got); }
  else { fail++; console.log('  [실패] ' + name + '  ->  ' + got + '   (기대 ' + want + ')'); }
}

/* --- 필요한 조각만 떼어낸다 --- */
function block(head, openCh, closeCh){
  const i = src.indexOf(head);
  if(i < 0){ console.log('  [실패] 앵커를 찾지 못했습니다: ' + head); process.exit(1); }
  let d = 0;
  for(let k = src.indexOf(openCh, i); k < src.length; k++){
    if(src[k] === openCh) d++;
    else if(src[k] === closeCh){ d--; if(!d) return src.slice(i, k + 1); }
  }
  console.log('  [실패] 닫는 괄호를 찾지 못했습니다: ' + head);
  process.exit(1);
}
const fn = n => block('function ' + n + '(', '{', '}');

const code = [
  fn('taskSortKey'), fn('taskTimeKey'), fn('taskDueKey'), fn('taskAddedKey'), fn('cmpStr'),
  block('const TASK_SORTS = [', '[', ']') + ';',
  block('const TASK_CMP = {', '{', '}') + ';',
  fn('taskSortMode'), fn('sortTasksBy'), fn('sortTasks'),
].join('\n');

const DB = { cfg: {} };
const M = new Function('DB', code +
  '; return {sortTasks, sortTasksBy, taskSortMode, TASK_SORTS, taskDueKey, taskAddedKey};')(DB);

/* --- 자료: 기준마다 답이 달라지도록 섞어 둔다 --- */
const T = (stamp, title, date, until, time, done) =>
  ({ id: '서영석-' + stamp + '-0000', title: title, date: date, until: until, time: time, done: !!done });

const list = [
  T('20260901-090000', '나 보고서 마무리', '2026-09-11', '',           '15:00'),
  T('20260902-090000', '가 예산안 검토',   '2026-09-08', '2026-09-30', ''),
  T('20260903-090000', '다 회의 준비',     '2026-09-11', '',           '09:30'),
  T('20260904-090000', '라 감사 대응',     '2026-09-01', '2026-09-12', '13:00'),
  T('20260905-090000', '마 끝낸 일',       '2026-09-11', '',           '08:00', true),
];
const order = l => l.map(r => r.title.slice(0, 1)).join(' ');
const by = m => { DB.cfg.taskSort = m; return order(M.sortTasks(list)); };

console.log('1. 기준마다 제대로 세우는가  (완료한 "마" 는 어느 기준에서도 맨 뒤)');
chk('마감일 빠른 순', by('due'),   '다 나 라 가 마');
chk('시간 빠른 순',   by('time'),  '다 라 나 가 마');
chk('등록한 순',      by('added'), '나 가 다 라 마');
chk('제목 가나다순',  by('title'), '가 나 다 라 마');
chk('직접 정한 순서 — order 가 없으면 시간 순', by('manual'), '다 라 나 가 마');

console.log('');
console.log('2. 키가 뜻대로 나오는가');
chk('하루짜리는 그날이 곧 마감', M.taskDueKey(list[0]), '2026-09-11');
chk('기간 업무는 마감일자',      M.taskDueKey(list[1]), '2026-09-30');
chk('날짜가 아예 없으면 맨 뒤',  M.taskDueKey({}),      '9999-99-99');
chk('등록 시각은 id 에서 꺼낸다', M.taskAddedKey(list[0]), '20260901090000');
chk('형식이 다른 옛 id 는 맨 뒤', M.taskAddedKey({ id: 't1' }), '99999999999999');

console.log('');
console.log('3. 기준은 보는 순서만 바꾼다');
DB.cfg.taskSort = 'title';
M.sortTasks(list);
chk('세워 봐도 order 가 생기지 않는다',
    list.filter(r => r.order !== undefined).length, 0);

list.forEach((r, i) => { r.order = list.length - 1 - i; });   // 손으로 맞춰 둔 순서
DB.cfg.taskSort = 'due';   M.sortTasks(list);
DB.cfg.taskSort = 'title'; M.sortTasks(list);
chk('기준을 오간 뒤에도 직접 순서는 그대로', by('manual'), '라 다 가 나 마');

console.log('');
console.log('4. 값이 이상해도 안 깨진다');
DB.cfg.taskSort = '아무말';
chk('모르는 기준은 직접 순서로 돌린다', M.taskSortMode(), 'manual');
delete DB.cfg.taskSort;
chk('기준을 고른 적이 없으면 직접 순서', M.taskSortMode(), 'manual');
chk('고를 수 있는 기준은 ' + M.TASK_SORTS.length + '가지',
    M.TASK_SORTS.map(x => x[0]).join(','), 'manual,due,time,added,title');

console.log('');
if(fail === 0){ console.log('===== 전부 통과 (' + pass + '건) ====='); process.exit(0); }
else { console.log('===== 성공 ' + pass + ' / 실패 ' + fail + ' ====='); process.exit(1); }
