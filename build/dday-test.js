/* =============================================================
 *  dday-test.js - 날짜 / D-day 계산 검증
 *
 *  실행:  node build\dday-test.js
 *
 *  web\스케줄러.html 의 날짜 유틸을 그대로 떼어내 돌린다.
 *
 *  배경: viewDate 가 현재 시각까지 들고 있으면 daysBetween 이 반올림되면서
 *        오후에 열었을 때 "오늘이 기한"인 건이 D-DAY 가 아니라 D+1(기한 지남)로
 *        표시됐다. 오전엔 정상, 오후엔 오답이라 눈치채기 어려운 종류였다.
 *        날짜 비교값을 전부 자정(startOfDay)으로 맞춰 고쳤고, 여기서 지킨다.
 * ============================================================= */
'use strict';
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const htmlPath = path.join(root, 'web', '스케줄러.html');

let pass = 0, fail = 0;
function chk(name, cond, detail){
  if(cond){ pass++; console.log('  [통과] ' + name); }
  else { fail++; console.log('  [실패] ' + name + '  -> ' + (detail === undefined ? '' : detail)); }
}

const src = fs.readFileSync(htmlPath, 'utf8');
const a = src.indexOf('const pad = (n, w)');
const b = src.indexOf('const LS_KEY');
if(a < 0 || b < 0){ console.log('  [실패] 앵커를 찾지 못했습니다'); process.exit(1); }

const modPath = path.join(__dirname, '.fixture', '_date.js');
fs.mkdirSync(path.dirname(modPath), { recursive: true });
fs.writeFileSync(modPath,
  src.slice(a, b) + '\nmodule.exports = { dateStr, todayStr, parseDate, addDays, daysBetween, startOfDay, normTime };\n',
  'utf8');

let M;
try{ M = require(modPath); }
catch(e){ console.log('  [실패] 떼어낸 코드가 실행되지 않습니다 -> ' + e.message); process.exit(1); }

console.log('');
console.log('===== 날짜 / D-day 검증 =====');
console.log('');

/* ---- 1. 하루 종일 같은 답이 나오는가 ---- */
console.log('1. 시각에 휘둘리지 않는가');

const due = M.parseDate('2026-08-27');
const afternoon = new Date(2026, 7, 27, 15, 30, 0);

console.log('   (참고) 시각을 그대로 쓰면: daysBetween = ' + M.daysBetween(due, afternoon) + '  -> 화면에 D+1');
console.log('   (참고) 자정으로 맞추면:    daysBetween = ' + M.daysBetween(due, M.startOfDay(afternoon)) + '  -> D-DAY');

chk('오후에도 오늘 기한은 D-DAY(0)', M.daysBetween(due, M.startOfDay(afternoon)) === 0);

let allSame = true, bad = [];
for(let h = 0; h < 24; h++){
  const t = M.startOfDay(new Date(2026, 7, 27, h, 45, 0));
  const dd = M.daysBetween(due, t);
  if(dd !== 0){ allSame = false; bad.push(h + '시 -> ' + dd); }
}
chk('하루 24개 시각 전부 같은 결과', allSame, bad.join(', '));

chk('내일이 기한이면 D-1', M.daysBetween(M.parseDate('2026-08-28'), M.startOfDay(afternoon)) === 1);
chk('어제가 기한이면 D+1', M.daysBetween(M.parseDate('2026-08-26'), M.startOfDay(afternoon)) === -1);
chk('3일 뒤가 기한이면 D-3', M.daysBetween(M.parseDate('2026-08-30'), M.startOfDay(afternoon)) === 3);

/* ---- 2. 날짜 이동 ---- */
console.log('');
console.log('2. 앞뒤 날짜 이동');

let v = M.startOfDay(new Date(2026, 7, 27, 23, 59, 59));
for(let i = 0; i < 40; i++) v = M.addDays(v, 1);
chk('40일 이동해도 자정 유지', v.getHours() === 0 && v.getMinutes() === 0 && v.getSeconds() === 0, v.toString());
chk('40일 뒤 날짜가 정확', M.dateStr(v) === '2026-10-06', M.dateStr(v));

let back = M.startOfDay(new Date(2026, 7, 27));
for(let i = 0; i < 100; i++) back = M.addDays(back, -1);
for(let i = 0; i < 100; i++) back = M.addDays(back, 1);
chk('100일 뒤로 갔다 돌아오면 제자리', M.dateStr(back) === '2026-08-27', M.dateStr(back));

chk('2월 말 -> 3월 1일', M.dateStr(M.addDays(M.startOfDay(new Date(2026, 1, 28)), 1)) === '2026-03-01');
chk('윤년 2월 28 -> 29', M.dateStr(M.addDays(M.startOfDay(new Date(2024, 1, 28)), 1)) === '2024-02-29');
chk('연말 -> 다음 해 1월 1일', M.dateStr(M.addDays(M.startOfDay(new Date(2026, 11, 31)), 1)) === '2027-01-01');

/* ---- 3. 문자열 변환 ---- */
console.log('');
console.log('3. 날짜·시간 문자열');

chk('dateStr 이 두 자리로 채움', M.dateStr(new Date(2026, 0, 5)) === '2026-01-05', M.dateStr(new Date(2026, 0, 5)));
chk('parseDate 왕복', M.dateStr(M.parseDate('2026-03-09')) === '2026-03-09');
chk('형식이 아니면 null', M.parseDate('2026/03/09') === null && M.parseDate('') === null && M.parseDate('abc') === null);
chk('parseDate 는 자정으로 돌려줌', (() => { const d = M.parseDate('2026-03-09'); return d.getHours() === 0 && d.getMinutes() === 0; })());

chk('normTime 09:05', M.normTime('9:5') === '09:05', M.normTime('9:5'));
chk('normTime 1430 -> 14:30', M.normTime('1430') === '14:30', M.normTime('1430'));
chk('normTime 9 -> 09:00', M.normTime('9') === '09:00', M.normTime('9'));
chk('normTime 빈 값은 빈 값', M.normTime('') === '' && M.normTime('   ') === '');
chk('normTime 잘못된 값은 null', M.normTime('25:00') === null && M.normTime('12:99') === null && M.normTime('아무말') === null);

console.log('');
if(fail === 0){ console.log('===== 전부 통과 (' + pass + '건) ====='); process.exit(0); }
else { console.log('===== 성공 ' + pass + ' / 실패 ' + fail + ' ====='); process.exit(1); }
