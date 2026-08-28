const pad = (n, w) => String(n).padStart(w || 2, '0');

function dateStr(d){ return d.getFullYear() + '-' + pad(d.getMonth()+1) + '-' + pad(d.getDate()); }
function todayStr(){ return dateStr(new Date()); }
function nowStamp(){
  const d = new Date();
  return dateStr(d) + ' ' + pad(d.getHours()) + ':' + pad(d.getMinutes()) + ':' + pad(d.getSeconds());
}
function parseDate(s){
  s = (s || '').trim();
  if(!/^\d{4}-\d{2}-\d{2}/.test(s)) return null;
  const d = new Date(+s.slice(0,4), +s.slice(5,7) - 1, +s.slice(8,10));
  return isNaN(d) ? null : d;
}
function addDays(d, n){ const x = new Date(d); x.setDate(x.getDate() + n); return x; }
function daysBetween(a, b){ return Math.round((a - b) / 86400000); }
// 날짜 비교에 쓰는 값은 반드시 자정으로 맞춘다.
// 시각이 붙어 있으면 daysBetween 이 반올림되면서 D-DAY 가 D+1 로 보이는 등 하루씩 틀어진다.
function startOfDay(d){ const x = new Date(d); x.setHours(0, 0, 0, 0); return x; }
const WD = ['일','월','화','수','목','금','토'];

function normTime(s){
  s = (s || '').trim();
  if(!s) return '';
  let m = s.match(/^(\d{1,2}):(\d{1,2})$/);
  if(!m) m = s.match(/^(\d{2})(\d{2})$/);
  if(!m && /^\d{1,2}$/.test(s)) m = [null, s, '0'];
  if(!m) return null;
  const h = +m[1], mi = +m[2];
  if(h < 0 || h > 23 || mi < 0 || mi > 59) return null;
  return pad(h) + ':' + pad(mi);
}

// ID: 작성자-yyyymmdd-hhnnss-4자리난수
// 자동증가 숫자는 PC마다 따로 증가해 반드시 충돌하므로 쓰지 않는다.
function newId(author){
  const d = new Date();
  const key = (author || '사용자').replace(/[|\\\t\r\n]/g, '') || '사용자';
  return key + '-' + d.getFullYear() + pad(d.getMonth()+1) + pad(d.getDate())
       + '-' + pad(d.getHours()) + pad(d.getMinutes()) + pad(d.getSeconds())
       + '-' + pad(Math.floor(Math.random() * 10000), 4);
}

/* --- 이스케이프: 반드시 이 순서 --- */
function esc(s){
  return String(s == null ? '' : s)
    .split('\\').join('\\\\')
    .split('|').join('\\p')
    .split('\r').join('')
    .split('\n').join('\\n')
    .split('\t').join('\\t');
}
// 단순 역순 치환은 "\\p" 같은 입력에서 깨진다. 한 문자씩 훑는다.
function unesc(s){
  s = String(s == null ? '' : s);
  let out = '', i = 0;
  while(i < s.length){
    const ch = s[i];
    if(ch === '\\' && i < s.length - 1){
      const nx = s[i+1];
      out += nx === '\\' ? '\\' : nx === 'p' ? '|' : nx === 'n' ? '\n' : nx === 't' ? '\t' : nx;
      i += 2;
    } else { out += ch; i++; }
  }
  return out;
}
// sum = (sum*31 + 문자코드) % 65536 → 4자리 대문자 16진수
function checksum4(s){
  let sum = 0;
  for(let i = 0; i < s.length; i++) sum = (sum * 31 + s.charCodeAt(i)) % 65536;
  return sum.toString(16).toUpperCase().padStart(4, '0');
}

function escHtml(s){
  return String(s == null ? '' : s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function ellipsis(s, n){
  s = String(s || '').replace(/[\r\n]+/g, ' ');
  return s.length <= n ? s : s.slice(0, n) + '…';
}

let toastTimer = null;
function toast(msg){
  const t = $('#toast');
  t.textContent = msg;
  t.classList.add('on');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove('on'), 2600);
}

async function copyText(text){
  try{
    if(navigator.clipboard && navigator.clipboard.writeText){
      await navigator.clipboard.writeText(text);
      return true;
    }
  }catch(e){ /* file:// 에서 막히면 아래로 */ }
  try{
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    const ok = document.execCommand('copy');
    document.body.removeChild(ta);
    return ok;
  }catch(e){ return false; }
}

function download(filename, content, mime){
  const blob = new Blob([content], {type: (mime || 'text/plain') + ';charset=utf-8'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}

/* ============================================================
   2. 데이터
   ============================================================ */

module.exports = { dateStr, todayStr, parseDate, addDays, daysBetween, startOfDay, normTime };
