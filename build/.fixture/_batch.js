let DB = { share: [], cfg: {}, log: [] };
const alive = l => l.filter(r => !r.del);
const myName = () => DB.cfg.name || "사용자";
const markDirty = () => {};
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

const PROTO = 'WSCH1';
const PART_MAX_DEFAULT = 800;      // 사내 메신저 한 번에 붙여넣기 가능한 글자 수
// 메신저가 한 번에 받아 주는 글자 수. 설정에서 바꾼다.
// 이 값이 클수록 조각 수가 줄어든다 (압축보다 훨씬 큰 차이를 만든다).
function partMax(){
  const v = +(DB.cfg.partMax || PART_MAX_DEFAULT);
  return (isFinite(v) && v >= 200) ? Math.floor(v) : PART_MAX_DEFAULT;
}
const FD = { NOHEADER:0, BADSUM:1, PARTIAL:2, COMPLETE:3, DUPLICATE:4 };

function shareToLine(r){
  return ['S', r.id, r.docNo, r.title, r.content, r.regDate, r.dueDate,
          r.author, r.mtime, r.del ? '1' : '0',
          r.replyType || '', r.owner || ''].map((v, i) => i === 0 ? v : esc(v)).join('|');
}
function lineToShare(ln){
  const f = ln.split('|');
  if(f.length < 10 || f[0] !== 'S') return null;
  return {
    id: unesc(f[1]), docNo: unesc(f[2]), title: unesc(f[3]), content: unesc(f[4]),
    regDate: unesc(f[5]), dueDate: unesc(f[6]), author: unesc(f[7]),
    mtime: unesc(f[8]), del: unesc(f[9]) === '1',
    // 아래 두 칸은 나중에 생겼다. 옛 버전이 보낸 줄에는 없으므로 빈 값으로 둔다.
    replyType: f.length > 10 ? unesc(f[10]) : '',
    owner: f.length > 11 ? unesc(f[11]) : ''
  };
}

/* ------------------------------------------------------------
   배치 ID
   같은 방에서 두 사람이 조각을 보내면 섞인다. 배치 ID가 없으면
   나중에 들어온 것이 앞의 것을 덮어쓰거나 조용히 버려진다.
   내보낼 때 한 번 만들어 그 배치의 모든 조각에 똑같이 붙인다.
   ------------------------------------------------------------ */
function newBatchId(){
  const d = new Date();
  return SanitizeBatch(myName()) + '_' +
         d.getFullYear() + pad(d.getMonth()+1) + pad(d.getDate()) +
         pad(d.getHours()) + pad(d.getMinutes()) + pad(d.getSeconds());
}
function SanitizeBatch(s){
  return String(s || '사용자').replace(/[|=#\s\\]/g, '') || '사용자';
}
function senderOf(batchId){
  const i = String(batchId).lastIndexOf('_');
  return i > 0 ? batchId.slice(0, i) : batchId;
}

// 마지막으로 만든 조각들의 정보 (너무 긴 레코드 경고용)
let lastBuildInfo = { oversize: [], batchId: '', bodyMax: 0 };

// 레코드 중간에서 자르지 않는다. 한 건이 최대치를 넘으면 그 건만 단독 파트.
function buildParts(records){
  const lines = records.map(shareToLine);
  lastBuildInfo = { oversize: [], batchId: '', bodyMax: 0 };
  if(!lines.length) return [];

  const bid = newBatchId();
  // 머리말·꼬리말도 붙여넣는 글자 수에 포함된다. 그만큼 본문을 줄여야
  // 사용자가 설정한 값이 실제 한 조각의 길이와 맞는다.
  const probeH = '###' + PROTO + '|B=' + bid + '|P99/99|N=999###';
  const probeF = '###END|B=' + bid + '|P99/99|C=FFFF###';
  const overhead = probeH.length + probeF.length + 4;      // 줄바꿈 2개
  const max = Math.max(120, partMax() - overhead);
  lastBuildInfo.batchId = bid;
  lastBuildInfo.bodyMax = max;

  const parts = []; let cur = []; let len = 0;
  for(let i = 0; i < lines.length; i++){
    const ln = lines[i];
    if(ln.length > max){
      // 한 줄이 통째로 최대치를 넘는다 = 이 건은 조각 하나에 안 들어간다
      lastBuildInfo.oversize.push({ title: records[i].title, docNo: records[i].docNo, len: ln.length });
    }
    const add = cur.length === 0 ? ln.length : ln.length + 1;
    if(cur.length > 0 && len + add > max){ parts.push(cur); cur = [ln]; len = ln.length; }
    else { cur.push(ln); len += add; }
  }
  if(cur.length) parts.push(cur);

  return parts.map((body, i) => {
    const b = body.join('\n');
    const tag = 'B=' + bid + '|P' + (i+1) + '/' + parts.length;
    return '###' + PROTO + '|' + tag + '|N=' + body.length + '###\r\n'
         + b.split('\n').join('\r\n') + '\r\n'
         + '###END|' + tag + '|C=' + checksum4(b) + '###';
  });
}

function scopeFilter(scope, days){
  const all = DB.share;
  const today = new Date(); today.setHours(0,0,0,0);
  // 지난번 내보낸 뒤로 바뀐 것만. 평소에는 이게 1~2건이라 조각도 1개로 끝난다.
  if(scope === 'changed'){
    const since = DB.cfg.lastExportAt || '';
    return all.filter(r => (r.mtime || '') > since);
  }
  if(scope === 'mine')   return all.filter(r => (r.author || '') === myName());
  if(scope === 'recent') return all.filter(r => { const d = parseDate(r.regDate); return d && daysBetween(d, today) >= -days; });
  if(scope === 'open')   return all.filter(r => { const d = parseDate(r.dueDate); return d && d >= today && !r.del; });
  return all;
}

/* ------------------------------------------------------------
   수신 버퍼
   inbox[배치ID] = { sender, total, parts:{파트번호: 본문} }
   여러 사람이 보낸 배치를 동시에 담아 둘 수 있다.
   ------------------------------------------------------------ */
let inbox = {};
function resetInbox(){ inbox = {}; }

function batchList(){
  return Object.keys(inbox).map(id => {
    const b = inbox[id];
    const got = Object.keys(b.parts).length;
    const miss = [];
    for(let i = 1; i <= b.total; i++) if(!b.parts[i]) miss.push(i);
    return { id, sender:b.sender, total:b.total, got, miss, complete: got >= b.total };
  }).sort((a, b) => a.id < b.id ? -1 : 1);
}
function completeBatches(){ return batchList().filter(b => b.complete); }
function inboxCount(){ return batchList().reduce((n, b) => n + b.got, 0); }

function parsePartNo(s){
  if(s[0] !== 'P') return null;
  const p = s.indexOf('/');
  if(p < 0) return null;
  const a = +s.slice(1, p), b = +s.slice(p + 1);
  if(!Number.isInteger(a) || !Number.isInteger(b) || a < 1 || b < 1 || a > b) return null;
  return {no:a, tot:b};
}
// 머리말: ###WSCH1|B=배치ID|P1/3|N=5###   (B= 는 옛 형식엔 없다)
function parseHeader(ln){
  if(!ln.endsWith('###')) return null;
  const f = ln.slice(3, -3).split('|');
  if(f.length < 3 || f[0] !== PROTO) return null;
  let k = 1, bid = '';
  if(f[k] && f[k].startsWith('B=')){ bid = f[k].slice(2); k++; }
  const pn = parsePartNo(f[k] || ''); k++;
  if(!pn) return null;
  const nf = f[k] || '';
  if(!nf.startsWith('N=')) return null;
  const n = +nf.slice(2);
  if(!Number.isInteger(n)) return null;
  return {bid, no:pn.no, tot:pn.tot, n};
}
function parseFooter(ln){
  if(!ln.endsWith('###')) return null;
  const f = ln.slice(3, -3).split('|');
  if(f.length < 3 || f[0] !== 'END') return null;
  let k = 1, bid = '';
  if(f[k] && f[k].startsWith('B=')){ bid = f[k].slice(2); k++; }
  const pn = parsePartNo(f[k] || ''); k++;
  if(!pn) return null;
  const cf = f[k] || '';
  if(!cf.startsWith('C=')) return null;
  const sum = cf.slice(2);
  if(sum.length !== 4) return null;
  return {bid, no:pn.no, tot:pn.tot, sum};
}

// 붙여넣은 텍스트를 먹인다. 조각 순서 무관, 같은 조각 재투입 안전(멱등).
// 서로 다른 사람이 보낸 배치가 섞여 들어와도 각자 따로 모인다.
function feed(text){
  const lines = String(text).replace(/\r\n/g,'\n').replace(/\r/g,'\n').split('\n');
  let inPart = false, body = [], hdr = null;
  let gotAny = false, gotNew = false, badSum = false, truncated = false;
  let touched = null;

  for(const raw of lines){
    const ln = raw.trim();
    if(!ln) continue;

    if(ln.startsWith('###' + PROTO + '|')){
      if(inPart) truncated = true;
      hdr = parseHeader(ln);
      if(hdr){ inPart = true; body = []; gotAny = true; }
      else inPart = false;
    }
    else if(ln.startsWith('###END|')){
      if(inPart && hdr){
        const ft = parseFooter(ln);
        if(!ft || ft.no !== hdr.no || ft.tot !== hdr.tot || (ft.bid || '') !== (hdr.bid || '')){
          badSum = true;
        }else{
          const b = body.join('\n');
          if(checksum4(b) !== ft.sum || body.length !== hdr.n){
            badSum = true;
          }else{
            // 옛 형식(B= 없음)은 조각 수로 묶는다. 예전 동작 그대로.
            const bid = hdr.bid || ('(이전형식)_' + hdr.tot);
            if(!inbox[bid]) inbox[bid] = { sender: senderOf(bid), total: hdr.tot, parts: {} };
            if(inbox[bid].total !== hdr.tot) inbox[bid].total = hdr.tot;
            if(!inbox[bid].parts[hdr.no]){ inbox[bid].parts[hdr.no] = b; gotNew = true; }
            touched = bid;
          }
        }
        inPart = false;
      }
    }
    else if(inPart) body.push(ln);
  }
  if(inPart) truncated = true;

  if(!gotAny) return {st:FD.NOHEADER, msg:'인식할 수 없는 텍스트예요. 머리말(###' + PROTO + ')까지 통째로 복사했는지 확인해 주세요.'};
  if(truncated || badSum) return {st:FD.BADSUM, msg:'글이 잘리거나 바뀌었어요. 보낸 분께 다시 부탁드리세요.'};

  const b = touched ? batchList().find(x => x.id === touched) : null;
  const who = b ? (b.sender + '님이 보낸 ') : '';
  const waiting = batchList().filter(x => !x.complete);

  if(b && b.complete){
    const done = completeBatches();
    const extra = waiting.length ? ' (' + waiting.length + '개 묶음은 아직 조각이 모자라요)' : '';
    return {st:FD.COMPLETE,
      msg: who + '조각 ' + b.total + '개를 모두 받았어요.' +
           (done.length > 1 ? ' 반영할 묶음 ' + done.length + '개.' : '') + extra};
  }
  if(!gotNew){
    return {st:FD.DUPLICATE,
      msg: who + '조각은 이미 받았어요. ' + (b ? b.got + '/' + b.total + ' 받은 상태 — 남은 조각: ' + b.miss.join(', ') : '')};
  }
  return {st:FD.PARTIAL,
    msg: who + '조각 ' + b.got + '/' + b.total + ' 까지 받았어요. 남은 조각: ' + b.miss.join(', ') +
         ' — 지금까지 받은 것만 먼저 [반영] 해도 됩니다.'};
}

/* 받은 조각에서 읽을 수 있는 건 전부 모은다. 같은 ID 는 수정시각이 최신인 것만.

   전에는 배치가 다 차야만(조각 2개면 2개 다 와야) 반영됐다. 그런데 메신저로
   주고받다 보면 한 조각이 늦거나 빠지는 일이 흔하고, 그동안 받은 것마저
   못 넣는 것은 손해다. buildParts 가 레코드를 조각 사이에서 자르지 않으므로,
   조각 하나만 있어도 그 안의 건들은 온전하다. 그래서 받은 만큼 바로 반영한다.
   나중에 남은 조각이 와서 다시 반영해도 수정시각으로 겨루므로 안전하다. */
function inboxRecords(){
  const d = {};
  for(const b of batchList()){
    const parts = inbox[b.id].parts;
    for(let i = 1; i <= b.total; i++){
      if(!parts[i]) continue;
      for(const ln of parts[i].split('\n')){
        if(!ln.trim()) continue;
        const r = lineToShare(ln);
        if(!r || !r.id) continue;
        if(!d[r.id] || r.mtime >= d[r.id].mtime) d[r.id] = r;
      }
    }
  }
  return d;
}

/* 병합 판정 — 삭제도 하나의 상태로 보고 수정시각으로 겨룬다.
     · 들어온 것이 더 최신  -> 그대로 채택 (지웠던 건이면 되살아남)
     · 시각이 같거나 과거   -> 내 것을 지킨다. 단 삭제는 같은 시각에서 우선.
   이렇게 하면 남이 옛날 데이터를 통째로 내보내도 내가 지운 게 되살아나지 않고,
   반대로 누가 지운 뒤 다시 손본 건은 제대로 돌아온다. */
function mergeVerdict(r){
  const cur = DB.share.find(x => x.id === r.id);
  if(!cur) return r.del ? '삭제' : '신규';
  if(r.mtime > cur.mtime){
    if(r.del) return '삭제';
    return cur.del ? '되살아남' : '갱신';
  }
  if(r.del && !cur.del) return '삭제';
  return '변화없음';
}

function syncPreview(){
  const d = inboxRecords();
  return Object.keys(d).map(k => ({rec:d[k], verdict:mergeVerdict(d[k])}));
}

function syncApply(){
  const d = inboxRecords();
  let nAdd = 0, nUpd = 0, nSame = 0, nDel = 0, nBack = 0;
  const dupDocNo = new Set();
  const docIndex = {};
  alive(DB.share).forEach(r => { if(r.docNo && !docIndex[r.docNo]) docIndex[r.docNo] = r.id; });

  Object.keys(d).forEach(id => {
    const r = d[id];
    const i = DB.share.findIndex(x => x.id === id);
    if(i < 0){
      DB.share.push(r);
      if(r.del){ nDel++; addLog('삭제표식수신', id, '', shareSummary(r), r.author); }
      else { nAdd++; addLog('신규', id, '', shareSummary(r), r.author); }
    }else{
      const cur = DB.share[i];
      if(r.mtime > cur.mtime){
        if(r.del){
          addLog('삭제', id, shareSummary(cur), '', r.author);
          cur.del = true; cur.mtime = r.mtime; nDel++;
        }else if(cur.del){
          // 내가 지운 뒤에 상대가 더 최신으로 손본 건 -> 되살린다
          addLog('되살림', id, '(삭제된 상태)', shareSummary(r), r.author);
          DB.share[i] = r; nBack++;
        }else{
          addLog('갱신', id, shareSummary(cur), shareSummary(r), r.author);
          DB.share[i] = r; nUpd++;
        }
      }
      else if(r.del && !cur.del){
        // 같은 시각이면 삭제가 이긴다
        addLog('삭제', id, shareSummary(cur), '', r.author);
        cur.del = true; nDel++;
      }
      else nSame++;
    }
    if(r.docNo){
      if(docIndex[r.docNo] && docIndex[r.docNo] !== id) dupDocNo.add(r.docNo);
      else docIndex[r.docNo] = id;
    }
  });

  DB.cfg.lastSync = nowStamp();
  markDirty();

  // 다 받은 배치는 비운다. 조각이 모자란 배치는 남은 조각을 더 받을 수 있게 그대로 둔다.
  // (이미 반영한 건이 다시 들어와도 수정시각으로 겨루므로 '변화없음' 이 된다)
  completeBatches().forEach(b => { delete inbox[b.id]; });

  let msg = '신규 ' + nAdd + '건, 갱신 ' + nUpd + '건, 변화없음 ' + nSame + '건';
  if(nBack) msg += ', 되살림 ' + nBack + '건';
  if(nDel) msg += ', 삭제 ' + nDel + '건';
  msg += ' 반영했어요.';
  const left = batchList();
  if(left.length) msg += '\n\n받은 조각까지는 반영했어요. 아직 조각이 모자란 묶음이 ' + left.length + '개 있어요: ' +
    left.map(b => b.sender + ' (' + b.got + '/' + b.total + ', 남은 조각 ' + b.miss.join('·') + ')').join(', ') +
    '\n남은 조각을 받으면 그때 다시 붙여넣으면 됩니다.';
  if(dupDocNo.size){
    msg += '\n\n※ 문서번호가 중복된 건이 있어요: ' + Array.from(dupDocNo).join(', ') +
           '\n   (병합은 그대로 했습니다. 어느 쪽이 맞는지 확인해 주세요.)';
  }
  return msg;
}

/* 내가 회신했는지 — 사람마다 다른 값이라 로컬에만 둔다.
   김대리는 회신했고 박과장은 안 했을 수 있다.
   shareToLine 에는 절대 들어가지 않는다. */
function myReplied(id){ return !!(DB.cfg.shareDone && DB.cfg.shareDone[id]); }
function setReplied(id, on){
  if(!DB.cfg.shareDone) DB.cfg.shareDone = {};
  if(on) DB.cfg.shareDone[id] = 1; else delete DB.cfg.shareDone[id];
  markDirty();
}

/* [닫기] 로 이월 배너에서 내린 지난 미완료.
   업무 자체는 건드리지 않는다 — done 도 그대로, 원래 날짜에 미완료로 남는다.
   달력·검색·그날 목록에는 계속 보이고 배너에서만 빠진다.
   그래서 회신 표시(myReplied)처럼 업무 레코드가 아니라 cfg 에 둔다. */
function carryOff(id){ return !!(DB.cfg.carryOff && DB.cfg.carryOff[id]); }
function setCarryOff(id, on){
  if(!DB.cfg.carryOff) DB.cfg.carryOff = {};
  if(on) DB.cfg.carryOff[id] = 1; else delete DB.cfg.carryOff[id];
}

function shareSummary(r){ return (r.docNo || '') + ' / ' + (r.title || '') + ' / 기한 ' + (r.dueDate || ''); }
function addLog(action, id, before, after, src){
  DB.log.unshift({time:nowStamp(), action, id, before:ellipsis(before,200), after:ellipsis(after,200), src:src||''});
  if(DB.log.length > 500) DB.log.length = 500;
}
function docNoOwner(docNo, exceptId){
  docNo = (docNo || '').trim();
  if(!docNo) return '';
  const hit = alive(DB.share).find(r => r.docNo === docNo && r.id !== exceptId);
  return hit ? (hit.author || '(작성자 없음)') : '';
}

/* ============================================================
   5. 모달
   ============================================================ */
/* 날짜를 비워 두는 게 기본인 폼(연락·점심)에서 쓰는 빠른 날짜 버튼.
   비워두기가 기본이면 매번 달력을 여는 게 번거로우니 한 번에 넣게 한다. */
function quickDateBtns(){
  return '<div class="quickdate">' +
    '<button type="button" class="btn sm" data-qd="0">오늘</button>' +
    '<button type="button" class="btn sm" data-qd="1">내일</button>' +
    '<button type="button" class="btn sm" data-qd="clear">비우기</button>' +
  '</div>';
}
function bindQuickDate(){
  $$('[data-qd]').forEach(b => b.onclick = () => {
    const v = b.dataset.qd;
    $('#f_date').value = (v === 'clear') ? '' : dateStr(addDays(startOfDay(new Date()), +v));
  });
}

/* 한 줄짜리 입력에서 Enter -> 저장.
   내용(textarea)에서는 Ctrl+Enter. 하루에 수십 번 하는 동작이라 기본에 가깝다. */
function bindEnterSave(okId){
  const ok = $('#' + (okId || 'f_ok'));
  if(!ok) return;
  // 창을 열 때마다 #modalRoot 에 붙이면 리스너가 쌓인다.
  // 지금 열린 창 안에만 붙여서 창이 닫히면 함께 사라지게 한다.
  const box = ok.closest('.modal') || $('#modalRoot');
  box.addEventListener('keydown', e => {
    if(e.key !== 'Enter') return;
    const t = e.target;
    if(t.tagName === 'TEXTAREA'){
      if(e.ctrlKey || e.metaKey){ e.preventDefault(); ok.click(); }
      return;
    }
    if(t.tagName === 'INPUT' && t.type !== 'file'){ e.preventDefault(); ok.click(); }
  });
}


module.exports = { buildParts, feed, inboxRecords, syncPreview, syncApply, resetInbox, FD, setDB: d => { DB = d; }, getDB: () => DB };
