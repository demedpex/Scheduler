const document={addEventListener(){},createElement:()=>({style:{},appendChild(){},remove(){}}),body:{appendChild(){}}};
const window={getSelection:()=>({removeAllRanges(){},addRange(){}})};
const $=()=>null; const $$=()=>[];
let DB={task:[],contact:[],share:[],lunch:[],emp:[],fav:[],cfg:{},log:[]};
const alive=l=>l.filter(r=>!r.del);
const myName=()=>DB.cfg.name||"사용자";
const markDirty=()=>{};
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

function empLookup(empName, dept){
  const name = (empName || '').trim();
  const d = (dept || '').trim();
  if(!name) return { state:'none', ext:'' };

  if(d){
    const exact = DB.emp.find(e => (e.name || '').trim() === name && (e.dept || '').trim() === d);
    if(exact) return { state:'ok', ext: exact.ext || '' };
  }
  const byName = DB.emp.filter(e => (e.name || '').trim() === name);
  if(byName.length === 1) return { state:'ok', ext: byName[0].ext || '', dept: byName[0].dept || '' };
  if(byName.length > 1)   return { state:'ambiguous', ext:'', count: byName.length, list: byName };
  return { state:'none', ext:'' };
}

function empExt(empName, dept){
  const r = empLookup(empName, dept);
  return r.state === 'ok' ? r.ext : '';
}

// 머리글 이름 → 우리가 쓰는 4개 항목

const EMP_HEAD = {
  '이름':0,'성명':0,'사원명':0,
  '내선번호':1,'내선':1,'전화번호':1,'연락처':1,'사내번호':1,'전화':1,
  '부서':2,'부서명':2,'소속':2,'팀':2,
  '직급':3,'직위':3
};

// 마지막으로 명부를 읽은 결과. 왜 몇 명만 읽혔는지 사용자에게 보여 주려고 남긴다.
let empParseInfo = { total:0, blank:0, noName:0, header:false, cols:'' };

// 행 배열(문자열 2차원) → 명부 레코드
// 머리글이 있으면 이름을 보고 열 위치를 잡고, 없으면 순서대로 본다.
//
// 중간에 빈 줄이 있어도 멈추지 않고 건너뛴다.
// (부서 구분용 빈 줄 하나 때문에 명부가 조용히 잘리는 사고가 있었다)
function rowsToEmp(rows){
  let map = [0,1,2,3], start = 0, header = false;
  const first = rows[0] || [];
  if(first.some(c => EMP_HEAD[String(c).trim()] !== undefined)){
    const m = [-1,-1,-1,-1];
    first.forEach((c, i) => {
      const k = EMP_HEAD[String(c).trim()];
      if(k !== undefined && m[k] < 0) m[k] = i;
    });
    map = m.map((v, i) => v < 0 ? i : v);
    start = 1;
    header = true;
  }

  const out = [];
  let blank = 0, noName = 0;
  for(let i = start; i < rows.length; i++){
    const f = rows[i] || [];
    if(f.every(c => !String(c == null ? '' : c).trim())){ blank++; continue; }
    const g = k => String(f[map[k]] == null ? '' : f[map[k]]).trim();
    const rec = { name:g(0), ext:g(1), dept:g(2), rank:g(3) };
    if(rec.name) out.push(rec); else noName++;
  }

  empParseInfo = {
    total: Math.max(0, rows.length - start),
    blank: blank, noName: noName, header: header,
    cols: (first || []).map(c => String(c == null ? '' : c).trim()).filter(Boolean).join(' | ')
  };
  return out;
}

function csvToRows(text){
  const lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  const head = lines[0] || '';
  const sep = head.split('\t').length > head.split(',').length ? '\t' : ',';
  return lines.map(l => splitCsvLine(l, sep));
}
function splitCsvLine(line, sep){
  const out = []; let cur = '', q = false;
  for(let i = 0; i < line.length; i++){
    const ch = line[i];
    if(q){
      if(ch === '"'){ if(line[i+1] === '"'){ cur += '"'; i++; } else q = false; }
      else cur += ch;
    }else{
      if(ch === '"') q = true;
      else if(ch === sep){ out.push(cur.trim()); cur = ''; }
      else cur += ch;
    }
  }
  out.push(cur.trim());
  return out;
}
// BOM 이 있으면 UTF-8, 없으면 CP949(euc-kr) 로 읽는다.
// 0212 같은 앞자리 0 은 문자열로 다루므로 그대로 보존된다.
function decodeCsv(buf){
  const b = new Uint8Array(buf);
  if(b.length >= 3 && b[0] === 0xEF && b[1] === 0xBB && b[2] === 0xBF){
    return new TextDecoder('utf-8').decode(b.subarray(3));
  }
  try{
    const t = new TextDecoder('euc-kr', {fatal:false}).decode(b);
    // euc-kr 로 읽었는데 대체문자가 많으면 utf-8 로 재시도
    const bad = (t.match(/�/g) || []).length;
    if(bad > t.length * 0.02) return new TextDecoder('utf-8').decode(b);
    return t;
  }catch(e){
    return new TextDecoder('utf-8').decode(b);
  }
}

/* ------------------------------------------------------------
   엑셀(.xlsx) 직접 읽기
   .xlsx 는 XML 몇 개를 담은 ZIP 이다. 외부 라이브러리 없이
   ZIP 을 직접 풀고(브라우저 내장 DecompressionStream) XML 을 읽는다.
   ------------------------------------------------------------ */
function zipEntries(buf){
  const u8 = new Uint8Array(buf);
  const dv = new DataView(buf);
  // 끝에서부터 중앙 디렉터리 끝(EOCD) 표식을 찾는다
  let eocd = -1;
  for(let i = u8.length - 22; i >= 0 && i >= u8.length - 22 - 65536; i--){
    if(dv.getUint32(i, true) === 0x06054b50){ eocd = i; break; }
  }
  if(eocd < 0) throw new Error('ZIP 형식이 아닙니다');
  const count = dv.getUint16(eocd + 10, true);
  let p = dv.getUint32(eocd + 16, true);

  const out = new Map();
  const dec = new TextDecoder('utf-8');
  for(let i = 0; i < count; i++){
    if(dv.getUint32(p, true) !== 0x02014b50) break;
    const method   = dv.getUint16(p + 10, true);
    const compSize = dv.getUint32(p + 20, true);
    const nameLen  = dv.getUint16(p + 28, true);
    const extraLen = dv.getUint16(p + 30, true);
    const cmtLen   = dv.getUint16(p + 32, true);
    const localOff = dv.getUint32(p + 42, true);
    const name = dec.decode(u8.subarray(p + 46, p + 46 + nameLen));

    // 로컬 헤더에서 실제 데이터 시작 위치를 다시 계산한다
    const lNameLen  = dv.getUint16(localOff + 26, true);
    const lExtraLen = dv.getUint16(localOff + 28, true);
    const dataOff = localOff + 30 + lNameLen + lExtraLen;

    out.set(name, { method, data: u8.subarray(dataOff, dataOff + compSize) });
    p += 46 + nameLen + extraLen + cmtLen;
  }
  return out;
}

async function unzipText(entry){
  if(!entry) return '';
  if(entry.method === 0) return new TextDecoder('utf-8').decode(entry.data);
  if(entry.method !== 8) throw new Error('지원하지 않는 압축 방식입니다 (' + entry.method + ')');
  if(typeof DecompressionStream === 'undefined'){
    throw new Error('이 브라우저는 xlsx 를 못 읽어요. CSV 로 저장해 주세요.');
  }
  const ds = new DecompressionStream('deflate-raw');
  const stream = new Blob([entry.data]).stream().pipeThrough(ds);
  return await new Response(stream).text();
}

function colIndex(ref){                       // "AB12" -> 27
  let n = 0;
  for(let i = 0; i < ref.length; i++){
    const c = ref.charCodeAt(i);
    if(c < 65 || c > 90) break;
    n = n * 26 + (c - 64);
  }
  return n - 1;
}

// 태그를 이름만 보고 찾는다.
// 엑셀은 <row> 로 쓰지만 한셀 등은 <x:row> 처럼 접두사를 붙인다.
// getElementsByTagName 은 접두사까지 맞아야 해서 한쪽을 통째로 놓친다.
function tags(node, name){
  return Array.from(node.getElementsByTagNameNS('*', name));
}

async function xlsxToRows(buf){
  const zip = zipEntries(buf);

  // 첫 번째 시트가 실제로 어느 파일인지 관계 파일에서 찾는다
  let sheetPath = 'xl/worksheets/sheet1.xml';
  try{
    const wbXml = await unzipText(zip.get('xl/workbook.xml'));
    const relXml = await unzipText(zip.get('xl/_rels/workbook.xml.rels'));
    const wb = new DOMParser().parseFromString(wbXml, 'application/xml');
    const rel = new DOMParser().parseFromString(relXml, 'application/xml');
    const s0 = tags(wb, 'sheet')[0];
    const rid = s0 && (s0.getAttribute('r:id') || s0.getAttribute('id'));
    if(rid){
      for(const r of tags(rel, 'Relationship')){
        if(r.getAttribute('Id') === rid){
          let t = r.getAttribute('Target') || '';
          if(t.startsWith('/')) t = t.slice(1);
          else if(!t.startsWith('xl/')) t = 'xl/' + t;
          if(zip.has(t)) sheetPath = t;
          break;
        }
      }
    }
  }catch(e){ /* 못 찾으면 sheet1.xml 로 간다 */ }

  // 공유 문자열 표
  let shared = [];
  if(zip.has('xl/sharedStrings.xml')){
    const doc = new DOMParser().parseFromString(await unzipText(zip.get('xl/sharedStrings.xml')), 'application/xml');
    shared = tags(doc, 'si').map(si => tags(si, 't').map(t => t.textContent).join(''));
  }

  const sheetXml = await unzipText(zip.get(sheetPath));
  if(!sheetXml) throw new Error('시트를 찾지 못했습니다');
  const doc = new DOMParser().parseFromString(sheetXml, 'application/xml');

  const rows = [];
  for(const row of tags(doc, 'row')){
    const arr = [];
    for(const c of tags(row, 'c')){
      const ci = colIndex(c.getAttribute('r') || '');
      const t = c.getAttribute('t');
      let v = '';
      if(t === 's'){
        const iEl = tags(c, 'v')[0];
        v = iEl ? (shared[+iEl.textContent] || '') : '';
      }else if(t === 'inlineStr'){
        v = tags(c, 't').map(x => x.textContent).join('');
      }else{
        const vEl = tags(c, 'v')[0];
        v = vEl ? vEl.textContent : '';
      }
      if(ci >= 0) arr[ci] = v;
    }
    for(let i = 0; i < arr.length; i++) if(arr[i] == null) arr[i] = '';
    rows.push(arr);
  }
  return rows;
}

/* ============================================================
   4. 동기화 프로토콜 — 엑셀 버전 modSync.bas 와 완전히 동일
      ###WSCH1|P1/3|N=5###
      S|ID|문서번호|제목|내용|등록일|회신기한|작성자|수정시각|삭제여부
      ###END|P1/3|C=A3F1###
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

function closeModal(){ $('#modalRoot').innerHTML = ''; }
// 맨 위 창 하나만 닫는다. 동명이인 선택창이 떠 있을 때 Esc 로
// 뒤의 등록창이 사라지던 문제를 막는다.
function closeTopModal(){
  const root = $('#modalRoot');
  if(root.lastElementChild) root.lastElementChild.remove();
}
function openModal(html, opts){
  opts = opts || {};
  const root = $('#modalRoot');
  root.innerHTML = '';
  const wrap = document.createElement('div');
  wrap.innerHTML = '<div class="backdrop"><div class="modal ' + (opts.wide ? 'wide' : '') + '">' + html + '</div></div>';
  root.appendChild(wrap);
  const bd = $('.backdrop', wrap);
  bd.addEventListener('mousedown', e => { if(e.target === bd && !opts.sticky) closeTopModal(); });
  const f = $('[data-focus]', root);
  if(f) setTimeout(() => f.focus(), 30);
  return root;
}
// Esc 는 항상 맨 위 창 하나만 닫는다
document.addEventListener('keydown', e => {
  if(e.key === 'Escape') closeTopModal();
});
function confirmBox(msg){ return window.confirm(msg); }

/* --- 업무 --- */
function taskModal(id){
  const r = id ? DB.task.find(x => x.id === id) : null;
  openModal(
    '<h3>' + (r ? '업무 수정' : '업무 등록') + '</h3>' +
    '<div class="body">' +
      '<div id="err"></div>' +
      '<div class="frow"><label>제목</label><input type="text" id="f_title" data-focus value="' + escHtml(r ? r.title : '') + '"></div>' +
      '<div class="frow two">' +
        '<div><label>날짜</label><input type="date" id="f_date" value="' + escHtml(r ? r.date : dateStr(viewDate)) + '"></div>' +
        '<div><label>시간 (선택)</label><input type="time" id="f_time" value="' + escHtml(r ? r.time : '') + '"></div>' +
      '</div>' +
      '<div class="frow"><label>마감일자 (선택)</label>' +
        '<input type="date" id="f_until" value="' + escHtml(r ? (r.until || '') : '') + '">' +
        '<div class="hint">마감일자를 정하면 <b>그날까지 매일</b> 목록에 보여요. ' +
        '하루만 하는 일이면 비워 두세요.</div></div>' +
      '<div class="frow"><label>내용</label><textarea id="f_content">' + escHtml(r ? r.content : '') + '</textarea></div>' +
    '</div>' +
    '<div class="foot">' +
      (r ? '<button class="btn danger" id="f_del">삭제</button>' : '') +
      '<span class="spacer"></span>' +
      '<button class="btn" id="f_cancel">취소</button>' +
      '<button class="btn primary" id="f_ok">저장</button>' +
    '</div>', {wide:true});

  bindEnterSave();
  $('#f_cancel').onclick = closeModal;
  if(r) $('#f_del').onclick = () => { if(softDelete('task', r.id)) { closeModal(); renderAll(); } };
  $('#f_ok').onclick = () => {
    const title = $('#f_title').value.trim();
    const date = $('#f_date').value;
    const time = normTime($('#f_time').value);
    const until = $('#f_until').value;
    if(!title) return showErr('제목을 입력해 주세요.');
    if(!date)  return showErr('날짜를 입력해 주세요.');
    if(time === null) return showErr('시간 형식이 올바르지 않습니다.');
    if(until && until < date) return showErr('마감일자는 시작 날짜보다 뒤여야 해요.');
    const rec = r || {id:newId(myName()), author:myName(), done:false, doneAt:'', del:false};
    Object.assign(rec, {title, date, time, until, content:$('#f_content').value, mtime:nowStamp()});
    if(!r){
      const ord = nextOrderFor(date);     // 순서를 매긴 날이면 맨 뒤로
      if(ord !== undefined) rec.order = ord;
      DB.task.push(rec);
    }
    markDirty(); closeModal(); renderAll();
    toast(r ? '수정했습니다.' : '등록했습니다.');
  };
}

/* --- 연락하기 --- */
function contactModal(id){
  const r = id ? DB.contact.find(x => x.id === id) : null;
  openModal(
    '<h3>' + (r ? '연락하기 수정' : '연락하기 등록') + '</h3>' +
    '<div class="body">' +
      '<div id="err"></div>' +
      (DB.emp.length ? '' :
        '<div class="notebox">' +
          '<span>직원 명부가 아직 없어서 <b>이름으로 내선번호를 찾을 수 없어요.</b><br>' +
          '내선번호를 직접 입력하시거나, 명부를 등록해 주세요.</span>' +
          '<span class="spacer"></span>' +
          '<button class="btn sm primary" id="f_regemp">명부 등록하기</button>' +
        '</div>') +
      '<div class="frow two">' +
        '<div><label>이름</label><input type="text" id="f_name" data-focus value="' + escHtml(r ? r.name : '') + '" placeholder="일부만 입력해도 찾습니다"></div>' +
        '<div><label>부서</label><input type="text" id="f_dept" value="' + escHtml(r ? r.dept : '') + '"></div>' +
      '</div>' +
      '<div class="frow" id="f_favwrap" hidden><label>자주 쓰는 연락처</label>' +
        '<div class="favlist" id="f_favs"></div></div>' +
      '<div class="frow"><button class="btn sm" id="f_find">명부에서 찾기</button> <span id="f_extinfo" class="hint" style="display:inline"></span></div>' +
      '<div class="frow"><label>내선 직접입력 (명부에 없을 때만)</label>' +
        '<div style="display:flex;gap:8px;align-items:center">' +
          '<input type="text" id="f_ext" value="' + escHtml(r ? r.extCache : '') + '">' +
          '<button class="btn sm" id="f_favadd" style="white-space:nowrap">★ 자주 쓰는 연락처로</button>' +
        '</div></div>' +
      '<div class="frow two">' +
        '<div><label>날짜 (선택)</label>' +
          '<input type="date" id="f_date" value="' + escHtml(r ? r.date : '') + '">' +
          quickDateBtns() +
        '</div>' +
        '<div><label>시간 (선택)</label><input type="time" id="f_time" value="' + escHtml(r ? r.time : '') + '"></div>' +
      '</div>' +
      '<div class="hint">날짜를 비워 두면 <b>날짜 미정</b>으로 오늘 화면에 계속 보여요.</div>' +
    '</div>' +
    '<div class="foot">' +
      (r ? '<button class="btn danger" id="f_del">삭제</button>' : '') +
      '<span class="spacer"></span>' +
      '<button class="btn" id="f_cancel">취소</button>' +
      '<button class="btn primary" id="f_ok">저장</button>' +
    '</div>');

  bindQuickDate();
  const regBtn = $('#f_regemp');
  if(regBtn) regBtn.onclick = () => { closeModal(); pickEmpFile(); };

  /* --- 자주 쓰는 연락처 ---
     명부에 없거나 명부가 아직 없어도, 자주 거는 곳은 직접 모아 쓸 수 있게 한다. */
  function drawFavs(){
    if(!DB.fav.length){ $('#f_favwrap').hidden = true; return; }
    $('#f_favwrap').hidden = false;
    $('#f_favs').innerHTML = DB.fav.map((v, i) =>
      '<span class="favchip" data-i="' + i + '">' +
        '<b>' + escHtml(v.name) + '</b>' +
        (v.dept ? '<em>' + escHtml(v.dept) + '</em>' : '') +
        '<em>' + escHtml(v.ext || '-') + '</em>' +
        '<span class="x" data-del="' + i + '" title="지우기">×</span>' +
      '</span>').join('');
  }

  function bindFavs(){
    drawFavs();
    $('#f_favs').addEventListener('click', ev => {
      const del = ev.target.closest('[data-del]');
      if(del){
        ev.stopPropagation();
        DB.fav.splice(+del.dataset.del, 1);
        markDirty(); drawFavs();
        return;
      }
      const chip = ev.target.closest('.favchip');
      if(!chip) return;
      const v = DB.fav[+chip.dataset.i];
      $('#f_name').value = v.name;
      $('#f_dept').value = v.dept || '';
      $('#f_ext').value = v.ext || '';
      refreshExt();
    });
    $('#f_favadd').onclick = () => {
      const name = $('#f_name').value.trim();
      if(!name) return showErr('이름을 먼저 입력해 주세요.');
      const dept = $('#f_dept').value.trim();
      const ext = $('#f_ext').value.trim() || empExt(name, dept);
      if(!ext) return showErr('내선번호를 입력해 주세요. 자주 쓰는 연락처에는 번호가 필요해요.');
      const dup = DB.fav.findIndex(v => v.name === name && (v.dept || '') === dept);
      if(dup >= 0) DB.fav[dup] = { name, dept, ext };
      else DB.fav.push({ name, dept, ext });
      markDirty(); drawFavs();
      toast(name + '님을 자주 쓰는 연락처에 담았어요.');
    };
  }

  /* 이름을 다 치면 부서까지 알아서 채운다.
     명부에 그 이름이 한 명뿐일 때만. 동명이인이면 함부로 정하지 않는다.
     사용자가 부서를 직접 적어 뒀으면 건드리지 않는다. */
  function autoFillDept(){
    const nameV = $('#f_name').value.trim();
    const deptV = $('#f_dept').value.trim();
    if(!nameV || deptV) return;
    const look = empLookup(nameV, '');
    if(look.state === 'ok' && look.dept) $('#f_dept').value = look.dept;
  }

  function refreshExt(){
    const ext = empExt($('#f_name').value, $('#f_dept').value);
    const box = $('#f_extinfo');
    if(ext){ box.textContent = '명부: 내선 ' + ext; box.className = 'hint ok'; $('#f_ext').disabled = true; }
    else { box.textContent = '명부에 없음 — 아래에 직접 입력'; box.className = 'hint'; $('#f_ext').disabled = false; }
  }
  // 부서 자동 채움은 이름 칸에서만. 부서 칸에서 하면 지우는 즉시 다시 채워진다.
  $('#f_name').oninput = () => { autoFillDept(); refreshExt(); };
  $('#f_dept').oninput = refreshExt;
  autoFillDept();
  refreshExt();
  bindFavs();

  // 0건 → 직접 입력 / 1건 → 팝업 없이 자동 세팅 / 2건 이상 → 선택 팝업
  $('#f_find').onclick = () => {
    const q = $('#f_name').value.trim();
    if(!q) return showErr('이름을 먼저 입력해 주세요. 일부만 입력해도 찾습니다.');
    const hits = empSearch(q);
    if(!hits.length) return showErr("'" + q + "' 은(는) 명부에 없습니다. 내선번호를 직접 입력해 주세요.");
    if(hits.length === 1){ applyPick(hits[0]); return; }
    pickerModal(hits, applyPick);
  };
  function applyPick(e){
    $('#f_name').value = e.name;
    $('#f_dept').value = e.dept || '';
    $('#f_ext').value = '';
    refreshExt();
    toast(e.name + (e.dept ? ' / ' + e.dept : '') + ' / 내선 ' + (e.ext || '-'));
  }

  $('#f_cancel').onclick = closeModal;
  if(r) $('#f_del').onclick = () => { if(softDelete('contact', r.id)) { closeModal(); renderAll(); } };
  $('#f_ok').onclick = () => {
    const name = $('#f_name').value.trim();
    const dept = $('#f_dept').value.trim();
    const time = normTime($('#f_time').value);
    if(!name) return showErr('이름을 입력해 주세요.');
    if(time === null) return showErr('시간 형식이 올바르지 않습니다.');
    let extCache = $('#f_ext').value.trim();
    if(empExt(name, dept)) extCache = '';       // 명부가 항상 우선
    const rec = r || {id:newId(myName()), author:myName(), done:false, del:false};
    Object.assign(rec, {name, dept, extCache, date:$('#f_date').value, time, mtime:nowStamp()});
    if(!r) DB.contact.push(rec);
    markDirty(); closeModal(); renderAll();
    toast(r ? '수정했습니다.' : '등록했습니다.');
  };
}

/* --- 동명이인 선택 --- */
function pickerModal(list, cb){
  const root = document.createElement('div');
  root.innerHTML =
    '<div class="backdrop" style="z-index:110"><div class="modal">' +
    '<h3>같은 이름이 ' + list.length + '명 있습니다</h3>' +
    '<div class="body"><div class="picklist">' +
      list.map((e, i) =>
        '<div class="p" data-i="' + i + '"><b>' + escHtml(e.name) + '</b>' +
        '<span>' + escHtml(e.dept || '-') + '</span>' +
        '<span>' + escHtml(e.rank || '-') + '</span>' +
        '<span style="margin-left:auto">내선 ' + escHtml(e.ext || '-') + '</span></div>'
      ).join('') +
    '</div></div>' +
    '<div class="foot"><span class="spacer"></span><button class="btn" data-cancel>취소</button></div>' +
    '</div></div>';
  $('#modalRoot').appendChild(root);
  const kill = () => root.remove();
  $('[data-cancel]', root).onclick = kill;
  $$('.p', root).forEach(el => el.onclick = () => { kill(); cb(list[+el.dataset.i]); });
}

/* --- 업무공유 --- */
function shareModal(id){
  const r = id ? DB.share.find(x => x.id === id) : null;
  openModal(
    '<h3>' + (r ? '업무공유 수정' : '업무공유 등록') + '</h3>' +
    '<div class="body">' +
      '<div id="err"></div>' +
      '<div class="frow two">' +
        '<div><label>문서번호</label><input type="text" id="f_doc" data-focus value="' + escHtml(r ? r.docNo : '') + '"></div>' +
        '<div><label>&nbsp;</label><div class="hint" id="f_docwarn"></div></div>' +
      '</div>' +
      '<div class="frow"><label>제목</label><input type="text" id="f_title" value="' + escHtml(r ? r.title : '') + '"></div>' +
      '<div class="frow two">' +
        '<div><label>회신 방식</label>' +
          '<div class="segbtns" id="f_rtype">' +
            '<button type="button" data-rt="문서">문서회신</button>' +
            '<button type="button" data-rt="메일">메일회신</button>' +
          '</div></div>' +
        '<div><label>담당자</label>' +
          '<input type="text" id="f_owner" value="' + escHtml(r ? (r.owner || '') : '') + '" placeholder="회신할 사람"></div>' +
      '</div>' +
      '<div class="frow two">' +
        '<div><label>등록일</label><input type="date" id="f_reg" value="' + escHtml(r ? r.regDate : todayStr()) + '"></div>' +
        '<div><label>회신기한</label><input type="date" id="f_due" value="' + escHtml(r ? r.dueDate : dateStr(addDays(new Date(), 3))) + '"></div>' +
      '</div>' +
      '<div class="frow"><label>내용</label><textarea id="f_content">' + escHtml(r ? r.content : '') + '</textarea></div>' +
      '<div class="hint">이 항목만 팀원과 공유됩니다.</div>' +
    '</div>' +
    '<div class="foot">' +
      (r ? '<button class="btn danger" id="f_del">삭제</button>' : '') +
      '<span class="spacer"></span>' +
      '<button class="btn" id="f_cancel">취소</button>' +
      '<button class="btn primary" id="f_ok">저장</button>' +
    '</div>');

  // 회신 방식 — 문서로 회신하는지 메일로 회신하는지
  let rtype = (r && r.replyType) || '문서';
  function paintRt(){
    $$('#f_rtype [data-rt]').forEach(b => b.classList.toggle('on', b.dataset.rt === rtype));
  }
  $('#f_rtype').addEventListener('click', e => {
    const b = e.target.closest('[data-rt]');
    if(!b) return;
    rtype = b.dataset.rt;
    paintRt();
  });
  paintRt();

  // 문서번호 중복은 막지 않고 미리 알려만 준다
  $('#f_doc').oninput = () => {
    const owner = docNoOwner($('#f_doc').value, r ? r.id : '');
    const w = $('#f_docwarn');
    w.textContent = owner ? '이미 등록된 문서번호입니다 (등록자: ' + owner + ')' : '';
    w.className = owner ? 'hint warn' : 'hint';
  };
  $('#f_doc').oninput();

  bindEnterSave();
  $('#f_cancel').onclick = closeModal;
  if(r) $('#f_del').onclick = () => { if(softDelete('share', r.id)) { closeModal(); renderAll(); } };
  $('#f_ok').onclick = () => {
    const docNo = $('#f_doc').value.trim();
    const title = $('#f_title').value.trim();
    const due = $('#f_due').value;
    if(!docNo) return showErr('문서번호를 입력해 주세요.');
    if(!title) return showErr('제목을 입력해 주세요.');
    if(!due)   return showErr('회신기한을 입력해 주세요. (필수 항목입니다)');
    const owner = docNoOwner(docNo, r ? r.id : '');
    if(owner && !confirmBox('문서번호 ' + docNo + ' 은(는) 이미 등록돼 있습니다.\n(등록자: ' + owner + ')\n\n그래도 등록할까요?')) return;
    const rec = r || {id:newId(myName()), author:myName(), del:false};
    Object.assign(rec, {docNo, title, content:$('#f_content').value,
                        regDate:$('#f_reg').value || todayStr(), dueDate:due,
                        replyType:rtype, owner:$('#f_owner').value.trim(), mtime:nowStamp()});
    if(!r) DB.share.push(rec);
    markDirty(); closeModal(); renderAll();
    toast(r ? '수정했습니다.' : '등록했습니다.');
  };
}

/* --- 점심약속 --- */
function lunchModal(id){
  const r = id ? DB.lunch.find(x => x.id === id) : null;
  openModal(
    '<h3>' + (r ? '점심약속 수정' : '점심약속 등록') + '</h3>' +
    '<div class="body">' +
      '<div id="err"></div>' +
      '<div class="frow"><label>대상 이름</label><input type="text" id="f_target" data-focus value="' + escHtml(r ? r.target : '') + '"></div>' +
      '<div class="frow two">' +
        '<div><label>날짜</label><input type="date" id="f_date" value="' + escHtml(r ? r.date : dateStr(viewDate)) + '">' + quickDateBtns() + '</div>' +
        '<div><label>시간 (선택)</label><input type="time" id="f_time" value="' + escHtml(r ? r.time : '') + '"></div>' +
      '</div>' +
      '<div class="frow"><label>장소</label><input type="text" id="f_place" value="' + escHtml(r ? r.place : '') + '"></div>' +
      '<div class="hint" id="f_same"></div>' +
    '</div>' +
    '<div class="foot">' +
      (r ? '<button class="btn danger" id="f_del">삭제</button>' : '') +
      '<span class="spacer"></span>' +
      '<button class="btn" id="f_cancel">취소</button>' +
      '<button class="btn primary" id="f_ok">저장</button>' +
    '</div>');

  bindQuickDate();

  function sameDay(){
    const d = $('#f_date').value;
    const other = alive(DB.lunch).find(x => x.date === d && x.id !== (r ? r.id : ''));
    const box = $('#f_same');
    box.textContent = other ? (d.slice(5).replace('-','/') + '에 ' + other.target + '님과 약속이 있습니다.') : '';
    box.className = other ? 'hint warn' : 'hint';
  }
  $('#f_date').onchange = sameDay;
  sameDay();

  $('#f_cancel').onclick = closeModal;
  if(r) $('#f_del').onclick = () => { if(softDelete('lunch', r.id)) { closeModal(); renderAll(); } };
  $('#f_ok').onclick = () => {
    const target = $('#f_target').value.trim();
    const date = $('#f_date').value;
    const time = normTime($('#f_time').value);
    if(!target) return showErr('약속 대상 이름을 입력해 주세요.');
    if(!date)   return showErr('날짜를 입력해 주세요.');
    if(time === null) return showErr('시간 형식이 올바르지 않습니다.');

    const myId = r ? r.id : '';
    const dup = alive(DB.lunch).find(x => x.date === date && x.target === target && x.id !== myId);
    if(dup){
      if(!confirmBox(date.slice(5).replace('-','/') + '에 ' + target + '님과의 약속이 이미 등록돼 있습니다.\n중복 등록할까요?')) return;
    }else{
      const other = alive(DB.lunch).find(x => x.date === date && x.id !== myId);
      if(other && !confirmBox(date.slice(5).replace('-','/') + '에 ' + other.target + '님과 약속이 있습니다.\n그래도 등록할까요?')) return;
    }
    const rec = r || {id:newId(myName()), author:myName(), del:false};
    Object.assign(rec, {target, date, time, place:$('#f_place').value.trim(), mtime:nowStamp()});
    if(!r) DB.lunch.push(rec);
    markDirty(); closeModal(); renderAll();
    toast(r ? '수정했습니다.' : '등록했습니다.');
  };
}

/* ------------------------------------------------------------
   출장
   이동수단·출장지역은 자주 쓰는 것을 버튼으로 두되,
   목록에 없는 경우가 반드시 있으므로 직접입력을 항상 남겨 둔다.
   갈 때 KTX, 올 때 자차처럼 여러 개를 고를 수 있어야 해서 복수 선택이다.
   ------------------------------------------------------------ */
const TRIP_TRANSPORT = ['자차', 'KTX', '새마을', '무궁화호'];
const TRIP_REGION = ['우정사업본부(세종)', '우정정보관리원(나주)', '우정인재개발원(천안)'];
const TRIP_MAX_MEMBER = 10;

// 저장값("KTX, 자차, 셔틀") 을 고른 것 / 직접 적은 것으로 가른다
function splitMulti(list, value){
  const parts = String(value || '').split(',').map(s => s.trim()).filter(Boolean);
  const sel = parts.filter(v => list.includes(v));
  const etc = parts.filter(v => !list.includes(v)).join(', ');
  return { sel, etc };
}
// 고른 것 + 직접 적은 것을 한 문자열로
function joinMulti(sel, etc){
  return sel.concat(String(etc || '').split(',').map(s => s.trim()).filter(Boolean)).join(', ');
}

function tripModal(id){
  const r = id ? DB.trip.find(x => x.id === id) : null;
  const tr = splitMulti(TRIP_TRANSPORT, r && r.transport);
  const rg = splitMulti(TRIP_REGION, r && r.region);
  let members = (r && Array.isArray(r.members)) ? r.members.slice() : [];

  // 여러 개 고를 수 있는 버튼 묶음
  const segs = (name, list, sel, etc) =>
    '<div class="segbtns multi" data-seg="' + name + '">' +
      list.map(v =>
        '<button type="button" data-v="' + escHtml(v) + '"' +
        (sel.includes(v) ? ' class="on"' : '') + '>' + escHtml(v) + '</button>').join('') +
      '<button type="button" data-etc="1"' + (etc ? ' class="on"' : '') + '>직접입력</button>' +
    '</div>';

  openModal(
    '<h3>' + (r ? '출장 수정' : '출장 등록') + '</h3>' +
    '<div class="body">' +
      '<div id="err"></div>' +
      '<div class="frow"><label>제목</label>' +
        '<input type="text" id="f_title" data-focus value="' + escHtml(r ? r.title : '') + '"></div>' +

      '<div class="frow two">' +
        '<div><label>날짜</label><input type="date" id="f_date" value="' +
          escHtml(r ? r.date : dateStr(viewDate)) + '">' + quickDateBtns() + '</div>' +
        '<div><label>회의 시작 시간</label><input type="time" id="f_time" value="' +
          escHtml(r ? r.time : '') + '"></div>' +
      '</div>' +

      '<div class="frow"><label>이동수단 <span class="lhint">여러 개 고를 수 있어요</span></label>' +
        segs('tr', TRIP_TRANSPORT, tr.sel, tr.etc) +
        '<input type="text" id="f_tretc" class="segetc" placeholder="예: 셔틀버스, 시외버스" value="' +
          escHtml(tr.etc) + '"' + (tr.etc ? '' : ' hidden') + '></div>' +

      '<div class="frow"><label>출장 지역 <span class="lhint">여러 곳도 됩니다</span></label>' +
        segs('rg', TRIP_REGION, rg.sel, rg.etc) +
        '<input type="text" id="f_rgetc" class="segetc" placeholder="가는 곳을 적어 주세요" value="' +
          escHtml(rg.etc) + '"' + (rg.etc ? '' : ' hidden') + '></div>' +

      '<div class="frow"><label>출장자 <span id="f_mcount" class="mcount"></span></label>' +
        '<div class="memadd">' +
          '<input type="text" id="f_mname" list="empNames" placeholder="이름을 적고 Enter">' +
          '<button type="button" class="btn sm" id="f_madd">추가</button>' +
        '</div>' +
        '<div class="memlist" id="f_members"></div>' +
        '<div class="hint">명부에 있는 분은 이름을 치면 목록이 떠요. 최대 ' + TRIP_MAX_MEMBER + '명.</div>' +
      '</div>' +

      '<div class="frow"><label>내용 (무슨 출장인지)</label>' +
        '<textarea id="f_content">' + escHtml(r ? r.content : '') + '</textarea></div>' +
    '</div>' +
    '<div class="foot">' +
      (r ? '<button class="btn danger" id="f_del">삭제</button>' : '') +
      '<span class="spacer"></span>' +
      '<button class="btn" id="f_cancel">취소</button>' +
      '<button class="btn primary" id="f_ok">저장</button>' +
    '</div>', {wide:true});

  bindQuickDate();
  bindEnterSave();

  // 버튼을 눌러 켜고 끈다. 직접입력은 글칸을 여닫는다.
  $$('[data-seg]').forEach(box => {
    box.addEventListener('click', e => {
      const b = e.target.closest('button');
      if(!b) return;
      const etcBox = $('#f_' + box.dataset.seg + 'etc');
      if(b.dataset.etc){
        b.classList.toggle('on');
        etcBox.hidden = !b.classList.contains('on');
        if(!etcBox.hidden) etcBox.focus(); else etcBox.value = '';
      }else{
        b.classList.toggle('on');
      }
    });
  });

  function drawMembers(){
    $('#f_members').innerHTML = members.map((v, i) =>
      '<span class="memchip"><b>' + escHtml(v) + '</b>' +
        '<span class="x" data-mi="' + i + '" title="빼기">×</span></span>').join('');
    $('#f_mcount').textContent = members.length + ' / ' + TRIP_MAX_MEMBER;
    $('#f_madd').disabled = members.length >= TRIP_MAX_MEMBER;
    $('#f_mname').disabled = members.length >= TRIP_MAX_MEMBER;
  }
  function addMember(){
    const v = $('#f_mname').value.trim();
    if(!v) return;
    if(members.length >= TRIP_MAX_MEMBER) return showErr('출장자는 ' + TRIP_MAX_MEMBER + '명까지예요.');
    if(members.includes(v)){ $('#f_mname').value = ''; return showErr(v + '님은 이미 넣으셨어요.'); }
    members.push(v);
    $('#f_mname').value = '';
    drawMembers();
    $('#f_mname').focus();
  }
  $('#f_madd').onclick = addMember;
  $('#f_mname').addEventListener('keydown', e => {
    // 여기서 Enter 는 저장이 아니라 "한 명 추가" 다
    if(e.key === 'Enter'){ e.preventDefault(); e.stopPropagation(); addMember(); }
  });
  $('#f_members').addEventListener('click', e => {
    const x = e.target.closest('[data-mi]');
    if(!x) return;
    members.splice(+x.dataset.mi, 1);
    drawMembers();
  });
  drawMembers();

  function segValue(name){
    const box = $('[data-seg="' + name + '"]');
    const sel = $$('button.on[data-v]', box).map(b => b.dataset.v);
    return joinMulti(sel, $('#f_' + name + 'etc').value);
  }

  $('#f_cancel').onclick = closeModal;
  if(r) $('#f_del').onclick = () => { if(softDelete('trip', r.id)) { closeModal(); renderAll(); } };
  $('#f_ok').onclick = () => {
    const title = $('#f_title').value.trim();
    const date = $('#f_date').value;
    const time = normTime($('#f_time').value);
    if(!title) return showErr('제목을 입력해 주세요.');
    if(!date)  return showErr('날짜를 입력해 주세요.');
    if(time === null) return showErr('회의 시작 시간이 올바르지 않아요.');
    const transport = segValue('tr');
    const region = segValue('rg');
    if(!transport) return showErr('이동수단을 하나 이상 골라 주세요.');
    if(!region) return showErr('출장 지역을 하나 이상 골라 주세요.');

    const rec = r || { id:newId(myName()), author:myName(), del:false };
    Object.assign(rec, {
      title, date, time, transport, region,
      members: members.slice(),
      content: $('#f_content').value, mtime: nowStamp()
    });
    if(!r) DB.trip.push(rec);
    markDirty(); closeModal(); renderAll();
    toast(r ? '수정했어요.' : '등록했어요.');
  };
}

function showErr(msg){
  const box = $('#err');
  if(box) box.innerHTML = '<div class="errbox">' + escHtml(msg) + '</div>';
  return false;
}
function softDelete(kind, id){
  if(!confirmBox('이 항목을 삭제할까요?\n(공유 건은 다음 내보내기 때 상대방에게도 삭제로 전달됩니다.)')) return false;
  const r = DB[kind].find(x => x.id === id);
  if(!r) return false;
  r.del = true; r.mtime = nowStamp();
  markDirty();
  return true;
}

/* --- 내용 보기 --- */
/* ------------------------------------------------------------
   달력 칸을 누르면 먼저 보여 주는 하루 미리보기
   전에는 바로 그 날짜 화면으로 넘어갔다. 그런데 달력에서 하는 일은
   대부분 "그날 뭐 있었지?" 하고 확인하는 것이라, 넘어갔다가 다시
   달력으로 돌아오는 일이 반복됐다. 그래서 먼저 펼쳐 보여 주고,
   정말 그날로 갈 때만 [이동하기] 를 누르게 한다.
   ------------------------------------------------------------ */
function dayPeekModal(k){
  const sd = parseDate(k);
  const when = sd ? (sd.getMonth() + 1) + '월 ' + sd.getDate() + '일 (' + WD[sd.getDay()] + ')' : k;
  const isToday = k === todayStr();

  const tasks = sortByTime(alive(DB.task).filter(r => r.date === k));
  const trips = sortByTime(alive(DB.trip).filter(r => r.date === k));
  const calls = sortByTime(alive(DB.contact).filter(r => r.date === k));
  const meals = sortByTime(alive(DB.lunch).filter(r => r.date === k));
  const dueOn = alive(DB.share).filter(r => r.dueDate === k);
  // 기한 당일에 이미 나온 건은 빼고, 기간에만 걸친 것을 따로 모은다
  const inSpan = alive(DB.share).filter(r => r.dueDate !== k && shareSpanDays(r).indexOf(k) >= 0);

  const sec = (name, n, rows) => !n ? '' :
    '<div class="dpsec"><div class="dphead">' + name +
    '<span class="n">' + n + '</span></div>' + rows + '</div>';

  const body =
    sec('업무', tasks.length, tasks.map(r =>
      '<div class="dprow' + (r.done ? ' done' : '') + '">' +
        '<span class="mk">' + (r.done ? '✓' : '·') + '</span>' +
        '<span class="tt">' + escHtml(r.title) + '</span>' +
        (r.time ? '<span class="tm">' + escHtml(r.time) + '</span>' : '') +
      '</div>').join('')) +

    sec('업무공유 · 회신기한', dueOn.length, dueOn.map(r =>
      '<div class="dprow">' +
        '<span class="mk">·</span>' +
        '<span class="tt">' + escHtml(r.title) +
          (r.docNo ? '<em>' + escHtml(r.docNo) + '</em>' : '') + '</span>' +
        (myReplied(r.id) ? '<span class="pill green">회신함</span>'
                         : '<span class="pill red">기한</span>') +
      '</div>').join('')) +

    sec('업무공유 · 처리 중', inSpan.length, inSpan.map(r =>
      '<div class="dprow soft">' +
        '<span class="mk">·</span>' +
        '<span class="tt">' + escHtml(r.title) +
          (r.docNo ? '<em>' + escHtml(r.docNo) + '</em>' : '') + '</span>' +
        '<span class="tm">~ ' + escHtml(r.dueDate) + '</span>' +
      '</div>').join('')) +

    sec('출장', trips.length, trips.map(r => {
      const mem = r.members || [];
      return '<div class="dprow">' +
        '<span class="mk">·</span>' +
        '<span class="tt">' + escHtml(r.title) +
          '<em>' + escHtml(r.region || '-') + ' · ' + escHtml(r.transport || '-') +
          (mem.length ? ' · ' + mem.length + '명' : '') + '</em></span>' +
        (r.time ? '<span class="tm">' + escHtml(r.time) + '</span>' : '') +
      '</div>';
    }).join('')) +

    sec('연락', calls.length, calls.map(r => {
      const ext = empExt(r.name, r.dept) || r.extCache || '';
      return '<div class="dprow' + (r.done ? ' done' : '') + '">' +
        '<span class="mk">' + (r.done ? '✓' : '·') + '</span>' +
        '<span class="tt">' + escHtml(r.name) +
          (r.dept ? '<em>' + escHtml(r.dept) + '</em>' : '') + '</span>' +
        (ext ? '<span class="tm">내선 ' + escHtml(ext) + '</span>' : '') +
      '</div>';
    }).join('')) +

    sec('점심약속', meals.length, meals.map(r =>
      '<div class="dprow">' +
        '<span class="mk">·</span>' +
        '<span class="tt">' + escHtml(r.target) +
          (r.place ? '<em>' + escHtml(r.place) + '</em>' : '') + '</span>' +
        (r.time ? '<span class="tm">' + escHtml(r.time) + '</span>' : '') +
      '</div>').join(''));

  openModal(
    '<h3>' + escHtml(when) + (isToday ? ' <span class="pill blue">오늘</span>' : '') + '</h3>' +
    '<div class="body">' +
      (body || '<p class="lead">이 날은 등록된 것이 없어요.<br>이동해서 바로 등록하실 수 있습니다.</p>') +
    '</div>' +
    '<div class="foot">' +
      '<span class="spacer"></span>' +
      '<button class="btn" id="dp_close">닫기</button>' +
      '<button class="btn primary" id="dp_go">이 날짜로 이동하기</button>' +
    '</div>', {wide:true});

  $('#dp_close').onclick = closeModal;
  $('#dp_go').onclick = () => {
    closeModal();
    viewDate = sd || startOfDay(new Date());
    showView('today');
    renderToday();
    window.scrollTo(0, 0);
  };
}

function detailModal(title, body){
  openModal(
    '<h3>' + escHtml(title) + '</h3>' +
    '<div class="body"><div class="detailBody">' + escHtml(body) + '</div></div>' +
    '<div class="foot"><span class="spacer"></span><button class="btn" id="d_close">닫기</button></div>',
    {wide:true});
  $('#d_close').onclick = closeModal;
}


module.exports={escHtml,esc,unesc,buildParts,feed,inboxRecords,syncApply,resetInbox,rowsToEmp,csvToRows,decodeCsv,xlsxToRows,zipEntries,setDB:d=>{DB=d},getDB:()=>DB};
