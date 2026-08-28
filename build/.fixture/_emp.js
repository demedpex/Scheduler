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


module.exports = { rowsToEmp, csvToRows, splitCsvLine, decodeCsv, zipEntries, unzipText, colIndex, xlsxToRows };
