/* =============================================================
 *  xlsx-test.js - 명부 읽기 검증 (엑셀 .xlsx / CSV)
 *
 *  실행:  node build\xlsx-test.js
 *
 *  dist\배포\sched.html 에서 명부 관련 코드를 그대로 떼어내 돌린다.
 *  Node 에는 DOMParser 가 없으므로 시험용 XML 파서를 끼워 준다.
 *  (브라우저에서는 내장 DOMParser 를 쓴다. 여기서 검증하는 것은
 *   ZIP 해석 / 열 번호 계산 / 공유문자열 / 머리글 매핑 로직이다)
 * ============================================================= */
'use strict';
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

const root = path.resolve(__dirname, '..');
const htmlPath = path.join(root, 'dist', '배포', 'sched.html');

let pass = 0, fail = 0;
function chk(name, cond, detail){
  if(cond){ pass++; console.log('  [통과] ' + name); }
  else { fail++; console.log('  [실패] ' + name + '  -> ' + (detail === undefined ? '' : detail)); }
}

/* ---------- 시험용 XML 파서 (브라우저 DOMParser 대체) ---------- */
function makeNode(tag){
  return {
    tagName: tag, attrs: {}, children: [], text: '',
    getAttribute(n){ return this.attrs[n] === undefined ? null : this.attrs[n]; },
    get textContent(){
      return this.text + this.children.map(c => c.textContent).join('');
    },
    getElementsByTagName(n){
      const out = [];
      (function walk(node){
        for(const c of node.children){
          if(n === '*' || c.tagName === n) out.push(c);
          walk(c);
        }
      })(this);
      return out;
    },
    // 접두사를 떼고 이름만 비교 (브라우저 getElementsByTagNameNS('*', n) 과 같은 동작)
    getElementsByTagNameNS(ns, n){
      const out = [];
      (function walk(node){
        for(const c of node.children){
          const local = c.tagName.includes(':') ? c.tagName.split(':').pop() : c.tagName;
          if(n === '*' || local === n) out.push(c);
          walk(c);
        }
      })(this);
      return out;
    }
  };
}
function parseXml(src){
  const rootNode = makeNode('#doc');
  const stack = [rootNode];
  const re = /<([!?/]?)([A-Za-z_][\w:.-]*)([^>]*?)(\/?)>|([^<]+)/g;
  let m;
  while((m = re.exec(src))){
    if(m[5] !== undefined){
      const t = m[5];
      if(t.trim() || t.includes(' ')) stack[stack.length-1].text += decodeEnt(t);
      continue;
    }
    const [, pre, tag, attrStr, selfClose] = m;
    if(pre === '!' || pre === '?') continue;
    if(pre === '/'){
      if(stack.length > 1) stack.pop();
      continue;
    }
    const node = makeNode(tag);
    const ar = /([\w:.-]+)\s*=\s*"([^"]*)"/g;
    let a;
    while((a = ar.exec(attrStr))) node.attrs[a[1]] = decodeEnt(a[2]);
    stack[stack.length-1].children.push(node);
    if(!selfClose) stack.push(node);
  }
  return rootNode;
}
function decodeEnt(s){
  return s.replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"')
          .replace(/&apos;/g,"'").replace(/&#(\d+);/g,(_,d)=>String.fromCharCode(+d))
          .replace(/&amp;/g,'&');
}
global.DOMParser = class { parseFromString(s){ return parseXml(s); } };

/* ---------- HTML 에서 명부 코드 떼어내기 ---------- */
const src = fs.readFileSync(htmlPath, 'utf8');
const a = src.indexOf('const EMP_HEAD');
const b = src.indexOf('4. 동기화 프로토콜');
if(a < 0 || b < 0){ console.log('  [실패] 앵커를 찾지 못했습니다'); process.exit(1); }
const bStart = src.lastIndexOf('/* ====', b);
const code = src.slice(a, bStart);

const modPath = path.join(__dirname, '.fixture', '_emp.js');
fs.mkdirSync(path.dirname(modPath), { recursive: true });
fs.writeFileSync(modPath,
  code + '\nmodule.exports = { rowsToEmp, csvToRows, splitCsvLine, decodeCsv, zipEntries, unzipText, colIndex, xlsxToRows };\n',
  'utf8');

let M;
try{ M = require(modPath); }
catch(e){ console.log('  [실패] 떼어낸 코드가 실행되지 않습니다 -> ' + e.message); process.exit(1); }

console.log('');
console.log('===== 명부 읽기 검증 =====');
console.log('');
console.log('1. 코드 적재');
chk('HTML 에서 명부 코드를 떼어내 실행', true);

/* ---------- 시험용 .xlsx 만들기 ---------- */
function makeZip(files){
  const chunks = [], central = [];
  let offset = 0;
  for(const [name, content] of files){
    const nameBuf = Buffer.from(name, 'utf8');
    const raw = Buffer.from(content, 'utf8');
    const comp = zlib.deflateRawSync(raw);
    const crc = zlib.crc32 ? zlib.crc32(raw) : crc32(raw);

    const lh = Buffer.alloc(30);
    lh.writeUInt32LE(0x04034b50, 0); lh.writeUInt16LE(20, 4); lh.writeUInt16LE(0, 6);
    lh.writeUInt16LE(8, 8); lh.writeUInt16LE(0, 10); lh.writeUInt16LE(0, 12);
    lh.writeUInt32LE(crc, 14); lh.writeUInt32LE(comp.length, 18); lh.writeUInt32LE(raw.length, 22);
    lh.writeUInt16LE(nameBuf.length, 26); lh.writeUInt16LE(0, 28);
    chunks.push(lh, nameBuf, comp);

    const ch = Buffer.alloc(46);
    ch.writeUInt32LE(0x02014b50, 0); ch.writeUInt16LE(20, 4); ch.writeUInt16LE(20, 6);
    ch.writeUInt16LE(0, 8); ch.writeUInt16LE(8, 10);
    ch.writeUInt32LE(crc, 16); ch.writeUInt32LE(comp.length, 20); ch.writeUInt32LE(raw.length, 24);
    ch.writeUInt16LE(nameBuf.length, 28); ch.writeUInt32LE(offset, 42);
    central.push(ch, nameBuf);

    offset += lh.length + nameBuf.length + comp.length;
  }
  const cd = Buffer.concat(central);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(files.length, 8); eocd.writeUInt16LE(files.length, 10);
  eocd.writeUInt32LE(cd.length, 12); eocd.writeUInt32LE(offset, 16);
  return Buffer.concat([...chunks, cd, eocd]);
}
function crc32(buf){
  let c, table = crc32.t || (crc32.t = (() => {
    const t = [];
    for(let n = 0; n < 256; n++){ c = n; for(let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
    return t;
  })());
  let crc = 0xFFFFFFFF;
  for(const byte of buf) crc = (crc >>> 8) ^ table[(crc ^ byte) & 0xFF];
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

// 공유 문자열: 머리글(순서를 일부러 뒤섞음) + 사람 이름 + 0212
const SS = ['부서','이름','직급','내선번호','기획팀','홍길동','과장','1234',
            '전산팀','영벌이','대리','0212','총무팀','김영희',''];
const ssXml = '<?xml version="1.0"?><sst count="' + SS.length + '" uniqueCount="' + SS.length + '">'
            + SS.map(s => '<si><t>' + s.replace(/&/g,'&amp;').replace(/</g,'&lt;') + '</t></si>').join('')
            + '</sst>';

function cS(ref, i){ return '<c r="' + ref + '" t="s"><v>' + i + '</v></c>'; }
function cN(ref, v){ return '<c r="' + ref + '"><v>' + v + '</v></c>'; }

const sheetXml = '<?xml version="1.0"?><worksheet><sheetData>'
  // 머리글 순서를 일부러 뒤섞었다: 부서 | 이름 | 직급 | 내선번호
  + '<row r="1">' + cS('A1',0) + cS('B1',1) + cS('C1',2) + cS('D1',3) + '</row>'
  + '<row r="2">' + cS('A2',4) + cS('B2',5) + cS('C2',6) + cS('D2',7) + '</row>'
  + '<row r="3">' + cS('A3',8) + cS('B3',9) + cS('C3',10) + cS('D3',11) + '</row>'
  // 부서·직급이 빈 행 + 내선이 숫자로 저장된 경우
  + '<row r="4">' + cS('B4',13) + cN('D4', 5678) + '</row>'
  // 중간 빈 행 → 여기서 멈춰야 한다
  + '<row r="5"></row>'
  + '<row r="6">' + cS('B6',5) + cS('D6',7) + '</row>'
  + '</sheetData></worksheet>';

const xlsx = makeZip([
  ['[Content_Types].xml', '<?xml version="1.0"?><Types/>'],
  ['xl/workbook.xml', '<?xml version="1.0"?><workbook><sheets><sheet name="명부" sheetId="1" r:id="rId1"/></sheets></workbook>'],
  ['xl/_rels/workbook.xml.rels', '<?xml version="1.0"?><Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>'],
  ['xl/sharedStrings.xml', ssXml],
  ['xl/worksheets/sheet1.xml', sheetXml]
]);

/* ---------- 2. xlsx 읽기 ---------- */
console.log('');
console.log('2. 엑셀(.xlsx) 읽기');

(async () => {
  let rows;
  try{
    const ab = xlsx.buffer.slice(xlsx.byteOffset, xlsx.byteOffset + xlsx.byteLength);
    rows = await M.xlsxToRows(ab);
    chk('ZIP 해석 + 시트 읽기', true);
  }catch(e){
    chk('ZIP 해석 + 시트 읽기', false, e.message);
    finish(); return;
  }

  chk('행을 6개 읽음 (빈 행 포함)', rows.length === 6, '행 수 ' + rows.length);
  chk('공유 문자열이 풀림', rows[1] && rows[1][1] === '홍길동', JSON.stringify(rows[1]));

  const emp = M.rowsToEmp(rows);
  chk('머리글 순서가 뒤섞여도 열을 찾아냄', emp.length >= 3, '명 수 ' + emp.length);
  // ★ 중간 빈 줄에서 멈추면 명부가 조용히 잘린다 (70명 중 32명만 읽히는 사고).
  //    빈 줄은 건너뛰고 끝까지 읽어야 한다.
  chk('중간 빈 줄을 건너뛰고 끝까지 읽음 (' + emp.length + '명)', emp.length === 4,
      '빈 줄에서 멈췄을 가능성 — 읽힌 명단: ' + emp.map(e => e.name).join(', '));
  chk('빈 줄 뒤에 있던 사람도 읽힘', emp.some(e => e.name === '홍길동' && !e.dept),
      emp.map(e => e.name + '/' + e.dept).join(', '));
  if(emp.length >= 3){
    chk('이름·부서·직급 매핑', emp[0].name === '홍길동' && emp[0].dept === '기획팀' && emp[0].rank === '과장',
        JSON.stringify(emp[0]));
    chk('내선 0212 가 212 로 바뀌지 않음', emp[1].ext === '0212', JSON.stringify(emp[1]));
    chk('부서·직급이 비어도 읽힘', emp[2].name === '김영희' && emp[2].dept === '' && emp[2].rank === '',
        JSON.stringify(emp[2]));
    chk('숫자로 저장된 내선도 읽힘', emp[2].ext === '5678', JSON.stringify(emp[2]));
  }

  /* ---------- 3. 실제 오피스가 만든 파일 ---------- */
  console.log('');
  console.log('3. 실제 오피스가 만든 파일 (한셀 빌드 결과물)');
  const realPath = path.join(root, 'dist', 'Scheduler_한셀테스트.xlsm');
  if(fs.existsSync(realPath)){
    try{
      const buf = fs.readFileSync(realPath);
      const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
      const r = await M.xlsxToRows(ab);
      // 이 파일의 첫 시트(달력)에는 값이 들어 있다. 0행이면 파서가 못 읽은 것이다.
      chk('진짜 오피스 파일에서 행을 읽어냄 (' + r.length + '행)', r.length > 0,
          '0행 — 네임스페이스 접두사(<x:row>)를 놓쳤을 가능성');
      const flat = r.flat().filter(Boolean);
      chk('셀 값까지 풀림 (' + flat.length + '개)', flat.length > 0, JSON.stringify(r.slice(0,3)));
    }catch(e){
      chk('진짜 오피스 파일 읽기', false, e.message);
    }
  }else{
    console.log('  (건너뜀 - ' + realPath + ' 없음)');
  }

  /* ---------- 4. CSV ---------- */
  console.log('');
  console.log('4. CSV 읽기');
  const csvBom = Buffer.concat([Buffer.from([0xEF,0xBB,0xBF]),
    Buffer.from('이름,내선번호,부서,직급\r\n영벌이,0212,전산팀,대리\r\n김영희,1234,,\r\n,,,\r\n홍길동,9999,기획팀,과장\r\n','utf8')]);
  const csvRows = M.csvToRows(M.decodeCsv(csvBom.buffer.slice(csvBom.byteOffset, csvBom.byteOffset + csvBom.byteLength)));
  const csvEmp = M.rowsToEmp(csvRows);
  chk('BOM 붙은 UTF-8 CSV 를 읽음 (' + csvEmp.length + '명)', csvEmp.length === 3, JSON.stringify(csvEmp));
  chk('CSV 에서도 0212 보존', csvEmp[0] && csvEmp[0].ext === '0212', JSON.stringify(csvEmp[0]));
  chk('CSV 도 중간 빈 줄 뒤까지 읽음 (홍길동 포함)', csvEmp.some(e => e.name === '홍길동'),
      csvEmp.map(e => e.name).join(', '));

  const csv949 = Buffer.from('이름,내선번호\n홍길동,0212\n', 'utf8');
  const r949 = M.rowsToEmp(M.csvToRows(M.decodeCsv(csv949.buffer.slice(csv949.byteOffset, csv949.byteOffset + csv949.byteLength))));
  chk('BOM 없는 UTF-8 CSV 도 처리', r949.length === 1 && r949[0].name === '홍길동', JSON.stringify(r949));

  const tsv = Buffer.from('이름\t내선번호\t부서\n홍길동\t0212\t기획팀\n', 'utf8');
  const rTsv = M.rowsToEmp(M.csvToRows(M.decodeCsv(tsv.buffer.slice(tsv.byteOffset, tsv.byteOffset + tsv.byteLength))));
  chk('탭 구분 파일도 처리', rTsv.length === 1 && rTsv[0].ext === '0212', JSON.stringify(rTsv));

  finish();
})();

function finish(){
  console.log('');
  if(fail === 0){ console.log('===== 전부 통과 (' + pass + '건) ====='); process.exit(0); }
  else { console.log('===== 성공 ' + pass + ' / 실패 ' + fail + ' ====='); process.exit(1); }
}
