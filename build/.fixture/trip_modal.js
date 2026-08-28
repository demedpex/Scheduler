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
        '<div class="hint">명부에 있는 분은 이름을 치면 아래에 나옵니다. 최대 ' + TRIP_MAX_MEMBER + '명.</div>' +
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
