function renderAll(){ renderToday(); renderCal(); renderSetup(); paintSaveState(); }

function sortByTime(list){
  const withT = list.filter(r => r.time).sort((a,b) => a.time < b.time ? -1 : a.time > b.time ? 1 : 0);
  return withT.concat(list.filter(r => !r.time));
}

// 업무 건수에 따라 블록 크기를 바꾼다. 적으면 크게, 많으면 조밀하게.
function taskDensity(n){ return n <= 4 ? 'd1' : n <= 10 ? 'd2' : 'd3'; }

function emptyBox(head, tail){
  return '<div class="empty"><b>' + escHtml(head) + '</b>' + escHtml(tail) + '</div>';
}

function renderToday(){
  const d = viewDate, ds = dateStr(d);
  const isToday = ds === todayStr();
  const diff = daysBetween(d, startOfDay(new Date()));
  let suffix = '';
  if(diff === 0) suffix = ' · 오늘';
  else if(diff === 1) suffix = ' · 내일';
  else if(diff === -1) suffix = ' · 어제';
  else suffix = diff > 0 ? ' · ' + diff + '일 뒤' : ' · ' + (-diff) + '일 전';

  $('#dLabel').textContent = ds + ' (' + WD[d.getDay()] + ')';
  $('#dToday').disabled = isToday;
  $('#hdrWhere').textContent = '오늘 · ' + ds + ' (' + WD[d.getDay()] + ')' + suffix;

  /* ---------- 1) 업무 (주역) ---------- */
  const tasks = sortByTime(alive(DB.task).filter(r => r.date === ds));
  const doneN = tasks.filter(r => r.done).length;

  const grid = $('#listTask');
  grid.className = 'rows taskgrid ' + taskDensity(tasks.length);
  grid.innerHTML = tasks.length ? tasks.map(r =>
    '<div class="row' + (r.done ? ' done' : '') + '" data-kind="task" data-id="' + escHtml(r.id) + '">' +
      '<span class="bar" style="background:' + (r.done ? 'var(--line-2)' : 'var(--task)') + '"></span>' +
      '<input type="checkbox" class="chk"' + (r.done ? ' checked' : '') + '>' +
      '<div style="min-width:0">' +
        '<div class="title">' + escHtml(r.title) + '</div>' +
        (r.content ? '<div class="sub">' + escHtml(r.content) + '</div>' : '') +
      '</div>' +
      '<span class="time">' + escHtml(r.time || '') + '</span>' +
      '<span class="acts"><button class="btn sm" data-edit>수정</button></span>' +
    '</div>').join('')
    : emptyBox('오늘 등록된 업무가 없어요.', '[＋ 업무 추가]를 누르면 여기에 보여요.');

  $('#taskCount').textContent = tasks.length
    ? tasks.length + '건 중 ' + doneN + '건 했어요'
    : '0건';

  const prog = $('#taskProgress');
  if(tasks.length){
    prog.hidden = false;
    $('#taskProgressText').textContent = '오늘 업무 ' + tasks.length + '건 중 ' + doneN + '건 했어요';
    $('#taskProgressBar').style.width = Math.round(doneN / tasks.length * 100) + '%';
  }else{
    prog.hidden = true;
  }

  /* ---------- 2) 업무공유 ---------- */
  const shares = alive(DB.share)
    .filter(r => { const due = parseDate(r.dueDate); return !due || daysBetween(due, d) >= -7; })
    .sort((a,b) => (a.dueDate || '') < (b.dueDate || '') ? -1 : 1);

  $('#listShare').innerHTML = shares.length ? shares.map(r => {
    const due = parseDate(r.dueDate);
    let badge = '<span class="pill gray">기한 없음</span>', tone = '';
    if(due){
      const dd = daysBetween(due, d);
      if(dd === 0){ badge = '<span class="pill red-solid">D-DAY</span>'; tone = ' due-now'; }
      else if(dd < 0){ badge = '<span class="pill red">D+' + (-dd) + '</span>'; tone = ' due-over'; }
      else if(dd <= 3){ badge = '<span class="pill orange">D-' + dd + '</span>'; tone = ' due-soon'; }
      else badge = '<span class="pill gray">D-' + dd + '</span>';
    }
    return '<div class="row' + tone + '" data-kind="share" data-id="' + escHtml(r.id) + '">' +
      badge +
      '<div style="min-width:0">' +
        '<div class="title">' + escHtml(r.title) + '</div>' +
        '<div class="sub">문서번호 ' + escHtml(r.docNo) +
          ' · 회신기한 ' + escHtml((r.dueDate || '').slice(5).replace('-', '/')) +
          ' · ' + escHtml(r.author || '') + '</div>' +
      '</div>' +
      '<span class="acts"><button class="btn sm" data-edit>수정</button></span>' +
    '</div>';
  }).join('') : emptyBox('공유된 업무가 없어요.', '[＋ 공유 추가]로 등록하거나 [가져오기]로 받아오세요.');
  $('#shareCount').textContent = shares.length + '건';

  /* ---------- 3) 유관부서 연락하기 ----------
     내선번호는 저장값이 아니라 이름+부서로 명부에서 매번 조회한 결과다. */
  const contacts = sortByTime(alive(DB.contact).filter(r =>
    r.date === ds || (!r.date && isToday && !r.done)));

  $('#listContact').innerHTML = contacts.length ? contacts.map(r => {
    const ext = empExt(r.name, r.dept) || r.extCache;
    const extHtml = ext
      ? '<span class="ext">내선 ' + escHtml(ext) + '</span>'
      : '<span class="ext missing">내선번호를 직접 넣으면 돼요</span>';
    return '<div class="row' + (r.done ? ' done' : '') + '" data-kind="contact" data-id="' + escHtml(r.id) + '">' +
      '<input type="checkbox" class="chk"' + (r.done ? ' checked' : '') + '>' +
      '<div style="min-width:0;display:flex;align-items:baseline;gap:8px;flex-wrap:wrap">' +
        '<span class="title">' + escHtml(r.name) + (r.dept ? ' (' + escHtml(r.dept) + ')' : '') + '</span>' +
        extHtml +
      '</div>' +
      '<span class="time">' + escHtml(r.time || '') + '</span>' +
    '</div>';
  }).join('') : emptyBox('연락할 곳이 없어요.', '[＋ 연락 추가]를 누르면 여기에 보여요.');
  $('#contactCount').textContent = contacts.length + '건';

  /* ---------- 4) 점심약속 ---------- */
  const lunches = sortByTime(alive(DB.lunch).filter(r => r.date === ds));
  $('#listLunch').innerHTML = lunches.length ? lunches.map(r =>
    '<div class="row" data-kind="lunch" data-id="' + escHtml(r.id) + '">' +
      '<div style="min-width:0;display:flex;align-items:baseline;gap:9px;flex-wrap:wrap">' +
        '<span class="title">' + escHtml(r.target) + '</span>' +
        (r.place ? '<span class="place">' + escHtml(r.place) + '</span>' : '') +
      '</div>' +
      '<span class="time">' + escHtml(r.time || '') + '</span>' +
    '</div>').join('') : emptyBox('점심약속이 없어요.', '[＋ 점심 추가]로 잡아 두세요.');
  $('#lunchCount').textContent = lunches.length + '건';
}

function renderCal(){
  const first = new Date(calMonth.getFullYear(), calMonth.getMonth(), 1);
  const now = new Date();
  $('#mLabel').textContent = first.getFullYear() + '년 ' + (first.getMonth() + 1) + '월';
  $('#mToday').disabled = (first.getFullYear() === now.getFullYear() && first.getMonth() === now.getMonth());
  $('#hdrWhere').textContent = '달력 · ' + first.getFullYear() + '년 ' + (first.getMonth() + 1) + '월';

  // 각 표를 한 번씩만 순회해 날짜별로 집계한다
  const cnt = {};
  const bump = (k, f) => { if(!k) return; (cnt[k] = cnt[k] || {t:0,c:0,s:0,l:0})[f]++; };
  alive(DB.task).forEach(r => bump(r.date, 't'));
  alive(DB.contact).forEach(r => bump(r.date, 'c'));
  alive(DB.lunch).forEach(r => bump(r.date, 'l'));
  const urgent = {};
  const today0 = startOfDay(new Date());
  alive(DB.share).forEach(r => {
    bump(r.dueDate, 's');
    const dueD = parseDate(r.dueDate);
    if(dueD && dueD >= today0 && daysBetween(dueD, today0) <= 3) urgent[r.dueDate] = true;
  });

  let cur = addDays(first, -first.getDay());
  const ts = todayStr();
  let html = '';
  for(let i = 0; i < 42; i++){
    const k = dateStr(cur);
    const out = cur.getMonth() !== first.getMonth();
    const c = cnt[k];

    // 칩 두 줄: 위는 업무/연락, 아래는 공유/점심. 기한 임박이면 붉은 칩.
    let chips = '';
    if(c){
      const a = [c.t ? '업무 ' + c.t : '', c.c ? '연락 ' + c.c : ''].filter(Boolean).join(' · ');
      const b = [c.s ? '공유 ' + c.s : '', c.l ? '점심 ' + c.l : ''].filter(Boolean).join(' · ');
      if(a) chips += '<span class="chip task">' + a + '</span>';
      if(b) chips += '<span class="chip ' + (!out && urgent[k] ? 'due' : c.s ? 'share' : 'lunch') + '">' +
                     b + (!out && urgent[k] ? ' · 임박' : '') + '</span>';
    }

    html += '<div class="cell' + (out ? ' out' : '') + (k === ts ? ' today' : '')
          + (cur.getDay() === 0 ? ' sun' : cur.getDay() === 6 ? ' sat' : '')
          + '" data-date="' + k + '">'
          + '<span class="d">' + cur.getDate() + '</span>'
          + chips
          + '</div>';
    cur = addDays(cur, 1);
  }
  $('#calGrid').innerHTML = html;
}

function renderSetup(){
  $('#cfgName').value = DB.cfg.name || '';
  $('#cfgPartMax').value = partMax();
  $('#empInfo').textContent = DB.emp.length
    ? DB.emp.length + '명' + (DB.cfg.empAt ? '  ·  ' + DB.cfg.empAt + '에 올렸어요' : '')
    : '아직 명부를 안 올렸어요.';

  // 지금 데이터로 실제 몇 조각이 되는지 바로 보여 준다
  const n = alive(DB.share).length;
  $('#partHint').textContent = n
    ? '지금 공유 ' + n + '건을 전체로 내보내면 ' + buildParts(DB.share).length + '조각이에요.'
    : '';

  $('#lastSync').textContent = DB.cfg.lastSync || '아직 없어요';
  $('#countInfo').textContent =
    '업무 ' + alive(DB.task).length + ' · 연락 ' + alive(DB.contact).length +
    ' · 공유 ' + alive(DB.share).length + ' · 점심 ' + alive(DB.lunch).length;
}
