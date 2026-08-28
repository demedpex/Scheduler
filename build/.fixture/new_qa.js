/* ------------------------------------------------------------
   한 줄 빠른 입력 파서
   "보고서 초안 오후 2시" -> 제목 "보고서 초안", 시간 14:00
   "다음주 화요일 팀 회의" -> 다음주 화요일
   못 알아들은 말은 전부 제목으로 둔다. 입력이 실패하는 일은 없어야 한다.
   ------------------------------------------------------------ */
const QA_WEEK = { '월':1, '화':2, '수':3, '목':4, '금':5, '토':6, '일':0 };

function parseQuickAdd(text, baseDate){
  let t = ' ' + String(text || '').replace(/\s+/g, ' ').trim() + ' ';
  const today = startOfDay(new Date());
  let date = dateStr(baseDate || today);
  let time = '';
  let usedDate = false;

  const cut = m => { t = t.replace(m, ' '); };

  /* ---------- 시간 ----------
     오전/오후 + 시:분 / N시 M분 / N시반.
     14:00~15:00 처럼 범위로 적어도 앞의 것만 쓴다. */
  let m = t.match(/ (오전|오후|아침|저녁|밤)?\s?(\d{1,2}):(\d{2})(?:\s?[~\-–]\s?\d{1,2}:\d{2})? /);
  if(m && +m[2] <= 23 && +m[3] <= 59){
    time = qaHour(m[1], +m[2]) + ':' + pad(+m[3]);
    cut(m[0]);
  }else{
    m = t.match(/ (오전|오후|아침|저녁|밤)?\s?(\d{1,2})\s?시\s?(반|(\d{1,2})\s?분)?(?:\s?[~\-–]\s?\d{1,2}\s?시)? /);
    if(m && +m[2] <= 23){
      const mi = m[3] === '반' ? 30 : (m[4] ? +m[4] : 0);
      if(mi <= 59){ time = qaHour(m[1], +m[2]) + ':' + pad(mi); cut(m[0]); }
    }
  }

  /* ---------- 날짜: 낱말 ---------- */
  const rel = [['글피', 3], ['모레', 2], ['내일', 1], ['오늘', 0]];
  for(const [w, off] of rel){
    const re = new RegExp(' ' + w + '(?:까지|부터|에|중|안에)? ');
    const mm = t.match(re);
    if(mm){ date = dateStr(addDays(today, off)); cut(mm[0]); usedDate = true; break; }
  }

  /* ---------- 날짜: N일 뒤 / N주 뒤 ---------- */
  if(!usedDate){
    m = t.match(/ (\d{1,2})\s?(일|주)\s?(뒤|후)(?:에)? /);
    if(m){
      const n = +m[1] * (m[2] === '주' ? 7 : 1);
      date = dateStr(addDays(today, n)); cut(m[0]); usedDate = true;
    }
  }

  /* ---------- 날짜: (이번주|다음주) X요일 ---------- */
  if(!usedDate){
    m = t.match(/ (이번\s?주|금주|다음\s?주|차주|담주|담달)?\s?([월화수목금토일])요일?(?:까지|부터|에|중)? /);
    if(m && QA_WEEK[m[2]] !== undefined){
      const wantDow = QA_WEEK[m[2]];
      const scope = (m[1] || '').replace(/\s/g, '');
      let d;
      if(scope === '다음주' || scope === '차주' || scope === '담주'){
        d = qaWeekday(addDays(qaMonday(today), 7), wantDow);
      }else if(scope === '이번주' || scope === '금주'){
        d = qaWeekday(qaMonday(today), wantDow);
      }else{
        // 그냥 "금요일" 이면 오늘 포함 다음에 오는 그 요일
        d = today;
        for(let i = 0; i < 7; i++){ if(d.getDay() === wantDow) break; d = addDays(d, 1); }
      }
      date = dateStr(d); cut(m[0]); usedDate = true;
    }
  }

  /* ---------- 날짜: 숫자 표기 ---------- */
  if(!usedDate){
    // 9/3, 9.3, 9-3, 09/03
    m = t.match(/ (\d{1,2})\s?[\/.\-]\s?(\d{1,2})(?:까지|부터|에)? /);
    if(!m) m = t.match(/ (\d{1,2})\s?월\s?(\d{1,2})\s?일(?:까지|부터|에)? /);
    if(m){
      const mo = +m[1], dd = +m[2];
      if(mo >= 1 && mo <= 12 && dd >= 1 && dd <= 31){
        date = qaYearOf(mo, dd, today); cut(m[0]); usedDate = true;
      }
    }
  }
  if(!usedDate){
    // "27일" -> 이번 달 27일 (이미 지났으면 다음 달)
    m = t.match(/ (\d{1,2})\s?일(?:까지|부터|에)? /);
    if(m){
      const dd = +m[1];
      if(dd >= 1 && dd <= 31){
        let y = today.getFullYear(), mo = today.getMonth() + 1;
        if(dd < today.getDate()){ mo++; if(mo > 12){ mo = 1; y++; } }
        const cand = y + '-' + pad(mo) + '-' + pad(dd);
        if(parseDate(cand)){ date = cand; cut(m[0]); usedDate = true; }
      }
    }
  }

  return { title: t.replace(/\s+/g, ' ').trim(), date, time, usedDate, usedTime: !!time };
}

// 오전/오후를 반영한 24시간제 시각
function qaHour(ampm, h){
  if(ampm === '오후' || ampm === '저녁' || ampm === '밤'){ if(h < 12) h += 12; }
  else if(ampm === '오전' || ampm === '아침'){ if(h === 12) h = 0; }
  return pad(h);
}
// 그 주의 월요일
function qaMonday(d){
  const dow = d.getDay();               // 0=일
  return addDays(d, dow === 0 ? -6 : 1 - dow);
}
// 월요일 기준 주에서 원하는 요일
function qaWeekday(monday, wantDow){
  return addDays(monday, wantDow === 0 ? 6 : wantDow - 1);
}
// 월/일만 적었을 때의 연도 — 이미 많이 지났으면 내년으로 본다
function qaYearOf(mo, dd, today){
  const y = today.getFullYear();
  const cand = new Date(y, mo - 1, dd);
  if(daysBetween(cand, today) < -180) return (y + 1) + '-' + pad(mo) + '-' + pad(dd);
  return y + '-' + pad(mo) + '-' + pad(dd);
}
