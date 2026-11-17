/* 신입생 반편성 도우미 - 화면 동작. 모든 처리는 브라우저 안에서만 한다. */
(function () {
  'use strict';
  const $ = s => document.querySelector(s);
  const KO = new Intl.Collator('ko');
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const digits = s => String(s == null ? '' : s).replace(/\D/g, '');
  const normH = h => String(h == null ? '' : h).replace(/\s+/g, '');
  const radio = n => (document.querySelector('input[name="' + n + '"]:checked') || {}).value;
  const fmt = (v, d = 2) => (v == null || !isFinite(v)) ? '' : (Math.round(v * 10 ** d) / 10 ** d).toFixed(d);

  const FIELDS = [
    ['name', '성명', true], ['sex', '성별', true], ['school', '출신학교', true], ['score', '성적', true],
    ['ptel', '보호자 연락처', false], ['exam', '수험번호', false], ['note', '비고', false]
  ];

  const S = { wb: null, fileName: '', sheets: {}, mainName: '', extraName: '', map: {}, students: [], pairs: [], fixedMap: {}, result: null, K: 7, T: 10, seed: 1, tab: 1, autoPairs: [], manualPairs: [], running: false };

  // ---------- 파일 읽기 ----------
  function parseSheet(name, ws) {
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true, blankrows: false });
    let h = -1;
    for (let i = 0; i < Math.min(rows.length, 20); i++) {
      if (rows[i].some(c => /^(성명|이름|학생명|학생이름|성\s*명)$/.test(normH(c)))) { h = i; break; }
    }
    if (h < 0) return { name, ok: false, rows };
    const headers = rows[h].map(normH);
    const raw = rows[h].map(c => String(c == null ? '' : c).replace(/\s*\n\s*/g, ' ').trim());
    const width = Math.max(headers.length, ...rows.slice(h + 1).map(r => r.length));
    while (headers.length < width) { headers.push(''); raw.push(''); }
    const data = rows.slice(h + 1).filter(r => r.some(c => String(c).trim() !== ''));
    return { name, ok: true, headers, raw, data };
  }

  function guessMap(headers) {
    const find = (...tests) => { for (const t of tests) { const i = headers.findIndex(t); if (i >= 0) return i; } return -1; };
    return {
      name: find(h => /^(성명|이름|학생명|학생이름)$/.test(h), h => h.includes('성명') || h.includes('이름')),
      sex: find(h => h.includes('성별'), h => h === '남녀'),
      school: find(h => h.includes('출신'), h => /중학교$/.test(h), h => h === '학교' || h === '학교명'),
      score: find(h => h.includes('백분율'), h => /석차|내신|환산|성적|점수|총점/.test(h) && !/순번|번호/.test(h)),
      ptel: find(h => /보호자|학부모|부모|가정/.test(h)),
      exam: find(h => /수험|접수번호/.test(h)),
      note: find(h => h.includes('비고'))
    };
  }

  function normSex(v) { const s = String(v == null ? '' : v).trim(); if (/^(남|남자|남학생|m|male)$/i.test(s) || s.startsWith('남')) return '남'; if (/^(여|여자|여학생|f|female)$/i.test(s) || s.startsWith('여')) return '여'; return ''; }
  function normSchool(v) { let s = String(v == null ? '' : v).replace(/\s+/g, ''); s = s.replace(/학교$/, ''); return s || '(빈칸)'; }
  function toNum(v) { if (v === '' || v == null) return null; if (typeof v === 'number') return isFinite(v) ? v : null; const s = String(v).replace(/[^\d.\-]/g, ''); if (!s) return null; const n = Number(s); return isFinite(n) ? n : null; }

  async function loadFile(file) {
    const buf = await file.arrayBuffer();
    let wb;
    try { wb = XLSX.read(buf, { type: 'array', cellDates: false }); } catch (e) { alert('파일을 읽지 못했어요. 엑셀 파일이 맞는지 확인해 주세요.\n' + e.message); return; }
    loadWorkbook(wb, file.name);
  }

  function loadWorkbook(wb, fileName) {
    S.wb = wb; S.fileName = fileName; S.sheets = {}; S._kTouched = false;
    wb.SheetNames.forEach(n => S.sheets[n] = parseSheet(n, wb.Sheets[n]));
    const usable = wb.SheetNames.filter(n => S.sheets[n].ok && S.sheets[n].data.length > 0);
    if (!usable.length) { alert('성명(이름) 열이 있는 시트를 찾지 못했어요. 첫 줄 근처에 "성명" 제목이 있어야 해요.'); return; }
    // 학생 명단: 성적 열이 있는 첫 시트
    S.mainName = usable.find(n => { const m = guessMap(S.sheets[n].headers); return m.score >= 0 && !/정원\s*외|특수/.test(n); }) || usable[0];
    S.extraName = usable.find(n => n !== S.mainName && /정원\s*외|특수/.test(n)) || '';
    const mainSel = $('#mainSheet'), exSel = $('#extraSheet');
    mainSel.innerHTML = usable.map(n => `<option ${n === S.mainName ? 'selected' : ''}>${esc(n)}</option>`).join('');
    exSel.innerHTML = '<option value="">없음</option>' + usable.map(n => `<option ${n === S.extraName ? 'selected' : ''}>${esc(n)}</option>`).join('');
    $('#fileName').textContent = '파일: ' + fileName;
    $('#mapping').hidden = false;
    setMain(S.mainName);
  }

  function setMain(name) {
    S.mainName = name;
    S.map = guessMap(S.sheets[name].headers);
    const h = normH(S.sheets[name].headers[S.map.score] || '');
    document.querySelector('input[name=dir][value=' + (/점수|총점|환산/.test(h) && !/백분율|석차|등급/.test(h) ? 'high' : 'low') + ']').checked = true;
    renderColMap();
    rebuild();
  }

  function renderColMap() {
    const sh = S.sheets[S.mainName];
    const opts = sel => '<option value="-1">(없음)</option>' + sh.raw.map((r, i) => `<option value="${i}" ${i === sel ? 'selected' : ''}>${esc(r || '(제목 없음 ' + (i + 1) + '열)')}</option>`).join('');
    $('#colMap').className = 'grid4 colmap';
    $('#colMap').innerHTML = FIELDS.map(([k, label, req]) => `<label class="${req ? 'req' : ''}"><span>${label}</span><select data-k="${k}" class="${req && S.map[k] < 0 ? 'missing' : ''}">${opts(S.map[k])}</select></label>`).join('');
  }

  function rowsToStudents(sh, map, extra) {
    const out = [];
    sh.data.forEach((r, ri) => {
      const name = String(r[map.name] == null ? '' : r[map.name]).trim();
      if (!name || /^(합계|계|소계|총계)$/.test(name)) return;
      const sc = (!extra && map.score >= 0) ? toNum(r[map.score]) : null;
      out.push({
        name, sex: normSex(r[map.sex]), sexRaw: map.sex >= 0 ? String(r[map.sex]) : '',
        school: map.school >= 0 ? normSchool(r[map.school]) : '(빈칸)', score: sc,
        ptel: map.ptel >= 0 ? digits(r[map.ptel]) : '', exam: map.exam >= 0 ? String(r[map.exam] == null ? '' : r[map.exam]).trim() : '',
        note: map.note >= 0 ? String(r[map.note] == null ? '' : r[map.note]).trim() : '',
        extra, noScore: !extra && sc == null, row: r, sheet: sh.name, ri
      });
    });
    return out;
  }

  function rebuild() {
    const main = S.sheets[S.mainName];
    let st = rowsToStudents(main, S.map, false);
    if (S.extraName && S.sheets[S.extraName]) {
      const ex = S.sheets[S.extraName];
      st = st.concat(rowsToStudents(ex, guessMap(ex.headers), true));
    }
    st.forEach((s, i) => s.id = i);
    S.students = st; S.result = null; $('#step3').hidden = true;
    renderPreview();
    $('#step2').hidden = !(S.map.name >= 0 && S.map.sex >= 0 && S.map.school >= 0 && S.map.score >= 0 && st.length);
    if (!$('#step2').hidden) { suggestK(); refreshAutoPairs(); }
    if (S.calib) renderCalib();
  }

  function renderPreview() {
    const st = S.students, reg = st.filter(s => !s.extra), msgs = [];
    const missing = FIELDS.filter(f => f[2] && S.map[f[0]] < 0).map(f => f[1]);
    if (missing.length) msgs.push(['bad', '꼭 필요한 열을 골라 주세요: ' + missing.join(', ')]);
    const badSex = st.filter(s => !s.sex);
    if (S.map.sex >= 0 && badSex.length) msgs.push(['warn', `성별을 알 수 없는 학생 ${badSex.length}명: ${badSex.slice(0, 8).map(s => s.name + '(' + (s.sexRaw || '빈칸') + ')').join(', ')}${badSex.length > 8 ? ' 외' : ''}`]);
    const noSc = reg.filter(s => s.noScore);
    if (S.map.score >= 0 && noSc.length) msgs.push(['warn', `성적이 비어 있는 학생 ${noSc.length}명은 정원 외 학생처럼 성적 없이 배치해요: ${noSc.slice(0, 8).map(s => s.name).join(', ')}${noSc.length > 8 ? ' 외' : ''}`]);
    // 두 시트에 모두 있는 학생
    if (S.extraName) {
      const keys = new Set(reg.map(s => s.name + '|' + s.school));
      const dup = st.filter(s => s.extra && keys.has(s.name + '|' + s.school));
      if (dup.length) msgs.push(['bad', `두 시트에 모두 있는 학생이 있어요 (같은 이름·학교): ${dup.map(s => s.name).join(', ')}. 정원 외 시트를 확인해 주세요.`]);
    }
    const male = st.filter(s => s.sex === '남').length, female = st.filter(s => s.sex === '여').length;
    const schools = {}; st.forEach(s => schools[s.school] = (schools[s.school] || 0) + 1);
    const schoolTxt = Object.entries(schools).sort((a, b) => b[1] - a[1]).map(([k, v]) => `${k} ${v}`).join(' · ');
    const scores = reg.map(s => s.score).filter(v => v != null);
    const sh = S.sheets[S.mainName];
    const cols = FIELDS.filter(f => S.map[f[0]] >= 0);
    const sample = reg.slice(0, 4).map(s => '<tr>' + cols.map(f => `<td>${esc(f[0] === 'sex' ? s.sex : f[0] === 'school' ? s.school : f[0] === 'score' ? (s.score == null ? '' : s.score) : f[0] === 'ptel' ? (s.ptel ? '••••' + s.ptel.slice(-4) : '') : s[f[0]])}</td>`).join('') + '</tr>').join('');
    $('#preview').innerHTML = `<div class="preview">
      <div class="facts"><span>학생 ${st.length}명</span><span>정원 내 ${reg.length}명</span><span>정원 외 ${st.length - reg.length}명</span><span>남 ${male} · 여 ${female}</span>${scores.length ? `<span>성적 ${fmt(Math.min(...scores))} ~ ${fmt(Math.max(...scores))}</span>` : ''}<span>출신학교 ${Object.keys(schools).length}곳</span></div>
      <div class="hint">${esc(schoolTxt)}</div>
      <ul class="msgs">${msgs.map(m => `<li class="${m[0]}">${esc(m[1])}</li>`).join('')}</ul>
      ${sample ? `<div class="tablewrap"><table><tr>${cols.map(f => `<th>${f[1]}</th>`).join('')}</tr>${sample}</table></div><div class="hint">앞의 4명만 보여 줘요. 읽은 내용이 맞는지 확인해 주세요.</div>` : ''}
    </div>`;
    void sh;
  }

  function suggestK() {
    const n = S.students.length;
    if (!S._kTouched) $('#K').value = Math.max(2, Math.round(S.students.filter(s => !s.extra).length / 30));
    updatePerClass();
    void n;
  }
  function updatePerClass() {
    const K = +$('#K').value || 1, n = S.students.length;
    $('#perClass').textContent = `전체 ${n}명 → 반당 ${Math.floor(n / K)}${n % K ? '~' + Math.ceil(n / K) : ''}명`;
    $('#gatherClass').max = K;
  }

  // ---------- 나누기 조건 ----------
  function refreshAutoPairs() {
    const st = S.students, out = [];
    if ($('#sameName').checked) {
      const g = {}; st.forEach(s => (g[s.name] = g[s.name] || []).push(s));
      Object.values(g).filter(a => a.length > 1).forEach(a => out.push({ type: '동명이인', ids: a.map(s => s.id) }));
    }
    if ($('#sibling').checked) {
      const g = {}; st.forEach(s => { if (s.ptel.length >= 8) (g[s.ptel] = g[s.ptel] || []).push(s); });
      Object.values(g).filter(a => a.length > 1).forEach(a => out.push({ type: '형제자매', ids: a.map(s => s.id) }));
    }
    S.autoPairs = out;
    const nm = out.filter(p => p.type === '동명이인'), sb = out.filter(p => p.type === '형제자매');
    const desc = p => p.ids.map(i => `${st[i].name}(${st[i].school})`).join(' · ');
    let html = '';
    if ($('#sameName').checked) html += `동명이인 ${nm.length}쌍${nm.length ? ': ' + esc(nm.map(desc).join(' / ')) : ''}<br>`;
    if ($('#sibling').checked) html += S.map.ptel < 0 ? '보호자 연락처 열이 없어서 형제자매를 자동으로 찾을 수 없어요. 아래 칸에 직접 적어 주세요.' : `형제자매 ${sb.length}쌍${sb.length ? ': ' + esc(sb.map(desc).join(' / ')) : ''}`;
    $('#autoPairsInfo').innerHTML = html;
  }

  function resolveToken(tok) {
    tok = tok.trim(); if (!tok) return { err: null, id: null };
    const st = S.students;
    let m = tok.match(/^(.+?)[(（](.+?)[)）]$/), name = tok, hint = '';
    if (m) { name = m[1].trim(); hint = m[2].trim(); }
    let c = st.filter(s => s.name === name);
    if (!c.length && /^\d+$/.test(tok)) c = st.filter(s => s.exam === tok);
    if (hint) { const h = normSchool(hint); c = c.filter(s => s.school === h || s.school.includes(hint) || s.exam === hint || s.sex === hint); }
    if (!c.length) return { err: `"${tok}" 학생을 찾지 못했어요.` };
    if (c.length > 1) return { err: `"${tok}" 같은 이름이 ${c.length}명이에요. 이름(학교)로 적어 주세요: ${c.map(s => s.name + '(' + s.school + ')').join(', ')}` };
    return { id: c[0].id };
  }

  function parseManual() {
    const errs = [], groups = [];
    $('#manualPairs').value.split(/\n/).forEach((line, li) => {
      const toks = line.match(/[^\s,，、\/·;]+(?:\s*[(（][^)）]*[)）])?/g) || [];
      const ids = [];
      toks.forEach(t => { const r = resolveToken(t.replace(/\s+/g, '')); if (r.err) errs.push(`${li + 1}줄: ${r.err}`); else if (r.id != null) ids.push(r.id); });
      if (ids.length === 1 && toks.length === 1) errs.push(`${li + 1}줄: 2명 이상 적어 주세요.`);
      if (ids.length > 1) groups.push({ type: '직접 입력', ids: [...new Set(ids)] });
    });
    $('#manualPairsMsg').textContent = errs.join('\n');
    S.manualPairs = groups;
    const K = +$('#K').value, fixed = {}, ferr = [];
    $('#fixed').value.split(/\n/).forEach((line, li) => {
      if (!line.trim()) return;
      const m = line.trim().match(/^(.+?)\s*(?:[=:→>]|\s)\s*(\d+)\s*반?$/);
      if (!m) { ferr.push(`${li + 1}줄: "이름 = 반" 형식으로 적어 주세요.`); return; }
      const r = resolveToken(m[1].replace(/\s+/g, ''));
      if (r.err) { ferr.push(`${li + 1}줄: ${r.err}`); return; }
      const k = +m[2]; if (k < 1 || k > K) { ferr.push(`${li + 1}줄: 반은 1~${K} 사이로 적어 주세요.`); return; }
      fixed[r.id] = k;
    });
    $('#fixedMsg').textContent = ferr.join('\n');
    S.fixedMap = fixed;
    const adj = {}, aerr = [];
    $('#adjust').value.split(/\n/).forEach((line, li) => {
      if (!line.trim()) return;
      const m = line.trim().match(/^(.+?)\s*[=:]\s*([+\-−]?\d+(?:\.\d+)?)\s*$/);
      if (!m) { aerr.push(`${li + 1}줄: "학교 = 숫자" 형식으로 적어 주세요.`); return; }
      const sc = normSchool(m[1]);
      if (!S.students.some(s => s.school === sc)) aerr.push(`${li + 1}줄: "${m[1]}" 출신 학생이 명단에 없어요.`);
      adj[sc] = Number(m[2].replace('−', '-'));
    });
    $('#adjustMsg').textContent = aerr.join('\n');
    S.adjust = adj;
    let terr = '';
    S.targets = null;
    const tv = $('#targets').value.trim();
    if (tv) {
      const nums = tv.split(/[^\d]+/).filter(Boolean).map(Number);
      const sum = nums.reduce((a, b) => a + b, 0);
      if (nums.length !== K) terr = `학급 수(${K})만큼 적어 주세요. 지금 ${nums.length}개예요.`;
      else if (sum !== S.students.length) terr = `합이 전체 학생 수(${S.students.length}명)와 같아야 해요. 지금 ${sum}명이에요.`;
      else S.targets = nums;
    }
    $('#targetsMsg').textContent = terr;
    return !errs.length && !ferr.length && !aerr.length && !terr;
  }

  // ---------- 지난 입학생 자료로 보정값 계산 ----------
  function guessCalib(h) {
    const find = (...tests) => { for (const t of tests) { const i = h.findIndex(t); if (i >= 0) return i; } return -1; };
    const SC = /성적|백분율|석차|내신|점수|등급|평균|환산|등수/;
    const school = find(x => x.includes('출신'), x => /중학교$/.test(x), x => x === '학교' || x === '학교명' || x === '중학교명');
    const mid = find(x => x !== h[school] && /중학|중학교|^중/.test(x) && SC.test(x), x => /내신|입학성적|입시/.test(x) && !/고교|고등|고1/.test(x));
    const high = find(x => /고교|고등|고1|고일|입학후|1학년/.test(x) && !/중학/.test(x) && SC.test(x), x => /학년말|고교|고1/.test(x));
    const year = find(x => /연도|학년도|년도|입학년/.test(x));
    return { school, mid: mid === school ? -1 : mid, high: (high === mid || high === school) ? -1 : high, year };
  }

  function readCalibBook(wb) {
    const recs = [], info = [], errs = [];
    wb.SheetNames.forEach(name => {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '', raw: true, blankrows: false });
      let hi = -1, m = null;
      for (let i = 0; i < Math.min(rows.length, 20); i++) {
        const g = guessCalib(rows[i].map(normH));
        if (g.school >= 0 && g.mid >= 0 && g.high >= 0) { hi = i; m = g; break; }
      }
      if (hi < 0) { if (rows.length) errs.push(`"${name}" 시트: 출신학교·중학교 성적·고1 성적 열을 찾지 못해 건너뛰었어요.`); return; }
      const raw = rows[hi].map(c => String(c == null ? '' : c).replace(/\s*\n\s*/g, ' ').trim());
      let n = 0, skip = 0;
      rows.slice(hi + 1).forEach(r => {
        const school = r[m.school], mid = toNum(r[m.mid]), high = toNum(r[m.high]);
        if (String(school).trim() === '' && mid == null && high == null) return;
        if (String(school).trim() === '' || mid == null || high == null) { skip++; return; }
        const y = m.year >= 0 && String(r[m.year]).trim() !== '' ? String(r[m.year]).trim() : name;
        recs.push({ year: y, school: normSchool(school), mid, high }); n++;
      });
      info.push(`"${name}" 시트 ${n}명 (${raw[m.school]} / ${raw[m.mid]} / ${raw[m.high]}${m.year >= 0 ? ' / ' + raw[m.year] : ''})${skip ? ` · 빈칸 있는 ${skip}줄 버림` : ''}`);
    });
    return { recs, info, errs };
  }

  async function loadCalibFile(file) {
    let wb;
    try { wb = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: false }); } catch (e) { alert('파일을 읽지 못했어요.\n' + e.message); return; }
    S.calib = Object.assign(readCalibBook(wb), { fileName: file.name });
    renderCalib();
  }

  function renderCalib() {
    const C = S.calib, out = $('#calibOut');
    if (!C) { out.innerHTML = ''; return; }
    const head = `<p class="filename">지난 자료: ${esc(C.fileName)}</p>` +
      (C.info.length ? `<p class="hint">${C.info.map(esc).join('<br>')}</p>` : '') +
      (C.errs.length ? `<p class="warn">${C.errs.map(esc).join('\n')}</p>` : '');
    if (!C.recs.length) { out.innerHTML = head + '<p class="warn">읽은 학생이 없어요. 제목 줄에 출신학교, 중학교 성적, 고1 성적이 있는지 확인해 주세요. 계산용 양식을 받아 써도 돼요.</p>'; return; }
    const r = Calib.compute(C.recs, { midLow: radio('dir') === 'low', highLow: radio('calibDir') === 'low' });
    C.result = r;
    if (!r.ok) { out.innerHTML = head + `<p class="warn">${esc(r.error)}</p>` + (r.warnings || []).map(w => `<p class="hint">${esc(w)}</p>`).join(''); return; }
    const now = {}; (S.students || []).forEach(s => { if (!s.extra) now[s.school] = (now[s.school] || 0) + 1; });
    const hasNow = Object.keys(now).length > 0;
    const V = { clear: ['뚜렷함', 'ok'], weak: ['우연일 수 있음', ''], few: ['자료 적음', ''] };
    const sign = v => (v > 0 ? '+' : '') + fmt(v, 1);
    let t = '<div class="tablewrap"><table><tr><th>출신학교</th><th>지난 자료<br>학생 수</th><th>고1 등수 차이<br><small>(같은 내신 대비, %)</small></th><th>해마다<br>같은 방향</th><th>판단</th><th>보정값</th>' + (hasNow ? '<th>올해<br>명단</th>' : '') + '</tr>';
    r.schools.forEach(s => {
      const v = V[s.verdict];
      t += `<tr><td class="l">${esc(s.school)}</td><td>${s.n}</td><td>${sign(s.mean)}</td><td>${s.yearsTotal >= 2 ? s.yearsSame + '/' + s.yearsTotal + '년' : '-'}</td><td class="${v[1] ? 'okc' : 'sub'}">${v[0]}</td><td><b>${s.verdict === 'clear' ? sign(s.value) : '<span class="sub">' + sign(s.value) + '</span>'}</b></td>` + (hasNow ? `<td>${now[s.school] ? now[s.school] + '명' : '<span class="sub">없음</span>'}</td>` : '') + '</tr>';
    });
    t += '</table></div>';
    const missing = hasNow ? Object.keys(now).filter(k => !r.schools.some(s => s.school === k)) : [];
    const pick = r.schools.filter(s => s.verdict === 'clear' && Math.abs(s.value) >= 0.5 && (!hasNow || now[s.school]));
    const facts = `<div class="preview"><div class="facts"><span>계산한 학생 ${r.N}명</span><span>${r.years.map(y => esc(y.year) + ' ' + y.n + '명').join(' · ')}</span><span>중학교→고1 성적 상관 ${fmt(r.r, 2)}</span><span>학교 사이 실제 차이 약 ${fmt(r.tau, 1)}%</span></div></div>`;
    const note = r.tau < 0.05
      ? '<ul class="msgs"><li class="ok">학교마다 보이는 차이가 우연으로 생길 수 있는 크기예요. 보정하지 않고 출신 중학교만 고르게 나누는 것을 권해요.</li></ul>'
      : `<p class="hint">· <b>고1 등수 차이</b>: 중학교 성적이 같은 학생과 비교해 고1 학년 등수(백분율)가 평균 몇 % 앞섰는지(+) 뒤졌는지(−)예요.<br>· <b>보정값</b>: 학생 수가 적을수록 0 쪽으로 줄여 중학교 성적 단위로 바꾼 값이에요(최대 ±15%). <b>뚜렷함</b>(학생 10명 이상, 우연으로 보기 어렵고, 해마다 방향이 같음)인 학교만 넣기를 권해요.</p>`;
    const btn = r.tau >= 0.05 ? (pick.length
      ? `<div class="row-btns left"><button class="btn primary small" id="calibApply" type="button">뚜렷한 학교 ${pick.length}곳 보정값을 아래 칸에 넣기</button></div>`
      : '<p class="hint">뚜렷한 학교가 없어서 넣을 보정값이 없어요. 보정하지 않는 것을 권해요.</p>') : '';
    out.innerHTML = head + facts + t + note + btn +
      (missing.length ? `<p class="hint">올해 명단에 있지만 지난 자료에 없는 학교: ${esc(missing.join(', '))} (보정 없이 배정)</p>` : '') +
      (r.warnings.length ? `<p class="hint">${r.warnings.map(esc).join('<br>')}</p>` : '');
    S.calibPick = pick;
  }

  function applyCalib() {
    const pick = S.calibPick || [];
    if (!pick.length) return;
    const ta = $('#adjust');
    if (ta.value.trim() && !confirm('보정 칸에 이미 적힌 내용을 계산 결과로 바꿀까요?')) return;
    ta.value = pick.map(s => `${s.school} = ${s.value > 0 ? '+' : ''}${fmt(s.value, 1)}`).join('\n');
    if (S.students.length) parseManual();
    ta.focus();
  }

  function calibDemo() {
    let seed = 20231117; const R = () => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed / 2147483648; };
    const N = () => { let u = 0; for (let i = 0; i < 12; i++) u += R(); return u - 6; };
    // 가상 학교 효과: 가람중은 같은 내신이면 고교에서 더 잘하고, 나래중은 덜 한다
    const schools = [['가람중', 48, 0.45], ['나래중', 37, -0.45], ['다솜중', 31, 0], ['라온중', 25, 0.05], ['마루중', 21, -0.1], ['바다중', 13, 0], ['사랑중', 6, 0.2]];
    const wb = XLSX.utils.book_new();
    [2023, 2024, 2025].forEach(y => {
      const list = [];
      schools.forEach(([sc, cnt, eff]) => { for (let i = 0; i < cnt; i++) { const ab = N(); list.push({ sc, mid: ab + 0.55 * N(), high: ab + eff + 0.55 * N() }); } });
      const rk = (k) => { const o = list.map((x, i) => i).sort((a, b) => list[b][k] - list[a][k]); const p = []; o.forEach((idx, r) => p[idx] = Math.round((r + 0.5) / list.length * 10000) / 100); return p; };
      const mp = rk('mid'), hp = rk('high');
      const rows = [['출신학교', '중학교 성적(백분율)', '고1 성적(학년 석차 백분율)']].concat(list.map((x, i) => [x.sc, mp[i], hp[i]]));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), String(y));
    });
    $('input[name=calibDir][value=low]').checked = true;
    S.calib = Object.assign(readCalibBook(wb), { fileName: '예시 지난 자료(가상, 3년).xlsx' });
    renderCalib();
  }

  function calibTemplate() {
    const wb = XLSX.utils.book_new();
    const a = XLSX.utils.aoa_to_sheet([['입학연도', '출신학교', '중학교 성적(백분율)', '고1 성적(학년 석차 백분율)'], [2024, '○○중', 12.5, 20.3], [2024, '△△중', 33.4, 28.1], [2025, '○○중', 41.0, 37.6]]);
    styleAll(a, [0]); a['!cols'] = [10, 12, 18, 24].map(w => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, a, '지난 입학생');
    const b = XLSX.utils.aoa_to_sheet([['쓰는 법'], ['한 줄에 학생 한 명. 이름은 필요 없어요.'], ['중학교 성적은 올해 반편성 명단과 같은 종류(예: 학년석차 백분율)로 넣어 주세요.'], ['고1 성적은 석차, 백분율, 등급 평균, 점수 모두 됩니다. 화면에서 낮을수록/높을수록 우수를 골라 주세요.'], ['해마다 시트를 나눠도 되고(시트 이름이 연도가 됩니다), 한 시트에 입학연도 열을 둬도 됩니다.'], ['전학·자퇴로 고1 성적이 없는 학생은 빼거나 칸을 비워 두면 됩니다.']]);
    b['!cols'] = [{ wch: 90 }];
    XLSX.utils.book_append_sheet(wb, b, '쓰는 법');
    XLSX.writeFile(wb, '보정값 계산용 양식.xlsx');
  }

  // ---------- 실행 ----------
  let worker = null;
  function getWorker() {
    if (worker) return worker;
    if (window.__OFFLINE__) return null;
    try { worker = new Worker('engine.js'); } catch (e) { worker = null; }
    return worker;
  }

  function run(newSeed) {
    if (S.running) return;
    if (!parseManual()) { alert('입력한 조건에 확인할 부분이 있어요. 빨간 글씨를 확인해 주세요.'); return; }
    const K = Math.max(2, Math.min(30, Math.round(+$('#K').value || 2)));
    const T = Math.max(2, Math.min(30, Math.round(+$('#tiers').value || 10)));
    const dir = radio('dir');
    if (newSeed) S.seed = (S.seed * 7919 + 13) % 1000003; else S.seed = 1;
    S.K = K; S.T = T;
    const st = S.students;
    const pairs = [];
    [...S.autoPairs, ...S.manualPairs].forEach(g => { for (let a = 0; a < g.ids.length; a++) for (let b = a + 1; b < g.ids.length; b++) pairs.push([g.ids[a], g.ids[b]]); });
    const input = {
      K, tiers: T, seed: S.seed, restarts: 8, iterations: 300000,
      sexLevel: radio('sexLevel'), schoolLevel: radio('schoolLevel'),
      extraMode: radio('extraMode'), gatherClass: Math.min(K, Math.max(1, +$('#gatherClass').value || 1)),
      pairs, targets: S.targets,
      students: st.map(s => {
        let score = null;
        if (!s.extra && s.score != null) { const a = S.adjust[s.school] || 0; score = dir === 'low' ? s.score - a : -(s.score + a); }
        return { score, sex: s.sex, school: s.school, extra: s.extra || s.noScore, fixed: S.fixedMap[s.id] || null };
      })
    };
    S.running = true; $('#runBtn').disabled = true; $('#againBtn').disabled = true;
    const bar = $('#progress'); bar.hidden = false; bar.firstElementChild.style.width = '3%';
    const done = out => {
      S.running = false; $('#runBtn').disabled = false; $('#againBtn').disabled = false; bar.hidden = true;
      st.forEach((s, i) => { s.cls = out.cls[i]; s.autoCls = out.cls[i]; s.rank = out.rank[i]; s.tier = out.tier[i]; });
      S.T = out.tiers; S.result = out; S.tab = 1;
      $('#step3').hidden = false; renderResult();
      $('#step3').scrollIntoView({ behavior: 'smooth' });
    };
    const w = getWorker();
    if (w) {
      w.onmessage = e => { const m = e.data; if (m.type === 'progress') bar.firstElementChild.style.width = Math.round(m.p * 100) + '%'; else if (m.type === 'done') done(m.out); else { S.running = false; $('#runBtn').disabled = false; bar.hidden = true; alert('배정 중 오류: ' + m.message); } };
      w.onerror = () => { worker = null; w.terminate(); setTimeout(() => done(ClassEngine.assign(input)), 30); };
      w.postMessage(input);
    } else setTimeout(() => done(ClassEngine.assign(input)), 30);
  }

  // ---------- 통계 ----------
  function stats() {
    const K = S.K, T = S.T, st = S.students;
    const schools = Object.entries(st.reduce((m, s) => (m[s.school] = (m[s.school] || 0) + 1, m), {})).sort((a, b) => b[1] - a[1] || KO.compare(a[0], b[0])).map(e => e[0]);
    const cls = [];
    for (let k = 1; k <= K; k++) {
      const m = st.filter(s => s.cls === k), sc = m.filter(s => !s.extra && s.score != null).map(s => s.score).sort((a, b) => a - b);
      const mu = sc.length ? sc.reduce((a, b) => a + b, 0) / sc.length : null;
      const sd = sc.length ? Math.sqrt(sc.reduce((a, b) => a + (b - mu) ** 2, 0) / sc.length) : null;
      const med = sc.length ? (sc.length % 2 ? sc[(sc.length - 1) / 2] : (sc[sc.length / 2 - 1] + sc[sc.length / 2]) / 2) : null;
      const tiers = new Array(T).fill(0); m.forEach(s => { if (s.tier != null) tiers[s.tier]++; });
      const sch = {}; m.forEach(s => sch[s.school] = (sch[s.school] || 0) + 1);
      cls.push({ k, n: m.length, reg: m.filter(s => !s.extra).length, extra: m.filter(s => s.extra).length, male: m.filter(s => s.sex === '남').length, female: m.filter(s => s.sex === '여').length, mu, sd, med, min: sc[0], max: sc[sc.length - 1], tiers, sch });
    }
    const all = st.filter(s => !s.extra && s.score != null).map(s => s.score);
    const gmu = all.reduce((a, b) => a + b, 0) / (all.length || 1);
    const range = f => { const v = cls.map(f).filter(x => x != null); return v.length ? Math.max(...v) - Math.min(...v) : 0; };
    const schRange = {}; schools.forEach(s => schRange[s] = range(c => c.sch[s] || 0));
    return { cls, schools, gmu, gsd: Math.sqrt(all.reduce((a, b) => a + (b - gmu) ** 2, 0) / (all.length || 1)), range, schRange };
  }

  function checks() {
    const st = S.students, out = [];
    [...S.autoPairs, ...S.manualPairs].forEach(g => {
      const ids = g.ids, clash = [];
      for (let a = 0; a < ids.length; a++) for (let b = a + 1; b < ids.length; b++) if (st[ids[a]].cls === st[ids[b]].cls) clash.push([ids[a], ids[b]]);
      out.push({ type: g.type, text: ids.map(i => `${st[i].name}(${st[i].school}) ${st[i].cls}반`).join(' / '), ok: !clash.length });
    });
    Object.entries(S.fixedMap).forEach(([id, k]) => { const s = st[id]; out.push({ type: '반 지정', text: `${s.name}(${s.school}) → ${k}반, 결과 ${s.cls}반`, ok: s.cls === k }); });
    return out;
  }

  function numbering() {
    const mode = radio('numbering');
    const ord = s => mode === 'female' ? (s.sex === '여' ? 0 : 1) : mode === 'male' ? (s.sex === '남' ? 0 : 1) : 0;
    const by = {};
    S.students.forEach(s => (by[s.cls] = by[s.cls] || []).push(s));
    Object.values(by).forEach(a => { a.sort((x, y) => ord(x) - ord(y) || KO.compare(x.name, y.name) || x.id - y.id); a.forEach((s, i) => s.num = i + 1); });
    return by;
  }

  function tags(s) {
    const t = [];
    if (s.extra) t.push('정원 외');
    if (s.noScore) t.push('성적 없음');
    S.autoPairs.forEach(g => { if (g.ids.includes(s.id)) t.push(g.type); });
    S.manualPairs.forEach(g => { if (g.ids.includes(s.id)) t.push('분리 지정'); });
    if (S.fixedMap[s.id]) t.push('반 지정');
    if (s.cls !== s.autoCls) t.push(`직접 옮김(${s.autoCls}반→)`);
    return [...new Set(t)];
  }

  // ---------- 결과 화면 ----------
  function renderResult() {
    const X = stats(), K = S.K, T = S.T, ck = checks();
    numbering();
    const muR = X.range(c => c.mu), sexR = Math.max(X.range(c => c.male), X.range(c => c.female));
    const schMax = Math.max(0, ...Object.values(X.schRange)), bad = ck.filter(c => !c.ok).length;
    const sizeR = X.range(c => c.n);
    const chip = (k, v, cls) => `<div class="chip ${cls}"><div class="k">${k}</div><div class="v">${v}</div></div>`;
    const dirLow = radio('dir') === 'low';
    $('#summary').innerHTML =
      chip(Object.keys(S.adjust || {}).length ? '반 평균 성적 차이 (보정 전 원래 성적)' : '반 평균 성적 차이 (최고-최저)', fmt(muR), muR <= X.gsd * 0.05 ? 'ok' : muR <= X.gsd * 0.15 ? 'mid' : 'bad') +
      chip('남녀 수 반 간 차이', sexR + '명', sexR <= 1 ? 'ok' : sexR <= 2 ? 'mid' : 'bad') +
      chip('같은 중학교 반 간 최대 차이', schMax + '명', schMax <= 1 ? 'ok' : schMax <= 2 ? 'mid' : 'bad') +
      chip('나누기 조건', bad ? `${bad}건 어긋남` : `${ck.length}건 모두 지킴`, bad ? 'bad' : 'ok');

    // 통계 표
    const hi = (v, r, lim) => r > lim ? ' class="hi"' : '';
    const tierRange = Array.from({ length: T }, (_, t) => X.range(c => c.tiers[t]));
    let h = `<table><tr><th>반</th><th>총원</th><th>정원 외</th><th>남</th><th>여</th><th>평균</th><th>표준편차</th><th>중앙값</th><th>${dirLow ? '최상위(최소)' : '최상위(최대)'}</th><th>${dirLow ? '최하위(최대)' : '최하위(최소)'}</th>`;
    for (let t = 0; t < T; t++) h += `<th title="성적 구간 ${t + 1} (1이 최상위)">구간${t + 1}</th>`;
    X.schools.forEach(s => h += `<th>${esc(s)}</th>`);
    h += '</tr>';
    X.cls.forEach(c => {
      h += `<tr><td><b>${c.k}반</b></td><td${hi(0, sizeR, 1)}>${c.n}</td><td>${c.extra}</td><td${hi(0, X.range(x => x.male), 1)}>${c.male}</td><td${hi(0, X.range(x => x.female), 1)}>${c.female}</td><td>${fmt(c.mu)}</td><td>${fmt(c.sd)}</td><td>${fmt(c.med)}</td><td>${fmt(dirLow ? c.min : c.max)}</td><td>${fmt(dirLow ? c.max : c.min)}</td>`;
      c.tiers.forEach((v, t) => h += `<td${hi(0, tierRange[t], 1)}>${v}</td>`);
      X.schools.forEach(s => h += `<td${hi(0, X.schRange[s], 1)}>${c.sch[s] || 0}</td>`);
      h += '</tr>';
    });
    const tot = X.cls.reduce((a, c) => ({ n: a.n + c.n, extra: a.extra + c.extra, male: a.male + c.male, female: a.female + c.female }), { n: 0, extra: 0, male: 0, female: 0 });
    h += `<tr class="total"><td>전체</td><td>${tot.n}</td><td>${tot.extra}</td><td>${tot.male}</td><td>${tot.female}</td><td>${fmt(X.gmu)}</td><td>${fmt(X.gsd)}</td><td></td><td></td><td></td>`;
    for (let t = 0; t < T; t++) h += `<td>${X.cls.reduce((a, c) => a + c.tiers[t], 0)}</td>`;
    X.schools.forEach(s => h += `<td>${X.cls.reduce((a, c) => a + (c.sch[s] || 0), 0)}</td>`);
    h += `</tr><tr class="range"><td>반 간 차이</td><td>${sizeR}</td><td></td><td>${X.range(c => c.male)}</td><td>${X.range(c => c.female)}</td><td>${fmt(muR)}</td><td>${fmt(X.range(c => c.sd))}</td><td>${fmt(X.range(c => c.med))}</td><td></td><td></td>`;
    tierRange.forEach(r => h += `<td>${r}</td>`);
    X.schools.forEach(s => h += `<td>${X.schRange[s]}</td>`);
    h += '</tr></table>';
    $('#stats').innerHTML = h;

    $('#checks').innerHTML = ck.length ? `<ul class="msgs">${ck.map(c => `<li class="${c.ok ? 'ok' : 'bad'}">${c.ok ? '지킴' : '같은 반!'} · ${esc(c.type)}: ${esc(c.text)}</li>`).join('')}</ul>` : '<p class="hint">나눌 조건이 없어요.</p>';

    const moved = S.students.filter(s => s.cls !== s.autoCls).length;
    $('#resetMoves').hidden = !moved;
    $('#resetMoves').textContent = `직접 옮긴 ${moved}명 되돌리기`;
    renderTabs(X); renderRoster();
  }

  function renderTabs(X) {
    $('#tabs').innerHTML = X.cls.map(c => `<button role="tab" aria-selected="${c.k === S.tab}" data-k="${c.k}">${c.k}반<span class="c">${c.n}명</span></button>`).join('') + `<button role="tab" aria-selected="${S.tab === 0}" data-k="0">전체</button>`;
  }

  function renderRoster() {
    const K = S.K, show = $('#showScore').checked;
    const list = S.students.filter(s => S.tab === 0 || s.cls === S.tab).sort((a, b) => a.cls - b.cls || a.num - b.num);
    const opt = cur => Array.from({ length: K }, (_, i) => `<option value="${i + 1}" ${i + 1 === cur ? 'selected' : ''}>${i + 1}반</option>`).join('');
    let h = `<table><tr><th>반</th><th>번호</th><th>성명</th><th>성별</th><th>출신학교</th>${show ? '<th>성적</th><th>전체 순위</th><th>구간</th>' : ''}<th>비고</th><th>반 바꾸기</th></tr>`;
    list.forEach(s => {
      const t = tags(s), note = [s.note, ...t].filter(Boolean).join(', ');
      h += `<tr class="${s.cls !== s.autoCls ? 'moved' : t.length ? 'flag' : ''}"><td>${s.cls}</td><td>${s.num}</td><td>${esc(s.name)}</td><td>${esc(s.sex)}</td><td>${esc(s.school)}</td>${show ? `<td>${s.score == null ? '' : fmt(s.score)}</td><td>${s.rank || ''}</td><td>${s.tier == null ? '' : s.tier + 1}</td>` : ''}<td class="l">${esc(note)}</td><td><select data-id="${s.id}">${opt(s.cls)}</select></td></tr>`;
    });
    $('#roster').innerHTML = h + '</table>';
  }

  // ---------- 엑셀 ----------
  const BD = { style: 'thin', color: { rgb: 'A6A6A6' } };
  const ST_H = { font: { name: '맑은 고딕', sz: 10, bold: true }, fill: { fgColor: { rgb: 'DCE6F7' } }, alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: { top: BD, bottom: BD, left: BD, right: BD } };
  const ST_C = { font: { name: '맑은 고딕', sz: 10 }, alignment: { horizontal: 'center', vertical: 'center' }, border: { top: BD, bottom: BD, left: BD, right: BD } };
  const ST_L = { font: { name: '맑은 고딕', sz: 10 }, alignment: { horizontal: 'left', vertical: 'center' }, border: { top: BD, bottom: BD, left: BD, right: BD } };
  const ST_FLAG = Object.assign({}, ST_C, { fill: { fgColor: { rgb: 'FFF2CC' } } });
  const ST_TITLE = { font: { name: '맑은 고딕', sz: 13, bold: true } };
  const ST_NOTE = { font: { name: '맑은 고딕', sz: 9, color: { rgb: '595959' } } };

  function styleAll(ws, headerRows, opts = {}) {
    const ref = XLSX.utils.decode_range(ws['!ref']);
    for (let r = ref.s.r; r <= ref.e.r; r++) for (let c = ref.s.c; c <= ref.e.c; c++) {
      const a = XLSX.utils.encode_cell({ r, c }); if (!ws[a]) continue;
      if (opts.skipRows && opts.skipRows.includes(r)) continue;
      const base = headerRows.includes(r) ? ST_H : (opts.leftCols && opts.leftCols.includes(c)) ? ST_L : (opts.flagRows && opts.flagRows.has(r)) ? ST_FLAG : ST_C;
      ws[a].s = Object.assign({}, base, { numFmt: ws[a].z || (ws[a].t === 'n' && Number.isInteger(ws[a].v) ? '0' : 'General') });
    }
  }
  const qs = n => "'" + n.replace(/'/g, "''") + "'";

  function exportXlsx() {
    const X = stats(), K = S.K, T = S.T, by = numbering(), show = $('#showScore').checked, dirLow = radio('dir') === 'low';
    const wb = XLSX.utils.book_new();
    const sorted = S.students.slice().sort((a, b) => a.cls - b.cls || a.num - b.num);
    // 1. 반편성 결과
    const R1 = [['반', '번호', '성명', '성별', '출신학교', '성적', '전체 순위', '성적 구간', '수험번호', '비고']];
    const flag1 = new Set();
    sorted.forEach((s, i) => { const t = tags(s); if (t.length) flag1.add(i + 1); R1.push([s.cls, s.num, s.name, s.sex, s.school, s.score == null ? '' : s.score, s.rank || '', s.tier == null ? '' : s.tier + 1, s.exam, [s.note, ...t].filter(Boolean).join(', ')]); });
    const ws1 = XLSX.utils.aoa_to_sheet(R1);
    for (let r = 1; r < R1.length; r++) { const a = 'F' + (r + 1); if (ws1[a] && ws1[a].t === 'n') ws1[a].z = '0.00'; }
    styleAll(ws1, [0], { leftCols: [9], flagRows: flag1 });
    ws1['!cols'] = [5, 5, 10, 5, 11, 9, 8, 8, 9, 30].map(w => ({ wch: w }));
    ws1['!autofilter'] = { ref: 'A1:J' + R1.length };
    for (let r = 1; r < R1.length; r++) { const a = 'F' + (r + 1); if (ws1[a] && ws1[a].t === 'n') ws1[a].z = '0.00'; }
    XLSX.utils.book_append_sheet(wb, ws1, '반편성 결과');

    // 2. 반편성표 (반별 나란히)
    const maxN = Math.max(...Object.values(by).map(a => a.length));
    const R2 = [['번호'], ['']];
    for (let k = 1; k <= K; k++) { R2[0].push(k + '반', ''); R2[1].push('성명', '성별'); }
    for (let i = 0; i < maxN; i++) { const row = [i + 1]; for (let k = 1; k <= K; k++) { const s = (by[k] || [])[i]; row.push(s ? s.name : '', s ? s.sex : ''); } R2.push(row); }
    const ws2 = XLSX.utils.aoa_to_sheet(R2);
    ws2['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 1, c: 0 } }];
    for (let k = 0; k < K; k++) ws2['!merges'].push({ s: { r: 0, c: 1 + 2 * k }, e: { r: 0, c: 2 + 2 * k } });
    styleAll(ws2, [0, 1]);
    ws2['!cols'] = [{ wch: 5 }].concat(...Array.from({ length: K }, () => [{ wch: 9 }, { wch: 4.5 }]));
    XLSX.utils.book_append_sheet(wb, ws2, '반편성표');

    // 3. 반별 통계 (엑셀에서 반을 고치면 다시 계산되도록 수식 + 현재 값)
    const n1 = R1.length, rg = c => `${qs('반편성 결과')}!$${c}$2:$${c}$${n1}`;
    const head = ['반', '총원', '정원 외', '남', '여', '평균', '표준편차', dirLow ? '최상위(최소)' : '최상위(최대)', dirLow ? '최하위(최대)' : '최하위(최소)'];
    for (let t = 0; t < T; t++) head.push('구간' + (t + 1));
    X.schools.forEach(s => head.push(s));
    const ws3 = XLSX.utils.aoa_to_sheet([['반별 통계  (반편성 결과 시트의 [반] 값을 고치면 개수·평균은 엑셀에서 다시 계산돼요)'], [], head]);
    const put = (r, c, v, f, z) => { const a = XLSX.utils.encode_cell({ r, c }); ws3[a] = { t: typeof v === 'number' ? 'n' : 's', v: v == null ? '' : v }; if (f) ws3[a].f = f; if (z) ws3[a].z = z; };
    X.cls.forEach((c, i) => {
      const r = 3 + i, A = 'A' + (r + 1);
      put(r, 0, c.k);
      put(r, 1, c.n, `COUNTIF(${rg('A')},${A})`);
      put(r, 2, c.extra, `COUNTIFS(${rg('A')},${A},${rg('J')},"*정원 외*")`);
      put(r, 3, c.male, `COUNTIFS(${rg('A')},${A},${rg('D')},"남")`);
      put(r, 4, c.female, `COUNTIFS(${rg('A')},${A},${rg('D')},"여")`);
      put(r, 5, c.mu == null ? '' : +c.mu.toFixed(4), c.mu == null ? null : `AVERAGEIFS(${rg('F')},${rg('A')},${A})`, '0.00');
      put(r, 6, c.sd == null ? '' : +c.sd.toFixed(4), null, '0.00');
      put(r, 7, dirLow ? c.min : c.max, null, '0.00'); put(r, 8, dirLow ? c.max : c.min, null, '0.00');
      c.tiers.forEach((v, t) => put(r, 9 + t, v, `COUNTIFS(${rg('A')},${A},${rg('H')},${t + 1})`));
      X.schools.forEach((s, j) => put(r, 9 + T + j, c.sch[s] || 0, `COUNTIFS(${rg('A')},${A},${rg('E')},"${s.replace(/"/g, '""')}")`));
    });
    const rT = 3 + K, rD = rT + 1, lastC = 8 + T + X.schools.length;
    put(rT, 0, '전체'); put(rD, 0, '반 간 차이');
    for (let c = 1; c <= lastC; c++) {
      const L = XLSX.utils.encode_col(c), rngTxt = `${L}4:${L}${rT}`;
      if (c === 5) put(rT, c, +X.gmu.toFixed(4), null, '0.00');
      else if (c === 6) put(rT, c, +X.gsd.toFixed(4), null, '0.00');
      else if (c === 7 || c === 8) put(rT, c, '');
      else put(rT, c, X.cls.reduce((a, x) => a + (c <= 4 ? [0, x.n, x.extra, x.male, x.female][c] : c - 9 < T ? x.tiers[c - 9] : (x.sch[X.schools[c - 9 - T]] || 0)), 0), `SUM(${rngTxt})`);
      if (c !== 7 && c !== 8 && c !== 2) {
        const vals = X.cls.map(x => [null, x.n, x.extra, x.male, x.female, x.mu, x.sd][c] ?? (c >= 9 ? (c - 9 < T ? x.tiers[c - 9] : (x.sch[X.schools[c - 9 - T]] || 0)) : 0));
        put(rD, c, +(Math.max(...vals) - Math.min(...vals)).toFixed(4), `MAX(${rngTxt})-MIN(${rngTxt})`, c === 5 || c === 6 ? '0.00' : undefined);
      }
    }
    ws3['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rD + 4, c: lastC } });
    styleAll(ws3, [2], { skipRows: [0, 1, rD + 2, rD + 3, rD + 4] });
    ws3.A1.s = ST_TITLE;
    const notes = [`※ 성적 구간: 정원 내 학생을 성적순으로 ${T}구간으로 나눔 (구간1 = 최상위). 정원 외·성적 없는 학생은 성적 통계에서 빠짐.`,
      `※ 조건: 학급 ${K}개 · 성비 ${{ strong: '강하게', normal: '보통', off: '안 함' }[radio('sexLevel')]} · 출신교 ${{ strong: '강하게', normal: '보통', off: '안 함' }[radio('schoolLevel')]} · 성적 ${dirLow ? '낮을수록 우수' : '높을수록 우수'}${Object.keys(S.adjust || {}).length ? ' · 출신교 보정 ' + Object.entries(S.adjust).map(([k, v]) => k + ' ' + v).join(', ') : ''}`,
      `※ 만든 날: ${new Date().toLocaleString('ko-KR')} · 원본 파일: ${S.fileName}`];
    notes.forEach((t, i) => { const a = 'A' + (rD + 3 + i); ws3[a] = { t: 's', v: t, s: ST_NOTE }; });
    ws3['!cols'] = [{ wch: 10 }].concat(Array.from({ length: lastC }, (_, i) => ({ wch: i >= 4 && i <= 7 ? 10 : 7 })));
    XLSX.utils.book_append_sheet(wb, ws3, '반별 통계');

    // 4. 나누기 확인
    const ck = checks();
    const R4 = [['구분', '학생 (배정 반)', '결과']].concat(ck.map(c => [c.type, c.text, c.ok ? '지킴' : '같은 반 (확인 필요)']));
    const exL = S.students.filter(s => s.extra || s.noScore);
    if (exL.length) R4.push(['정원 외·성적 없음', exL.map(s => `${s.name}(${s.cls}반)`).join(', '), radio('extraMode') === 'gather' ? '한 반에 모음' : '나눠 배치']);
    const moved = S.students.filter(s => s.cls !== s.autoCls);
    if (moved.length) R4.push(['직접 옮김', moved.map(s => `${s.name}(${s.autoCls}반→${s.cls}반)`).join(', '), '']);
    if (R4.length === 1) R4.push(['', '나눌 조건이 없어요', '']);
    const ws4 = XLSX.utils.aoa_to_sheet(R4);
    styleAll(ws4, [0], { leftCols: [1] });
    ws4['!cols'] = [{ wch: 16 }, { wch: 70 }, { wch: 18 }];
    XLSX.utils.book_append_sheet(wb, ws4, '나누기 확인');

    // 5. 원본 + 배정 반
    const main = S.sheets[S.mainName];
    const R5 = [main.raw.concat(['배정 반', '반 번호'])];
    S.students.filter(s => s.sheet === S.mainName).forEach(s => R5.push(main.raw.map((_, c) => s.row[c] == null ? '' : s.row[c]).concat([s.cls, s.num])));
    const hdrRows = [0];
    if (S.extraName) {
      const ex = S.sheets[S.extraName];
      R5.push([]); hdrRows.push(R5.length); R5.push(ex.raw.concat(['배정 반', '반 번호']));
      S.students.filter(s => s.sheet === S.extraName && s.extra).forEach(s => R5.push(ex.raw.map((_, c) => s.row[c] == null ? '' : s.row[c]).concat([s.cls, s.num])));
    }
    const ws5 = XLSX.utils.aoa_to_sheet(R5);
    styleAll(ws5, hdrRows);
    ws5['!cols'] = R5[0].map(() => ({ wch: 12 }));
    XLSX.utils.book_append_sheet(wb, ws5, '원본+배정반');

    // 6. 반별 명단
    for (let k = 1; k <= K; k++) {
      const hdr = ['번호', '성명', '성별', '출신학교'].concat(show ? ['성적'] : [], ['비고']);
      const R = [[`${k}반 명단`], hdr];
      const flag = new Set();
      (by[k] || []).forEach((s, i) => { const note = [s.note, ...tags(s).filter(t => !/^직접 옮김/.test(t))].filter(Boolean).join(', '); if (note) flag.add(i + 2); R.push([s.num, s.name, s.sex, s.school].concat(show ? [s.score == null ? '' : s.score] : [], [note])); });
      const ws = XLSX.utils.aoa_to_sheet(R);
      styleAll(ws, [1], { skipRows: [0], leftCols: [hdr.length - 1], flagRows: flag });
      ws.A1.s = ST_TITLE;
      ws['!cols'] = [6, 10, 6, 11].concat(show ? [9] : [], [26]).map(w => ({ wch: w }));
      XLSX.utils.book_append_sheet(wb, ws, `${k}반`);
    }
    const base = S.fileName.replace(/\.[^.]+$/, '');
    XLSX.writeFile(wb, `반편성 결과(${K}학급) - ${base}.xlsx`);
  }

  function template() {
    const wb = XLSX.utils.book_new();
    const a = XLSX.utils.aoa_to_sheet([['순번', '수험번호', '성명', '성별', '출신학교', '학년석차 백분율(%)', '보호자 연락처', '비고'], [1, 1, '홍길동', '남', '○○중', 12.5, '010-0000-0000', ''], [2, 2, '김영희', '여', '△△중', 33.4, '010-1111-1111', '']]);
    styleAll(a, [0]); a['!cols'] = [6, 8, 10, 6, 12, 16, 16, 20].map(w => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, a, '학생 명단');
    const b = XLSX.utils.aoa_to_sheet([['순번', '성명', '성별', '출신학교', '보호자 연락처', '비고'], [1, '이철수', '남', '○○중', '010-2222-2222', '특수교육대상자']]);
    styleAll(b, [0]); b['!cols'] = [6, 10, 6, 12, 16, 20].map(w => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, b, '정원 외');
    XLSX.writeFile(wb, '반편성 명단 양식.xlsx');
  }

  function demo() {
    let seed = 20261117; const R = () => { seed = (seed * 1103515245 + 12345) % 2147483648; return seed / 2147483648; };
    const pick = a => a[Math.floor(R() * a.length)];
    const sur = ['김', '이', '박', '최', '정', '강', '조', '윤', '장', '임', '한', '오', '서', '신', '권', '황', '안', '송', '류', '홍'];
    const g1 = ['서', '민', '지', '하', '도', '수', '예', '시', '유', '주', '은', '현', '준', '재', '태', '나', '소', '채', '다', '승'];
    const g2 = ['연', '준', '우', '윤', '아', '진', '빈', '원', '현', '린', '율', '호', '은', '희', '성', '훈', '경', '민', '영', '온'];
    const schools = [['가람중', 50, 0], ['나래중', 38, 6], ['다솜중', 30, -5], ['라온중', 26, 3], ['마루중', 20, -3], ['바다중', 14, 0], ['사랑중', 6, 4], ['검정고시', 1, 0]];
    const rows = [['순번', '수험번호', '성명', '성별', '출신학교', '학년석차 백분율(%)', '보호자 연락처', '비고']];
    const tel = () => '010-' + String(1000 + Math.floor(R() * 9000)) + '-' + String(1000 + Math.floor(R() * 9000));
    let n = 0;
    schools.forEach(([sc, cnt, off]) => { for (let i = 0; i < cnt; i++) { n++; const p = Math.min(99.9, Math.max(0.3, R() * 100 + off * (R() - 0.3))); rows.push([n, n, pick(sur) + pick(g1) + pick(g2), R() < 0.5 ? '남' : '여', sc, Math.round(p * 100) / 100, tel(), '']); } });
    rows[12][2] = rows[80][2]; rows[40][2] = rows[150][2];          // 동명이인
    rows[21][6] = rows[22][6]; rows[22][2] = rows[21][2][0] + '다' + '은'; // 형제자매
    rows[100][6] = rows[133][6]; rows[133][2] = rows[100][2][0] + '하' + '린';
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), '접수대장(예시)');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['순번', '성명', '성별', '출신학교', '보호자 연락처', '비고'], [1, '노하준', '남', '가람중', tel(), '특수교육대상자'], [2, '문서아', '여', '다솜중', tel(), '특수교육대상자'], [3, '배지호', '남', '나래중', tel(), '특수교육대상자']]), '정원 외');
    loadWorkbook(wb, '예시 자료(가상 학생).xlsx');
  }

  // ---------- 이벤트 ----------
  const drop = $('#drop');
  $('#file').addEventListener('change', e => { const f = e.target.files[0]; if (f) loadFile(f); e.target.value = ''; });
  ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', e => { const f = e.dataTransfer.files[0]; if (f) loadFile(f); });
  $('#tplBtn').addEventListener('click', template);
  $('#demoBtn').addEventListener('click', demo);
  $('#mainSheet').addEventListener('change', e => { setMain(e.target.value); });
  $('#extraSheet').addEventListener('change', e => { S.extraName = e.target.value; rebuild(); });
  $('#colMap').addEventListener('change', e => { const k = e.target.dataset.k; if (!k) return; S.map[k] = +e.target.value; renderColMap(); rebuild(); });
  $('#K').addEventListener('input', () => { S._kTouched = true; updatePerClass(); });
  ['sameName', 'sibling'].forEach(id => $('#' + id).addEventListener('change', refreshAutoPairs));
  document.querySelectorAll('input[name=extraMode]').forEach(r => r.addEventListener('change', () => { $('#gatherBox').hidden = radio('extraMode') !== 'gather'; }));
  ['manualPairs', 'fixed', 'adjust', 'targets'].forEach(id => $('#' + id).addEventListener('change', () => { if (S.students.length) parseManual(); }));
  $('#runBtn').addEventListener('click', () => run(false));
  $('#againBtn').addEventListener('click', () => run(true));
  $('#xlsxBtn').addEventListener('click', exportXlsx);
  $('#resetMoves').addEventListener('click', () => { S.students.forEach(s => s.cls = s.autoCls); renderResult(); });
  $('#tabs').addEventListener('click', e => { const b = e.target.closest('button'); if (!b) return; S.tab = +b.dataset.k; renderTabs(stats()); renderRoster(); });
  $('#roster').addEventListener('change', e => { const id = e.target.dataset.id; if (id == null) return; S.students[+id].cls = +e.target.value; renderResult(); });
  document.querySelectorAll('input[name=numbering]').forEach(r => r.addEventListener('change', () => { if (S.result) renderResult(); }));
  $('#showScore').addEventListener('change', () => { if (S.result) renderRoster(); });
  $('#calibFile').addEventListener('change', e => { const f = e.target.files[0]; if (f) loadCalibFile(f); e.target.value = ''; });
  $('#calibTpl').addEventListener('click', calibTemplate);
  $('#calibDemo').addEventListener('click', calibDemo);
  $('#calibOut').addEventListener('click', e => { if (e.target.id === 'calibApply') applyCalib(); });
  document.querySelectorAll('input[name=calibDir], input[name=dir]').forEach(r => r.addEventListener('change', () => { if (S.calib) renderCalib(); }));
  window.addEventListener('beforeunload', e => { if (S.result) { e.preventDefault(); e.returnValue = ''; } });
})();




