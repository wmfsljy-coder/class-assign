/* 신입생 반편성 도우미 - 지난 입학생 자료로 출신 중학교별 성적 보정값 계산.
   브라우저 안에서만 계산한다. node 에서도 불러 시험할 수 있다. */
(function (root) {
  'use strict';

  // 같은 해 학생끼리 등수를 백분율로 바꾼다. 0 에 가까울수록 우수. 같은 값은 평균 등수.
  function pctRank(vals, low) {
    const idx = vals.map((v, i) => i).sort((a, b) => low ? vals[a] - vals[b] : vals[b] - vals[a]);
    const n = vals.length, out = new Array(n);
    for (let i = 0; i < n;) {
      let j = i;
      while (j + 1 < n && vals[idx[j + 1]] === vals[idx[i]]) j++;
      const r = (i + j) / 2 + 0.5;               // 0.5 ~ n-0.5
      for (let k = i; k <= j; k++) out[idx[k]] = r / n * 100;
      i = j + 1;
    }
    return out;
  }

  function ols(x, y) {
    const n = x.length; let mx = 0, my = 0;
    for (let i = 0; i < n; i++) { mx += x[i]; my += y[i]; }
    mx /= n; my /= n;
    let sxx = 0, sxy = 0, syy = 0;
    for (let i = 0; i < n; i++) { const dx = x[i] - mx, dy = y[i] - my; sxx += dx * dx; sxy += dx * dy; syy += dy * dy; }
    const b = sxx ? sxy / sxx : 0;
    return { a: my - b * mx, b, r: sxx && syy ? sxy / Math.sqrt(sxx * syy) : 0, mx, my, sxx };
  }

  /**
   * records: [{ year, school, mid, high }]  (mid/high 는 원래 숫자, school 은 정리된 이름)
   * opts: { midLow: 중학교 성적이 낮을수록 우수인가, highLow: 고교 성적이 낮을수록 우수인가, cap: 보정 상한(백분율 등수 단위, 기본 15) }
   * 결과의 value 는 반편성 화면 보정 칸에 그대로 넣는 값(중학교 성적 단위, + 는 고교에서 더 잘함).
   */
  function compute(records, opts = {}) {
    const midLow = opts.midLow !== false, highLow = opts.highLow !== false, cap = opts.cap || 15;
    const recs = records.filter(r => r && r.school && isFinite(r.mid) && isFinite(r.high) && r.mid !== null && r.high !== null);
    const warnings = [];
    if (recs.length < 30) return { ok: false, error: `계산할 학생이 ${recs.length}명뿐이에요. 적어도 30명(가능하면 2~3년치)이 있어야 해요.`, N: recs.length };

    // 해마다 따로 등수 백분율
    const byYear = new Map();
    recs.forEach(r => { const y = String(r.year == null ? '' : r.year); if (!byYear.has(y)) byYear.set(y, []); byYear.get(y).push(r); });
    const years = [];
    const pts = [];
    byYear.forEach((list, y) => {
      if (list.length < 10) { warnings.push(`${y || '(연도 없음)'}: ${list.length}명뿐이라 계산에서 뺐어요.`); return; }
      const mp = pctRank(list.map(r => r.mid), midLow), hp = pctRank(list.map(r => r.high), highLow);
      list.forEach((r, i) => pts.push({ year: y, school: r.school, mid: r.mid, mp: mp[i], hp: hp[i] }));
      years.push({ year: y, n: list.length });
    });
    if (pts.length < 30) return { ok: false, error: '해마다 10명 이상인 자료가 30명이 안 돼요.', N: pts.length, warnings };
    years.sort((a, b) => String(a.year).localeCompare(String(b.year), 'ko', { numeric: true }));

    // 중학교 등수로 고교 등수를 예측 (모든 해 합쳐서 직선 하나)
    const fit = ols(pts.map(p => p.mp), pts.map(p => p.hp));
    if (fit.b < 0.2) return { ok: false, error: `중학교 성적과 고1 성적의 관련이 너무 약해요(상관 ${fit.r.toFixed(2)}). 고1 성적의 '낮을수록/높을수록 우수'가 맞는지 확인해 주세요.`, N: pts.length, r: fit.r, warnings };
    let sse = 0;
    pts.forEach(p => { p.res = (fit.a + fit.b * p.mp) - p.hp; sse += p.res * p.res; });   // + 면 예상보다 고교에서 잘함
    const sigma = Math.sqrt(sse / Math.max(1, pts.length - 2));

    // 원래 중학교 성적 단위로 바꾸는 비율 (등수 백분율 1 이 원래 성적 몇 점인지)
    const conv = ols(pts.map(p => p.mp), pts.map(p => p.mid));
    const unit = Math.abs(conv.b) || 1;

    // 학교별 평균 잔차
    const bySch = new Map();
    pts.forEach(p => { if (!bySch.has(p.school)) bySch.set(p.school, []); bySch.get(p.school).push(p); });
    const schools = [];
    bySch.forEach((list, school) => {
      const n = list.length, mean = list.reduce((a, p) => a + p.res, 0) / n;
      const yearMeans = years.map(y => { const l = list.filter(p => p.year === y.year); return l.length >= 3 ? l.reduce((a, p) => a + p.res, 0) / l.length : null; }).filter(v => v != null);
      schools.push({ school, n, mean, se: sigma / Math.sqrt(n), yearMeans });
    });

    // 학교 사이 진짜 차이의 크기(tau²) = 학교 평균들의 흩어짐 - 우연으로 생기는 흩어짐
    const big = schools.filter(s => s.n >= 2);
    let tau2 = 0;
    if (big.length >= 2) {
      const W = big.reduce((a, s) => a + s.n, 0), gm = big.reduce((a, s) => a + s.n * s.mean, 0) / W;
      const between = big.reduce((a, s) => a + (s.mean - gm) ** 2, 0) / (big.length - 1);
      const noise = big.reduce((a, s) => a + s.se * s.se, 0) / big.length;
      tau2 = Math.max(0, between - noise);
    }

    schools.forEach(s => {
      s.shrink = tau2 > 0 ? tau2 / (tau2 + s.se * s.se) : 0;
      let eff = s.mean * s.shrink;                      // 고교 등수 백분율 단위
      eff = eff / fit.b;                                // 같은 고교 성적이 되려면 중학교 등수가 얼마나 달라야 하나
      eff = Math.max(-cap, Math.min(cap, eff));
      s.value = Math.round(eff * unit * 10) / 10;       // 중학교 성적 단위
      s.z = s.se ? s.mean / s.se : 0;
      s.yearsTotal = s.yearMeans.length;
      s.yearsSame = s.yearMeans.filter(v => Math.sign(v) === Math.sign(s.mean) && v !== 0).length;
      s.verdict = s.n < 10 ? 'few' : (Math.abs(s.z) >= 2.5 && tau2 > 0 && Math.abs(s.value) >= 0.5 && (s.yearsTotal < 2 || s.yearsSame >= Math.ceil(s.yearsTotal * 2 / 3))) ? 'clear' : 'weak';
    });
    schools.sort((a, b) => b.n - a.n || a.school.localeCompare(b.school, 'ko'));

    return { ok: true, N: pts.length, years, b: fit.b, r: fit.r, sigma, tau: Math.sqrt(tau2), unit, schools, warnings };
  }

  const api = { compute, pctRank };
  root.Calib = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof self !== 'undefined' ? self : this);
