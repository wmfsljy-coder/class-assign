/* 반편성 배정 엔진. 브라우저(Web Worker)와 node 양쪽에서 쓴다. 외부 전송 없음. */
(function (root) {
  'use strict';

  function rng(seed) {
    let s = seed >>> 0;
    return function () { s = (s + 0x6D2B79F5) >>> 0; let t = s; t = Math.imul(t ^ (t >>> 15), t | 1); t ^= t + Math.imul(t ^ (t >>> 7), t | 61); return ((t ^ (t >>> 14)) >>> 0) / 4294967296; };
  }

  const LEVEL = {
    sex: { off: [0, 0], normal: [10, 30], strong: [20, 200] },
    school: { off: [0, 0], normal: [4, 20], strong: [10, 100] }
  };

  /**
   * input = {
   *   K, tiers,
   *   students: [{ key, score|null (낮을수록 우수로 맞춘 값), sex, school, fixed: 1..K|null, extra: bool }],
   *   pairs: [[i, j]],            // 같은 반 금지 (동명이인, 형제자매, 직접 입력)
   *   sexLevel, schoolLevel,      // 'off' | 'normal' | 'strong'
   *   extraMode: 'spread' | 'gather', gatherClass,
   *   restarts, iterations, seed
   * }
   * 반환: { cls: [1..K], cost, ranks: [], tier: [] }
   */
  function assign(input, onProgress) {
    const K = input.K, ST = input.students, N = ST.length;
    const wSex = LEVEL.sex[input.sexLevel || 'strong'], wSch = LEVEL.school[input.schoolLevel || 'strong'];

    // 점수 있는 학생 순위·구간
    const scored = [];
    for (let i = 0; i < N; i++) if (ST[i].score != null && !ST[i].extra) scored.push(i);
    scored.sort((a, b) => ST[a].score - ST[b].score || a - b);
    const rank = new Array(N).fill(null), tier = new Array(N).fill(null);
    const T = Math.max(1, Math.min(input.tiers || 10, scored.length));
    scored.forEach((i, r) => { rank[i] = r + 1; tier[i] = Math.floor(r * T / scored.length); });

    // z 점수
    let mu = 0; scored.forEach(i => mu += ST[i].score); mu /= Math.max(1, scored.length);
    let sd = 0; scored.forEach(i => sd += (ST[i].score - mu) ** 2); sd = Math.sqrt(sd / Math.max(1, scored.length)) || 1;
    const z = ST.map(s => (s.score == null || s.extra) ? 0 : (s.score - mu) / sd);

    // 학교 색인
    const schools = []; const schIdx = ST.map(s => { let k = schools.indexOf(s.school); if (k < 0) { schools.push(s.school); k = schools.length - 1; } return k; });
    const S = schools.length;
    const male = ST.map(s => s.sex === '남' ? 1 : 0);
    const maleTot = male.reduce((a, b) => a + b, 0);
    const schTot = new Array(S).fill(0); schIdx.forEach(k => schTot[k]++);

    // 금지 쌍 인접 목록
    const adj = Array.from({ length: N }, () => []);
    (input.pairs || []).forEach(([a, b]) => { if (a !== b) { adj[a].push(b); adj[b].push(a); } });

    // ---- 초기 배치 ----
    function initial() {
      const cls = new Array(N).fill(0);
      const size = new Array(K + 1).fill(0);
      const free = [];
      for (let i = 0; i < N; i++) {
        const f = ST[i].fixed;
        if (f && f >= 1 && f <= K) { cls[i] = f; size[f]++; }
      }
      // 성적 없음/정원 외
      const extras = [];
      for (let i = 0; i < N; i++) if (!cls[i] && (ST[i].extra || ST[i].score == null)) extras.push(i);
      if (input.extraMode === 'gather') {
        const g = Math.min(K, Math.max(1, input.gatherClass || 1));
        extras.forEach(i => { cls[i] = g; size[g]++; });
      } else {
        const tg = (Array.isArray(input.targets) && input.targets.length === K) ? input.targets : null;
        const room = k => tg ? tg[k - 1] - size[k] : -size[k];
        extras.forEach(i => { let best = 1; for (let k = 2; k <= K; k++) if (room(k) > room(best)) best = k; cls[i] = best; size[best]++; });
      }
      // 반별 목표 인원: 직접 정했으면 그대로, 아니면 전체를 고르게
      const target = new Array(K + 1).fill(0);
      if (Array.isArray(input.targets) && input.targets.length === K) {
        for (let k = 1; k <= K; k++) target[k] = Math.max(0, input.targets[k - 1] | 0);
      } else {
        const base = Math.floor(N / K); let rem = N - base * K;
        const byLoad = Array.from({ length: K }, (_, k) => k + 1).sort((a, b) => size[b] - size[a] || a - b);
        target.fill(base, 1); byLoad.forEach(k => { if (rem > 0) { target[k]++; rem--; } });
      }
      // 성적 있는 학생 자리 = 목표 - (성적 없는 학생 + 성적 없이 고정된 학생)
      const preNoScore = new Array(K + 1).fill(0);
      for (let i = 0; i < N; i++) if (cls[i] && rank[i] == null) preNoScore[cls[i]]++;
      const cap0 = target.map((t, k) => Math.max(0, t - preNoScore[k]));
      let capSum = 0; for (let k = 1; k <= K; k++) capSum += cap0[k];
      let short = scored.length - capSum, kk = 1;
      while (short > 0) { cap0[kk]++; short--; kk = kk % K + 1; }
      for (let k = 1; k <= K; k++) if (cap0[k] === 0) cap0[k] = 1e-9;
      // 성적순으로 차례대로: 자리 대비 덜 찬 반부터 (같으면 지그재그 순서)
      const got = new Array(K + 1).fill(0);
      scored.forEach((i, pos) => {
        if (cls[i]) { got[cls[i]]++; return; } // 고정 학생은 그 자리에서 셈
        const pass = Math.floor(pos / K), asc = pass % 2 === 0;
        let best = -1, bestR = Infinity;
        for (let t = 0; t < K; t++) {
          const k = asc ? t + 1 : K - t;
          if (got[k] >= cap0[k] - 1e-12) continue;
          const r = (got[k] + 0.5) / cap0[k];
          if (r < bestR - 1e-12) { bestR = r; best = k; }
        }
        if (best < 0) { best = 1; for (let k = 2; k <= K; k++) if (got[k] / cap0[k] < got[best] / cap0[best]) best = k; }
        cls[i] = best; got[best]++; size[best]++;
      });
      return cls;
    }

    const tierTot = new Array(T).fill(0); scored.forEach(i => tierTot[tier[i]]++);
    // ---- 비용 ----
    function makeState(cls) {
      const st = { n: new Float64Array(K + 1), sz: new Float64Array(K + 1), sum: new Float64Array(K + 1), sq: new Float64Array(K + 1), m: new Float64Array(K + 1), sc: Array.from({ length: K + 1 }, () => new Float64Array(S)), tc: Array.from({ length: K + 1 }, () => new Float64Array(T)) };
      for (let i = 0; i < N; i++) { const k = cls[i]; st.sz[k]++; st.m[k] += male[i]; st.sc[k][schIdx[i]]++; if (rank[i] != null) { st.n[k]++; st.sum[k] += z[i]; st.sq[k] += z[i] * z[i]; st.tc[k][tier[i]]++; } }
      return st;
    }
    function classCost(st, k) {
      let c = 0;
      if (st.n[k] > 0) { const m = st.sum[k] / st.n[k]; const v = Math.max(0, st.sq[k] / st.n[k] - m * m); c += 25000 * m * m + 2500 * (Math.sqrt(v) - 1) ** 2;
        const tc = st.tc[k], share = st.n[k] / scored.length; for (let t = 0; t < T; t++) { const d = tc[t] - tierTot[t] * share; c += 40 * d * d + (Math.abs(d) >= 1 ? 80 : 0); } }
      const sz = st.sz[k];
      if (wSex[0]) { const t = sz * maleTot / N, d = st.m[k] - t; c += wSex[0] * d * d + (Math.abs(d) >= 1 ? wSex[1] : 0); }
      if (wSch[0]) { const sc = st.sc[k]; for (let s = 0; s < S; s++) { const t = schTot[s] * sz / N, d = sc[s] - t; c += wSch[0] * d * d + (Math.abs(d) >= 1 ? wSch[1] : 0); } }
      return c;
    }
    function totalCost(cls, st) {
      let c = 0; for (let k = 1; k <= K; k++) c += classCost(st, k);
      for (let i = 0; i < N; i++) for (const j of adj[i]) if (j > i && cls[i] === cls[j]) c += 10000;
      return c;
    }
    function moveOut(st, i, k) { st.m[k] -= male[i]; st.sc[k][schIdx[i]]--; if (rank[i] != null) { st.n[k]--; st.sum[k] -= z[i]; st.sq[k] -= z[i] * z[i]; st.tc[k][tier[i]]--; } }
    function moveIn(st, i, k) { st.m[k] += male[i]; st.sc[k][schIdx[i]]++; if (rank[i] != null) { st.n[k]++; st.sum[k] += z[i]; st.sq[k] += z[i] * z[i]; st.tc[k][tier[i]]++; } }
    function conflictsOf(cls, i, k, skip) { let c = 0; for (const j of adj[i]) if (j !== skip && cls[j] === k) c++; return c; }

    // 교환 후보 묶음: 같은 성적 구간끼리, 성적 없음 학생끼리
    const groups = [];
    const byTier = Array.from({ length: T }, () => []);
    const extraFree = [];
    for (let i = 0; i < N; i++) {
      if (ST[i].fixed) continue;
      if (rank[i] != null) byTier[tier[i]].push(i);
      else if (input.extraMode !== 'gather') extraFree.push(i);
    }
    byTier.forEach(g => { if (g.length > 1) groups.push(g); });
    const allScoredFree = [].concat(...byTier);
    if (extraFree.length > 1) groups.push(extraFree);

    const restarts = input.restarts || 6, iters = input.iterations || 300000;
    const R = rng(input.seed || 12345);
    let best = null, bestC = Infinity;
    for (let r = 0; r < restarts; r++) {
      const cls = initial();
      if (!groups.length) { best = cls; bestC = 0; break; }
      const st = makeState(cls);
      let cur = totalCost(cls, st), runBest = cls.slice(), runBestC = cur;
      for (let it = 0; it < iters; it++) {
        const temp = 50 * Math.pow(0.0004, it / iters);
        const g = (allScoredFree.length > 1 && R() < 0.25) ? allScoredFree : groups[Math.floor(R() * groups.length)];
        const i = g[Math.floor(R() * g.length)], j = g[Math.floor(R() * g.length)];
        const a = cls[i], b = cls[j];
        if (a === b) continue;
        const before = classCost(st, a) + classCost(st, b) + 10000 * (conflictsOf(cls, i, a, j) + conflictsOf(cls, j, b, i));
        moveOut(st, i, a); moveOut(st, j, b); moveIn(st, i, b); moveIn(st, j, a); cls[i] = b; cls[j] = a;
        const after = classCost(st, a) + classCost(st, b) + 10000 * (conflictsOf(cls, i, b, j) + conflictsOf(cls, j, a, i));
        const d = after - before;
        if (d <= 0 || R() < Math.exp(-d / temp)) { cur += d; if (cur < runBestC - 1e-9) { runBestC = cur; runBest = cls.slice(); } }
        else { moveOut(st, i, b); moveOut(st, j, a); moveIn(st, i, a); moveIn(st, j, b); cls[i] = a; cls[j] = b; }
      }
      runBestC = totalCost(runBest, makeState(runBest));
      if (runBestC < bestC) { bestC = runBestC; best = runBest; }
      if (onProgress) onProgress((r + 1) / restarts);
    }
    return { cls: best, cost: bestC, rank, tier, tiers: T };
  }

  const api = { assign };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.ClassEngine = api;
  // Web Worker
  if (typeof self !== 'undefined' && typeof window === 'undefined' && typeof importScripts === 'function') {
    self.onmessage = function (e) {
      try {
        const out = assign(e.data, p => self.postMessage({ type: 'progress', p }));
        self.postMessage({ type: 'done', out });
      } catch (err) { self.postMessage({ type: 'error', message: String(err && err.message || err) }); }
    };
  }
})(typeof self !== 'undefined' ? self : this);

