/* =========================================================
   VOAR AI 광고 대시보드 · Vanilla JS
   ---------------------------------------------------------
   데이터 흐름
     1) 사용자가 CPC 광고 XLSX 파일을 업로드
     2) SheetJS로 파싱 → 광고그룹×키워드 단위 집계
     3) 브랜드 / 기간 필터 적용
     4) AI 룰 기반 분석 → 7개 섹션 렌더
     5) Chart.js로 상위 그룹의 ROAS/CTR/CVR 시각화
   ---------------------------------------------------------
   XLSX 컬럼 (기존 index.html 파서와 동일)
     캠페인명 / 광고그룹 / 광고집행 상품명 / 캠페인 ID / 광고유형 / 과금 방식
     광고 노출 지면 / 키워드 / 노출수 / 클릭수 / 광고비
     직접 판매수량(1일) / 직접 전환매출액(1일)
     총 판매수량(1일) / 총 전환매출액(1일)
   ========================================================= */

'use strict';

// ---------- 상태 ----------
const state = {
  rows: [],          // 파일에서 파싱한 원본 rows (날짜 포함)
  dates: new Set(),  // 데이터가 있는 날짜들
  period: 7,         // 최근 N일 필터 (사용자가 다중 파일 업로드 시 유효)
  brand: '전체',
  chartMetric: 'roas',
};

let mainChart = null;

// ---------- 유틸 ----------
const toNum = (v) => {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return v;
  const n = Number(String(v).replace(/[,\s]/g, ''));
  return isNaN(n) ? 0 : n;
};
const fmtInt = (n) => Math.round(n).toLocaleString('ko-KR');
const fmtWon = (n) => Math.round(n).toLocaleString('ko-KR') + '원';
const fmtPct = (n, d = 1) => (n == null || isNaN(n)) ? '-' : n.toFixed(d) + '%';
const fmt1 = (n, d = 1) => (n == null || isNaN(n)) ? '-' : n.toFixed(d);

/** 키워드 정규화: 공백 제거·소문자화. 같은 키워드 변형 통합용 */
const normKw = (kw) => (kw || '').trim().toLowerCase().replace(/\s+/g, '');

/** 브랜드 텍스트 분류 (매칭 SKU 없이 blob 검사) */
function classifyBrand(blob) {
  if (/보아르|voar/i.test(blob)) return '보아르';
  if (/오아|OA-|OOR/i.test(blob)) return '오아';
  return '기타';
}

/** 파일명에서 날짜 추출 (..._YYYYMMDD_YYYYMMDD.xlsx) */
function extractDate(fileName) {
  const m = fileName.match(/(\d{8})_\d{8}\.xlsx?$/i) || fileName.match(/(\d{8})/);
  if (!m) return null;
  const s = m[1];
  return `${s.slice(0,4)}-${s.slice(4,6)}-${s.slice(6,8)}`;
}

// =========================================================
// 1) XLSX 파싱
// =========================================================
/** 워크북 → 표준 rows */
function parseWorkbook(wb, fileName) {
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const raws = XLSX.utils.sheet_to_json(sheet, { defval: null, raw: true });
  const date = extractDate(fileName) || new Date().toISOString().slice(0, 10);
  const out = [];
  for (const r of raws) {
    const impressions = toNum(r['노출수']);
    const clicks = toNum(r['클릭수']);
    const adCost = toNum(r['광고비']);
    // 완전 빈 행 제외
    if (impressions === 0 && clicks === 0 && adCost === 0) continue;
    out.push({
      date,
      campaignName: r['캠페인명'] || '',
      adGroup: r['광고그룹'] || '',
      productName: r['광고집행 상품명'] || '',
      placement: r['광고 노출 지면'] || '',
      keyword: r['키워드'] || '',
      keywordNorm: normKw(r['키워드']),
      chargeMethod: r['과금 방식'] || '',
      adType: r['광고유형'] || '',
      impressions,
      clicks,
      adCost,
      directQty: toNum(r['직접 판매수량(1일)']),
      directRevenue: toNum(r['직접 전환매출액(1일)']),
      totalQty: toNum(r['총 판매수량(1일)']),
      totalRevenue: toNum(r['총 전환매출액(1일)']),
    });
  }
  return out;
}

/** File 객체를 읽어 rows 반환 */
async function readFile(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array' });
  return parseWorkbook(wb, file.name);
}

// =========================================================
// 2) 집계: 광고그룹 × 키워드 단위
// =========================================================
/** 브랜드 필터 통과 여부 */
function passBrand(row) {
  if (state.brand === '전체') return true;
  const blob = `${row.campaignName} ${row.adGroup} ${row.productName}`;
  return classifyBrand(blob) === state.brand;
}

/** rows → 광고그룹×키워드 집계된 배열 */
function aggregate() {
  const map = new Map();
  for (const r of state.rows) {
    if (!passBrand(r)) continue;
    const key = `${r.adGroup}||${r.keywordNorm || '(자동노출)'}`;
    let g = map.get(key);
    if (!g) {
      g = {
        key,
        adGroup: r.adGroup,
        keyword: r.keyword || '(자동노출)',
        keywordNorm: r.keywordNorm,
        campaignName: r.campaignName,
        productName: r.productName,
        impressions: 0, clicks: 0, adCost: 0,
        directQty: 0, directRevenue: 0,
        totalQty: 0, totalRevenue: 0,
      };
      map.set(key, g);
    }
    g.impressions += r.impressions;
    g.clicks += r.clicks;
    g.adCost += r.adCost;
    g.directQty += r.directQty;
    g.directRevenue += r.directRevenue;
    g.totalQty += r.totalQty;
    g.totalRevenue += r.totalRevenue;
  }
  // 파생 지표 계산
  const list = [...map.values()].map(g => {
    const ctr = g.impressions > 0 ? g.clicks / g.impressions * 100 : 0;
    const cpc = g.clicks > 0 ? g.adCost / g.clicks : 0;
    const cvr = g.clicks > 0 ? g.directQty / g.clicks * 100 : 0;
    const roas = g.adCost > 0 ? g.directRevenue / g.adCost * 100 : 0;
    return { ...g, ctr, cpc, cvr, roas };
  });
  return list;
}

/** 광고그룹 단위 집계 (상품별 점수용) */
function aggregateByGroup(rows) {
  const map = new Map();
  for (const r of rows) {
    let g = map.get(r.adGroup);
    if (!g) {
      g = {
        adGroup: r.adGroup,
        impressions: 0, clicks: 0, adCost: 0,
        directQty: 0, directRevenue: 0,
        totalQty: 0, totalRevenue: 0,
      };
      map.set(r.adGroup, g);
    }
    g.impressions += r.impressions;
    g.clicks += r.clicks;
    g.adCost += r.adCost;
    g.directQty += r.directQty;
    g.directRevenue += r.directRevenue;
    g.totalQty += r.totalQty;
    g.totalRevenue += r.totalRevenue;
  }
  return [...map.values()].map(g => {
    const ctr = g.impressions > 0 ? g.clicks / g.impressions * 100 : 0;
    const cpc = g.clicks > 0 ? g.adCost / g.clicks : 0;
    const cvr = g.clicks > 0 ? g.directQty / g.clicks * 100 : 0;
    const roas = g.adCost > 0 ? g.directRevenue / g.adCost * 100 : 0;
    return { ...g, ctr, cpc, cvr, roas };
  });
}

// =========================================================
// 3) AI 룰 엔진
// =========================================================
/*
   ── 입찰가 추천 규칙 ──
   ① 올림  : ROAS >= 300 AND CVR >= 평균CVR*1.2 AND 클릭 >= 5 → +5~15%
   ② 내림  : (클릭 >= 10 AND 구매 == 0) OR (ROAS < 100 AND 클릭 >= 5) → -5~20%
   ③ 삭제  : 클릭 >= 20 AND 구매 == 0 AND 광고비 >= 30000 → 삭제
   ④ 유지  : ROAS 100~300 사이, 클릭 >= 3 (안정 구간)
   ⑤ 신규  : 상위 키워드의 파생/조합
   ⑥ 등급  : 종합 점수 → A~F
*/

/** ① 입찰가 올릴 키워드 */
function ruleBidUp(rows, avgCvr) {
  return rows
    .filter(g => g.keyword && g.keyword !== '(자동노출)')
    .filter(g => g.roas >= 300 && g.cvr >= Math.max(3, avgCvr * 1.2) && g.clicks >= 5)
    .map(g => {
      // 증액 폭: ROAS 값에 따라 5~15%
      let pct = 5;
      if (g.roas >= 800) pct = 15;
      else if (g.roas >= 500) pct = 10;
      const cur = Math.round(g.cpc);
      const rec = Math.round(cur * (1 + pct / 100));
      const reason = `ROAS ${fmtPct(g.roas, 0)} · CVR ${fmtPct(g.cvr)} → +${pct}%`;
      return { ...g, curBid: cur, recBid: rec, reason };
    })
    .sort((a, b) => b.roas - a.roas)
    .slice(0, 20);
}

/** ② 입찰가 내릴 키워드 */
function ruleBidDown(rows) {
  return rows
    .filter(g => g.keyword && g.keyword !== '(자동노출)')
    .filter(g => (g.clicks >= 10 && g.directQty === 0) || (g.roas > 0 && g.roas < 100 && g.clicks >= 5))
    .map(g => {
      let pct = 5;
      let reason;
      if (g.clicks >= 10 && g.directQty === 0) {
        pct = g.clicks >= 30 ? 20 : g.clicks >= 20 ? 15 : 10;
        reason = `클릭 ${g.clicks} 대비 구매 0 → -${pct}%`;
      } else {
        pct = g.roas < 50 ? 15 : 10;
        reason = `ROAS ${fmtPct(g.roas, 0)} 저조 → -${pct}%`;
      }
      const cur = Math.round(g.cpc);
      const rec = Math.round(cur * (1 - pct / 100));
      return { ...g, curBid: cur, recBid: rec, reason };
    })
    .sort((a, b) => b.adCost - a.adCost)
    .slice(0, 20);
}

/** ③ 삭제 후보 */
function ruleKill(rows) {
  return rows
    .filter(g => g.clicks >= 20 && g.directQty === 0 && g.adCost >= 30000)
    .sort((a, b) => b.adCost - a.adCost)
    .slice(0, 20);
}

/** ④ 유지 추천 */
function ruleKeep(rows) {
  return rows
    .filter(g => g.roas >= 100 && g.roas < 300 && g.clicks >= 3 && g.directQty > 0)
    .sort((a, b) => b.roas - a.roas)
    .slice(0, 20);
}

/** ⑤ 신규 테스트 추천 키워드
 *  로직: 성과 상위(ROAS>=300) 키워드에서 조합 어휘를 파생.
 *  파생 어휘: 브랜드/카테고리 힌트 + 상위 키워드의 형태소 재조합
 */
function ruleNewKeywords(rows) {
  const goodRows = rows
    .filter(g => g.keyword && g.keyword !== '(자동노출)')
    .filter(g => g.roas >= 300 && g.clicks >= 3);

  // 상위 키워드 토큰 수집
  const tokens = new Map(); // token → 누적 매출
  const goodKeywords = new Set();
  for (const g of goodRows) {
    goodKeywords.add(g.keyword);
    // 한글/영문 토큰 분리
    const parts = g.keyword.split(/\s+/).flatMap(p => p.split(/(?<=[가-힣])(?=[a-zA-Z0-9])|(?<=[a-zA-Z0-9])(?=[가-힣])/));
    for (const t of parts) {
      if (t.length < 2) continue;
      tokens.set(t, (tokens.get(t) || 0) + g.directRevenue);
    }
  }
  const topTokens = [...tokens.entries()].sort((a, b) => b[1] - a[1]).map(x => x[0]);

  // 접두/접미 어휘 (뷰티/헤어 소도구 도메인)
  const prefixes = ['미니', '휴대용', '무선', '여행용', '전문가용', 'LED'];
  const suffixes = ['추천', '순위', '가성비', '세트', '거치대', '파우치'];

  const suggested = new Set();
  // (1) 상위 토큰 × 접미어
  for (const t of topTokens.slice(0, 4)) {
    for (const s of suffixes) suggested.add(`${t} ${s}`);
  }
  // (2) 접두어 × 상위 토큰
  for (const p of prefixes) {
    for (const t of topTokens.slice(0, 3)) suggested.add(`${p} ${t}`);
  }
  // 이미 있는 키워드는 제외
  const existing = new Set(rows.map(r => normKw(r.keyword)));
  const results = [...suggested].filter(k => !existing.has(normKw(k))).slice(0, 12);

  return results.map(k => ({
    keyword: k,
    reason: '상위 성과 키워드에서 파생',
  }));
}

/** ⑥ 상품(광고그룹)별 광고 점수
 *  종합 점수 = ROAS(35) + CTR(20) + CVR(25) + 비용효율(20)
 *  각 지표는 상대적 percentile 기준으로 0~1 로 정규화
 */
function ruleGrade(groups) {
  if (groups.length === 0) return [];
  // 최소 노출 임계값 — 노출 100 미만은 신뢰도 낮으므로 별도 처리
  const scored = groups.filter(g => g.impressions >= 100 && g.adCost > 0);
  if (scored.length === 0) return [];

  // 각 지표의 최대값 (percentile 대체 — 간단화)
  const maxRoas = Math.max(...scored.map(g => g.roas), 100);
  const maxCtr  = Math.max(...scored.map(g => g.ctr),  1);
  const maxCvr  = Math.max(...scored.map(g => g.cvr),  1);
  const maxCost = Math.max(...scored.map(g => g.adCost), 1);

  const items = scored.map(g => {
    const roasS = Math.min(1, g.roas / maxRoas);
    const ctrS  = Math.min(1, g.ctr  / maxCtr);
    const cvrS  = Math.min(1, g.cvr  / maxCvr);
    // 비용효율: 광고비 대비 매출 (ROAS와 유사하나 규모까지 반영)
    const efficiency = g.adCost > 0 ? Math.min(1, (g.directRevenue - g.adCost) / (maxCost)) : 0;
    const total = roasS * 35 + ctrS * 20 + cvrS * 25 + Math.max(0, efficiency) * 20;
    const grade =
      total >= 80 ? 'A' :
      total >= 65 ? 'B' :
      total >= 50 ? 'C' :
      total >= 35 ? 'D' :
      total >= 20 ? 'E' : 'F';
    const stars = Math.max(1, Math.min(5, Math.round(total / 20)));
    return { ...g, score: total, grade, stars };
  });
  return items.sort((a, b) => b.score - a.score);
}

/** ⑦ AI 코멘트 자동 생성 */
function makeComment(rows, bidUp, bidDown, killList, kpi) {
  const parts = [];
  const best = [...rows].filter(g => g.clicks >= 5).sort((a, b) => b.roas - a.roas)[0];
  const worst = [...rows].filter(g => g.clicks >= 10 && g.directQty === 0).sort((a, b) => b.adCost - a.adCost)[0];

  parts.push(`이번 기간 총 광고비 <strong>${fmtWon(kpi.adCost)}</strong>, ROAS <strong>${fmtPct(kpi.roas, 0)}</strong>, 전환수 ${fmtInt(kpi.directQty)}건이었어요.`);
  if (best && best.roas > 200) {
    const upPct = best.roas >= 800 ? 15 : best.roas >= 500 ? 10 : 5;
    parts.push(`특히 <strong>‘${best.keyword}’</strong> 키워드가 ROAS ${fmtPct(best.roas, 0)}로 매우 좋은 효율을 보이고 있어, 입찰가를 <strong>+${upPct}%</strong> 인상하는 것을 추천합니다.`);
  }
  if (worst) {
    parts.push(`반대로 <strong>‘${worst.keyword}’</strong>는 클릭 ${worst.clicks}회 대비 구매가 0건이라 예산 축소 또는 소재/상세페이지 점검이 필요합니다.`);
  }
  if (killList.length >= 3) {
    parts.push(`광고비 낭비로 분류된 키워드가 <strong>${killList.length}개</strong>입니다. 우선 상위 3개(<strong>${killList.slice(0,3).map(k => k.keyword).join(', ')}</strong>) 확인 후 삭제 검토를 권장합니다.`);
  }
  if (bidUp.length === 0 && bidDown.length === 0 && killList.length === 0) {
    parts.push(`전반적으로 안정적인 성과를 유지하고 있어, 현재 세팅을 유지하시면 됩니다.`);
  }
  return parts.join('\n\n');
}

// =========================================================
// 4) 렌더링
// =========================================================
function computeKpi(rows) {
  const s = rows.reduce((a, g) => ({
    impressions: a.impressions + g.impressions,
    clicks: a.clicks + g.clicks,
    adCost: a.adCost + g.adCost,
    directQty: a.directQty + g.directQty,
    directRevenue: a.directRevenue + g.directRevenue,
  }), { impressions: 0, clicks: 0, adCost: 0, directQty: 0, directRevenue: 0 });
  s.ctr = s.impressions > 0 ? s.clicks / s.impressions * 100 : 0;
  s.cvr = s.clicks > 0 ? s.directQty / s.clicks * 100 : 0;
  s.cpc = s.clicks > 0 ? s.adCost / s.clicks : 0;
  s.roas = s.adCost > 0 ? s.directRevenue / s.adCost * 100 : 0;
  return s;
}

function renderKpi(k) {
  document.getElementById('kpiCost').textContent = fmtWon(k.adCost);
  document.getElementById('kpiImp').textContent  = fmtInt(k.impressions);
  document.getElementById('kpiClick').innerHTML  = `${fmtInt(k.clicks)} <span class="kpi-sub">${fmtPct(k.ctr)}</span>`;
  document.getElementById('kpiConv').innerHTML   = `${fmtInt(k.directQty)} <span class="kpi-sub">${fmtPct(k.cvr)}</span>`;
  document.getElementById('kpiRoas').textContent = fmtPct(k.roas, 0);
  document.getElementById('kpiCpc').textContent  = fmtWon(k.cpc);
}

/** 표 body에 rows를 그린다. rowRender: (row) => [cells...] */
function renderTable(tblId, rows, rowRender, emptyMsg = '해당하는 키워드가 없습니다.') {
  const tbody = document.querySelector(`#${tblId} tbody`);
  if (rows.length === 0) {
    const colCount = document.querySelectorAll(`#${tblId} thead th`).length;
    tbody.innerHTML = `<tr><td class="empty-row" colspan="${colCount}">${emptyMsg}</td></tr>`;
    return;
  }
  tbody.innerHTML = rows.map(r => `<tr>${rowRender(r).join('')}</tr>`).join('');
}

function renderBidUp(list) {
  document.getElementById('cnt1').textContent = `${list.length}개`;
  renderTable('tblBidUp', list, r => [
    `<td class="kw">${r.keyword}</td>`,
    `<td>${r.adGroup}</td>`,
    `<td class="num">${fmtWon(r.curBid)}</td>`,
    `<td class="num up-price">${fmtWon(r.recBid)}</td>`,
    `<td class="num">${fmtPct(r.roas, 0)}</td>`,
    `<td class="num">${fmtPct(r.cvr)}</td>`,
    `<td class="reason">${r.reason}</td>`,
  ]);
}

function renderBidDown(list) {
  document.getElementById('cnt2').textContent = `${list.length}개`;
  renderTable('tblBidDown', list, r => [
    `<td class="kw">${r.keyword}</td>`,
    `<td>${r.adGroup}</td>`,
    `<td class="num">${fmtWon(r.curBid)}</td>`,
    `<td class="num down-price">${fmtWon(r.recBid)}</td>`,
    `<td class="num">${fmtInt(r.clicks)}</td>`,
    `<td class="num">${fmtInt(r.directQty)}</td>`,
    `<td class="reason">${r.reason}</td>`,
  ]);
}

function renderKill(list) {
  document.getElementById('cnt3').textContent = `${list.length}개`;
  renderTable('tblKill', list, r => [
    `<td class="kw">${r.keyword}</td>`,
    `<td>${r.adGroup}</td>`,
    `<td class="num">${fmtInt(r.clicks)}</td>`,
    `<td class="num">${fmtInt(r.directQty)}</td>`,
    `<td class="num">${fmtWon(r.adCost)}</td>`,
    `<td class="reason">클릭 대비 구매 0 · 광고비 낭비</td>`,
  ]);
}

function renderKeep(list) {
  document.getElementById('cnt4').textContent = `${list.length}개`;
  renderTable('tblKeep', list, r => [
    `<td class="kw">${r.keyword}</td>`,
    `<td>${r.adGroup}</td>`,
    `<td class="num">${fmtPct(r.roas, 0)}</td>`,
    `<td class="num">${fmtPct(r.ctr)}</td>`,
    `<td class="num">${fmtPct(r.cvr)}</td>`,
    `<td class="num">${fmtWon(r.adCost)}</td>`,
  ]);
}

function renderNewKeywords(list) {
  const grid = document.getElementById('newKeywordGrid');
  document.getElementById('cnt5').textContent = `${list.length}개`;
  if (list.length === 0) {
    grid.innerHTML = `<div class="hint">파생 대상이 될 상위 성과 키워드가 부족합니다.</div>`;
    return;
  }
  grid.innerHTML = list.map(k =>
    `<span class="kw-chip">${k.keyword}<span class="kw-meta">${k.reason}</span></span>`
  ).join('');
}

function renderGrade(list) {
  const rows = list.slice(0, 20);
  renderTable('tblGrade', rows, r => {
    const stars = '★'.repeat(r.stars) + `<span class="stars-empty">${'★'.repeat(5 - r.stars)}</span>`;
    return [
      `<td class="kw">${r.adGroup}</td>`,
      `<td class="grade-${r.grade}">${r.grade}</td>`,
      `<td class="stars">${stars}</td>`,
      `<td class="num">${fmtPct(r.roas, 0)}</td>`,
      `<td class="num">${fmtPct(r.ctr)}</td>`,
      `<td class="num">${fmtPct(r.cvr)}</td>`,
      `<td class="num">${fmtWon(r.adCost)}</td>`,
      `<td class="num">${fmt1(r.score, 1)}</td>`,
    ];
  }, '점수 산정에 필요한 노출 데이터가 부족합니다.');
}

function renderComment(html) {
  document.getElementById('aiComment').innerHTML = html;
}

/** 차트 렌더링: 상위 광고그룹의 선택 지표를 막대 그래프로 */
function renderChart(groups) {
  const metric = state.chartMetric;
  const top = [...groups]
    .filter(g => g.impressions >= 100)
    .sort((a, b) => b.adCost - a.adCost)
    .slice(0, 8);
  const labels = top.map(g => g.adGroup.length > 14 ? g.adGroup.slice(0, 14) + '…' : g.adGroup);
  const data = top.map(g => Number((g[metric] || 0).toFixed(1)));
  const color = { roas: '#E11D2E', ctr: '#2563EB', cvr: '#16A34A' }[metric];
  const label = { roas: 'ROAS (%)', ctr: 'CTR (%)', cvr: 'CVR (%)' }[metric];

  if (mainChart) mainChart.destroy();
  const ctx = document.getElementById('mainChart').getContext('2d');
  mainChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label, data,
        backgroundColor: color + '99',
        borderColor: color,
        borderWidth: 1.5,
        borderRadius: 6,
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (c) => ` ${label}: ${c.parsed.y.toFixed(1)}%`,
          }
        }
      },
      scales: {
        y: { beginAtZero: true, ticks: { callback: v => v + '%' } },
        x: { ticks: { maxRotation: 40, minRotation: 0, font: { size: 11 } } },
      }
    }
  });
}

// =========================================================
// 5) 파이프라인 실행
// =========================================================
function runAnalysis() {
  const rows = aggregate();
  const groupRows = aggregateByGroup(state.rows.filter(passBrand));
  const kpi = computeKpi(rows);
  const avgCvr = kpi.cvr;

  const bidUp = ruleBidUp(rows, avgCvr);
  const bidDown = ruleBidDown(rows);
  const killList = ruleKill(rows);
  const keepList = ruleKeep(rows);
  const newKws = ruleNewKeywords(rows);
  const grades = ruleGrade(groupRows);
  const commentHtml = makeComment(rows, bidUp, bidDown, killList, kpi);

  renderKpi(kpi);
  renderBidUp(bidUp);
  renderBidDown(bidDown);
  renderKill(killList);
  renderKeep(keepList);
  renderNewKeywords(newKws);
  renderGrade(grades);
  renderComment(commentHtml);
  renderChart(groupRows);

  // 표시 전환
  document.getElementById('emptyState').hidden = true;
  document.getElementById('dashboard').hidden = false;
  document.getElementById('loadInfo').textContent =
    `${state.rows.length.toLocaleString()}건 · ${state.dates.size}일 · 브랜드 ${state.brand} · 그룹 ${groupRows.length}`;
}

// =========================================================
// 6) 이벤트 바인딩
// =========================================================
async function handleFiles(fileList) {
  const files = [...fileList];
  if (files.length === 0) return;
  const allRows = [];
  const dates = new Set();
  for (const f of files) {
    try {
      const rs = await readFile(f);
      allRows.push(...rs);
      rs.forEach(r => dates.add(r.date));
    } catch (e) {
      alert(`${f.name} 파싱 실패: ${e.message}`);
    }
  }
  if (allRows.length === 0) {
    alert('데이터를 읽을 수 없습니다. 쿠팡 CPC 광고 XLSX인지 확인하세요.');
    return;
  }
  state.rows = allRows;
  state.dates = dates;
  runAnalysis();
}

/** 브랜드/기간 탭 클릭 */
function bindTabs() {
  document.getElementById('brandTabs').addEventListener('click', (e) => {
    const b = e.target.closest('.btag');
    if (!b) return;
    document.querySelectorAll('#brandTabs .btag').forEach(x => x.classList.remove('active'));
    b.classList.add('active');
    state.brand = b.dataset.brand;
    if (state.rows.length > 0) runAnalysis();
  });
  document.getElementById('periodTabs').addEventListener('click', (e) => {
    const b = e.target.closest('.period');
    if (!b) return;
    document.querySelectorAll('#periodTabs .period').forEach(x => x.classList.remove('active'));
    b.classList.add('active');
    state.period = Number(b.dataset.period);
    // 다중 파일이 아니면 기간은 표시용
    if (state.rows.length > 0) runAnalysis();
  });
  document.querySelectorAll('.chart-toggle .chip').forEach(chip => {
    chip.addEventListener('click', () => {
      document.querySelectorAll('.chart-toggle .chip').forEach(c => c.classList.remove('active'));
      chip.classList.add('active');
      state.chartMetric = chip.dataset.metric;
      const groupRows = aggregateByGroup(state.rows.filter(passBrand));
      renderChart(groupRows);
    });
  });
}

function bindFiles() {
  const on = (id) => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('change', (e) => handleFiles(e.target.files));
  };
  on('fileInput');
  on('fileInput2');
}

// =========================================================
// 원격 자동 로드
// ---------------------------------------------------------
// 기본 경로: 같은 origin의 ./data/cpc/YYYYMMDD.xlsx
// 이 파일을 coupang-dashboard 저장소와 함께 배포하면
// 별도 CORS 설정 없이 곧바로 fetch 됩니다.
// =========================================================
// 데이터 소스 (KSK 계정의 coupang-dashboard 저장소가 원본)
// 시도 순서:
//   1) 같은 origin 상대경로 (같은 저장소에 배포된 경우)
//   2) coupang-dashboard.pages.dev 직접 (같은 origin이면 성공, 크로스면 CORS 차단)
//   3) corsproxy.io 프록시 (크로스 오리진에서도 성공, 프록시 서비스 의존)
//   4) allorigins.win 프록시 (2차 폴백)
const KSK_ORIGIN = 'https://coupang-dashboard.pages.dev';
const REMOTE_BASE = './data/cpc/';
const REMOTE_DIRECT = `${KSK_ORIGIN}/data/cpc/`;
const CORS_PROXIES = [
  (url) => `https://corsproxy.io/?url=${encodeURIComponent(url)}`,
  (url) => `https://api.allorigins.win/raw?url=${encodeURIComponent(url)}`,
];
const REMOTE_LOOKBACK_DAYS = 60;
const REMOTE_MAX_CONSEC_MISS = 7;

function dateKeyMinus(days) {
  const d = new Date();
  d.setDate(d.getDate() - days);
  const y = d.getFullYear(), m = String(d.getMonth() + 1).padStart(2, '0'), dd = String(d.getDate()).padStart(2, '0');
  return { key: `${y}-${m}-${dd}`, ymd: `${y}${m}${dd}` };
}

function setRemoteStatus(msg, isErr = false) {
  const el = document.getElementById('remoteStatus');
  if (!el) return;
  el.textContent = msg;
  el.classList.toggle('err', isErr);
}

/** 단일 URL에서 XLSX 시도. 성공 시 rows 배열, 실패 시 null */
async function tryFetchXlsx(url, dateKey) {
  try {
    const r = await fetch(url, { cache: 'no-cache' });
    if (!r.ok) return null;
    const ct = (r.headers.get('content-type') || '').toLowerCase();
    // Cloudflare 등이 없는 파일에 index.html을 돌려주는 케이스 거부
    if (ct.includes('text/html')) return null;
    const buf = await r.arrayBuffer();
    if (buf.byteLength < 100) return null;
    const wb = XLSX.read(buf, { type: 'array' });
    // 파일명을 dateKey로 위장해서 파서 재사용
    const rows = parseWorkbook(wb, `remote_${dateKey.replace(/-/g,'')}_${dateKey.replace(/-/g,'')}.xlsx`);
    return rows;
  } catch (e) {
    return null;
  }
}

/** 최근 N일 원격 XLSX 훑기 */
async function fetchRemote() {
  const allRows = [];
  const dates = new Set();
  let consecMiss = 0;
  let hit = 0;

  // period 필터 기준 며칠 훑을지: state.period 값 우선, 없으면 30일
  const targetDays = Math.max(state.period || 7, 30);

  for (let i = 1; i <= REMOTE_LOOKBACK_DAYS; i++) {
    const { key, ymd } = dateKeyMinus(i);
    const directUrl = `${REMOTE_DIRECT}${ymd}.xlsx`;
    const urls = [
      `${REMOTE_BASE}${ymd}.xlsx`,     // 같은 origin 상대경로
      directUrl,                        // KSK 저장소 직접
      ...CORS_PROXIES.map(p => p(directUrl)), // CORS 프록시 폴백
    ];
    setRemoteStatus(`원격 ${key} 확인 중... (${hit}일 수집)`);
    let rows = null;
    for (const u of urls) {
      rows = await tryFetchXlsx(u, key);
      if (rows) break;
    }
    if (rows && rows.length > 0) {
      allRows.push(...rows);
      dates.add(key);
      hit++;
      consecMiss = 0;
      // 목표 일수 도달 시 종료
      if (hit >= targetDays) break;
    } else {
      consecMiss++;
      if (consecMiss >= REMOTE_MAX_CONSEC_MISS && hit === 0) {
        setRemoteStatus('원격 데이터를 찾을 수 없습니다. XLSX 업로드나 샘플을 이용하세요.', true);
        return;
      }
      if (consecMiss >= REMOTE_MAX_CONSEC_MISS) break;
    }
  }

  if (allRows.length === 0) {
    setRemoteStatus('원격 데이터 로드 실패. CORS 프록시가 응답 안 하거나 파일이 없습니다. XLSX 업로드를 이용하세요.', true);
    return;
  }

  state.rows = allRows;
  state.dates = dates;
  setRemoteStatus(`원격 로드 완료: ${hit}일 · ${allRows.length.toLocaleString()}건`);
  runAnalysis();
}

function bindRemote() {
  const btn = document.getElementById('remoteBtn');
  const btn2 = document.getElementById('remoteBtn2');
  if (btn) btn.addEventListener('click', fetchRemote);
  if (btn2) btn2.addEventListener('click', fetchRemote);
}

function bindSample() {
  const load = () => { state.rows = buildSampleData(); state.dates = new Set(state.rows.map(r => r.date)); runAnalysis(); };
  document.getElementById('sampleBtn').addEventListener('click', load);
  document.getElementById('sampleBtn2').addEventListener('click', load);
}

// =========================================================
// 7) 샘플 데이터 (미리보기용)
// =========================================================
function buildSampleData() {
  const samples = [
    // adGroup, keyword, imp, click, cost, dqty, drev
    ['보아르 미니고데기', '미니고데기',    12000,  480,  144000,  0,       0],
    ['보아르 미니고데기', '휴대용고데기',    8000,  260,   78000,  2,   36000],
    ['보아르 빗고데기',   '빗고데기',      18000, 1050,  315000, 42,  1180000],
    ['보아르 빗고데기',   '빗형고데기',    10500,  620,  186000, 22,   630000],
    ['보아르 빗고데기',   '보아르 빗고데기', 6200,  380,  114000, 30,   870000],
    ['오아 안마기',       '오아안마기',    22000,  980,  244000, 55,   980000],
    ['오아 안마기',       '전신안마기',    15500,  610,  152000,  8,   210000],
    ['오아 종아리 마사지기','종아리마사지기',9800,  340,   85000,  0,       0],
    ['오아 온열매트',     '온열매트',      13400,  520,  130000, 18,   540000],
    ['오아 온열매트',     '전기매트',       9800,  310,   77500,  6,   170000],
    ['보아르 헤어드라이어','고속드라이어',  11200,  430,  129000, 12,   380000],
    ['보아르 헤어드라이어','저소음드라이어', 4300,   90,   27000,  0,       0],
    ['오아 정수기',        '미니정수기',   28000,  790,  197000,  4,   120000],
    ['오아 정수기',        '휴대용정수기',  6100,  180,   45000,  9,   240000],
    ['보아르 매직기',      '매직기',       17800,  980,  294000, 51,  1470000],
    ['보아르 매직기',      '스팀매직기',    9400,  520,  156000, 14,   420000],
  ];
  const today = new Date().toISOString().slice(0, 10);
  return samples.map(([adGroup, keyword, imp, click, cost, dqty, drev]) => ({
    date: today,
    campaignName: adGroup,
    adGroup,
    productName: adGroup,
    placement: '검색영역',
    keyword,
    keywordNorm: normKw(keyword),
    chargeMethod: 'cpc',
    adType: '매출최적화',
    impressions: imp,
    clicks: click,
    adCost: cost,
    directQty: dqty,
    directRevenue: drev,
    totalQty: dqty + Math.floor(dqty * 0.3),
    totalRevenue: drev * 1.3,
  }));
}

// ---------- 부팅 ----------
document.addEventListener('DOMContentLoaded', () => {
  bindTabs();
  bindFiles();
  bindSample();
  bindRemote();
});
