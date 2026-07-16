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
  rows: [],              // 파일에서 파싱한 원본 rows (날짜 포함)
  dates: new Set(),      // 데이터가 있는 날짜들
  period: 7,             // 최근 N일 필터
  brand: '전체',
  selectedGroup: null,   // 선택된 광고그룹 (null = 전체)
  searchQuery: '',       // 상품 검색어
};

// 브랜드 키워드 판정 — 이런 키워드는 이미 검색 잘 되고 있어서
// ROAS 높다고 입찰가 올려도 확장 여지가 적음 → 입찰가 상승 추천에서 제외
const BRAND_PREFIX_RE = /^(오아|보아르|VOAR|OA-|OOR)/i;
function isBrandKeyword(kw) {
  return BRAND_PREFIX_RE.test((kw || '').trim());
}

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

// =========================================================
// 2) 집계: 광고그룹 × 키워드 단위
// =========================================================
/** 브랜드 필터 통과 여부 */
function passBrand(row) {
  if (state.brand === '전체') return true;
  const blob = `${row.campaignName} ${row.adGroup} ${row.productName}`;
  return classifyBrand(blob) === state.brand;
}

/** rows → 광고그룹×키워드 집계된 배열 (선택된 그룹만 or 전체) */
function aggregate() {
  const map = new Map();
  for (const r of state.rows) {
    if (!passBrand(r)) continue;
    if (state.selectedGroup && r.adGroup !== state.selectedGroup) continue;
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

/** ① 입찰가 올릴 키워드 — 객관적 시장 확장성 분석
 *
 *  판정 기준 (모두 통과해야 추천):
 *   (a) 브랜드 키워드 제외 — '오아/보아르/VOAR/OA-/OOR' 로 시작하면 핵심 확장 대상 아님
 *       (이미 브랜드 검색으로 유입되고 있어 입찰 올려도 신규 고객 확보 어려움)
 *   (b) 실제 전환 있음 — clicks ≥ 5, directQty ≥ 1 (우연이 아닌 실증)
 *   (c) 수익성 확보 — ROAS ≥ 200% (증액 후에도 흑자 유지)
 *   (d) 전환력 우수 — CVR ≥ 평균 CVR × 1.2 (상세페이지가 팔림)
 *   (e) 확장 여지 — 노출수가 전체 평균 이하 (더 많이 노출되면 매출 성장 여력)
 *
 *  증액 폭: 데이터 강도(전환력·수익성·확장여지) 기반 스코어링 5~20%
 */
function ruleBidUp(rows, avgCvr) {
  // 확장 여지 판정용 노출 중앙값
  const impList = rows.filter(g => g.clicks >= 3).map(g => g.impressions).sort((a, b) => a - b);
  const impMedian = impList[Math.floor(impList.length / 2)] || 0;

  return rows
    .filter(g => g.keyword && g.keyword !== '(자동노출)')
    .filter(g => !isBrandKeyword(g.keyword))       // (a) 브랜드 키워드 제외
    .filter(g => g.clicks >= 5 && g.directQty >= 1) // (b) 실증 데이터
    .filter(g => g.roas >= 200)                     // (c) 수익성
    .filter(g => g.cvr >= Math.max(2, avgCvr * 1.2))// (d) 전환력
    .map(g => {
      // 확장 여지 = 노출이 중앙값보다 낮을수록 높음
      const hasRoom = impMedian > 0 && g.impressions < impMedian;
      // 강도 점수: ROAS + CVR + 확장여지
      let pct = 5;
      const reasons = [];

      // ROAS 강도
      if (g.roas >= 800) { pct += 6; reasons.push(`ROAS ${fmtPct(g.roas, 0)} 매우 우수`); }
      else if (g.roas >= 500) { pct += 4; reasons.push(`ROAS ${fmtPct(g.roas, 0)} 우수`); }
      else { pct += 2; reasons.push(`ROAS ${fmtPct(g.roas, 0)} 양호`); }

      // CVR 강도
      if (g.cvr >= avgCvr * 2) { pct += 4; reasons.push(`CVR ${fmtPct(g.cvr)} 평균 2배 (강한 구매 의도)`); }
      else if (g.cvr >= avgCvr * 1.5) { pct += 2; reasons.push(`CVR ${fmtPct(g.cvr)} 평균 대비 우수`); }
      else { reasons.push(`CVR ${fmtPct(g.cvr)}`); }

      // 확장 여지
      if (hasRoom) { pct += 3; reasons.push(`노출 ${fmtInt(g.impressions)}회로 확장 여지 큼`); }

      pct = Math.min(20, pct);

      const cur = Math.round(g.cpc);
      const rec = Math.round(cur * (1 + pct / 100));
      const reason = `${reasons.join(' · ')} → +${pct}%`;
      // 우선순위 점수: 확장 여지 + 전환력
      const priority = (g.roas / 100) * g.cvr * (hasRoom ? 1.5 : 1);
      return { ...g, curBid: cur, recBid: rec, reason, priority };
    })
    .sort((a, b) => b.priority - a.priority)
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

// =========================================================
// 상품 검색 (광고그룹 리스트)
// =========================================================
/** 브랜드 필터를 통과한 rows에서 광고그룹별 요약 → 검색 필터 후 렌더 */
function renderProductPicker() {
  const listEl = document.getElementById('productList');
  const countEl = document.getElementById('productCount');
  const selectedInfo = document.getElementById('selectedInfo');
  const selectedName = document.getElementById('selectedName');
  if (!listEl) return;

  // 광고그룹별 지표 합산 (브랜드 필터 적용)
  const groupMap = new Map();
  for (const r of state.rows) {
    if (!passBrand(r)) continue;
    if (!r.adGroup) continue;
    const g = groupMap.get(r.adGroup) || {
      adGroup: r.adGroup,
      impressions: 0, clicks: 0, adCost: 0,
      directQty: 0, directRevenue: 0,
    };
    g.impressions += r.impressions;
    g.clicks += r.clicks;
    g.adCost += r.adCost;
    g.directQty += r.directQty;
    g.directRevenue += r.directRevenue;
    groupMap.set(r.adGroup, g);
  }
  const allGroups = [...groupMap.values()].map(g => ({
    ...g,
    roas: g.adCost > 0 ? g.directRevenue / g.adCost * 100 : 0,
  })).sort((a, b) => b.adCost - a.adCost);

  // 검색어 필터
  const q = state.searchQuery.trim().toLowerCase();
  const filtered = q
    ? allGroups.filter(g => g.adGroup.toLowerCase().includes(q))
    : allGroups;

  countEl.textContent = `총 ${allGroups.length}개${q ? ` (검색 ${filtered.length})` : ''}`;

  // 선택 배너
  if (state.selectedGroup) {
    selectedInfo.hidden = false;
    selectedName.textContent = state.selectedGroup;
  } else {
    selectedInfo.hidden = true;
  }

  if (allGroups.length === 0) {
    listEl.innerHTML = `<div class="picker-empty">데이터가 없습니다.</div>`;
    return;
  }
  if (filtered.length === 0) {
    listEl.innerHTML = `<div class="picker-empty">검색 결과가 없습니다.</div>`;
    return;
  }

  const roasClass = (roas) => roas >= 300 ? 'high' : roas >= 100 ? 'mid' : 'low';
  listEl.innerHTML = filtered.slice(0, 200).map(g => {
    const active = state.selectedGroup === g.adGroup;
    const escaped = g.adGroup.replace(/"/g, '&quot;');
    return `
      <div class="product-item ${active ? 'active' : ''}" data-group="${escaped}">
        <div class="name">${g.adGroup}</div>
        <div class="metric">광고비 <strong>${fmtWon(g.adCost)}</strong></div>
        <div class="metric">전환 <strong>${fmtInt(g.directQty)}</strong></div>
        <div class="roas-tag ${roasClass(g.roas)}">ROAS ${fmtPct(g.roas, 0)}</div>
      </div>`;
  }).join('');
}

function bindProductPicker() {
  const listEl = document.getElementById('productList');
  const searchEl = document.getElementById('productSearch');
  const clearBtn = document.getElementById('clearSearch');
  const clearSel = document.getElementById('clearSelection');

  // 리스트 클릭 → 광고그룹 선택
  listEl.addEventListener('click', (e) => {
    const item = e.target.closest('.product-item');
    if (!item) return;
    state.selectedGroup = item.dataset.group;
    renderProductPicker();
    runAnalysis();
    document.querySelector('.kpi-row')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  });

  // 검색어 입력 → 필터
  let debounceTimer = null;
  searchEl.addEventListener('input', (e) => {
    clearBtn.hidden = !e.target.value;
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(() => {
      state.searchQuery = e.target.value;
      renderProductPicker();
    }, 150);
  });

  clearBtn.addEventListener('click', () => {
    searchEl.value = '';
    state.searchQuery = '';
    clearBtn.hidden = true;
    renderProductPicker();
    searchEl.focus();
  });

  // 선택 해제
  clearSel.addEventListener('click', () => {
    state.selectedGroup = null;
    renderProductPicker();
    runAnalysis();
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

  renderKpi(kpi);
  renderBidUp(bidUp);
  renderBidDown(bidDown);
  renderKill(killList);
  renderKeep(keepList);
  renderNewKeywords(newKws);
  renderGrade(grades);

  const scopeText = state.selectedGroup ? `선택: ${state.selectedGroup}` : '전체 상품';
  document.getElementById('loadInfo').textContent =
    `${state.rows.length.toLocaleString()}건 · ${state.dates.size}일 · 브랜드 ${state.brand} · ${scopeText}`;
}

// =========================================================
// 6) 이벤트 바인딩
// =========================================================
/** 브랜드/기간 탭 클릭 */
function bindTabs() {
  document.getElementById('brandTabs').addEventListener('click', (e) => {
    const b = e.target.closest('.btag');
    if (!b) return;
    document.querySelectorAll('#brandTabs .btag').forEach(x => x.classList.remove('active'));
    b.classList.add('active');
    state.brand = b.dataset.brand;
    state.selectedGroup = null; // 브랜드 변경 시 선택 해제
    if (state.rows.length > 0) { renderProductPicker(); runAnalysis(); }
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
//   1) 같은 origin의 Pages Function (/api/cpc/YYYYMMDD) — 서버사이드 fetch, 초고속
//   2) 같은 origin 상대경로 (같은 저장소에 배포된 경우)
//   3) coupang-dashboard.pages.dev 직접 (같은 origin이면 성공, 크로스면 CORS 차단)
//   4) corsproxy.io 프록시 (Function 없을 때 최후 폴백)
const KSK_ORIGIN = 'https://coupang-dashboard.pages.dev';
const API_BASE = '/api/cpc/';                       // Pages Function 엔드포인트 (최우선)
const REMOTE_BASE = './data/cpc/';
const REMOTE_DIRECT = `${KSK_ORIGIN}/data/cpc/`;
const CORS_PROXIES = [
  (url) => `https://corsproxy.io/?url=${encodeURIComponent(url)}`,
];
const REMOTE_LOOKBACK_DAYS = 60;

function dateKeyMinus(days) {
  const d = new Date();
  d.setDate(d.getDate() - days);
  const y = d.getFullYear(), m = String(d.getMonth() + 1).padStart(2, '0'), dd = String(d.getDate()).padStart(2, '0');
  return { key: `${y}-${m}-${dd}`, ymd: `${y}${m}${dd}` };
}

/** 상단 헤더의 로딩/상태 배지 갱신
 *  status: 'loading' | 'ok' | 'err'
 */
function setHeaderStatus(msg, status = 'loading') {
  const el = document.getElementById('headerStatus');
  if (!el) return;
  el.textContent = msg;
  el.classList.remove('ok', 'err');
  if (status === 'ok') el.classList.add('ok');
  if (status === 'err') el.classList.add('err');
}
// 하위호환: 기존 setRemoteStatus 콜은 헤더 배지로 리디렉트
function setRemoteStatus(msg, isErr = false) {
  setHeaderStatus(msg, isErr ? 'err' : 'loading');
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

/** 한 날짜에 대해 Pages Function만 호출 (500ms 이내 응답)
 *  Function이 404면 파일 없음 — 폴백 시도로 시간 낭비하지 않고 즉시 종료
 */
async function fetchOneDay(key, ymd) {
  return await tryFetchXlsx(`${API_BASE}${ymd}`, key);
}

/** 최근 N일 원격 XLSX 병렬 훑기
 *  - 60일 범위를 배치 10개씩 병렬 처리 → 총 6회 왕복
 *  - 어느 날짜에 파일이 있든 없든 모두 시도 (연속 누락으로 조기 종료 안 함)
 */
async function fetchRemote() {
  const allRows = [];
  const dates = new Set();
  const BATCH = 20; // Pages Function만 사용 — 대역폭 여유 충분

  let scanned = 0;
  setHeaderStatus(`데이터 로딩 중... (0/${REMOTE_LOOKBACK_DAYS}일)`, 'loading');

  // 배치 병렬 실행 — i=1(어제)부터 i=60(60일 전)까지
  for (let start = 1; start <= REMOTE_LOOKBACK_DAYS; start += BATCH) {
    const jobs = [];
    for (let i = start; i < start + BATCH && i <= REMOTE_LOOKBACK_DAYS; i++) {
      const { key, ymd } = dateKeyMinus(i);
      jobs.push(fetchOneDay(key, ymd).then(rows => ({ key, rows })));
    }
    const results = await Promise.all(jobs);
    for (const { key, rows } of results) {
      scanned++;
      if (rows && rows.length > 0) {
        allRows.push(...rows);
        dates.add(key);
      }
    }
    setHeaderStatus(`데이터 로딩 중... (${scanned}/${REMOTE_LOOKBACK_DAYS}일, ${dates.size}일 수집)`, 'loading');
    // 원하는 만큼 모였으면 조기 종료
    if (dates.size >= 30) break;
  }

  if (allRows.length === 0) {
    setHeaderStatus('데이터를 찾을 수 없습니다', 'err');
    return;
  }

  state.rows = allRows;
  state.dates = dates;
  setHeaderStatus(`${dates.size}일 · ${allRows.length.toLocaleString()}건 로드됨`, 'ok');
  renderProductPicker();
  runAnalysis();
}

// ---------- 부팅 ----------
document.addEventListener('DOMContentLoaded', () => {
  bindTabs();
  bindProductPicker();
  // 페이지 진입 시 원격 데이터 자동 로드
  fetchRemote();
});
