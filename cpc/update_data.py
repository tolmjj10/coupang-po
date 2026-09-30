"""cpc/index.html 의 DATA 블록을 갱신하는 도구.

사용: 매주 새 엑셀이 들어오면 add_week()로 주차 데이터를 추가한 뒤 저장.
Claude가 대화 중에 편집하는 것이 기본 워크플로우이지만,
직접 실행할 수도 있게 CLI 경로도 열어둠.

DATA 스키마:
  {
    product, seller, targetRoas, breakevenRoas, lastUpdate,
    weeks: [
      {
        id, label, period, isBaseline,
        campaigns: { ai: {cost, revenue, clicks, orders, impressions, name}, manual: {...} },
        actionsBefore: [...],
        keywordMovement: { excluded: [...], kept: [...], new: [...] } | null
      }
    ]
  }
"""
import json
import re
import sys
from pathlib import Path
from collections import defaultdict

HERE = Path(__file__).parent
HTML = HERE / "index.html"


def num(x):
    if x is None:
        return 0
    if isinstance(x, (int, float)):
        return x
    s = str(x).replace(",", "").replace("%", "").strip()
    try:
        return float(s)
    except Exception:
        return 0


def summarize_excel(xlsx_path, ai_camp_name, manual_camp_name):
    """쿠팡 판매자센터 다운로드 엑셀(43컬럼) → 캠페인 요약 dict."""
    import openpyxl
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    rows = list(wb.active.iter_rows(values_only=True))[1:]
    agg = defaultdict(lambda: {"cost": 0, "revenue": 0, "clicks": 0, "orders": 0, "impressions": 0})
    for r in rows:
        c = r[4] or ""
        agg[c]["impressions"] += num(r[12])
        agg[c]["clicks"] += num(r[13])
        agg[c]["cost"] += num(r[14])
        agg[c]["orders"] += num(r[25])
        agg[c]["revenue"] += num(r[31])
    return {
        "ai":     {**agg[ai_camp_name],     "name": "AI (스마트타겟팅)"},
        "manual": {**agg[manual_camp_name], "name": "Manual (수동)"},
    }


def keyword_set(xlsx_path, campaign_name):
    """캠페인 내 검색 키워드 집합 (기호 '-' 제외)."""
    import openpyxl
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    rows = list(wb.active.iter_rows(values_only=True))[1:]
    return {(r[11] or "").strip() for r in rows if (r[4] or "") == campaign_name and (r[11] or "") != "-"}


def read_data():
    text = HTML.read_text(encoding="utf-8")
    m = re.search(r"const DATA = (\{.*?\n\});", text, re.DOTALL)
    if not m:
        raise RuntimeError("DATA 블록을 찾을 수 없음")
    # JS 객체 → JSON 변환 (키에 따옴표, 후행쉼표 제거)
    js = m.group(1)
    # 후행 쉼표 제거
    js = re.sub(r",\s*([\]}])", r"\1", js)
    # 키 감싸기 (간단한 케이스만)
    js = re.sub(r"([{,]\s*)([a-zA-Z_][a-zA-Z0-9_]*)\s*:", r'\1"\2":', js)
    # 주석 제거
    js = re.sub(r"//.*", "", js)
    return json.loads(js), text, m.span()


def write_data(data, full_text, span):
    js = json.dumps(data, ensure_ascii=False, indent=2)
    new_text = full_text[:span[0]] + "const DATA = " + js + ";" + full_text[span[1]:]
    HTML.write_text(new_text, encoding="utf-8")


def add_week(*, xlsx, week_id, label, period, ai_camp, manual_camp,
             actions_before, prev_xlsx=None):
    """
    xlsx           : 이번 주 엑셀 경로
    week_id        : "week1", "week2" 등 고유 ID
    label          : "1주차" 등 표시명
    period         : "2026-09-30 ~ 2026-10-06 (7일)"
    ai_camp        : 원본 엑셀의 AI 캠페인명 (정확히)
    manual_camp    : 원본 엑셀의 Manual 캠페인명
    actions_before : 이 주 시작 직전 실행한 조치 리스트
    prev_xlsx      : 직전 주 엑셀 경로 (키워드 이동 계산용, 선택)
    """
    data, full, span = read_data()
    camps = summarize_excel(xlsx, ai_camp, manual_camp)
    kwm = None
    if prev_xlsx and Path(prev_xlsx).exists():
        prev_ai = keyword_set(prev_xlsx, ai_camp)
        prev_man = keyword_set(prev_xlsx, manual_camp)
        cur_ai = keyword_set(xlsx, ai_camp)
        cur_man = keyword_set(xlsx, manual_camp)
        prev_all = prev_ai | prev_man
        cur_all = cur_ai | cur_man
        kwm = {
            "excluded": sorted(prev_all - cur_all),
            "kept":     sorted(prev_all & cur_all),
            "new":      sorted(cur_all - prev_all),
        }
    week = {
        "id": week_id,
        "label": label,
        "period": period,
        "isBaseline": False,
        "campaigns": camps,
        "actionsBefore": actions_before,
        "keywordMovement": kwm,
    }
    # 같은 id가 있으면 교체
    data["weeks"] = [w for w in data["weeks"] if w["id"] != week_id]
    data["weeks"].append(week)
    # lastUpdate 자동
    import datetime
    data["lastUpdate"] = datetime.date.today().isoformat()
    write_data(data, full, span)
    print(f"[OK] {week_id} 추가 완료. lastUpdate={data['lastUpdate']}")
    print(f"  AI 광고비 {camps['ai']['cost']:,.0f} · 매출 {camps['ai']['revenue']:,.0f}")
    print(f"  Manual 광고비 {camps['manual']['cost']:,.0f} · 매출 {camps['manual']['revenue']:,.0f}")
    if kwm:
        print(f"  키워드 이동: 제외 {len(kwm['excluded'])} · 유지 {len(kwm['kept'])} · 신규 {len(kwm['new'])}")


if __name__ == "__main__":
    print(__doc__)
