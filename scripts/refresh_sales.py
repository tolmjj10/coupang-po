#!/usr/bin/env python3
"""
refresh_sales.py — 사내 oa_ops DB에서 판매 데이터를 끌어와 JSON 파일을 갱신하고 git push.

업데이트 대상:
  data/sales-dash.json           ← 대시보드 KPI/6개월 차트 (매번 갱신)
  data/sales-sku/{YYYY-MM}.json  ← 판매 현황 상세 (DB에 새 데이터가 있을 때만)
  data/sales-sku/index.json      ← 월별 totals 인덱스

의존성:
  pip install requests

환경 변수 (기본값은 사용자 PC 전용):
  OA_OPS_MCP_URL  (기본 http://192.168.30.69:8790/mcp)
  OA_OPS_AUTH     (기본 Bearer 토큰 — oa_ops.mcp.team.json 참조)
  REPO_DIR        (기본 C:\\Users\\admin\\Desktop\\17. claude\\po)
"""
import os
import sys
import json
import datetime as dt
import subprocess
from pathlib import Path

# Windows 콘솔 cp949 회피 — 한글/em-dash 출력 깨짐 방지
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

try:
    import requests
except ImportError:
    print("[err] 'requests' 패키지 필요 — pip install requests", file=sys.stderr)
    sys.exit(1)

MCP_URL = os.environ.get("OA_OPS_MCP_URL", "http://192.168.30.69:8790/mcp")
AUTH    = os.environ.get("OA_OPS_AUTH",    "Bearer oaops_ro_9f3kQ2mZ7pLx8vNt3sYb6Hd4Rw")
REPO    = Path(os.environ.get("REPO_DIR",  r"C:\Users\admin\Desktop\17. claude\po"))


# ──────────────────────────────── MCP HTTP 클라이언트 ────────────────────────────────

class McpSession:
    def __init__(self, url, auth):
        self.url = url
        self.sess = requests.Session()
        self.sess.headers.update({
            "Authorization": auth,
            "Content-Type":  "application/json",
            "Accept":        "application/json, text/event-stream",
        })
        self._sid = None
        self._mid = 0
        self._initialize()

    def _next_id(self):
        self._mid += 1
        return self._mid

    def _post(self, payload, is_notification=False):
        headers = {}
        if self._sid:
            headers["mcp-session-id"] = self._sid
        r = self.sess.post(self.url, json=payload, headers=headers, timeout=120)
        r.raise_for_status()
        sid = r.headers.get("mcp-session-id")
        if sid:
            self._sid = sid
        if is_notification or not r.text.strip():
            return None
        ct = r.headers.get("content-type", "")
        if ct.startswith("text/event-stream"):
            # SSE: 이벤트 단위로 파싱. 하나의 이벤트는 여러 data: 라인(\n으로 재결합)으로 나올 수 있음.
            events, current = [], []
            for line in r.text.split("\n"):
                if line.startswith("data:"):
                    current.append(line[5:].lstrip())
                elif line.strip() == "":
                    if current:
                        try:
                            events.append(json.loads("\n".join(current)))
                        except json.JSONDecodeError:
                            pass
                        current = []
            if current:
                try:
                    events.append(json.loads("\n".join(current)))
                except json.JSONDecodeError:
                    pass
            # 이번 요청에 해당하는 id의 응답 반환
            for e in events:
                if "id" in e and e.get("id") == payload.get("id"):
                    return e
            if events:
                return events[-1]
            raise RuntimeError("event-stream 응답에 유효 JSON 없음")
        return r.json()

    def _initialize(self):
        self._post({
            "jsonrpc": "2.0",
            "id": self._next_id(),
            "method": "initialize",
            "params": {
                "protocolVersion": "2024-11-05",
                "capabilities": {},
                "clientInfo": {"name": "refresh_sales", "version": "1.0"},
            },
        })
        self._post({
            "jsonrpc": "2.0",
            "method": "notifications/initialized",
            "params": {},
        }, is_notification=True)

    def run_query(self, sql, limit=500):
        resp = self._post({
            "jsonrpc": "2.0",
            "id": self._next_id(),
            "method": "tools/call",
            "params": {
                "name": "run_query",
                "arguments": {"sql": sql, "limit": limit},
            },
        })
        if "error" in resp:
            raise RuntimeError(f"MCP 오류: {resp['error']}")
        result = resp["result"]
        if result.get("isError"):
            raise RuntimeError(f"쿼리 오류: {result}")
        content = result["content"][0]["text"]
        return json.loads(content)


# ──────────────────────────────── 비즈니스 로직 ────────────────────────────────

def refresh_sales_dash(mcp):
    """매출/원가 집계 → data/sales-dash.json"""
    today = dt.date.today()
    cm = today.strftime("%Y-%m")
    ly = f"{today.year - 1}-{today.month:02d}"

    # 최근 6개월 + 작년동월 — 쿼리 하나로
    first_needed = f"{today.year - 1}-{today.month:02d}-01"
    sql = (
        "SELECT DATE_FORMAT(stat_date,'%Y-%m') ym, "
        "       SUM(sales_amount) sales_amt, "
        "       SUM(cost_amount)  cost_amt "
        "FROM mart_sales_daily "
        f"WHERE partner_name='쿠팡' AND stat_date >= '{first_needed}' "
        "GROUP BY ym ORDER BY ym"
    )
    result = mcp.run_query(sql)
    by_ym = {r["ym"]: {"sales": float(r["sales_amt"] or 0),
                       "cost":  float(r["cost_amt"]  or 0)} for r in result["rows"]}

    # 최근 6개월 (오래된→최신)
    trend = []
    y, m = today.year, today.month
    for _ in range(6):
        trend.append(f"{y}-{m:02d}")
        m -= 1
        if m == 0:
            m = 12
            y -= 1
    trend.reverse()

    dash = {
        "updated_at":   today.isoformat(),
        "source":       "oa_ops.mart_sales_daily · partner_name=쿠팡",
        "current_month": cm,
        "kpi": {
            "sales_this": int(by_ym.get(cm, {}).get("sales", 0)),
            "cost_this":  int(by_ym.get(cm, {}).get("cost",  0)),
            "sales_ly":   int(by_ym.get(ly, {}).get("sales", 0)),
            "cost_ly":    int(by_ym.get(ly, {}).get("cost",  0)),
        },
        "trend_6m": [
            {"ym": ym, "cost": int(by_ym.get(ym, {}).get("cost", 0))}
            for ym in trend
        ],
    }
    out = REPO / "data" / "sales-dash.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(dash, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"[dash] 당월 매출 {dash['kpi']['sales_this']:,}원  입고 {dash['kpi']['cost_this']:,}원  (작년 동월 {dash['kpi']['sales_ly']:,} / {dash['kpi']['cost_ly']:,})")


def db_latest_coupang_date(mcp):
    r = mcp.run_query(
        "SELECT MAX(stat_date) d FROM crawl_coupang_daily WHERE channel_code='coupang'"
    )
    return str(r["rows"][0]["d"]) if r["rows"] else None


def file_latest_nonzero_date(ym):
    f = REPO / "data" / "sales-sku" / f"{ym}.json"
    if not f.exists():
        return None
    j = json.loads(f.read_text(encoding="utf-8"))
    dates = j["dates"]
    out = j["metrics"]["out"]
    last_idx = -1
    for i in range(len(dates)):
        if any((m["qty"][i] or 0) for m in out):
            last_idx = i
    return dates[last_idx] if last_idx >= 0 else None


def refresh_sales_sku_month(mcp, ym):
    """해당 월(YYYY-MM)의 sales-sku JSON 전체 재생성."""
    y, m = map(int, ym.split("-"))
    first = dt.date(y, m, 1)
    last = (dt.date(y + 1, 1, 1) if m == 12 else dt.date(y, m + 1, 1)) - dt.timedelta(days=1)
    dates = [(first + dt.timedelta(days=i)).isoformat() for i in range((last - first).days + 1)]
    date_idx = {d: i for i, d in enumerate(dates)}

    skus = {}  # sku_id -> accumulator

    for day in dates:
        offset = 0
        day_total = 0
        while True:
            sql = (
                "SELECT sku_id, "
                "SUM(COALESCE(outbound_qty,0)) out_q, "
                "SUM(COALESCE(inbound_qty,0))  in_q, "
                "MAX(sku_name)  sn, "
                "MAX(cat_large) cl, MAX(cat_mid) cm, MAX(cat_small) cs "
                "FROM crawl_coupang_daily "
                f"WHERE channel_code='coupang' AND stat_date='{day}' "
                "GROUP BY sku_id ORDER BY sku_id "
                f"LIMIT 500 OFFSET {offset}"
            )
            res = mcp.run_query(sql)
            rows = res["rows"]
            for r in rows:
                sku = str(r["sku_id"])
                if sku not in skus:
                    cat = ">".join([x for x in [r.get("cl"), r.get("cm"), r.get("cs")] if x])
                    skus[sku] = {
                        "id": sku, "c": sku, "sn": r.get("sn") or "",
                        "cat": cat,
                        "total_out": 0, "total_in": 0,
                        "qty_out": [0] * len(dates),
                        "qty_in":  [0] * len(dates),
                    }
                i = date_idx[day]
                out_q = int(r["out_q"] or 0)
                in_q  = int(r["in_q"]  or 0)
                skus[sku]["qty_out"][i] = out_q
                skus[sku]["qty_in"][i]  = in_q
                skus[sku]["total_out"] += out_q
                skus[sku]["total_in"]  += in_q
            day_total += len(rows)
            if len(rows) < 500:
                break
            offset += 500
        print(f"[sku] {day}  SKUs {day_total}")

    # 월별 매출/마진 합계는 mart_sales_daily
    totsql = (
        "SELECT SUM(sales_amount) amt, SUM(margin_amount) mgn "
        "FROM mart_sales_daily WHERE partner_name='쿠팡' "
        f"AND stat_date BETWEEN '{dates[0]}' AND '{dates[-1]}'"
    )
    tot = mcp.run_query(totsql)["rows"][0]

    total_out = sum(s["total_out"] for s in skus.values())
    total_in  = sum(s["total_in"]  for s in skus.values())

    out_metrics = sorted(skus.values(), key=lambda s: -s["total_out"])[:500]
    in_metrics  = sorted(skus.values(), key=lambda s: -s["total_in"])[:500]

    data = {
        "month": ym,
        "dates": dates,
        "totals": {
            "out": total_out,
            "in":  total_in,
            "amt": int(float(tot.get("amt") or 0)),
            "mgn": int(float(tot.get("mgn") or 0)),
        },
        "metrics": {
            "out": [{"id": s["id"], "c": s["c"], "sn": s["sn"], "cat": s["cat"],
                     "total": s["total_out"], "qty": s["qty_out"]} for s in out_metrics],
            "in":  [{"id": s["id"], "c": s["c"], "sn": s["sn"], "cat": s["cat"],
                     "total": s["total_in"],  "qty": s["qty_in"]}  for s in in_metrics],
        },
    }
    f = REPO / "data" / "sales-sku" / f"{ym}.json"
    f.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    print(f"[sku] {ym}.json 저장  out:{total_out:,}  in:{total_in:,}  amt:{data['totals']['amt']:,}  mgn:{data['totals']['mgn']:,}")


def update_index():
    sku_dir = REPO / "data" / "sales-sku"
    months = sorted(f.stem for f in sku_dir.glob("20??-??.json"))
    monthly_totals = {}
    for m in months:
        j = json.loads((sku_dir / f"{m}.json").read_text(encoding="utf-8"))
        monthly_totals[m] = j["totals"]
    idx = {
        "updated_at": dt.date.today().isoformat(),
        "source": "oa_ops.crawl_coupang_daily (출고/입고) + mart_sales_daily (매출/마진) · channel=coupang",
        "months": months,
        "monthly_totals": monthly_totals,
    }
    (sku_dir / "index.json").write_text(
        json.dumps(idx, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"[idx] {len(months)}개월 인덱스 갱신")


def refresh_inbound(mcp):
    """mart_purchase_incoming → 오아/보아르 + 하이픈≤1 필터 + 날짜별 집계 → /api/inbound-state 저장."""
    today = dt.date.today()
    # 오늘 이후 90일치
    end = today + dt.timedelta(days=90)
    sql = (
        "SELECT eta_warehouse, product_name, SUM(qty) qty "
        "FROM mart_purchase_incoming "
        f"WHERE eta_warehouse BETWEEN '{today}' AND '{end}' "
        "  AND (product_name LIKE '오아%' OR product_name LIKE '보아르%') "
        "  AND product_name NOT LIKE '%-%-%' "
        "GROUP BY eta_warehouse, product_name "
        "ORDER BY eta_warehouse, qty DESC"
    )
    res = mcp.run_query(sql)
    rows = res["rows"]

    # 날짜별 그룹
    WEEKDAYS = ["월", "화", "수", "목", "금", "토", "일"]
    by_date = {}
    for r in rows:
        d_str = str(r["eta_warehouse"])
        name = r["product_name"]
        qty = int(r["qty"] or 0)
        if qty <= 0:
            continue
        if d_str not in by_date:
            d_obj = dt.date.fromisoformat(d_str)
            label = f"{d_obj.month}/{d_obj.day:02d} ({WEEKDAYS[d_obj.weekday()]})"
            by_date[d_str] = {
                "key": f"{d_obj.month:02d}-{d_obj.day:02d}",
                "label": label,
                "items": [],
            }
        by_date[d_str]["items"].append({"name": name, "qty": qty, "foc": False})

    dates = [by_date[k] for k in sorted(by_date.keys())]
    parsed = {
        "sheetName": f"auto-{today.isoformat()}",
        "dates": dates,
    }

    body = {
        "filename": f"[자동] oa_ops.mart_purchase_incoming ({today.isoformat()})",
        "updated_at": dt.datetime.now().isoformat(),
        "parsed": parsed,
    }

    r = requests.post(
        "https://coupang-po.pages.dev/api/inbound-state",
        json=body, timeout=30,
    )
    r.raise_for_status()
    total_items = sum(len(d["items"]) for d in dates)
    total_qty = sum(it["qty"] for d in dates for it in d["items"])
    print(f"[inbound] {len(dates)}개 입고일, {total_items}품목, 총 {total_qty:,}개 → KV 저장")


def git_push():
    os.chdir(REPO)
    subprocess.run(["git", "add", "data/sales-dash.json", "data/sales-sku/"], check=True)
    r = subprocess.run(["git", "diff", "--cached", "--quiet"])
    if r.returncode == 0:
        print("[git] 변경 없음 — push 생략")
        return
    today = dt.date.today().isoformat()
    msg = f"판매 데이터 자동 갱신 ({today})"
    subprocess.run(["git", "commit", "-m", msg], check=True)
    # untracked 파일 때문에 rebase 실패하지 않도록 stash 처리
    stash = subprocess.run(
        ["git", "stash", "-u"],
        capture_output=True, text=True,
        encoding="utf-8", errors="replace",
    )
    stashed = "No local changes" not in (stash.stdout or "")
    try:
        subprocess.run(["git", "pull", "--rebase"], check=False)
        subprocess.run(["git", "push"], check=True)
    finally:
        if stashed:
            subprocess.run(["git", "stash", "pop"], check=False)
    print(f"[git] push 완료: {msg}")


def main():
    print(f"=== refresh_sales {dt.datetime.now().isoformat(timespec='seconds')} ===")
    mcp = McpSession(MCP_URL, AUTH)

    # 1) 대시보드 — 항상 갱신
    refresh_sales_dash(mcp)

    # 2) 입고리스트 — mart_purchase_incoming
    refresh_inbound(mcp)

    # 3) sales-sku — DB에 새 데이터가 있을 때만
    today = dt.date.today()
    cm = today.strftime("%Y-%m")
    db_latest = db_latest_coupang_date(mcp)
    file_latest = file_latest_nonzero_date(cm)
    print(f"[check] DB 최신 {db_latest}  로컬 최신 {file_latest}")
    if file_latest is None or (db_latest and str(db_latest) > str(file_latest)):
        refresh_sales_sku_month(mcp, cm)
        update_index()
    else:
        print("[sku] 새 데이터 없음 — 생략")

    # 3) git push
    git_push()
    print("=== 완료 ===")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        print("[err] 실패:", e, file=sys.stderr)
        traceback.print_exc()
        sys.exit(1)
