#!/usr/bin/env python3
"""
refresh_inbound_xlsx.py — GW 그룹웨어에서 최신 입고리스트 엑셀 다운로드 후 파싱해 /api/inbound-state KV 업데이트.

흐름:
  1) ego-browser로 download_gw_inbound.mjs 실행 → RESULT_JSON에서 파일 경로 획득
  2) 파일 openpyxl로 파싱 (inbound.html parseWorkbook JS 로직 포팅)
  3) 마진테이블(비단종) + 신제품 허용 필터
  4) /api/inbound-state POST
"""
import os, sys, re, json, subprocess, datetime as dt
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception: pass

import requests
import openpyxl

REPO = Path(os.environ.get("REPO_DIR", r"C:\Users\admin\Desktop\17. claude\po"))
MJS = REPO / "daily-sales" / "download_gw_inbound.mjs"


# ───────── ego-browser 실행 ─────────

def download_xlsx():
    url = "file:///" + str(MJS).replace("\\", "/").replace(" ", "%20")
    cmd = ["ego-browser", "nodejs", "-e", f"await import('{url}')"]
    print(f"[dl] {' '.join(cmd)}")
    # ego-browser는 Windows에서 stdout/stderr 모두 사용 — 둘 다 캡처
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=240)
    combined = (r.stdout or "") + "\n" + (r.stderr or "")
    print(combined)
    if r.returncode != 0:
        raise RuntimeError(f"ego-browser 실패 rc={r.returncode}")
    for line in combined.splitlines():
        if line.startswith("RESULT_JSON:"):
            return json.loads(line[len("RESULT_JSON:"):])
    raise RuntimeError("RESULT_JSON 라인 없음")


# ───────── xlsx 파서 (inbound.html parseWorkbook JS 로직 포팅) ─────────

def clean_name(raw):
    if not raw: return ""
    s = str(raw).split("\n")[0].strip()
    s = re.sub(r"\s*[\(（][A-Z0-9\-]+[\)）]\s*$", "", s).strip()
    s = re.sub(r"\s*무료제공분\s*$", "", s).strip()
    return s

def parse_date(cell):
    if not cell: return None
    matches = list(re.finditer(r"(\d{1,2})/(\d{1,2})", str(cell)))
    if not matches: return None
    m, d = int(matches[-1].group(1)), int(matches[-1].group(2))
    return {"m": m, "d": d, "key": f"{m:02d}-{d:02d}"}

WEEKDAYS = ["월", "화", "수", "목", "금", "토", "일"]
def label_for(date):
    today = dt.date.today()
    year = today.year
    if date["m"] < today.month - 2:
        year += 1
    try:
        wd = dt.date(year, date["m"], date["d"]).weekday()
    except ValueError:
        wd = 0
    return f"{date['m']}/{date['d']:02d} ({WEEKDAYS[wd]})"

def parse_workbook(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
    sheet_name = wb.sheetnames[-1]
    ws = wb[sheet_name]
    aoa = [list(row) for row in ws.iter_rows(values_only=True)]
    print(f"[parse] sheet='{sheet_name}'  {len(aoa)}행")

    # 1) "입고 예정일" 포함 셀 블록 추출
    blocks = []
    for r, row in enumerate(aoa):
        for c, v in enumerate(row or []):
            if v and "입고 예정일" in str(v):
                d = parse_date(v)
                if d:
                    blocks.append({"row": r, "col": c, "date": d})

    # 2) 블록 좌표: nameCol = col-1, qtyCol = nameCol+4, subCol = nameCol+1
    for b in blocks:
        b["nameCol"] = max(0, b["col"] - 1)
        b["qtyCol"] = b["nameCol"] + 4
        b["subCol"] = b["nameCol"] + 1

    by_date = {}      # dateKey → {name: {qty, foc}}
    date_label = {}   # dateKey → "11/24 (화)"

    for idx, b in enumerate(blocks):
        # 같은 nameCol(±2) 다음 블록까지의 수직 범위
        end_row = len(aoa)
        for j in range(idx + 1, len(blocks)):
            nb = blocks[j]
            if abs(nb["nameCol"] - b["nameCol"]) <= 2 and nb["row"] > b["row"]:
                end_row = nb["row"]
                break
        for r in range(b["row"] + 1, end_row):
            row = aoa[r] if r < len(aoa) else []
            if not row: continue
            def cell(i):
                return row[i] if 0 <= i < len(row) else None
            name_cell = cell(b["nameCol"])
            sub_cell = cell(b["subCol"]) or ""
            qty_cell = cell(b["qtyCol"])
            if not name_cell: continue
            name = clean_name(name_cell)
            if not name: continue
            if re.match(r"^(제품|구분|입수단위|수량|카톤수)$", name): continue
            if re.search(r"총\s*\d+\s*박스", str(name_cell)) or re.match(r"^\d+\s*박스", name): continue
            if re.search(r"NO\s*SALE", str(sub_cell), re.I) or "(부속)" in str(name_cell): continue
            try:
                qty = int(float(str(qty_cell).replace(",", "")))
            except (ValueError, TypeError):
                continue
            if not qty: continue
            date_key = b["date"]["key"]
            if date_key not in by_date:
                by_date[date_key] = {}
                date_label[date_key] = label_for(b["date"])
            bucket = by_date[date_key]
            is_foc = bool(re.search(r"무료제공분|FOC.*ADD", f"{name_cell} {sub_cell}", re.I))
            prev = bucket.get(name, {"qty": 0, "foc": False})
            prev["qty"] += qty
            if is_foc: prev["foc"] = True
            bucket[name] = prev

    dates = []
    for k in sorted(by_date.keys()):
        items = [{"name": n, "qty": v["qty"], "foc": v["foc"]}
                 for n, v in by_date[k].items()]
        items.sort(key=lambda x: -x["qty"])
        dates.append({"key": k, "label": date_label[k], "items": items})
    return {"sheetName": sheet_name, "dates": dates}


# ───────── 허용 상품명 필터 (마진 비단종 + 신제품) ─────────

def norm(s):
    return (s or "").strip().replace(" ", "").lower()

def load_allowed():
    allowed = set()
    try:
        mg = requests.get("https://coupang-po.pages.dev/api/margin-state", timeout=15).json()
        mg_rows = mg if isinstance(mg, list) else mg.get("rows", [])
        for r in mg_rows:
            nm, status = (r.get("name") or "").strip(), (r.get("status") or "").strip()
            if nm and status != "단종":
                allowed.add(norm(nm))
    except Exception as e:
        print(f"[allow] margin 로드 실패: {e}")
    try:
        ns = requests.get("https://coupang-po.pages.dev/api/new-state", timeout=15).json()
        ns_rows = ns if isinstance(ns, list) else ns.get("rows", [])
        for r in ns_rows:
            for key in ("realName", "regName"):
                nm = (r.get(key) or "").strip()
                if nm: allowed.add(norm(nm))
    except Exception as e:
        print(f"[allow] new 로드 실패: {e}")
    print(f"[allow] 허용 상품명 {len(allowed)}개")
    return allowed


def filter_parsed(parsed, allowed):
    if not allowed:
        return parsed
    out = {"sheetName": parsed["sheetName"], "dates": []}
    kept_items = 0; dropped_items = 0
    for d in parsed["dates"]:
        new_items = []
        for it in d["items"]:
            if norm(it["name"]) in allowed:
                new_items.append(it); kept_items += 1
            else:
                dropped_items += 1
        if new_items:
            out["dates"].append({**d, "items": new_items})
    print(f"[filter] 유지 {kept_items} · 제외 {dropped_items}")
    return out


# ───────── KV POST ─────────

def post_kv(parsed, filename):
    body = {
        "filename": f"[자동] {filename}",
        "updated_at": dt.datetime.now().isoformat(),
        "parsed": parsed,
    }
    r = requests.post(
        "https://coupang-po.pages.dev/api/inbound-state",
        data=json.dumps(body, ensure_ascii=False).encode("utf-8"),
        headers={"Content-Type": "application/json; charset=utf-8"},
        timeout=30,
    )
    r.raise_for_status()
    tot_items = sum(len(d["items"]) for d in parsed["dates"])
    tot_qty = sum(it["qty"] for d in parsed["dates"] for it in d["items"])
    print(f"[kv] {len(parsed['dates'])}일, {tot_items}품목, 총 {tot_qty:,}개 → KV 저장")


def main():
    print(f"=== refresh_inbound_xlsx {dt.datetime.now().isoformat(timespec='seconds')} ===")
    info = download_xlsx()
    print(f"[dl] 저장: {info['file']} ({info['size']:,}B)  — {info.get('topTitle')}")
    parsed = parse_workbook(info["file"])
    allowed = load_allowed()
    parsed = filter_parsed(parsed, allowed)
    post_kv(parsed, info["filename"])
    print("=== 완료 ===")


if __name__ == "__main__":
    try: main()
    except Exception as e:
        import traceback
        print("[err] 실패:", e, file=sys.stderr)
        traceback.print_exc()
        sys.exit(1)
