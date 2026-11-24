"""
정산 매핑 빌더: 폴더 하나(mapping/)에 아래 두 종류 파일을 두면 camp2viid/viid2sku를 생성해 KV(/api/settlement-maps)에 PUT.

입력 파일 (여러 월 넣어도 union 집계됨):
  - 매출성장광고보고서_*.xlsx  → 캠페인ID × VIID
  - 일간종합성과지표_*.csv     → VIID × SKU ID (TAB 구분, cp949)

실행: python build_maps.py  (또는 build_maps.bat)
"""
import os, sys, re, json, csv, urllib.request, urllib.error
from pathlib import Path
import openpyxl

HERE = Path(__file__).parent
FOLDER = HERE / 'mapping'
DAILY_FOLDER = HERE / '1'  # dailySettlement xlsx — 캠페인명 뒤 _SKUID 파싱용
API = 'https://coupang-po.pages.dev/api/settlement-maps'

SKU_SUFFIX_RE = re.compile(r'_(\d{5,})\s*$')  # 캠페인명 뒤 _<5자리이상숫자>

def to_str_id(v):
    if v is None: return None
    try:
        return str(int(v))
    except Exception:
        s = str(v).strip()
        return s or None

def parse_growth(path, camp2viid, camp_name):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    header_skipped = False
    n = 0
    for row in ws.iter_rows(values_only=True):
        if not header_skipped:
            header_skipped = True
            continue
        cid = to_str_id(row[3])
        cname = row[4]
        viid = to_str_id(row[7])
        if not cid or not viid: continue
        camp2viid.setdefault(cid, set()).add(viid)
        if cname and cid not in camp_name:
            camp_name[cid] = str(cname).strip()
        n += 1
    wb.close()
    return n

def parse_daily_settlement(path, camp_name_out):
    """dailySettlement xlsx에서 캠페인 ID→캠페인명 수집 (manual_camp_sku 유추용)"""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    n = 0
    for row in ws.iter_rows(min_row=3, values_only=True):
        if len(row) < 9: continue
        cid_raw = row[7]
        if not isinstance(cid_raw, (int, float)): continue
        cid = str(int(cid_raw))
        name = str(row[8] or '').strip()
        if name and cid not in camp_name_out:
            camp_name_out[cid] = name
            n += 1
    wb.close()
    return n

def parse_daily(path, viid2sku):
    n = 0
    with open(path, encoding='cp949', errors='replace') as f:
        rd = csv.reader(f, delimiter='\t')
        for i, r in enumerate(rd):
            if i == 0: continue
            if len(r) < 8: continue
            sku = to_str_id(r[3])
            name = (r[4] or '').strip()
            viid = to_str_id(r[5])
            if not sku or not viid: continue
            if viid not in viid2sku:
                viid2sku[viid] = {'sku': sku, 'name': name}
            n += 1
    return n

UA = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) build_maps/1.0'}

def fetch_existing():
    try:
        req = urllib.request.Request(API, headers=UA)
        with urllib.request.urlopen(req, timeout=15) as r:
            raw = r.read().decode('utf-8')
        return json.loads(raw) if raw and raw != 'null' else {}
    except Exception as e:
        print(f'[warn] 기존 KV 조회 실패: {e}')
        return {}

def main():
    if not FOLDER.exists():
        print(f'폴더 없음: {FOLDER}  (생성 후 파일을 넣으세요)')
        FOLDER.mkdir(parents=True, exist_ok=True)
        return 1

    growth_files = sorted(FOLDER.glob('매출성장광고보고서*.xlsx')) + sorted(FOLDER.glob('매출성장*.xlsx'))
    growth_files = list(dict.fromkeys(growth_files))
    daily_files = sorted(FOLDER.glob('일간종합성과지표*.csv')) + sorted(FOLDER.glob('일간종합*.csv'))
    daily_files = list(dict.fromkeys(daily_files))

    if not growth_files:
        print(f'[err] 매출성장광고보고서_*.xlsx 없음 in {FOLDER}')
        return 1
    if not daily_files:
        print(f'[err] 일간종합성과지표_*.csv 없음 in {FOLDER}')
        return 1

    camp2viid = {}; camp_name = {}
    for p in growth_files:
        n = parse_growth(p, camp2viid, camp_name)
        print(f'[1] {p.name}: {n}행')
    viid2sku = {}
    for p in daily_files:
        n = parse_daily(p, viid2sku)
        print(f'[2] {p.name}: {n}행')

    print(f'합계: 캠페인 {len(camp2viid)}개, VIID 총 {sum(len(v) for v in camp2viid.values())}개, VIID→SKU {len(viid2sku)}개')

    existing = fetch_existing()
    sku2cat = existing.get('sku2cat', {})
    manual_camp_sku = dict(existing.get('manual_camp_sku', {}))  # 기존 값 보존

    # dailySettlement 전체 xlsx에서 캠페인명 수집 → 뒤 _SKUID 파싱
    if DAILY_FOLDER.exists():
        daily_xlsx = [p for p in sorted(DAILY_FOLDER.glob('*dailySettlement*.xlsx'))
                      if not p.name.startswith('~$')]
        for p in daily_xlsx:
            n = parse_daily_settlement(p, camp_name)
            print(f'[0] {p.name}: 캠페인명 +{n}')

    # 캠페인명 뒤 _<숫자> 접미사 → manual_camp_sku 자동 추가
    auto_added = 0; already = 0
    for cid, nm in camp_name.items():
        if cid in manual_camp_sku:
            already += 1
            continue
        m = SKU_SUFFIX_RE.search(nm)
        if not m: continue
        sku_id = m.group(1)
        if sku_id in sku2cat:
            manual_camp_sku[cid] = {'skus': [sku_id]}
            auto_added += 1
    print(f'manual_camp_sku: 자동추가 {auto_added}건, 기존유지 {already}건, 총 {len(manual_camp_sku)}건')

    merged = {
        'updated_at': __import__('datetime').datetime.now().strftime('%Y-%m-%d %H:%M'),
        'camp2viid': {k: sorted(v) for k, v in camp2viid.items()},
        'camp_name': camp_name,
        'viid2sku': viid2sku,
        'sku2cat': sku2cat,
        'manual_camp_sku': manual_camp_sku,
        'camp_viid_cost': {},
    }

    body = json.dumps(merged, ensure_ascii=False).encode('utf-8')
    print(f'POST {API} ({len(body):,} bytes)')
    req = urllib.request.Request(API, data=body, method='POST',
                                 headers={'Content-Type': 'application/json; charset=utf-8', **UA})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            print('응답:', r.status, r.read().decode('utf-8'))
    except urllib.error.HTTPError as e:
        print('HTTPError:', e.code, e.read().decode('utf-8', errors='replace'))
        return 1
    print('완료')
    return 0

if __name__ == '__main__':
    sys.exit(main())
