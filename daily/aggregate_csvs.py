"""
data/daily/basic_operation_rocket_YYYYMMDD*.csv 파일들을 읽어
9월 출고 시트의 해당 날짜 컬럼에 SKU별 출고수량을 자동 반영.

- 여러 CSV 를 한번에 처리
- SKU ID 기준으로 xlsx 의 각 행 매칭
- 같은 날짜에 여러 센터(FC/RC)가 있으면 출고수량 합산
- 기존 값 덮어씀
"""
import csv, sys, io, re, datetime as dt
from pathlib import Path
from collections import defaultdict

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)

REPO_ROOT = Path(__file__).parent.parent   # po/
CSV_DIR = REPO_ROOT / 'data' / 'daily'
XLSB_PATH = REPO_ROOT / 'daily' / '일일판매재고.xlsb'
XLSX_OUT = Path.home() / 'AppData' / 'Local' / 'Temp' / '일일판매재고_updated.xlsx'


def parse_csv(csv_path: Path) -> tuple[dt.date, dict]:
    """CSV 하나 파싱 → (date, {sku_id: 출고수량_합, 재고수량, 입고수량})"""
    per_sku = defaultdict(lambda: {'출고수량': 0, '입고수량': 0, '재고수량': 0})
    date = None
    with open(csv_path, encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        for row in reader:
            d_str = row.get('날짜', '').strip()
            if d_str and not date:
                try:
                    date = dt.datetime.strptime(d_str, '%Y%m%d').date()
                except ValueError:
                    pass
            sku = row.get('SKU ID', '').strip()
            if not sku:
                continue
            try:
                out_qty = int(row.get('출고수량', '0') or 0)
                in_qty = int(row.get('입고수량', '0') or 0)
                stock = int(row.get('현재재고수량', '0') or 0)
            except ValueError:
                out_qty = in_qty = stock = 0
            per_sku[sku]['출고수량'] += out_qty
            per_sku[sku]['입고수량'] += in_qty
            # 재고수량은 가장 큰 값(마지막 FC 재고) 유지 — 합산 아님
            per_sku[sku]['재고수량'] = max(per_sku[sku]['재고수량'], stock)
    return date, dict(per_sku)


def xlsb_to_xlsx() -> Path:
    """Windows Excel COM 으로 xlsb → xlsx 변환 (임시)."""
    import win32com.client
    dst = Path.home() / 'AppData' / 'Local' / 'Temp' / '일일판매재고_from_xlsb.xlsx'
    if dst.exists(): dst.unlink()
    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = False; excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(XLSB_PATH), ReadOnly=True, UpdateLinks=0)
        for i in range(1, wb.Sheets.Count + 1):
            wb.Sheets(i).Visible = -1
        wb.SaveAs(str(dst), FileFormat=51)
        wb.Close(SaveChanges=False)
    finally:
        excel.Quit()
    return dst


def main():
    csvs = sorted(CSV_DIR.glob('basic_operation_rocket_*.csv'))
    print(f'CSV 파일 {len(csvs)}개 발견')

    # 파일별로 집계
    per_date = {}  # date → {sku: {출고수량, ...}}
    for csv_path in csvs:
        date, data = parse_csv(csv_path)
        if date:
            per_date[date] = data
    print(f'수집된 날짜: {min(per_date)} ~ {max(per_date)} ({len(per_date)}일)')

    # xlsb 로드
    print('xlsb → xlsx 변환 중...')
    xlsx_src = xlsb_to_xlsx()

    import openpyxl
    wb = openpyxl.load_workbook(xlsx_src)
    ws = wb['9월 출고']
    max_col = ws.max_column
    max_row = ws.max_row

    # 헤더 행 찾기 (SKU ID 있는 행)
    sku_col = None
    header_row = None
    for r in range(1, min(15, max_row + 1)):
        for c in range(1, max_col + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip() == 'SKU ID':
                sku_col = c
                header_row = r
                break
        if sku_col: break
    if not sku_col:
        print('!! SKU ID 컬럼 못 찾음')
        return
    print(f'SKU 컬럼: {sku_col} (헤더 행 {header_row})')

    # 날짜 → 컬럼 매핑 (헤더 행)
    date_to_col = {}
    for c in range(1, max_col + 1):
        v = ws.cell(row=header_row, column=c).value
        if isinstance(v, dt.datetime):
            date_to_col[v.date()] = c
        elif isinstance(v, dt.date):
            date_to_col[v] = c
    print(f'날짜 컬럼: {len(date_to_col)}개')

    # SKU → 행 매핑
    sku_to_row = {}
    for r in range(header_row + 1, max_row + 1):
        v = ws.cell(row=r, column=sku_col).value
        if v is None: continue
        sku = str(v).strip()
        if sku and sku not in sku_to_row:
            sku_to_row[sku] = r
    print(f'xlsx 내 SKU 행: {len(sku_to_row)}개')

    # 값 반영
    written = 0
    missing_sku = set()
    for date, sku_data in per_date.items():
        col = date_to_col.get(date)
        if not col:
            print(f'  ! 날짜 {date} 컬럼 없음 (스킵)')
            continue
        for sku, data in sku_data.items():
            row = sku_to_row.get(sku)
            if not row:
                missing_sku.add(sku)
                continue
            ws.cell(row=row, column=col).value = data['출고수량']
            written += 1
    print(f'셀 업데이트: {written:,}개, 매칭 안 된 SKU: {len(missing_sku)}개')

    wb.save(XLSX_OUT)
    print(f'저장: {XLSX_OUT}')


if __name__ == '__main__':
    main()
