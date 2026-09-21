"""
일일판매재고.xlsb + data/daily/*.csv → Luckysheet 뷰어(HTML) 생성.

- xlsb → xlsx (Excel COM)
- data/daily/*.csv 자동 집계 → 9월 출고 시트 해당 날짜 컬럼에 반영 (SKU별 출고수량 합)
- xlsx 전처리:
  * 최근 7일 데이터 있는 날짜 컬럼만 남기고 나머지 숨김
  * 특정 컬럼 서식 강제 적용
- 브라우저에서 LuckyExcel + Luckysheet 로 편집/수식 지원
"""
import os, sys, io, base64, tempfile, shutil, re, csv, datetime as dt
from pathlib import Path
from collections import defaultdict

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', write_through=True)

BASE = Path(__file__).parent
SRC_XLSB = BASE / '일일판매재고.xlsb'
CSV_DIR = BASE.parent / 'data' / 'daily'   # data/daily/*.csv 를 매일 합침
# 이 폴더 자체가 웹 배포 경로 (coupang-po.pages.dev/daily/)
OUT_HTML = BASE / 'index.html'

DATE_WINDOW_DAYS = 7
EXCEL_EPOCH = dt.date(1899, 12, 30)
TODAY = dt.date.today()
DATE_SERIAL_RANGE = (40000, 60000)


IS_WINDOWS = sys.platform.startswith('win')


def xlsb_to_xlsx(src: Path) -> Path:
    """Windows: Excel COM. Linux: LibreOffice headless."""
    dst = Path(tempfile.gettempdir()) / (src.stem + '_conv.xlsx')
    if dst.exists():
        dst.unlink()
    if IS_WINDOWS:
        import win32com.client
        excel = win32com.client.DispatchEx('Excel.Application')
        excel.Visible = False; excel.DisplayAlerts = False
        try:
            wb = excel.Workbooks.Open(str(src), ReadOnly=True, UpdateLinks=0)
            for i in range(1, wb.Sheets.Count + 1):
                wb.Sheets(i).Visible = -1
            wb.SaveAs(str(dst), FileFormat=51)
            wb.Close(SaveChanges=False)
        finally:
            excel.Quit()
    else:
        import subprocess
        out_dir = Path(tempfile.gettempdir())
        subprocess.run(
            ['libreoffice', '--headless', '--calc', '--convert-to', 'xlsx', '--outdir', str(out_dir), str(src)],
            check=True, timeout=300,
        )
        conv = out_dir / (src.stem + '.xlsx')
        if conv != dst:
            shutil.move(str(conv), str(dst))
    return dst


def excel_recalc(path: Path) -> None:
    """수식 결과값을 캐시로 저장. Windows: Excel. Linux: LibreOffice."""
    if IS_WINDOWS:
        import win32com.client
        excel = win32com.client.DispatchEx('Excel.Application')
        excel.Visible = False; excel.DisplayAlerts = False
        try:
            wb = excel.Workbooks.Open(str(path), UpdateLinks=0)
            excel.CalculateFull()
            wb.Save()
            wb.Close(SaveChanges=False)
        finally:
            excel.Quit()
    else:
        import subprocess
        out_dir = path.parent
        # 임시 폴더에 변환 → 원본 덮어쓰기
        tmp_out = Path(tempfile.gettempdir()) / '_recalc_out'
        tmp_out.mkdir(parents=True, exist_ok=True)
        subprocess.run(
            ['libreoffice', '--headless', '--calc', '--convert-to', 'xlsx', '--outdir', str(tmp_out), str(path)],
            check=True, timeout=300,
        )
        result = tmp_out / (path.stem + '.xlsx')
        shutil.move(str(result), str(path))


def detect_date_columns(ws, max_col: int) -> dict:
    result = {}
    scan_rows = min(10, ws.max_row or 1)
    for c in range(1, max_col + 1):
        for r in range(1, scan_rows + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, dt.datetime):
                result[c] = v.date()
                break
            if isinstance(v, dt.date):
                result[c] = v
                break
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                iv = int(v)
                if DATE_SERIAL_RANGE[0] <= iv <= DATE_SERIAL_RANGE[1] and abs(v - iv) < 0.001:
                    result[c] = EXCEL_EPOCH + dt.timedelta(days=iv)
                    break
    return result


def column_has_data(ws, c: int, header_end_row: int = 10) -> bool:
    for r in range(header_end_row + 1, (ws.max_row or 0) + 1):
        v = ws.cell(row=r, column=c).value
        if v not in (None, '', 0):
            return True
    return False


def compute_hidden_date_cols(ws, date_cols: dict) -> set:
    if not date_cols:
        return set()
    keep = []
    for c, d in date_cols.items():
        if d > TODAY:
            continue
        if not column_has_data(ws, c):
            continue
        keep.append((d, c))
    keep.sort(reverse=True)
    keep_set = {c for _, c in keep[:DATE_WINDOW_DAYS]}
    return set(date_cols.keys()) - keep_set


def collect_headers(ws, max_col: int, up_to_row: int = 10) -> dict:
    headers = {}
    for c in range(1, max_col + 1):
        parts = []
        for r in range(1, min(up_to_row + 1, (ws.max_row or 0) + 1)):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip():
                parts.append(v.strip())
        headers[c] = ' '.join(parts)
    return headers


def decimals_format_for(header: str) -> str | None:
    if '판매가능일수' in header:
        return '#,##0.0'
    if '평균판매량' in header:
        return '#,##0'
    if any(k in header for k in ('공급가', '원가', '단가', '판매가', '금액')):
        return '#,##0'
    return None


def col_letter(c: int) -> str:
    """1-based col index → Excel letter."""
    s = ''
    while c > 0:
        c, rem = divmod(c - 1, 26)
        s = chr(65 + rem) + s
    return s


def parse_daily_csv(csv_path: Path) -> tuple[dt.date | None, dict]:
    """CSV 파싱 → (date, {sku: {'out': 합, 'in': 합, 'stock': max}})"""
    per_sku = defaultdict(lambda: {'out': 0, 'in': 0, 'stock': 0})
    date = None
    try:
        with open(csv_path, encoding='utf-8-sig', newline='') as f:
            reader = csv.DictReader(f)
            for row in reader:
                if not date:
                    d = (row.get('날짜') or '').strip()
                    try: date = dt.datetime.strptime(d, '%Y%m%d').date()
                    except ValueError: pass
                sku = (row.get('SKU ID') or '').strip()
                if not sku: continue
                try:
                    per_sku[sku]['out'] += int(row.get('출고수량') or 0)
                    per_sku[sku]['in']  += int(row.get('입고수량') or 0)
                    per_sku[sku]['stock'] = max(per_sku[sku]['stock'], int(row.get('현재재고수량') or 0))
                except ValueError:
                    pass
    except Exception as e:
        print(f'  !! CSV 파싱 실패 {csv_path.name}: {e}')
    return date, dict(per_sku)


def aggregate_csvs_into_wb(wb, wb_vals):
    """data/daily/*.csv 를 읽어 9월 출고 / 9월 입고 시트 해당 날짜 컬럼에 반영."""
    if not CSV_DIR.exists():
        print(f'  CSV 폴더 없음 (스킵): {CSV_DIR}')
        return
    csvs = sorted(CSV_DIR.glob('basic_operation_rocket_*.csv'))
    if not csvs:
        print(f'  CSV 파일 없음: {CSV_DIR}')
        return

    per_date = {}
    for cp in csvs:
        d, data = parse_daily_csv(cp)
        if d: per_date[d] = data
    if not per_date:
        print('  CSV 유효 데이터 없음')
        return
    print(f'  CSV 집계: {len(per_date)}일 ({min(per_date)} ~ {max(per_date)})')

    for sheet_name, out_field in (('9월 출고', 'out'), ('9월 입고', 'in')):
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        ws_v = wb_vals[sheet_name]
        max_col = ws.max_column or 1
        max_row = ws.max_row or 1

        # SKU ID 컬럼 + 헤더 행 탐지
        sku_col = None
        header_row = None
        for r in range(1, min(15, max_row + 1)):
            for c in range(1, max_col + 1):
                v = ws_v.cell(row=r, column=c).value
                if isinstance(v, str) and v.strip() == 'SKU ID':
                    sku_col = c; header_row = r
                    break
            if sku_col: break
        if not sku_col:
            print(f'  [{sheet_name}] SKU ID 컬럼 못 찾음')
            continue

        # 날짜 → 컬럼 매핑
        date_to_col = {}
        for c in range(1, max_col + 1):
            v = ws_v.cell(row=header_row, column=c).value
            if isinstance(v, dt.datetime): date_to_col[v.date()] = c
            elif isinstance(v, dt.date):   date_to_col[v] = c

        # SKU → 행 매핑
        sku_to_row = {}
        for r in range(header_row + 1, max_row + 1):
            v = ws_v.cell(row=r, column=sku_col).value
            if v is None: continue
            sku = str(v).strip()
            if sku and sku not in sku_to_row:
                sku_to_row[sku] = r

        written = 0; missing_dates = 0; missing_skus = set()
        for date, sku_data in per_date.items():
            col = date_to_col.get(date)
            if not col: missing_dates += 1; continue
            for sku, data in sku_data.items():
                row = sku_to_row.get(sku)
                if not row: missing_skus.add(sku); continue
                ws.cell(row=row, column=col).value = data[out_field]
                written += 1
        print(f'  [{sheet_name}] 반영 {written:,}셀 · 날짜미매칭 {missing_dates} · SKU미매칭 {len(missing_skus)}')


def preprocess_xlsx(src: Path) -> Path:
    """xlsx 전처리:
    - 수식 → 캐시된 값으로 대체 (Luckysheet formula.js 파싱 오류 회피)
    - 옛 날짜 컬럼 hidden=True
    - 지정 컬럼 서식 override
    - 주요 요약 컬럼(7일/14일/8월 평균판매량, 판매가능일수)에는 =AVERAGE / 나눗셈 수식 재삽입
      → 사용자가 날짜 셀 값을 바꾸면 자동 재계산됨
    """
    import openpyxl
    dst = Path(tempfile.gettempdir()) / '일일판매재고_prep.xlsx'
    shutil.copy(src, dst)

    wb_vals = openpyxl.load_workbook(dst, data_only=True)
    wb = openpyxl.load_workbook(dst)

    # data/daily/*.csv 를 시트에 반영 (수식 변환 전에 실행 → 이후 Excel 재계산이 값 반영)
    print('CSV 자동 반영 중...')
    aggregate_csvs_into_wb(wb, wb_vals)

    # CSV 반영 후 wb 를 값 기준으로 다시 읽어 hidden 판정용으로 사용
    # (하지만 파일에 아직 저장 전이므로 wb 자체를 직접 참조)
    # openpyxl 은 .value 가 cell.value 로 노출되므로 wb 그대로 has_data 판정 가능

    # 9월 출고 / 9월 입고 시트에서 항상 숨길 컬럼
    ALWAYS_HIDDEN = {
        '9월 출고': {16, 17},   # P: 현재고*원가, Q: 현재고*공급가(+VAT)
        '9월 입고': {16, 17},
    }

    for name in wb.sheetnames:
        ws = wb[name]
        ws_v = wb_vals[name]
        max_col = ws.max_column or 1
        max_row = ws.max_row or 1
        date_cols = detect_date_columns(ws_v, max_col)
        # CSV 반영 후의 데이터로 판정하려면 ws(값이 최신) 사용
        hidden = compute_hidden_date_cols(ws, date_cols)
        # 시트별 항상 숨길 컬럼 추가
        hidden |= ALWAYS_HIDDEN.get(name, set())
        headers = collect_headers(ws_v, max_col)

        # 수식 → 값
        formula_replaced = 0
        for r in range(1, max_row + 1):
            for c in range(1, max_col + 1):
                cell = ws.cell(row=r, column=c)
                if isinstance(cell.value, str) and cell.value.startswith('='):
                    v = ws_v.cell(row=r, column=c).value
                    cell.value = v
                    formula_replaced += 1

        # 헤더 행 찾기 (텍스트가 가장 많은 상단 15행)
        header_row = 1
        best = -1
        for r in range(1, min(16, max_row + 1)):
            cnt = sum(1 for c in range(1, max_col + 1) if isinstance(ws_v.cell(row=r, column=c).value, str))
            if cnt > best:
                best = cnt
                header_row = r
        data_start = header_row + 1

        # 데이터 있는 날짜 컬럼 (오늘 이하) 오름차순
        dated_with_data = sorted(
            [(d, c) for c, d in date_cols.items()
             if d <= TODAY and column_has_data(ws_v, c)]
        )
        # 마지막 N 개
        last7_cols = [c for _, c in dated_with_data[-7:]]
        last14_cols = [c for _, c in dated_with_data[-14:]]
        # 8월 판매량용: 8월 날짜만
        aug_cols = [c for d, c in dated_with_data if d.month == 8]
        # 9월 총합계용: 9월 날짜만
        sep_cols = [c for d, c in dated_with_data if d.month == 9]

        # 요약 컬럼 찾기
        col_targets = {}  # col idx → formula_type
        for c in range(1, max_col + 1):
            h = headers.get(c, '')
            if '7일평균판매량' in h:
                col_targets[c] = '7avg'
            elif '14일평균판매량' in h:
                col_targets[c] = '14avg'
            elif '8월평균판매량' in h:
                col_targets[c] = '8avg'
            elif '8월 판매량' in h or '8월판매량' in h:
                col_targets[c] = '8sum'
            elif '판매가능일수' in h:
                col_targets[c] = 'days'
            elif h.strip() == '총합계':
                col_targets[c] = 'sepSum'

        # 재고량 컬럼: "재고량" 정확 매칭 or "재고일" (실제 재고 수량이 들어있는 컬럼) 을 찾음.
        # 잘못된 컬럼(재고현황 = 텍스트 카테고리) 피하려고 숫자 값인 것 우선.
        stock_col = None
        for c in range(1, max_col + 1):
            h = headers.get(c, '').strip()
            if h in ('재고량', '재고'):
                stock_col = c
                break
        if stock_col is None:
            # 폴백: 헤더 '재고일' 이면서 데이터가 정수형이면 실제 수량
            for c in range(1, max_col + 1):
                h = headers.get(c, '').strip()
                if h == '재고일':
                    sample = ws_v.cell(row=data_start, column=c).value
                    if isinstance(sample, (int, float)) and not isinstance(sample, bool):
                        stock_col = c
                        break

        formulas_added = 0

        def range_or_list(cols, row):
            """연속된 컬럼이면 A:B 범위, 아니면 콤마 리스트."""
            if not cols:
                return None
            s = sorted(cols)
            if s[-1] - s[0] == len(s) - 1:
                return f'{col_letter(s[0])}{row}:{col_letter(s[-1])}{row}'
            return ','.join(f'{col_letter(cc)}{row}' for cc in s)

        for r in range(data_start, max_row + 1):
            for c, kind in col_targets.items():
                # 원래 값이 없던 행(빈 행/서브총계 등)은 건너뜀
                orig = ws_v.cell(row=r, column=c).value
                if orig is None:
                    continue

                f = None
                if kind == '7avg':
                    rng = range_or_list(last7_cols, r)
                    if rng: f = f'=AVERAGE({rng})'
                elif kind == '14avg':
                    rng = range_or_list(last14_cols, r)
                    if rng: f = f'=AVERAGE({rng})'
                elif kind == '8avg':
                    rng = range_or_list(aug_cols, r)
                    if rng: f = f'=AVERAGE({rng})'
                elif kind == '8sum':
                    rng = range_or_list(aug_cols, r)
                    if rng: f = f'=SUM({rng})'
                elif kind == 'days' and stock_col and last7_cols:
                    stock_ref = f'{col_letter(stock_col)}{r}'
                    rng = range_or_list(last7_cols, r)
                    f = f'=IF(AVERAGE({rng})=0,0,{stock_ref}/AVERAGE({rng}))'
                elif kind == 'sepSum':
                    rng = range_or_list(sep_cols, r)
                    if rng: f = f'=SUM({rng})'
                if f:
                    ws.cell(row=r, column=c).value = f
                    formulas_added += 1

        # 숨김 + 서식 (hidden=True + width=0 → LibreOffice/Luckysheet 재변환에도 유지)
        for c in range(1, max_col + 1):
            letter = col_letter(c)
            if c in hidden:
                cd = ws.column_dimensions[letter]
                cd.hidden = True
                cd.width = 0
            fmt = decimals_format_for(headers.get(c, ''))
            if fmt:
                for r in range(1, max_row + 1):
                    cell = ws.cell(row=r, column=c)
                    if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                        cell.number_format = fmt

        print(f'  전처리 [{name}] hidden dates={len(hidden)} formulas→값={formula_replaced} 재삽입={formulas_added}')
    wb.save(dst)

    # 수식 결과값 캐시 (Luckysheet 가 즉시 화면에 표시)
    excel_recalc(dst)
    print('  재계산 + 캐시 저장')
    return dst


HTML_TEMPLATE = r'''<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8">
<title>일일판매재고</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/luckysheet@2.1.13/dist/plugins/css/pluginsCss.css">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/luckysheet@2.1.13/dist/plugins/plugins.css">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/luckysheet@2.1.13/dist/css/luckysheet.css">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/luckysheet@2.1.13/dist/assets/iconfont/iconfont.css">
<style>
  html, body { margin: 0; padding: 0; height: 100%; overflow: hidden; font-family: "맑은 고딕", sans-serif; }
  #luckysheet { position: absolute; top: 0; left: 0; right: 0; bottom: 0; }
  #loader { position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); font-size: 14px; color: #444; text-align: center; z-index: 10; }
  #loader .sp { display: inline-block; width: 20px; height: 20px; border: 3px solid #ddd; border-top-color: #2563eb; border-radius: 50%; animation: spin 1s linear infinite; vertical-align: middle; margin-right: 8px; }
  @keyframes spin { to { transform: rotate(360deg); } }
</style>
</head>
<body>
<div id="loader"><span class="sp"></span>엑셀 파싱 중...</div>
<div id="luckysheet"></div>
<script src="https://cdn.jsdelivr.net/npm/luckysheet@2.1.13/dist/plugins/js/plugin.js"></script>
<script src="https://cdn.jsdelivr.net/npm/luckysheet@2.1.13/dist/luckysheet.umd.js"></script>
<script src="https://cdn.jsdelivr.net/npm/luckyexcel@1.0.1/dist/luckyexcel.umd.js"></script>
<script>
const XLSX_B64 = "__XLSX_B64__";

function b64ToUint8(b64) {
  const bin = atob(b64);
  const arr = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
  return arr;
}

function initSheet() {
  const loader = document.getElementById('loader');
  if (typeof LuckyExcel === 'undefined') {
    loader.textContent = 'LuckyExcel 로드 실패 (CDN 접근 불가?)';
    console.error('LuckyExcel undefined');
    return;
  }
  if (typeof luckysheet === 'undefined') {
    loader.textContent = 'Luckysheet 로드 실패';
    console.error('luckysheet undefined');
    return;
  }
  const bytes = b64ToUint8(XLSX_B64);
  const file = new File([bytes], 'daily.xlsx', {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
  console.log('file size:', file.size);
  LuckyExcel.transformExcelToLucky(file, (exportJson, luckysheetfile) => {
    console.log('transform callback', exportJson);
    if (!exportJson || !exportJson.sheets || !exportJson.sheets.length) {
      loader.textContent = '엑셀 파싱 실패 (콘솔 확인)';
      return;
    }
    loader.style.display = 'none';
    luckysheet.create({
      container: 'luckysheet',
      data: exportJson.sheets,
      title: '일일판매재고',
      lang: 'en',
      showinfobar: false,
      showtoolbarConfig: {
        undoRedo: true, paintFormat: true, currencyFormat: true, percentageFormat: true,
        numberDecrease: true, numberIncrease: true, moreFormats: true,
        font: true, fontSize: true, bold: true, italic: true, strikethrough: true, underline: true,
        textColor: true, fillColor: true, border: true, mergeCell: true,
        horizontalAlignMode: true, verticalAlignMode: true, textWrapMode: true, textRotateMode: true,
        image: false, link: false, chart: false, postil: true, pivotTable: false,
        function: true, frozenMode: true, sortAndFilter: true, conditionalFormat: true, dataVerification: true,
        splitColumn: true, screenshot: false, findAndReplace: true, protection: true, print: false,
      },
      showsheetbarConfig: {
        add: true, menu: true, sheet: true,
      },
      cellRightClickConfig: {
        copy: true, copyAs: true, paste: true, insertRow: true, insertColumn: true,
        deleteRow: true, deleteColumn: true, deleteCell: true, hideRow: true, hideColumn: true,
        rowHeight: true, columnWidth: true, clear: true, matrix: true, sort: true, filter: true,
        chart: false, image: false, link: false, data: true, cellFormat: true,
      },
    });
  });
}
window.addEventListener('load', () => setTimeout(initSheet, 100));
window.addEventListener('error', (e) => {
  const l = document.getElementById('loader');
  if (l) l.innerHTML = '<b>스크립트 오류:</b><br>' + (e.message || e.error) + '<br><small>' + (e.filename || '') + ':' + (e.lineno || '') + '</small>';
});
</script>
</body>
</html>'''


def main():
    print(f'변환 중: {SRC_XLSB}')
    xlsx = xlsb_to_xlsx(SRC_XLSB)
    print(f'  → xlsx ({os.path.getsize(xlsx):,} bytes)')

    print('전처리 중...')
    prep = preprocess_xlsx(xlsx)
    print(f'  → {prep} ({os.path.getsize(prep):,} bytes)')

    data = prep.read_bytes()
    b64 = base64.b64encode(data).decode('ascii')
    print(f'base64 크기: {len(b64):,} chars')

    html = HTML_TEMPLATE.replace('__XLSX_B64__', b64)
    OUT_HTML.write_text(html, encoding='utf-8')
    print(f'완료: {OUT_HTML}  ({os.path.getsize(OUT_HTML):,} bytes)')


if __name__ == '__main__':
    main()
