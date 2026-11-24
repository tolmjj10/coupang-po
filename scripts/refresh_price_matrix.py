#!/usr/bin/env python3
"""
refresh_price_matrix.py — oa_ops에서 12개 채널 최신가 재수집 → price-matrix 재빌드.
스케줄: 매일 08/09/10/11/12/14/15시 (작업 스케줄러에서 호출)
출력: po/price-matrix/data.json + po/price-matrix/index.html (임베드 갱신)
"""
import os, sys, json, re, datetime as dt
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

import requests

MCP_URL = os.environ.get("OA_OPS_MCP_URL", "http://192.168.30.69:8790/mcp")
AUTH    = os.environ.get("OA_OPS_AUTH",    "Bearer oaops_ro_9f3kQ2mZ7pLx8vNt3sYb6Hd4Rw")
REPO    = Path(os.environ.get("REPO_DIR",  r"C:\Users\admin\Desktop\17. claude\po"))
OUT_DIR = REPO / "price-matrix"
CACHE   = Path(r"C:\tmp")

CHANNELS = ['naver','11st','gmarket','auction','ssg','lotteon','musinsa','ohou','kurly','kakao','zigzag','29cm']
# UI 노출 순서: 쿠팡 맨 앞 + 12채널
UI_CHANNELS = ['coupang'] + CHANNELS
SHOP_HISTORY = REPO / "price" / "_shop_history.json"

URL_TPL = {
    'naver':    'https://brand.naver.com/oastore/products/{}',
    '11st':     'https://www.11st.co.kr/products/{}',
    'gmarket':  'https://item.gmarket.co.kr/Item?goodscode={}',
    'auction':  'https://itempage3.auction.co.kr/DetailView.aspx?ItemNo={}',
    'ssg':      'https://www.ssg.com/item/itemView.ssg?itemId={}',
    'lotteon':  'https://www.lotteon.com/p/product/{}',
    'musinsa':  'https://www.musinsa.com/products/{}',
    'ohou':     'https://ohou.se/productions/{}',
    'kurly':    'https://www.kurly.com/goods/{}',
    'kakao':    'https://gift.kakao.com/product/{}',
    'zigzag':   'https://zigzag.kr/catalog/products/{}',
    '29cm':     'https://product.29cm.co.kr/catalog/{}',
}


class McpSession:
    def __init__(self, url, auth):
        self.url = url
        self.sess = requests.Session()
        self.sess.headers.update({"Authorization": auth, "Content-Type": "application/json",
                                  "Accept": "application/json, text/event-stream"})
        self._sid = None; self._mid = 0; self._init()
    def _nid(self): self._mid += 1; return self._mid
    def _post(self, payload, notif=False):
        h = {"mcp-session-id": self._sid} if self._sid else {}
        r = self.sess.post(self.url, json=payload, headers=h, timeout=180); r.raise_for_status(); r.encoding="utf-8"
        sid = r.headers.get("mcp-session-id"); self._sid = sid or self._sid
        if notif or not r.text.strip(): return None
        ct = r.headers.get("content-type","")
        if ct.startswith("text/event-stream"):
            events, cur = [], []
            for line in r.text.split("\n"):
                if line.startswith("data:"): cur.append(line[5:].lstrip())
                elif line.strip()=="" and cur:
                    try: events.append(json.loads("\n".join(cur)))
                    except: pass
                    cur=[]
            if cur:
                try: events.append(json.loads("\n".join(cur)))
                except: pass
            for e in events:
                if "id" in e and e.get("id")==payload.get("id"): return e
            return events[-1] if events else None
        return r.json()
    def _init(self):
        self._post({"jsonrpc":"2.0","id":self._nid(),"method":"initialize",
                    "params":{"protocolVersion":"2024-11-05","capabilities":{},
                              "clientInfo":{"name":"refresh_price_matrix","version":"1.0"}}})
        self._post({"jsonrpc":"2.0","method":"notifications/initialized","params":{}}, notif=True)
    def run_query(self, sql, limit=500):
        resp = self._post({"jsonrpc":"2.0","id":self._nid(),"method":"tools/call",
                           "params":{"name":"run_query","arguments":{"sql":sql,"limit":limit}}})
        if "error" in resp: raise RuntimeError(f"MCP err: {resp['error']}")
        result = resp["result"]
        if result.get("isError"): raise RuntimeError(f"query err: {result}")
        return json.loads(result["content"][0]["text"])


def load_erps():
    """Cache에서 SKU→ERP 매핑 로드. 없으면 에러."""
    p = CACHE / "sku_erp_map.json"
    if not p.exists():
        raise RuntimeError(f"missing {p} — 초기 매핑 파일 필요")
    with open(p, encoding="utf-8") as f: return json.load(f)

def load_skus():
    p = CACHE / "sku_1119_clean.json"
    if not p.exists():
        raise RuntimeError(f"missing {p}")
    with open(p, encoding="utf-8") as f: return json.load(f)


def fetch_coupang_prices(mcp, sku_ids):
    """SKU → vendor_item_id → _shop_history.json 최신가 매핑 반환. {sku_id: {price, base, url, collected}}"""
    if not SHOP_HISTORY.exists():
        print(f"[warn] {SHOP_HISTORY} 없음 — 쿠팡 가격 스킵", file=sys.stderr)
        return {}
    with open(SHOP_HISTORY, encoding="utf-8") as f: hist = json.load(f)
    ids_quoted = ",".join("'"+str(int(s))+"'" for s in sku_ids)
    rows = mcp.run_query(
        f"SELECT sku_id, vendor_item_id FROM dim_coupang_sku WHERE channel_code='coupang' AND sku_id IN ({ids_quoted})"
    ).get("rows", [])
    sku2vi = {r['sku_id']: r['vendor_item_id'] for r in rows if r.get('vendor_item_id')}
    out = {}
    for sku_id, vi in sku2vi.items():
        entry = hist.get(vi)
        if not entry or not entry.get('runs'): continue
        latest_key = sorted(entry['runs'].keys())[-1]
        r = entry['runs'][latest_key]
        out[sku_id] = {
            'price': r.get('price'),
            'base': r.get('basePrice'),
            'wow': r.get('wowPrice'),
            'url': entry.get('url',''),
            'name': entry.get('name',''),
            'collected': r.get('date','') + ' ' + r.get('ts',''),
        }
    return out


def fetch_matrix(mcp, erps):
    quoted = ",".join("'"+e+"'" for e in sorted(set(erps)))
    all_rows = []
    for ch in CHANNELS:
        sql = f"""SELECT m.erp_product_code, MAX(m.erp_product_name) erp_product_name, m.channel_code, m.channel_product_no,
MAX(m.product_name) product_name,
(SELECT p.price FROM v_price_latest p WHERE p.channel_code=m.channel_code AND p.channel_product_no=m.channel_product_no ORDER BY p.collected_at DESC LIMIT 1) price,
(SELECT p.collected_at FROM v_price_latest p WHERE p.channel_code=m.channel_code AND p.channel_product_no=m.channel_product_no ORDER BY p.collected_at DESC LIMIT 1) collected_at
FROM v_erp_map_all m
WHERE m.channel_code='{ch}' AND m.use_yn='Y' AND m.erp_product_code IN ({quoted})
GROUP BY m.erp_product_code, m.channel_code, m.channel_product_no"""
        res = mcp.run_query(sql)
        rows = res.get("rows", [])
        if res.get("truncated"):
            print(f"[warn] {ch} truncated at {len(rows)}", file=sys.stderr)
        all_rows.extend(rows)
        print(f"  {ch}: {len(rows)}")
    return all_rows


def build_data(rows, sku2erp, skus, coupang_px):
    from collections import defaultdict
    erp2skus = defaultdict(list)
    for s in skus:
        sid = str(int(s['sku']))
        erp = sku2erp.get(sid)
        if erp: erp2skus[erp].append(s)

    erp_data = defaultdict(lambda: defaultdict(list)); erp_name = {}
    for r in rows:
        erp, ch, cno = r['erp_product_code'], r['channel_code'], r['channel_product_no']
        url = URL_TPL.get(ch,'').format(cno) if ch in URL_TPL else ''
        price = r.get('price')
        try: price = float(price) if price is not None else None
        except: price = None
        erp_data[erp][ch].append({'url':url,'price':price,'name':r.get('product_name') or '','cno':cno,'collected':r.get('collected_at')})
        if r.get('erp_product_name'): erp_name[erp] = r['erp_product_name']

    products=[]
    for erp, slist in sorted(erp2skus.items()):
        rep = slist[0]
        # 쿠팡: 이 ERP에 속한 SKU들의 최신가. 여러 개면 list로 보존
        coupang_items = []
        for s in slist:
            sid = str(int(s['sku']))
            cp = coupang_px.get(sid)
            if cp and cp.get('price') is not None:
                try: cp_price = float(cp['price'])
                except: cp_price = None
                coupang_items.append({'url': cp.get('url',''), 'price': cp_price, 'name': cp.get('name',''),
                                      'cno': sid, 'collected': cp.get('collected','')})
        ch_data = dict(erp_data.get(erp, {}))
        if coupang_items: ch_data['coupang'] = coupang_items
        products.append({'erp':erp,'erp_name':erp_name.get(erp,rep['name']),'sku_name':rep['name'],
                         'sku_ids':[str(int(s['sku'])) for s in slist],'cat':rep['cat'],'manager':rep['manager'],
                         'channels': ch_data})
    return {'generated_at': dt.datetime.now().isoformat(timespec='seconds'),
            'channels': UI_CHANNELS, 'products': products}


def write_outputs(data):
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    data_path = OUT_DIR / "data.json"
    html_path = OUT_DIR / "index.html"
    with open(data_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False)
    # Re-embed
    if html_path.exists():
        with open(html_path, encoding='utf-8') as f: html = f.read()
        with open(data_path, encoding='utf-8') as f: data_str = f.read()
        html2 = re.sub(r'<script>window\.__DATA__ = .*?</script>\n', '', html, count=1, flags=re.S)
        inject = f'<script>window.__DATA__ = {data_str};</script>\n<script>'
        html3 = html2.replace('<script>', inject, 1)
        with open(html_path, 'w', encoding='utf-8') as f: f.write(html3)
        print(f"[ok] updated {html_path}")
    print(f"[ok] updated {data_path}")


def main():
    print(f"[{dt.datetime.now():%Y-%m-%d %H:%M:%S}] refresh_price_matrix start")
    sku2erp = load_erps(); skus = load_skus()
    erps = sorted(set(sku2erp.values()))
    print(f"  ERPs: {len(erps)}, SKUs: {len(skus)}")
    mcp = McpSession(MCP_URL, AUTH)
    rows = fetch_matrix(mcp, erps)
    print(f"  total rows: {len(rows)}")
    coupang_px = fetch_coupang_prices(mcp, [s['sku'] for s in skus])
    print(f"  쿠팡 가격: {len(coupang_px)}건")
    data = build_data(rows, sku2erp, skus, coupang_px)
    write_outputs(data)
    print(f"[{dt.datetime.now():%Y-%m-%d %H:%M:%S}] done ({len(data['products'])} products)")


if __name__ == "__main__":
    main()
