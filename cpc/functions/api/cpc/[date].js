/**
 * Cloudflare Pages Function
 * URL: /api/cpc/YYYYMMDD
 *
 * KSK의 coupang-dashboard.pages.dev/data/cpc/YYYYMMDD.xlsx 를
 * 서버 사이드에서 fetch하여 CORS 우회 없이 되돌려준다.
 *
 * KSK의 Pages는 없는 경로에도 SPA fallback(index.html)을 200으로 반환하므로,
 * 실제로 XLSX인지 확인 후 아니면 404를 반환한다.
 */
export async function onRequest(context) {
  const { date } = context.params;

  if (!/^\d{8}$/.test(date)) {
    return new Response('Invalid date', { status: 400 });
  }

  const upstream = `https://coupang-dashboard.pages.dev/data/cpc/${date}.xlsx`;

  const r = await fetch(upstream, {
    cf: { cacheTtl: 3600, cacheEverything: true },
  });

  if (!r.ok) {
    return new Response('Not found', { status: 404 });
  }

  // Content-Type이 HTML이면 SPA fallback → 파일 없음
  const ct = (r.headers.get('content-type') || '').toLowerCase();
  if (ct.includes('text/html') || ct.includes('xhtml')) {
    return new Response('Not found (fallback html)', { status: 404 });
  }

  const buf = await r.arrayBuffer();

  // 매직 바이트 검사: XLSX는 ZIP이므로 'PK' (0x50 0x4B) 로 시작해야 함
  if (buf.byteLength < 4) {
    return new Response('Empty', { status: 404 });
  }
  const head = new Uint8Array(buf, 0, 2);
  if (head[0] !== 0x50 || head[1] !== 0x4B) {
    return new Response('Not xlsx', { status: 404 });
  }

  return new Response(buf, {
    status: 200,
    headers: {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Access-Control-Allow-Origin': '*',
      'Cache-Control': 'public, max-age=3600',
    },
  });
}
