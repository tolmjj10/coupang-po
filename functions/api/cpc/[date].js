 /**
   * Cloudflare Pages Function
   * URL: /api/cpc/YYYYMMDD
   *
   * KSK의 coupang-dashboard.pages.dev/data/cpc/YYYYMMDD.xlsx 를
   * 서버 사이드에서 fetch하여 CORS 우회 없이 되돌려준다.
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

    const buf = await r.arrayBuffer();

    return new Response(buf, {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Access-Control-Allow-Origin': '*',
        'Cache-Control': 'public, max-age=3600',
      },
    });
  }
