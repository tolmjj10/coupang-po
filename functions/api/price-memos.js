/**
 * Cloudflare Pages Function
 * URL: /api/price-memos
 *
 * SKU ID → 메모 텍스트 매핑 저장소. price 대시보드에서 상품별 메모.
 * 형식: { "12846215": "이슈 내용...", ... }
 */
const KEY = 'price_memos';
const MAX_BYTES = 512 * 1024;

export async function onRequestGet({ env }) {
  if (!env.STATE_KV) return new Response('STATE_KV binding missing', { status: 500 });
  const raw = await env.STATE_KV.get(KEY);
  return new Response(raw || '{}', {
    headers: {
      'Content-Type': 'application/json; charset=utf-8',
      'Cache-Control': 'no-store',
    },
  });
}

export async function onRequestPost({ request, env }) {
  if (!env.STATE_KV) return new Response('STATE_KV binding missing', { status: 500 });
  const text = await request.text();
  if (text.length > MAX_BYTES) return new Response('Payload too large', { status: 413 });
  try { JSON.parse(text); } catch (e) { return new Response('Invalid JSON', { status: 400 }); }
  await env.STATE_KV.put(KEY, text);
  return new Response('ok');
}
