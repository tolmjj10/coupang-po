/**
 * Cloudflare Pages Function
 * URL: /api/price-state
 *
 * 마이샵 판매가 이력 저장소.
 * 형식: { history: {...상품별 이력...}, order: [...최신 페이지 순서 키...] }
 */
const KEY = 'price_state';
const MAX_BYTES = 2 * 1024 * 1024;  // 2MB

export async function onRequestGet({ env }) {
  if (!env.STATE_KV) {
    return new Response('STATE_KV binding missing', { status: 500 });
  }
  const raw = await env.STATE_KV.get(KEY);
  return new Response(raw || '{"history":{},"order":[]}', {
    headers: {
      'Content-Type': 'application/json; charset=utf-8',
      'Cache-Control': 'no-store'
    }
  });
}

export async function onRequestPost({ request, env }) {
  if (!env.STATE_KV) {
    return new Response('STATE_KV binding missing', { status: 500 });
  }
  const text = await request.text();
  if (text.length > MAX_BYTES) {
    return new Response('Payload too large', { status: 413 });
  }
  try { JSON.parse(text); } catch (e) {
    return new Response('Invalid JSON', { status: 400 });
  }
  await env.STATE_KV.put(KEY, text);
  return new Response('ok');
}
