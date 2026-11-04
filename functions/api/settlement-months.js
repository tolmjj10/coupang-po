/**
 * Cloudflare Pages Function
 * URL: /api/settlement-months
 *
 * 월별 상계처리 + 월별 cpc/coupon/sales 데이터 전체를 KV에 공유.
 * (localStorage의 'settle_months_v1' 전체 스토어를 동기화)
 * STATE_KV 재사용, 키 분리.
 */
const KEY = 'settlement_months';
const MAX_BYTES = 20 * 1024 * 1024;  // 20MB (10개월분 전체)

export async function onRequestGet({ env }) {
  if (!env.STATE_KV) return new Response('STATE_KV binding missing', { status: 500 });
  const raw = await env.STATE_KV.get(KEY);
  return new Response(raw || 'null', {
    headers: {
      'Content-Type': 'application/json; charset=utf-8',
      'Cache-Control': 'no-store'
    }
  });
}

export async function onRequestPost({ request, env }) {
  if (!env.STATE_KV) return new Response('STATE_KV binding missing', { status: 500 });
  const text = await request.text();
  if (text.length > MAX_BYTES) return new Response('Payload too large', { status: 413 });
  try { JSON.parse(text); } catch (e) {
    return new Response('Invalid JSON', { status: 400 });
  }
  await env.STATE_KV.put(KEY, text);
  return new Response('ok');
}
