/**
 * Cloudflare Pages Function
 * URL: /api/settlement-data
 *
 * 정산 탭에 업로드된 cpcData/couponData/salesData를 KV에 저장해서
 * 다른 사용자가 URL만 열어도 동일한 결과를 볼 수 있게 한다.
 * STATE_KV 재사용, 키 분리.
 */
const KEY = 'settlement_data';
const MAX_BYTES = 5 * 1024 * 1024;  // 5MB

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
