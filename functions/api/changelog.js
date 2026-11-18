/**
 * GET /api/changelog
 *
 * STATE_KV 의 'changelog' 키에 적재된 최근 변경 이력을 반환.
 * 쓰기는 각 state POST 엔드포인트가 _changelog.js logChange() 로 수행하므로
 * 여기선 읽기만 지원.
 */
const LOG_KEY = 'changelog';

export async function onRequestGet({ env }) {
  if (!env.STATE_KV) {
    return new Response('STATE_KV binding missing', { status: 500 });
  }
  const raw = await env.STATE_KV.get(LOG_KEY);
  return new Response(raw || '[]', {
    headers: {
      'Content-Type': 'application/json; charset=utf-8',
      'Cache-Control': 'no-store'
    }
  });
}
