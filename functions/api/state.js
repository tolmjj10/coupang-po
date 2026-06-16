const KEY = 'state';

  export async function onRequestGet({ env }) {
    if (!env.STATE_KV) {
      return json({ error: 'KV not bound — STATE_KV 환경변수 확인 필요' }, 500);
    }
    const data = await env.STATE_KV.get(KEY);
    return new Response(data || '{}', {
      headers: { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' }
    });
  }

  export async function onRequestPost({ request, env }) {
    if (!env.STATE_KV) {
      return json({ error: 'KV not bound' }, 500);
    }
    const text = await request.text();
    try { JSON.parse(text); } catch { return json({ error: 'bad json' }, 400); }
    if (text.length > 1_000_000) return json({ error: 'too large' }, 413);
    await env.STATE_KV.put(KEY, text);
    return json({ ok: true });
  }

  function json(obj, status = 200) {
    return new Response(JSON.stringify(obj), {
      status,
      headers: { 'Content-Type': 'application/json' }
    });
  }
