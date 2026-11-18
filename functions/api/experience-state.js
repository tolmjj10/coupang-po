/**
 * Cloudflare Pages Function
 * URL: /api/experience-state
 *
 * 체험단(new.html 체험단 뷰) 공유 상태 저장소.
 * STATE_KV binding 재사용, 키만 분리.
 */
import { logChange } from './_changelog.js';

const KEY = 'experience_state';
const MAX_BYTES = 4 * 1024 * 1024;

export async function onRequestGet({ env }) {
  if (!env.STATE_KV) {
    return new Response('STATE_KV binding missing', { status: 500 });
  }
  const raw = await env.STATE_KV.get(KEY);
  return new Response(raw || '[]', {
    headers: {
      'Content-Type': 'application/json; charset=utf-8',
      'Cache-Control': 'no-store'
    }
  });
}

export async function onRequestPost({ request, env, waitUntil }) {
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
  const prev = await env.STATE_KV.get(KEY);
  await env.STATE_KV.put(KEY, text);
  waitUntil(logChange(env, '체험단', prev, text));
  return new Response('ok');
}
