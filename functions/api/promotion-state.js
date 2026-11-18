/**
 * Cloudflare Pages Function
 * URL: /api/promotion-state
 *
 * 프로모션 일정 관리(promotion.html) 공유 상태 저장소.
 * 기존 STATE_KV binding 재사용, 키만 분리.
 *
 * POST 동작: 필드 보존 머지
 *  - incoming 행에 있는 id 기준으로 서버 기존 행과 병합
 *  - incoming 에 **없는** 필드(key 자체가 undefined)는 서버 값 유지
 *  - incoming 에 명시적으로 ''/null/0 으로 설정된 필드는 그대로 적용 (사용자 삭제 의도 존중)
 *  - incoming 에만 있는 신규 id 는 추가, 서버에만 있는 id 는 **삭제로 간주** (툼스톤 필터는 클라이언트 담당)
 *
 * 이 머지 덕에 오래된 탭이 reviewCount 같은 신규 필드를 모른 채로 저장해도
 * 기존 값이 날아가지 않는다.
 */
import { logChange } from './_changelog.js';

const KEY = 'promotion_state';
const MAX_BYTES = 1024 * 1024;

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

function mergeRows(existingArr, incomingArr) {
  const byId = new Map();
  for (const r of existingArr) {
    if (r && r.id != null) byId.set(String(r.id), r);
  }
  const out = [];
  for (const inc of incomingArr) {
    if (!inc || inc.id == null) { out.push(inc); continue; }
    const prev = byId.get(String(inc.id));
    if (!prev) { out.push(inc); continue; }
    // prev + incoming 병합: incoming 에 있는 키만 덮어씀 (undefined 는 서버 값 유지)
    const merged = { ...prev };
    for (const k of Object.keys(inc)) {
      merged[k] = inc[k];
    }
    out.push(merged);
  }
  return out;
}

export async function onRequestPost({ request, env, waitUntil }) {
  if (!env.STATE_KV) {
    return new Response('STATE_KV binding missing', { status: 500 });
  }
  const text = await request.text();
  if (text.length > MAX_BYTES) {
    return new Response('Payload too large', { status: 413 });
  }
  let incoming;
  try { incoming = JSON.parse(text); } catch (e) {
    return new Response('Invalid JSON', { status: 400 });
  }
  if (!Array.isArray(incoming)) {
    return new Response('Root must be an array', { status: 400 });
  }

  // 기존 값 읽어서 필드 보존 머지
  const prevRaw = await env.STATE_KV.get(KEY);
  let merged = incoming;
  if (prevRaw) {
    try {
      const prev = JSON.parse(prevRaw);
      if (Array.isArray(prev)) merged = mergeRows(prev, incoming);
    } catch (_) { /* 기존이 손상돼있으면 incoming 그대로 */ }
  }

  const mergedText = JSON.stringify(merged);
  await env.STATE_KV.put(KEY, mergedText);
  waitUntil(logChange(env, '프로모션', prevRaw, mergedText));
  return new Response('ok');
}
