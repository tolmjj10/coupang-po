/**
 * Cloudflare Pages Function
 * URL: /api/repo-files?path=<dir>&ref=<branch>
 *
 * GitHub Contents API 프록시.
 * 목적: 클라이언트에서 직접 api.github.com을 호출하면 비인증 60건/시간 제한에 걸림.
 *       서버에서 인증 토큰을 붙여 프록시하면 5000건/시간으로 격상.
 *
 * 사용법: 클라이언트가 '/api/repo-files?path=data&ref=main' 로 호출 →
 *         GitHub의 https://api.github.com/repos/{OWNER}/{REPO}/contents/data?ref=main 응답을 그대로 반환.
 *
 * 환경변수(Cloudflare Pages Settings → Environment variables):
 *   GITHUB_TOKEN — Fine-grained PAT, Contents:Read only, coupang-po repo만 허용
 *   (미설정 시 비인증으로 시도 → 60건/시간 제한 그대로)
 */
const OWNER = 'tolmjj10';
const REPO = 'coupang-po';
const CACHE_TTL_SECONDS = 60;
const KV_KEY_PREFIX = 'repo_files:';

export async function onRequestGet({ request, env }) {
  const url = new URL(request.url);
  const path = url.searchParams.get('path') || '';
  const ref = url.searchParams.get('ref') || 'main';

  // 경로 화이트리스트 (오픈 프록시 방지)
  if (!/^[\w./-]*$/.test(path) || path.includes('..')) {
    return jsonErr(400, 'Invalid path');
  }
  if (!/^[\w./-]+$/.test(ref)) {
    return jsonErr(400, 'Invalid ref');
  }

  const cacheKey = KV_KEY_PREFIX + ref + ':' + path;

  // 60초 KV 캐시 (여러 팀원이 동시에 열어도 GitHub API는 1분에 한 번만)
  if (env.STATE_KV) {
    try {
      const cached = await env.STATE_KV.get(cacheKey);
      if (cached) {
        return new Response(cached, {
          headers: {
            'Content-Type': 'application/json; charset=utf-8',
            'Cache-Control': 'no-store',
            'X-Cache': 'HIT'
          }
        });
      }
    } catch (e) { /* KV 캐시 실패는 무시하고 원본 요청 */ }
  }

  const ghUrl = 'https://api.github.com/repos/' + OWNER + '/' + REPO +
                '/contents/' + encodeURI(path) + '?ref=' + encodeURIComponent(ref);
  const headers = {
    'Accept': 'application/vnd.github+json',
    'User-Agent': 'coupang-po-pages-fn'
  };
  if (env.GITHUB_TOKEN) {
    headers['Authorization'] = 'Bearer ' + env.GITHUB_TOKEN;
  }

  const ghRes = await fetch(ghUrl, { headers });
  const body = await ghRes.text();

  if (ghRes.ok && env.STATE_KV) {
    try { await env.STATE_KV.put(cacheKey, body, { expirationTtl: CACHE_TTL_SECONDS }); } catch (e) {}
  }

  return new Response(body, {
    status: ghRes.status,
    headers: {
      'Content-Type': 'application/json; charset=utf-8',
      'Cache-Control': 'no-store',
      'X-Cache': 'MISS'
    }
  });
}

function jsonErr(status, message) {
  return new Response(JSON.stringify({ error: message }), {
    status,
    headers: { 'Content-Type': 'application/json; charset=utf-8' }
  });
}
