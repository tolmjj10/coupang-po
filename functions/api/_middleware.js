/**
 * Cloudflare Pages Middleware — 접근 인증
 *
 * 정책:
 *  1) CF-Connecting-IP 가 env.ALLOWED_IPS 목록에 있으면 통과 (사무실 IP)
 *  2) 그 외에는 서명된 쿠키(po_auth) 필요
 *  3) 쿠키 없으면 /login 로그인 화면 표시. 비밀번호 맞으면 쿠키 발급
 *
 * 필요 환경변수 (Cloudflare Pages → Settings → Environment variables):
 *  - SITE_PASSWORD  : 로그인 비밀번호
 *  - COOKIE_SECRET  : 쿠키 서명용 비밀 문자열 (32자 이상 랜덤 권장)
 *  - ALLOWED_IPS    : 쉼표구분 IP 또는 CIDR (예: "203.0.113.7, 203.0.113.0/24")
 */

const COOKIE_NAME = 'po_auth';
const MAX_AGE_SEC = 60 * 60 * 24 * 7; // 7일

function getClientIP(request) {
  return request.headers.get('CF-Connecting-IP') || '';
}

function parseList(raw) {
  return String(raw || '').split(',').map(s => s.trim()).filter(Boolean);
}

function ipv4ToInt(ip) {
  const parts = ip.split('.');
  if (parts.length !== 4) return null;
  let n = 0;
  for (const p of parts) {
    const v = parseInt(p, 10);
    if (isNaN(v) || v < 0 || v > 255) return null;
    n = ((n << 8) | v) >>> 0;
  }
  return n;
}

function cidrMatchV4(ip, cidr) {
  const [range, bitsStr] = cidr.split('/');
  const bits = parseInt(bitsStr, 10);
  if (isNaN(bits) || bits < 0 || bits > 32) return false;
  const ipInt = ipv4ToInt(ip);
  const rangeInt = ipv4ToInt(range);
  if (ipInt == null || rangeInt == null) return false;
  const mask = bits === 0 ? 0 : (0xffffffff << (32 - bits)) >>> 0;
  return (ipInt & mask) === (rangeInt & mask);
}

function ipAllowed(ip, list) {
  if (!ip) return false;
  for (const entry of list) {
    if (entry.includes('/')) {
      if (cidrMatchV4(ip, entry)) return true;
    } else if (entry === ip) {
      return true;
    }
  }
  return false;
}

async function hmacSha256Hex(secret, message) {
  const key = await crypto.subtle.importKey(
    'raw',
    new TextEncoder().encode(secret),
    { name: 'HMAC', hash: 'SHA-256' },
    false,
    ['sign']
  );
  const sig = await crypto.subtle.sign('HMAC', key, new TextEncoder().encode(message));
  return Array.from(new Uint8Array(sig)).map(b => b.toString(16).padStart(2, '0')).join('');
}

async function makeToken(secret) {
  const exp = Math.floor(Date.now() / 1000) + MAX_AGE_SEC;
  const sig = await hmacSha256Hex(secret, 'auth:' + exp);
  return `${exp}.${sig}`;
}

async function verifyToken(secret, token) {
  if (!token) return false;
  const dot = token.indexOf('.');
  if (dot < 0) return false;
  const expStr = token.slice(0, dot);
  const sig = token.slice(dot + 1);
  const exp = parseInt(expStr, 10);
  if (!exp || !sig) return false;
  if (Math.floor(Date.now() / 1000) > exp) return false;
  const expected = await hmacSha256Hex(secret, 'auth:' + exp);
  if (expected.length !== sig.length) return false;
  let diff = 0;
  for (let i = 0; i < expected.length; i++) {
    diff |= expected.charCodeAt(i) ^ sig.charCodeAt(i);
  }
  return diff === 0;
}

function getCookie(request, name) {
  const raw = request.headers.get('Cookie') || '';
  for (const part of raw.split(';')) {
    const trimmed = part.trim();
    const eq = trimmed.indexOf('=');
    if (eq < 0) continue;
    if (trimmed.slice(0, eq) === name) return trimmed.slice(eq + 1);
  }
  return null;
}

function loginPage(errorMsg, nextUrl) {
  const nextField = nextUrl ? `<input type="hidden" name="next" value="${escapeHtml(nextUrl)}">` : '';
  const html = `<!doctype html>
<html lang="ko"><head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>접근 인증 · PO Hub</title>
<style>
  :root { --bg:#f6f3ee; --card:#fff; --text:#3a3530; --muted:#8a8378; --line:#e4ded3; --accent:#3a3530; --accent-hover:#26221e; --err:#b04a3a; }
  * { box-sizing: border-box; }
  html, body { margin:0; padding:0; height:100%; }
  body { min-height:100vh; display:flex; align-items:center; justify-content:center;
         background: radial-gradient(circle at 20% 10%, #fbf9f5 0%, var(--bg) 60%);
         font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Noto Sans KR', sans-serif;
         color: var(--text); letter-spacing: -0.01em; }
  .card { background: var(--card); padding: 36px 32px 28px; border-radius: 16px;
          box-shadow: 0 20px 60px rgba(58, 53, 48, 0.10), 0 2px 6px rgba(58, 53, 48, 0.04);
          width: calc(100% - 32px); max-width: 360px; }
  .brand { font-size: 12px; color: var(--muted); font-weight: 600; letter-spacing: 0.08em;
           text-transform: uppercase; margin-bottom: 10px; }
  h1 { font-size: 20px; margin: 0 0 6px; font-weight: 700; letter-spacing: -0.02em; }
  p.sub { margin: 0 0 22px; color: var(--muted); font-size: 13px; line-height: 1.5; }
  label { display:block; font-size: 12px; color: var(--muted); margin-bottom: 6px; font-weight: 500; }
  input[type=password] { width: 100%; padding: 12px 14px; border: 1px solid var(--line);
         border-radius: 10px; font-size: 14px; outline: none; background: #fbf9f5;
         color: var(--text); transition: border-color .15s, background .15s; font-family: inherit; }
  input[type=password]:focus { border-color: #c9b48a; background: #fff; }
  button { margin-top: 14px; width: 100%; padding: 12px; border: 0; border-radius: 10px;
           background: var(--accent); color: #fff; font-size: 14px; font-weight: 600;
           cursor: pointer; transition: background .15s; font-family: inherit; letter-spacing: -0.01em; }
  button:hover { background: var(--accent-hover); }
  .err { color: var(--err); font-size: 12px; margin-top: 12px; min-height: 16px; text-align: center; }
</style></head>
<body>
  <form class="card" method="POST" action="/login" autocomplete="off">
    <div class="brand">PO HUB</div>
    <h1>접근 인증</h1>
    <p class="sub">허용된 네트워크가 아닙니다. 접근하려면 비밀번호를 입력하세요.</p>
    <label for="pw">비밀번호</label>
    <input id="pw" type="password" name="password" autofocus required>
    ${nextField}
    <button type="submit">확인</button>
    <div class="err">${escapeHtml(errorMsg || '')}</div>
  </form>
</body></html>`;
  return new Response(html, {
    status: errorMsg ? 401 : 200,
    headers: {
      'Content-Type': 'text/html; charset=utf-8',
      'Cache-Control': 'no-store, no-cache, must-revalidate',
      'X-Robots-Tag': 'noindex, nofollow'
    }
  });
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({
    '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'
  }[c]));
}

function safeNext(nextRaw) {
  if (!nextRaw) return '/';
  // 오픈 리다이렉트 방지: 같은 오리진의 절대경로만 허용
  if (typeof nextRaw !== 'string') return '/';
  if (!nextRaw.startsWith('/') || nextRaw.startsWith('//')) return '/';
  return nextRaw;
}

export async function onRequest(context) {
  const { request, env, next } = context;
  const url = new URL(request.url);
  const path = url.pathname;

  const secret = env.COOKIE_SECRET;
  const password = env.SITE_PASSWORD;
  if (!secret || !password) {
    return new Response(
      'Auth misconfigured: SITE_PASSWORD / COOKIE_SECRET 환경변수를 Cloudflare Pages에 설정하세요.',
      { status: 500, headers: { 'Content-Type': 'text/plain; charset=utf-8' } }
    );
  }

  // /login 엔드포인트
  if (path === '/login') {
    if (request.method === 'POST') {
      const form = await request.formData();
      const submitted = String(form.get('password') || '');
      const nextUrl = safeNext(form.get('next'));
      if (submitted && submitted === password) {
        const token = await makeToken(secret);
        return new Response(null, {
          status: 303,
          headers: {
            'Location': nextUrl,
            'Set-Cookie': `${COOKIE_NAME}=${token}; Path=/; Max-Age=${MAX_AGE_SEC}; HttpOnly; Secure; SameSite=Lax`
          }
        });
      }
      return loginPage('비밀번호가 올바르지 않습니다.', nextUrl);
    }
    // GET /login
    return loginPage(null, safeNext(url.searchParams.get('next')));
  }

  // /logout — 쿠키 만료
  if (path === '/logout') {
    return new Response(null, {
      status: 303,
      headers: {
        'Location': '/login',
        'Set-Cookie': `${COOKIE_NAME}=; Path=/; Max-Age=0; HttpOnly; Secure; SameSite=Lax`
      }
    });
  }

  // 1) IP 허용 목록 확인
  const ip = getClientIP(request);
  const allowedIPs = parseList(env.ALLOWED_IPS);
  if (ipAllowed(ip, allowedIPs)) {
    return next();
  }

  // 2) 쿠키 확인
  const token = getCookie(request, COOKIE_NAME);
  if (await verifyToken(secret, token)) {
    return next();
  }

  // 3) 미인증 — API는 401, 그 외는 로그인 화면
  if (path.startsWith('/api/')) {
    return new Response('Unauthorized', {
      status: 401,
      headers: { 'Content-Type': 'text/plain; charset=utf-8' }
    });
  }
  return loginPage(null, path + url.search);
}
