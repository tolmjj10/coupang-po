/**
 * 공용 변경이력 헬퍼.
 * 각 state POST 엔드포인트가 저장 직전에 호출 → old/new JSON 딥 diff →
 * STATE_KV 의 'changelog' 키에 링버퍼(최대 500건)로 적재.
 *
 * 포맷: [{ ts, tab, changes:[{path, before, after, op}], truncated }, ...] (최신 먼저)
 */
const LOG_KEY = 'changelog';
const MAX_ENTRIES = 500;
const MAX_CHANGES_PER_ENTRY = 50;
const MAX_LOG_BYTES = 4 * 1024 * 1024;

export function safeParse(text, fallback) {
  try { return JSON.parse(text); } catch { return fallback; }
}

function scalar(v) {
  if (v === undefined || v === null) return null;
  if (typeof v === 'object') return Array.isArray(v) ? `[${v.length}개]` : '{…}';
  const s = String(v);
  return s.length > 120 ? s.slice(0, 120) + '…' : s;
}

const NAME_FIELDS = ['name', 'realName', 'regName', 'title', 'skuName', 'productName', 'label', 'promoName', 'filename'];

function nameOf(v) {
  if (!v || typeof v !== 'object' || Array.isArray(v)) return null;
  for (const f of NAME_FIELDS) {
    if (v[f] != null && v[f] !== '') return String(v[f]).slice(0, 60);
  }
  return null;
}

function summarize(v) {
  if (v === undefined || v === null) return null;
  if (typeof v !== 'object') return scalar(v);
  if (Array.isArray(v)) return `[${v.length}개]`;
  const n = nameOf(v);
  if (n) return n;
  return '{…}';
}

export function diffDeep(oldVal, newVal, maxEntries = MAX_CHANGES_PER_ENTRY) {
  const out = [];
  function walk(a, b, path) {
    if (out.length >= maxEntries) return;
    if (a === b) return;
    if (a === undefined && b === undefined) return;
    const aIsArr = Array.isArray(a);
    const bIsArr = Array.isArray(b);
    if (aIsArr || bIsArr) {
      const arrA = aIsArr ? a : [];
      const arrB = bIsArr ? b : [];
      const keyed = (arr) => arr.length > 0 && arr.every(x => x && typeof x === 'object' && (x.id != null || x.sku != null));
      if (keyed(arrA) || keyed(arrB)) {
        const idOf = (x) => String(x.id != null ? x.id : x.sku);
        const mapA = new Map(arrA.filter(x => x && typeof x === 'object' && (x.id != null || x.sku != null)).map(x => [idOf(x), x]));
        const mapB = new Map(arrB.filter(x => x && typeof x === 'object' && (x.id != null || x.sku != null)).map(x => [idOf(x), x]));
        const ids = new Set([...mapA.keys(), ...mapB.keys()]);
        for (const id of ids) {
          if (out.length >= maxEntries) return;
          const va = mapA.get(id), vb = mapB.get(id);
          // id 대신 상품명/프로모션명 등 사람이 읽을 수 있는 라벨이 있으면 그걸 사용
          const label = nameOf(vb) || nameOf(va) || id;
          const p = path ? `${path}[${label}]` : `[${label}]`;
          if (va === undefined) { out.push({ path: p, before: null, after: summarize(vb), op: 'add' }); continue; }
          if (vb === undefined) { out.push({ path: p, before: summarize(va), after: null, op: 'del' }); continue; }
          walk(va, vb, p);
        }
      } else {
        const len = Math.max(arrA.length, arrB.length);
        for (let i = 0; i < len; i++) {
          if (out.length >= maxEntries) return;
          walk(arrA[i], arrB[i], path ? `${path}[${i}]` : `[${i}]`);
        }
      }
      return;
    }
    const aIsObj = a && typeof a === 'object';
    const bIsObj = b && typeof b === 'object';
    if (aIsObj && bIsObj) {
      const keys = new Set([...Object.keys(a), ...Object.keys(b)]);
      for (const k of keys) {
        if (out.length >= maxEntries) return;
        walk(a[k], b[k], path ? `${path}.${k}` : k);
      }
      return;
    }
    // leaf
    if (JSON.stringify(a) !== JSON.stringify(b)) {
      out.push({
        path,
        before: scalar(a),
        after: scalar(b),
        op: a === undefined ? 'add' : (b === undefined ? 'del' : 'chg')
      });
    }
  }
  walk(oldVal, newVal, '');
  return out;
}

export async function logChange(env, tab, oldText, newText) {
  if (!env.STATE_KV) return;
  try {
    const oldVal = safeParse(oldText, null);
    const newVal = safeParse(newText, null);
    const changes = diffDeep(oldVal, newVal, MAX_CHANGES_PER_ENTRY);
    if (changes.length === 0) return;
    const entry = {
      ts: new Date().toISOString(),
      tab,
      count: changes.length,
      truncated: changes.length >= MAX_CHANGES_PER_ENTRY,
      changes
    };
    const raw = await env.STATE_KV.get(LOG_KEY);
    let log = safeParse(raw || '[]', []);
    if (!Array.isArray(log)) log = [];
    log.unshift(entry);
    if (log.length > MAX_ENTRIES) log = log.slice(0, MAX_ENTRIES);
    let text = JSON.stringify(log);
    while (text.length > MAX_LOG_BYTES && log.length > 10) {
      log = log.slice(0, Math.floor(log.length * 0.8));
      text = JSON.stringify(log);
    }
    await env.STATE_KV.put(LOG_KEY, text);
  } catch {
    // 로깅 실패는 저장을 막지 않는다
  }
}
