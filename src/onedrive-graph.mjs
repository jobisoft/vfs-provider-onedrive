/**
 * Microsoft Graph HTTP layer.
 *
 * - URL builders for path-addressed operations that respect a connection's
 *   root (own drive vs. a shared-with-me subtree).
 * - `graphFetch` centralizes auth injection, refresh-on-401, 429 backoff,
 *   and Graph → VFS error mapping.
 */

export const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

// ── URL builders ────────────────────────────────────────────────────────────

function _encodePath(path) {
  if (!path || path === '/') return '';
  return path
    .replace(/^\/+/, '')
    .split('/')
    .filter(Boolean)
    .map(encodeURIComponent)
    .join('/');
}

/**
 * Returns the `/drives/{driveId}/root` or `/drives/{driveId}/items/{id}` base
 * for a connection — the anchor from which all path-addressed calls resolve.
 */
export function driveRootBase(conn) {
  return conn.rootItemId
    ? `${GRAPH_BASE}/drives/${encodeURIComponent(conn.driveId)}/items/${encodeURIComponent(conn.rootItemId)}`
    : `${GRAPH_BASE}/drives/${encodeURIComponent(conn.driveId)}/root`;
}

/**
 * Builds a Graph URL for an action on a specific VFS path within the
 * connection's root.
 *
 *   pathToGraphUrl(conn, '/a/b.txt', ':/content')
 *     → .../drives/{d}/root:/a/b.txt:/content
 *
 *   pathToGraphUrl(conn, '/', '/children')
 *     → .../drives/{d}/root/children
 *
 * Each path segment is `encodeURIComponent`-ed so colons inside file names
 * are escaped (Graph uses `:` as the path delimiter).
 */
export function pathToGraphUrl(conn, path, suffix = '') {
  const base    = driveRootBase(conn);
  const encoded = _encodePath(path);

  if (!encoded) {
    if (!suffix) return base;
    return suffix.startsWith(':') ? base : base + (suffix.startsWith('/') ? suffix : '/' + suffix);
  }

  return `${base}:/${encoded}${suffix}`;
}

/** Returns the parent path of a VFS path, or '/' for root-level items. */
export function parentOf(path) {
  if (!path || path === '/') return '/';
  const clean = path.replace(/\/+$/, '');
  const idx   = clean.lastIndexOf('/');
  return idx <= 0 ? '/' : clean.slice(0, idx);
}

/** Returns the basename (final segment) of a VFS path. */
export function basenameOf(path) {
  if (!path || path === '/') return '';
  return path.replace(/\/+$/, '').split('/').pop();
}

// ── Parent-reference → VFS path ─────────────────────────────────────────────

/**
 * Translates a Graph `parentReference.path` back into a VFS path relative to
 * the connection's configured root. Graph shapes seen in the wild:
 *
 *   '/drive/root:'             → own-drive root
 *   '/drive/root:/Folder/Sub'  → own-drive subpath
 *   '/drives/{id}/root:'       → idem (business tenants)
 *   '/drives/{id}/items/{r}:/Folder' → shared-item subpath
 */
export function parentRefToPath(conn, parentReference, itemName) {
  const raw = parentReference?.path ?? '';
  const colonIdx = raw.indexOf(':');
  const after = colonIdx >= 0 ? raw.slice(colonIdx + 1) : raw;

  let segments = after
    .split('/')
    .filter(Boolean)
    .map(seg => { try { return decodeURIComponent(seg); } catch { return seg; } });

  // Shared-item connections: Graph returns paths rooted at the owner's drive
  // (e.g. '/Owner Folder/Our Share/sub'), but the VFS caller only ever sees
  // paths relative to the shared root. Strip everything up through the first
  // segment matching the shared item's name. Best-effort; if the shared name
  // collides with an earlier segment, downstream out-of-band notifications
  // may miss until the user navigates, triggering a fresh onList.
  if (conn.rootItemId && conn.driveName) {
    const idx = segments.indexOf(conn.driveName);
    if (idx >= 0) segments = segments.slice(idx + 1);
  }

  const parentPath = '/' + segments.join('/');
  const normalizedParent = parentPath === '/' ? '/' : parentPath.replace(/\/+$/, '');

  if (!itemName) return normalizedParent || '/';
  return normalizedParent === '/' ? `/${itemName}` : `${normalizedParent}/${itemName}`;
}

// ── Graph item → VFS entry ─────────────────────────────────────────────────

/**
 * Maps a Graph driveItem to the VFS `Entry` shape expected by the toolkit.
 *
 * @param {object} conn - The connection (used to anchor the parent path).
 * @param {object} item - Graph driveItem JSON.
 * @returns {object} VFS entry with additional internal fields (`id`, `eTag`).
 */
export function mapGraphItem(conn, item) {
  const name = item.name ?? '';
  const path = parentRefToPath(conn, item.parentReference, name);
  const kind = item.folder ? 'directory' : 'file';
  return {
    name,
    path,
    kind,
    size:         kind === 'file' ? (item.size ?? undefined) : undefined,
    lastModified: item.lastModifiedDateTime ? Date.parse(item.lastModifiedDateTime) : undefined,
    id:           item.id,
    eTag:         item.eTag,
  };
}

// ── HTTP wrapper ────────────────────────────────────────────────────────────

const MAX_429_RETRIES     = 3;
const MAX_RETRY_AFTER_SEC = 60;

/**
 * @param {string} method
 * @param {string} url
 * @param {object} opts
 * @param {string} opts.accountId - Used by `getToken` to resolve/refresh.
 * @param {(accountId, signal, force?) => Promise<string>} opts.getToken
 * @param {object} [opts.headers]
 * @param {BodyInit|null} [opts.body]
 * @param {AbortSignal} [opts.signal]
 * @param {boolean} [opts.raw] - When true, the raw Response is returned and
 *   no error-mapping is applied to non-2xx responses. Use for endpoints where
 *   the caller needs the full Response (e.g. `onReadFile` streaming the blob,
 *   or upload-session chunk PUTs that accept 202 with partial-content bodies).
 */
export async function graphFetch(method, url, opts) {
  const { accountId, getToken, headers = {}, body, signal, raw = false } = opts;

  const resp = await _authedFetchWithRetry(
    (tok) => _doFetch(method, url, tok, headers, body, signal),
    accountId, getToken, signal
  );

  if (raw || resp.ok) return resp;
  await throwGraphError(resp);
}

/**
 * Shared auth/retry wrapper used by both `graphFetch` and `graphBatch`.
 * Handles initial token fetch, one forced-refresh retry on 401, and
 * Retry-After-aware retries on 429. The caller supplies a function that
 * performs the actual HTTP call given a bearer token.
 */
async function _authedFetchWithRetry(doFetchWithToken, accountId, getToken, signal) {
  let token = await getToken(accountId, signal);
  let resp  = await doFetchWithToken(token);

  if (resp.status === 401) {
    token = await getToken(accountId, signal, /* force */ true);
    resp  = await doFetchWithToken(token);
    if (resp.status === 401) {
      throw Object.assign(new Error(browser.i18n.getMessage('errorAuth')), { code: 'E:AUTH' });
    }
  }

  let attempts = 0;
  while (resp.status === 429 && attempts < MAX_429_RETRIES) {
    const ra = Math.min(MAX_RETRY_AFTER_SEC, parseInt(resp.headers.get('Retry-After') ?? '1', 10) || 1);
    await sleepAbortable(ra * 1000, signal);
    resp = await doFetchWithToken(token);
    attempts++;
  }

  return resp;
}

async function _doFetch(method, url, token, headers, body, signal) {
  return fetch(url, {
    method,
    headers: { Authorization: `Bearer ${token}`, ...headers },
    body,
    signal,
  });
}

// ── Batch ───────────────────────────────────────────────────────────────────

const BATCH_URL     = `${GRAPH_BASE}/$batch`;
export const BATCH_MAX_REQUESTS = 20;

/**
 * POSTs a batch of up to `BATCH_MAX_REQUESTS` sub-requests to Graph's $batch
 * endpoint and returns per-request results.
 *
 * Each `subrequests[i]` must be `{ id, method, url, headers?, body? }` where
 * `url` is relative to `/v1.0` (Graph $batch syntax). `body`, if given, is
 * passed through verbatim to the Graph service (typically a plain object).
 *
 * Returns an array aligned by `id`: `[{ id, ok, status, headers, body }]`.
 * Auth and retry handling (401 → refresh, 429 → Retry-After) are applied to
 * the **outer** batch call. Per-sub-request failures are surfaced in the
 * returned array, not thrown — callers decide how to react.
 *
 * Note: Graph also reports Retry-After on individual sub-responses. This
 * helper does not transparently retry those; let callers pick a strategy
 * if they need per-sub-request retries.
 */
export async function graphBatch(subrequests, opts) {
  const { accountId, getToken, signal } = opts;
  if (!Array.isArray(subrequests) || subrequests.length === 0) return [];
  if (subrequests.length > BATCH_MAX_REQUESTS) {
    throw Object.assign(
      new Error(`graphBatch: too many sub-requests (${subrequests.length} > ${BATCH_MAX_REQUESTS})`),
      { code: 'E:PROVIDER' }
    );
  }

  const envelope = JSON.stringify({ requests: subrequests });

  const resp = await _authedFetchWithRetry(
    (tok) => fetch(BATCH_URL, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${tok}`,
        'Content-Type': 'application/json',
      },
      body: envelope,
      signal,
    }),
    accountId, getToken, signal
  );

  if (!resp.ok) {
    await throwGraphError(resp);
  }

  const payload = await resp.json();
  const responses = Array.isArray(payload?.responses) ? payload.responses : [];

  return responses.map(r => ({
    id:      r.id,
    ok:      r.status >= 200 && r.status < 300,
    status:  r.status,
    headers: r.headers ?? {},
    body:    r.body,
  }));
}

/**
 * Sleep that rejects with AbortError if the signal fires. Used between
 * 429 retries and copy-monitor polls.
 */
export function sleepAbortable(ms, signal) {
  return new Promise((resolve, reject) => {
    if (signal?.aborted) return reject(new DOMException('Aborted', 'AbortError'));
    const t = setTimeout(resolve, ms);
    const onAbort = () => { clearTimeout(t); reject(new DOMException('Aborted', 'AbortError')); };
    signal?.addEventListener('abort', onAbort, { once: true });
  });
}

export async function throwGraphError(resp) {
  const status = resp.status;
  const bodyText = await resp.text().catch(() => '');

  // Parse Graph's error envelope — `{ error: { code, message, innerError: { code } } }` — once up front
  // so `err.message` surfaces as a clean human-readable string and `err.graphCode`
  // carries the machine-readable lookup key.
  let graphCode = null;
  let innerCode = null;
  let graphMessage = null;
  try {
    const parsed = JSON.parse(bodyText);
    graphCode    = parsed?.error?.code                 ?? null;
    innerCode    = parsed?.error?.innerError?.code     ?? null;
    graphMessage = parsed?.error?.message              ?? null;
  } catch { /* non-JSON body is fine — we fall back to a generic message below */ }

  // Message shown to UI: prefer Graph's human message, fall back to HTTP status.
  const message = graphMessage || `HTTP ${status}`;

  const mk = (code, details) => {
    const e = new Error(message);
    e.code = code;
    if (details) e.details = details;
    e.graphCode = graphCode;
    e.innerCode = innerCode;
    e.status = status;
    return e;
  };

  if (status === 403) {
    throw mk('E:AUTH');
  }
  if (status === 409 || status === 412 || graphCode === 'nameAlreadyExists') {
    const e = mk('E:EXIST');
    e.message = browser.i18n.getMessage('errorFileExists');
    throw e;
  }
  if (status === 423) {
    throw mk('E:PROVIDER', {
      id:          'locked',
      title:       browser.i18n.getMessage('errorLockedTitle'),
      description: browser.i18n.getMessage('errorLockedDescription'),
    });
  }
  if (status === 429) {
    throw mk('E:PROVIDER', {
      id:          'rate-limited',
      title:       browser.i18n.getMessage('errorRateLimitedTitle'),
      description: browser.i18n.getMessage('errorRateLimitedDescription'),
    });
  }
  if (status === 507 || graphCode === 'quotaLimitReached') {
    throw mk('E:PROVIDER', {
      id:          'quota-exceeded',
      title:       browser.i18n.getMessage('errorQuotaExceededTitle'),
      description: browser.i18n.getMessage('errorQuotaExceededDescription'),
    });
  }

  throw mk('E:PROVIDER', {
    id:          `http-${status}`,
    title:       browser.i18n.getMessage('errorHttpTitle', [String(status)]),
    description: browser.i18n.getMessage('errorHttpDescription', [String(status)]),
  });
}

/**
 * Convenience for callers that want JSON body directly. Propagates errors
 * from `graphFetch`.
 */
export async function graphJSON(method, url, opts) {
  const resp = await graphFetch(method, url, opts);
  if (resp.status === 204) return null;
  return resp.json();
}
