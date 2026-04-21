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

  let token = await getToken(accountId, signal);
  let resp  = await _doFetch(method, url, token, headers, body, signal);

  if (resp.status === 401) {
    token = await getToken(accountId, signal, /* force */ true);
    resp  = await _doFetch(method, url, token, headers, body, signal);
    if (resp.status === 401) {
      throw Object.assign(new Error(browser.i18n.getMessage('errorAuth')), { code: 'E:AUTH' });
    }
  }

  let attempts = 0;
  while (resp.status === 429 && attempts < MAX_429_RETRIES) {
    const ra = Math.min(MAX_RETRY_AFTER_SEC, parseInt(resp.headers.get('Retry-After') ?? '1', 10) || 1);
    await sleepAbortable(ra * 1000, signal);
    resp = await _doFetch(method, url, token, headers, body, signal);
    attempts++;
  }

  if (raw || resp.ok) return resp;
  await _throwGraphError(resp);
}

async function _doFetch(method, url, token, headers, body, signal) {
  return fetch(url, {
    method,
    headers: { Authorization: `Bearer ${token}`, ...headers },
    body,
    signal,
  });
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

async function _throwGraphError(resp) {
  let bodyText = '';
  let graphCode = null;
  try {
    bodyText = await resp.text();
    const parsed = JSON.parse(bodyText);
    graphCode = parsed?.error?.code ?? null;
  } catch { /* non-JSON body is fine */ }

  const status = resp.status;

  if (status === 403) {
    throw Object.assign(new Error(bodyText || browser.i18n.getMessage('errorAuth')), { code: 'E:AUTH' });
  }
  if (status === 409 || status === 412 || graphCode === 'nameAlreadyExists') {
    throw Object.assign(new Error(browser.i18n.getMessage('errorFileExists')), { code: 'E:EXIST' });
  }
  if (status === 423) {
    throw Object.assign(new Error(bodyText || `HTTP ${status}`), {
      code: 'E:PROVIDER',
      details: {
        id:          'locked',
        title:       browser.i18n.getMessage('errorLockedTitle'),
        description: browser.i18n.getMessage('errorLockedDescription'),
      },
    });
  }
  if (status === 429) {
    throw Object.assign(new Error(bodyText || `HTTP ${status}`), {
      code: 'E:PROVIDER',
      details: {
        id:          'rate-limited',
        title:       browser.i18n.getMessage('errorRateLimitedTitle'),
        description: browser.i18n.getMessage('errorRateLimitedDescription'),
      },
    });
  }
  if (status === 507 || graphCode === 'quotaLimitReached') {
    throw Object.assign(new Error(bodyText || `HTTP ${status}`), {
      code: 'E:PROVIDER',
      details: {
        id:          'quota-exceeded',
        title:       browser.i18n.getMessage('errorQuotaExceededTitle'),
        description: browser.i18n.getMessage('errorQuotaExceededDescription'),
      },
    });
  }

  throw Object.assign(new Error(bodyText || `HTTP ${status}`), {
    code: 'E:PROVIDER',
    details: {
      id:          `http-${status}`,
      title:       browser.i18n.getMessage('errorHttpTitle', [String(status)]),
      description: browser.i18n.getMessage('errorHttpDescription', [String(status)]),
    },
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
