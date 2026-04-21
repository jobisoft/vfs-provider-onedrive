/**
 * `/delta` driver for change detection.
 *
 * Each connection stores its own `deltaLink`. `pollDelta` walks pages starting
 * from that link; on HTTP 410 it silently re-primes from `token=latest` so a
 * stale link never surfaces as a user-visible error.
 */

import { driveRootBase, graphJSON, parentRefToPath } from './onedrive-graph.mjs';

/**
 * Builds the initial `/delta` URL for a connection. Uses `token=latest` so
 * Graph returns a baseline `@odata.deltaLink` without enumerating the whole
 * tree (empty `value` array on the terminating page).
 */
function _primeUrl(conn) {
  return `${driveRootBase(conn)}/delta?token=latest`;
}

/**
 * Walks delta pages starting from `startUrl` until an `@odata.deltaLink` is
 * returned. Returns `{ changes, newDeltaLink, resynced }`.
 *
 * `resynced` is true when we had to restart from `token=latest` due to 410.
 */
async function _walkDelta(conn, startUrl, opts) {
  let url    = startUrl;
  const out  = [];
  let resynced = false;

  while (url) {
    let body;
    try {
      body = await graphJSON('GET', url, opts);
    } catch (e) {
      if (e.code === 'E:PROVIDER' && /http-410/.test(e.details?.id ?? '')) {
        // Stale deltaLink — restart from latest. Don't surface to caller.
        url = _primeUrl(conn);
        resynced = true;
        continue;
      }
      throw e;
    }

    for (const item of body.value ?? []) {
      const change = _mapDeltaItem(conn, item);
      if (change) out.push(change);
    }

    if (body['@odata.deltaLink']) {
      return { changes: out, newDeltaLink: body['@odata.deltaLink'], resynced };
    }
    url = body['@odata.nextLink'] ?? null;
  }

  return { changes: out, newDeltaLink: null, resynced };
}

/**
 * Primes (or re-primes) the delta state for a connection. Returns the
 * baseline `@odata.deltaLink`. Items returned during priming are discarded
 * because callers enumerate the tree lazily via `onList`.
 */
export async function primeDelta(conn, opts) {
  const { newDeltaLink } = await _walkDelta(conn, _primeUrl(conn), opts);
  return newDeltaLink;
}

/**
 * Runs a poll tick from the connection's stored `deltaLink` (or primes if
 * not yet set). Returns `{ changes, newDeltaLink, resynced }`.
 */
export async function pollDelta(conn, opts) {
  const start = conn.deltaLink || _primeUrl(conn);
  return _walkDelta(conn, start, opts);
}

// ── Internals ───────────────────────────────────────────────────────────────

function _mapDeltaItem(conn, item) {
  // Skip the drive/item root entry that Graph emits as part of every delta
  // response — it represents the walk anchor, not a change.
  if (item.root) return null;

  const kind = item.folder ? 'directory' : (item.file ? 'file' : 'file');

  if (item.deleted) {
    const path = parentRefToPath(conn, item.parentReference, item.name);
    if (!path) return null;
    return { kind, action: 'deleted', target: { path } };
  }

  const path = parentRefToPath(conn, item.parentReference, item.name);
  if (!path) return null;
  // Graph does not distinguish 'created' vs 'modified' in delta records.
  // 'modified' is safe: the toolkit client treats modified-on-unknown-path
  // identically to created for cache invalidation purposes.
  return { kind, action: 'modified', target: { path } };
}
