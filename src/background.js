/**
 * Microsoft OneDrive storage provider for the VFS Toolkit.
 *
 * Storage layout (see onedrive-storage.mjs):
 *   onedrive-account-{accountId}  →  { clientId, displayName, userPrincipalName,
 *                                      accessToken, refreshToken, expiresAt, ... }
 *   onedrive-conn-{storageId}     →  { accountId, driveId, rootItemId,
 *                                      driveName, driveType, deltaLink,
 *                                      pollInterval }
 *
 * Multiple connections can share one account (one OAuth sign-in, many drives).
 * Poll timers use `browser.alarms` (persistent across event-page unloads);
 * module-level state is rebuilt on each wake.
 */

import { VfsProviderImplementation } from './vendor/vfs-provider.mjs';
import {
  ACCOUNT_PREFIX, CONNECTION_PREFIX,
  accountKey, connectionKey, loadConnections,
} from './onedrive-storage.mjs';
import {
  GRAPH_BASE, pathToGraphUrl, parentOf, basenameOf,
  mapGraphItem, graphFetch, graphJSON, graphBatch, BATCH_MAX_REQUESTS,
  sleepAbortable,
} from './onedrive-graph.mjs';
import { refreshAccessToken, resolveClientId } from './onedrive-auth.mjs';
import { pollDelta, primeDelta } from './onedrive-delta.mjs';

const ALARM_PREFIX       = 'onedrive-poll-';
const POLL_MIN_SEC       = 60;                 // alarms have 1-minute minimum
const SIMPLE_UPLOAD_MAX  = 4 * 1024 * 1024;    // 4 MiB
const UPLOAD_CHUNK_SIZE  = 5 * 1024 * 1024;    // multiple of 327_680 per Graph
const COPY_POLL_INTERVAL = 1000;               // ms between copy-monitor polls

function _chunks(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

function _err(code, message, details) {
  const e = new Error(message);
  e.code = code;
  if (details) e.details = details;
  return e;
}

// ────────────────────────────────────────────────────────────────────────────

class OneDriveProvider extends VfsProviderImplementation {
  /** requestId → AbortController */
  #aborts = new Map();
  /** accountId → Promise<string> (in-flight refresh) */
  #refreshing = new Map();
  /** driveId → rootItemId — avoids repeat GETs to /drives/{d}/root for root-parent ops */
  #rootIds = new Map();

  constructor() {
    super({
      name: 'OneDrive',
      setupPath:    '/setup/setup.html',
      setupWidth:   540,
      setupHeight:  680,
      configPath:   '/config/config.html',
      configWidth:  540,
      configHeight: 520,
    });
  }

  // ── Cancellation ──────────────────────────────────────────────────────────

  onCancel(canceledRequestId) {
    this.#aborts.get(canceledRequestId)?.abort();
  }

  #signal(requestId) {
    const ac = new AbortController();
    this.#aborts.set(requestId, ac);
    return ac.signal;
  }

  #done(requestId) {
    this.#aborts.delete(requestId);
  }

  async #withRequest(requestId, fn) {
    const signal = this.#signal(requestId);
    try { return await fn(signal); }
    finally { this.#done(requestId); }
  }

  // ── Account / connection lookup ───────────────────────────────────────────

  async #accountData(accountId) {
    const key = accountKey(accountId);
    const data = (await browser.storage.local.get(key))[key];
    if (!data) throw _err('E:AUTH', browser.i18n.getMessage('errorUnknownConnection'));
    return data;
  }

  async #connection(storageId) {
    const key = connectionKey(storageId);
    const conn = (await browser.storage.local.get(key))[key];
    if (!conn?.accountId || !conn?.driveId) {
      throw _err('E:AUTH', browser.i18n.getMessage('errorUnknownConnection'));
    }
    return conn;
  }

  async #bundle(storageId) {
    const conn    = await this.#connection(storageId);
    const account = await this.#accountData(conn.accountId);
    return { conn, account };
  }

  async #persistAccount(accountId, updates) {
    const key = accountKey(accountId);
    const cur = (await browser.storage.local.get(key))[key] ?? {};
    await browser.storage.local.set({ [key]: { ...cur, ...updates } });
  }

  async #persistConnection(storageId, updates) {
    const key = connectionKey(storageId);
    const cur = (await browser.storage.local.get(key))[key] ?? {};
    await browser.storage.local.set({ [key]: { ...cur, ...updates } });
  }

  // ── Peer discovery ────────────────────────────────────────────────────────

  /**
   * Returns every storageId bound to the same (accountId, driveId, rootItemId)
   * tuple — i.e. the peers that should receive a `storageChange` broadcast
   * when any one of them performs a write.
   */
  async #peerStorageIds(accountId, driveId, rootItemId) {
    const all = await browser.storage.local.get(null);
    return loadConnections(all)
      .filter(c =>
        c.accountId  === accountId &&
        c.driveId    === driveId   &&
        (c.rootItemId ?? null) === (rootItemId ?? null)
      )
      .map(c => c.storageId);
  }

  async #allStorageIdsForAccount(accountId) {
    const all = await browser.storage.local.get(null);
    return loadConnections(all).filter(c => c.accountId === accountId).map(c => c.storageId);
  }

  async #broadcastChanges(conn, changes) {
    if (!changes?.length) return;
    const ids = await this.#peerStorageIds(conn.accountId, conn.driveId, conn.rootItemId);
    for (const sid of ids) this.reportStorageChange(sid, changes);
  }

  // ── Token management ──────────────────────────────────────────────────────

  /**
   * Returns a valid access token for the account, refreshing it when
   * expired or within the safety skew. Concurrent callers share a single
   * in-flight refresh via the `#refreshing` map.
   */
  async #getAccessToken(accountId, signal, force = false) {
    if (this.#refreshing.has(accountId)) return this.#refreshing.get(accountId);

    const account = await this.#accountData(accountId);
    if (!force && account.accessToken && account.expiresAt && account.expiresAt > Date.now()) {
      return account.accessToken;
    }
    if (!account.refreshToken) {
      throw _err('E:AUTH', browser.i18n.getMessage('errorAuth'));
    }

    const clientId = resolveClientId(account);
    const p = (async () => {
      try {
        const fresh = await refreshAccessToken(clientId, account.refreshToken, signal);
        await this.#persistAccount(accountId, fresh);
        return fresh.accessToken;
      } finally {
        this.#refreshing.delete(accountId);
      }
    })();
    this.#refreshing.set(accountId, p);
    return p;
  }

  #callOpts(accountId, signal, extras = {}) {
    return {
      accountId,
      getToken: (id, sig, force) => this.#getAccessToken(id, sig, force),
      signal,
      ...extras,
    };
  }

  // ── Read operations ───────────────────────────────────────────────────────

  async onList(requestId, storageId, path) {
    const conn = await this.#connection(storageId);
    return this.#withRequest(requestId, async (signal) => {
      const results = await this.#listChildren(conn, path, signal);
      results.sort((a, b) => {
        if (a.kind !== b.kind) return a.kind === 'directory' ? -1 : 1;
        return a.name.localeCompare(b.name);
      });
      return results.map(({ path, name, kind, size, lastModified }) => ({ path, name, kind, size, lastModified }));
    });
  }

  async #listChildren(conn, path, signal) {
    const out = [];
    let url = pathToGraphUrl(conn, path,
      path === '/' ? '/children' : ':/children'
    ) + '?$select=id,name,size,folder,file,lastModifiedDateTime,parentReference,eTag&$top=200';
    while (url) {
      const page = await graphJSON('GET', url, this.#callOpts(conn.accountId, signal));
      for (const item of page.value ?? []) out.push(mapGraphItem(conn, item));
      url = page['@odata.nextLink'] ?? null;
    }
    return out;
  }

  async onReadFile(requestId, storageId, path) {
    const { conn } = await this.#bundle(storageId);
    return this.#withRequest(requestId, async (signal) => {
      const resp = await graphFetch('GET', pathToGraphUrl(conn, path, ':/content'),
        this.#callOpts(conn.accountId, signal));
      const blob = await resp.blob();
      return new File([blob], basenameOf(path), { type: blob.type || 'application/octet-stream' });
    });
  }

  async onStorageUsage(storageId) {
    const bundle = await this.#bundle(storageId).catch(() => null);
    if (!bundle) return { usage: null, quota: null };
    const { conn } = bundle;
    try {
      const body = await graphJSON('GET', `${GRAPH_BASE}/drives/${encodeURIComponent(conn.driveId)}?$select=quota`,
        this.#callOpts(conn.accountId, null));
      const q = body?.quota ?? {};
      return {
        usage: Number.isFinite(q.used)  ? q.used  : null,
        quota: Number.isFinite(q.total) ? q.total : null,
      };
    } catch {
      return { usage: null, quota: null };
    }
  }

  // ── Write operations ──────────────────────────────────────────────────────

  async onWriteFile(requestId, storageId, path, file, overwrite) {
    const { conn } = await this.#bundle(storageId);
    return this.#withRequest(requestId, async (signal) => {
      await this.#mkdirpParent(conn, path, signal);

      const size = file.size ?? (await file.arrayBuffer()).byteLength;
      const existedBefore = await this.#exists(conn, path, signal);
      if (!overwrite && existedBefore) {
        throw _err('E:EXIST', browser.i18n.getMessage('errorFileExists'));
      }

      if (size <= SIMPLE_UPLOAD_MAX) {
        await this.#uploadSmall(conn, path, file, overwrite, signal);
      } else {
        await this.#uploadLarge(conn, path, file, overwrite, signal, requestId);
      }

      await this.#broadcastChanges(conn, [{
        kind: 'file',
        action: existedBefore ? 'modified' : 'created',
        target: { path },
      }]);
    });
  }

  async onAddFolder(requestId, storageId, path) {
    const { conn } = await this.#bundle(storageId);
    return this.#withRequest(requestId, async (signal) => {
      await this.#mkdirpParent(conn, path, signal);
      await this.#createFolder(conn, path, /* mergeIfExists */ false, signal);
      await this.#broadcastChanges(conn, [{ kind: 'directory', action: 'created', target: { path } }]);
    });
  }

  async onDeleteFile(requestId, storageId, path) {
    return this.#delete(requestId, storageId, path, 'file');
  }

  async onDeleteFolder(requestId, storageId, path) {
    return this.#delete(requestId, storageId, path, 'directory');
  }

  async #delete(requestId, storageId, path, kind) {
    const { conn } = await this.#bundle(storageId);
    return this.#withRequest(requestId, async (signal) => {
      const resp = await graphFetch('DELETE', pathToGraphUrl(conn, path, ':'),
        this.#callOpts(conn.accountId, signal, { raw: true }));
      if (resp.status !== 204 && resp.status !== 404 && !resp.ok) {
        // Non-404 errors go through the mapper for a proper throw.
        await graphFetch('DELETE', pathToGraphUrl(conn, path, ':'),
          this.#callOpts(conn.accountId, signal));
      }
      await this.#broadcastChanges(conn, [{ kind, action: 'deleted', target: { path } }]);
    });
  }

  async onMoveFile(requestId, storageId, oldPath, newPath, overwrite) {
    return this.#moveOrCopy(requestId, storageId, oldPath, newPath, overwrite, { op: 'move', kind: 'file' });
  }

  async onCopyFile(requestId, storageId, oldPath, newPath, overwrite) {
    return this.#moveOrCopy(requestId, storageId, oldPath, newPath, overwrite, { op: 'copy', kind: 'file' });
  }

  async onMoveFolder(requestId, storageId, oldPath, newPath, merge) {
    if (merge) return this.#mergeOp(requestId, storageId, oldPath, newPath, 'move');
    return this.#moveOrCopy(requestId, storageId, oldPath, newPath, false, { op: 'move', kind: 'directory' });
  }

  async onCopyFolder(requestId, storageId, oldPath, newPath, merge) {
    if (merge) return this.#mergeOp(requestId, storageId, oldPath, newPath, 'copy');
    return this.#moveOrCopy(requestId, storageId, oldPath, newPath, false, { op: 'copy', kind: 'directory' });
  }

  async #moveOrCopy(requestId, storageId, oldPath, newPath, overwrite, { op, kind }) {
    const { conn } = await this.#bundle(storageId);
    return this.#withRequest(requestId, async (signal) => {
      await this.#mkdirpParent(conn, newPath, signal);

      // Personal OneDrive's /copy silently ignores `conflictBehavior` (either
      // in body or query) — it auto-renames and reports success regardless.
      // And `replace` doesn't actually replace; the original target keeps
      // its content. So we enforce both semantics ourselves by resolving the
      // target ID up front:
      //   - overwrite=true  → pre-delete the target, then copy onto a clean slot.
      //   - overwrite=false → throw E:EXIST immediately.
      // Move uses PATCH whose default is reliably "fail on conflict", so only
      // the pre-delete side matters for move.
      if (op === 'copy' || (op === 'move' && overwrite)) {
        const existingId = await this.#resolveItemId(conn, newPath, signal).catch(() => null);
        if (existingId) {
          if (op === 'copy' && !overwrite) {
            throw _err('E:EXIST', browser.i18n.getMessage('errorFileExists'));
          }
          await graphFetch('DELETE',
            `${GRAPH_BASE}/drives/${encodeURIComponent(conn.driveId)}/items/${encodeURIComponent(existingId)}`,
            this.#callOpts(conn.accountId, signal));
        }
      }

      const srcItem = await this.#resolveItemMeta(conn, oldPath, signal);
      const destParentId = await this.#resolveParentId(conn, newPath, signal);
      const newName = basenameOf(newPath);

      if (op === 'move') {
        const body = { parentReference: { id: destParentId }, name: newName };
        await graphJSON('PATCH',
          `${GRAPH_BASE}/drives/${encodeURIComponent(conn.driveId)}/items/${encodeURIComponent(srcItem.id)}`,
          this.#callOpts(conn.accountId, signal, {
            headers: { 'Content-Type': 'application/json' },
            body:    JSON.stringify(body),
          }));
        await this.#broadcastChanges(conn, [{
          kind, action: 'moved',
          source: { path: oldPath }, target: { path: newPath },
        }]);
      } else {
        // Existence + overwrite semantics are already enforced above by the
        // pre-check + pre-delete. At this point the target slot is clean.
        const body = {
          parentReference: { driveId: conn.driveId, id: destParentId },
          name: newName,
        };
        const resp = await graphFetch('POST',
          `${GRAPH_BASE}/drives/${encodeURIComponent(conn.driveId)}/items/${encodeURIComponent(srcItem.id)}/copy`,
          this.#callOpts(conn.accountId, signal, {
            headers: { 'Content-Type': 'application/json' },
            body:    JSON.stringify(body),
            raw:     true,
          }));
        if (resp.status !== 202) {
          // Surface anything unexpected via the error mapper.
          if (!resp.ok) {
            await graphFetch('POST',
              `${GRAPH_BASE}/drives/${encodeURIComponent(conn.driveId)}/items/${encodeURIComponent(srcItem.id)}/copy`,
              this.#callOpts(conn.accountId, signal, {
                headers: { 'Content-Type': 'application/json' },
                body:    JSON.stringify(body),
              }));
          }
        }
        const monitorUrl = resp.headers.get('Location');
        if (monitorUrl) await this.#awaitCopy(monitorUrl, requestId, signal);

        // For folder copies, Graph's /copy can report `completed` before the
        // child items are visible via path-based access on personal OneDrive.
        // Wait until dest child count matches src (or a short timeout) so
        // callers reading a just-copied descendant don't hit itemNotFound.
        if (kind === 'directory') {
          await this.#awaitFolderCopyConsistency(conn, oldPath, newPath, signal);
        }

        await this.#broadcastChanges(conn, [{
          kind, action: 'copied',
          source: { path: oldPath }, target: { path: newPath },
        }]);
      }
    });
  }

  async #awaitFolderCopyConsistency(conn, srcPath, destPath, signal) {
    const expected = await this.#collectAll(conn, srcPath, signal).catch(() => []);
    if (expected.length === 0) return;

    const POLL_MS      = 200;
    const MAX_WAIT_MS  = 15_000;
    const deadline     = Date.now() + MAX_WAIT_MS;

    while (Date.now() < deadline) {
      if (signal.aborted) return;
      const actual = await this.#collectAll(conn, destPath, signal).catch(() => []);
      if (actual.length >= expected.length) return;
      await sleepAbortable(POLL_MS, signal);
    }
  }

  async #mergeOp(requestId, storageId, srcPath, destPath, op) {
    const { conn } = await this.#bundle(storageId);
    return this.#withRequest(requestId, async (signal) => {
      await this.#mkdirpParent(conn, destPath, signal);
      await this.#createFolder(conn, destPath, /* mergeIfExists */ true, signal);

      const srcNorm = srcPath.replace(/\/$/, '');
      const dstNorm = destPath.replace(/\/$/, '');
      const entries = await this.#collectAll(conn, srcPath, signal);

      const dirs  = entries.filter(e => e.kind === 'directory').sort((a, b) => a.path.localeCompare(b.path));
      const files = entries.filter(e => e.kind === 'file');
      const completed = [];
      const totalItems = dirs.length + files.length;

      // Parent-id cache reused across the whole merge so the same parent
      // path isn't resolved N times when many files share a parent.
      const parentIdCache = new Map();
      const getParentId = async (destP) => {
        const parent = parentOf(destP);
        if (parentIdCache.has(parent)) return parentIdCache.get(parent);
        const id = await this.#resolveParentId(conn, destP, signal);
        parentIdCache.set(parent, id);
        return id;
      };

      try {
        // ── Mkdir phase ─────────────────────────────────────────────
        // Serial: usually few dirs, ordering matters (parent-before-child),
        // and the #createFolder path already handles the merge-case cleanly.
        for (const d of dirs) {
          if (signal.aborted) { this.#emitPartial(conn, completed); return; }
          const dest = dstNorm + d.path.slice(srcNorm.length);
          await this.#createFolder(conn, dest, /* mergeIfExists */ true, signal);
          completed.push({ kind: 'directory', action: 'created', target: { path: dest } });
          this.reportProgress(requestId, Math.floor((completed.length / totalItems) * 100));
        }

        // ── File phase (batched) ────────────────────────────────────
        for (const chunk of _chunks(files, BATCH_MAX_REQUESTS)) {
          if (signal.aborted) { this.#emitPartial(conn, completed); return; }

          // Per-file prep in parallel: parent-id + existing-target lookup.
          // Each file still issues up to 2 Graph calls, but within a chunk
          // they run concurrently, so end-to-end we wait on one round-trip,
          // not N.
          const plan = await Promise.all(chunk.map(async (f) => {
            if (signal.aborted) throw new DOMException('Aborted', 'AbortError');
            const destP   = dstNorm + f.path.slice(srcNorm.length);
            const parentId = await getParentId(destP);
            const existingId = await this.#resolveItemId(conn, destP, signal).catch(() => null);
            return { src: f, destP, parentId, existingId };
          }));

          // Batch-delete pre-existing targets (merge semantics = overwrite
          // per file). Graph's /copy on personal OneDrive silently
          // ignores conflictBehavior, so we have to pre-clear.
          const conflicts = plan.filter(p => p.existingId);
          if (conflicts.length > 0) {
            const delReqs = conflicts.map((p, i) => ({
              id:     `d${i}`,
              method: 'DELETE',
              url:    `/drives/${encodeURIComponent(conn.driveId)}/items/${encodeURIComponent(p.existingId)}`,
            }));
            // Best-effort — 404s (item already gone) are fine; we only care
            // that the slot is empty for the follow-up copy/move.
            await graphBatch(delReqs, this.#batchOpts(conn.accountId, signal));
          }

          if (op === 'copy') {
            const copyReqs = plan.map((p, i) => ({
              id:     `c${i}`,
              method: 'POST',
              url:    `/drives/${encodeURIComponent(conn.driveId)}/items/${encodeURIComponent(p.src.id)}/copy`,
              headers: { 'Content-Type': 'application/json' },
              body: {
                parentReference: { driveId: conn.driveId, id: p.parentId },
                name: basenameOf(p.destP),
              },
            }));
            const results = await graphBatch(copyReqs, this.#batchOpts(conn.accountId, signal));

            // Parallel monitor polling. Each /copy returns 202 + a monitor
            // URL; we poll them all concurrently instead of one-at-a-time.
            await Promise.all(results.map(async (r, i) => {
              if (r.status === 202) {
                const loc = r.headers?.Location ?? r.headers?.location;
                if (loc) await this.#awaitCopy(loc, requestId, signal);
              } else if (!r.ok) {
                throw this.#batchResponseToError(r);
              }
            }));

            for (const p of plan) {
              completed.push({ kind: 'file', action: 'copied', source: { path: p.src.path }, target: { path: p.destP } });
            }
          } else {
            const moveReqs = plan.map((p, i) => ({
              id:     `m${i}`,
              method: 'PATCH',
              url:    `/drives/${encodeURIComponent(conn.driveId)}/items/${encodeURIComponent(p.src.id)}`,
              headers: { 'Content-Type': 'application/json' },
              body: {
                parentReference: { id: p.parentId },
                name: basenameOf(p.destP),
              },
            }));
            const results = await graphBatch(moveReqs, this.#batchOpts(conn.accountId, signal));
            for (const r of results) {
              if (!r.ok) throw this.#batchResponseToError(r);
            }
            for (const p of plan) {
              completed.push({ kind: 'file', action: 'moved', source: { path: p.src.path }, target: { path: p.destP } });
            }
          }

          this.reportProgress(requestId, Math.floor((completed.length / totalItems) * 100));
        }
      } catch (e) {
        this.#emitPartial(conn, completed);
        if (e.name !== 'AbortError') throw e;
        return;
      }

      if (op === 'move') {
        // Remove the now-empty source tree. A single DELETE on the root
        // path handles the entire subtree server-side.
        const delResp = await graphFetch('DELETE', pathToGraphUrl(conn, srcPath, ':'),
          this.#callOpts(conn.accountId, signal, { raw: true }));
        if (delResp.status !== 204 && delResp.status !== 404 && !delResp.ok) {
          await graphFetch('DELETE', pathToGraphUrl(conn, srcPath, ':'),
            this.#callOpts(conn.accountId, signal));
        }
      }

      await this.#broadcastChanges(conn, completed);
    });
  }

  #emitPartial(conn, completed) {
    if (!completed.length) return;
    this.#broadcastChanges(conn, completed).catch(() => { });
  }

  #batchOpts(accountId, signal) {
    return {
      accountId,
      getToken: (id, sig, force) => this.#getAccessToken(id, sig, force),
      signal,
    };
  }

  #batchResponseToError(r) {
    const errBody   = r.body?.error ?? {};
    const graphCode = errBody.code;
    if (graphCode === 'nameAlreadyExists' || r.status === 409) {
      return _err('E:EXIST', browser.i18n.getMessage('errorFileExists'));
    }
    if (r.status === 403) {
      return _err('E:AUTH', errBody.message || 'Forbidden');
    }
    const msg = errBody.message || `HTTP ${r.status}`;
    const e = _err('E:PROVIDER', msg, { id: `batch-${r.status}` });
    e.graphCode = graphCode;
    e.status    = r.status;
    return e;
  }

  // ── Upload helpers ────────────────────────────────────────────────────────

  async #uploadSmall(conn, path, file, overwrite, signal) {
    const conflict = overwrite ? 'replace' : 'fail';
    await graphFetch('PUT',
      pathToGraphUrl(conn, path, `:/content?@microsoft.graph.conflictBehavior=${conflict}`),
      this.#callOpts(conn.accountId, signal, {
        headers: { 'Content-Type': file.type || 'application/octet-stream' },
        body:    file,
      }));
  }

  async #uploadLarge(conn, path, file, overwrite, signal, requestId) {
    const conflict = overwrite ? 'replace' : 'fail';
    const session = await graphJSON('POST',
      pathToGraphUrl(conn, path, ':/createUploadSession'),
      this.#callOpts(conn.accountId, signal, {
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          item: {
            '@microsoft.graph.conflictBehavior': conflict,
            name: basenameOf(path),
          },
        }),
      }));

    const uploadUrl = session.uploadUrl;
    if (!uploadUrl) throw _err('E:PROVIDER', 'No uploadUrl returned');

    const total = file.size;
    let offset = 0;

    try {
      while (offset < total) {
        if (signal.aborted) throw new DOMException('Aborted', 'AbortError');
        const end = Math.min(offset + UPLOAD_CHUNK_SIZE, total);
        const chunk = file.slice(offset, end);
        const resp = await fetch(uploadUrl, {
          method: 'PUT',
          headers: {
            'Content-Length': String(end - offset),
            'Content-Range':  `bytes ${offset}-${end - 1}/${total}`,
          },
          body: chunk,
          signal,
        });
        if (!resp.ok) {
          if (resp.status === 409) {
            throw _err('E:EXIST', browser.i18n.getMessage('errorFileExists'));
          }
          const text = await resp.text().catch(() => '');
          throw _err('E:PROVIDER', text || `HTTP ${resp.status}`, {
            id:          `http-${resp.status}`,
            title:       browser.i18n.getMessage('errorHttpTitle', [String(resp.status)]),
            description: browser.i18n.getMessage('errorHttpDescription', [String(resp.status)]),
          });
        }
        offset = end;
        this.reportProgress(requestId, Math.floor((offset / total) * 100));
      }
    } catch (e) {
      // Best-effort cleanup of the upload session on abort/error.
      fetch(uploadUrl, { method: 'DELETE' }).catch(() => { });
      throw e;
    }
  }

  // ── Copy-monitor polling ──────────────────────────────────────────────────

  async #awaitCopy(monitorUrl, requestId, signal) {
    let lastPct = -1;
    while (true) {
      if (signal.aborted) throw new DOMException('Aborted', 'AbortError');
      const resp = await fetch(monitorUrl, { signal });
      if (!resp.ok) {
        const text = await resp.text().catch(() => '');
        throw _err('E:PROVIDER', text || `HTTP ${resp.status}`, { id: `copy-monitor-${resp.status}` });
      }
      const body = await resp.json();
      const pct = Math.floor(body.percentageComplete ?? 0);
      if (pct !== lastPct) {
        this.reportProgress(requestId, pct);
        lastPct = pct;
      }
      if (body.status === 'completed') return body.resourceId ?? null;
      if (body.status === 'failed') {
        const graphErrCode = body.error?.code;
        if (graphErrCode === 'nameAlreadyExists') {
          throw _err('E:EXIST', browser.i18n.getMessage('errorFileExists'));
        }
        throw _err('E:PROVIDER', body.error?.message ?? 'Copy failed', { id: 'copy-failed' });
      }
      await sleepAbortable(COPY_POLL_INTERVAL, signal);
    }
  }

  // ── Path resolution ───────────────────────────────────────────────────────

  async #exists(conn, path, signal) {
    try {
      await this.#resolveItemMeta(conn, path, signal);
      return true;
    } catch (e) {
      if (e.code === 'E:PROVIDER' && /http-404/.test(e.details?.id ?? '')) return false;
      if (e.code === 'E:AUTH') throw e;
      return false;
    }
  }

  async #resolveItemMeta(conn, path, signal) {
    if (!path || path === '/') {
      // The connection root — always exists.
      return { id: conn.rootItemId ?? 'root' };
    }
    const body = await graphJSON('GET',
      pathToGraphUrl(conn, path, ':?$select=id,parentReference,name,folder,file'),
      this.#callOpts(conn.accountId, signal));
    return body;
  }

  async #resolveItemId(conn, path, signal) {
    const meta = await this.#resolveItemMeta(conn, path, signal);
    return meta.id;
  }

  async #resolveParentId(conn, childPath, signal) {
    const parent = parentOf(childPath);
    if (parent === '/') {
      if (conn.rootItemId) return conn.rootItemId;
      const cached = this.#rootIds.get(conn.driveId);
      if (cached) return cached;
      const rootMeta = await graphJSON('GET',
        `${GRAPH_BASE}/drives/${encodeURIComponent(conn.driveId)}/root?$select=id`,
        this.#callOpts(conn.accountId, signal));
      if (rootMeta?.id) this.#rootIds.set(conn.driveId, rootMeta.id);
      return rootMeta.id;
    }
    return this.#resolveItemId(conn, parent, signal);
  }

  async #createFolder(conn, path, mergeIfExists, signal) {
    const parent = parentOf(path);
    const name = basenameOf(path);
    const parentSuffix = parent === '/' ? '/children' : ':/children';
    const url = pathToGraphUrl(conn, parent, parentSuffix);

    // Always `fail` on conflict. For merge-mode we interpret the resulting
    // E:EXIST as "folder is already there — that's fine" in the catch. This
    // deliberately avoids Graph's `replace` conflict-behaviour on folders,
    // which can delete existing folder contents and appears to trigger
    // Microsoft's `itemDisabledDueToUserContentMigration` state when fired
    // repeatedly against the same folder (e.g. via `#mkdirpParent` on every
    // write).
    const body = {
      name,
      folder: {},
      '@microsoft.graph.conflictBehavior': 'fail',
    };

    try {
      await graphJSON('POST', url, this.#callOpts(conn.accountId, signal, {
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify(body),
      }));
    } catch (e) {
      if (e.code === 'E:EXIST' && mergeIfExists) return;
      throw e;
    }
  }

  async #mkdirpParent(conn, path, signal) {
    const parent = parentOf(path);
    if (parent === '/') return;
    const segs = parent.split('/').filter(Boolean);
    let cur = '';
    for (const seg of segs) {
      cur += '/' + seg;
      await this.#createFolder(conn, cur, /* mergeIfExists */ true, signal);
    }
  }

  async #collectAll(conn, root, signal) {
    const out = [];
    const stack = [root];
    while (stack.length) {
      const entries = await this.#listChildren(conn, stack.pop(), signal);
      for (const entry of entries) {
        out.push(entry);
        if (entry.kind === 'directory') stack.push(entry.path);
      }
    }
    return out;
  }

  // ── Alarm-driven polling ──────────────────────────────────────────────────

  /** Creates or replaces the poll alarm for a connection. */
  async ensurePollAlarm(storageId, pollInterval) {
    const name = ALARM_PREFIX + storageId;
    await browser.alarms.clear(name);
    const sec = Math.max(POLL_MIN_SEC, pollInterval | 0);
    if (!pollInterval || pollInterval <= 0) return;
    browser.alarms.create(name, { periodInMinutes: sec / 60 });
  }

  async clearPollAlarm(storageId) {
    await browser.alarms.clear(ALARM_PREFIX + storageId);
  }

  /**
   * Resolves a fired alarm to its storageId, runs a poll tick, persists the
   * new deltaLink, and broadcasts any changes.
   */
  async handleAlarm(name) {
    if (!name.startsWith(ALARM_PREFIX)) return;
    const storageId = name.slice(ALARM_PREFIX.length);

    const { conn } = await this.#bundle(storageId).catch(() => ({ conn: null }));
    if (!conn) {
      await this.clearPollAlarm(storageId);
      return;
    }

    try {
      const { changes, newDeltaLink } = await pollDelta(conn,
        this.#callOpts(conn.accountId, null));

      if (newDeltaLink && newDeltaLink !== conn.deltaLink) {
        await this.#persistConnection(storageId, { deltaLink: newDeltaLink });
      }
      if (changes?.length) {
        for (const sid of await this.#peerStorageIds(conn.accountId, conn.driveId, conn.rootItemId)) {
          this.reportStorageChange(sid, changes);
        }
      }
    } catch {
      // Swallow per-tick failures. Next tick will retry; 410s resync inside pollDelta.
    }
  }

  /**
   * Reconciles alarms against stored connections. Called on install, startup,
   * and connection storage changes. Idempotent.
   */
  async reconcileAlarms() {
    const all = await browser.storage.local.get(null);
    const conns = loadConnections(all);
    const wanted = new Set();

    for (const c of conns) {
      if (c.pollInterval && c.pollInterval > 0) {
        wanted.add(ALARM_PREFIX + c.storageId);
        await this.ensurePollAlarm(c.storageId, c.pollInterval);
      } else {
        await this.clearPollAlarm(c.storageId);
      }

      // Prime missing deltaLink lazily so the first poll tick has a baseline.
      if (!c.deltaLink) {
        try {
          const link = await primeDelta(c, this.#callOpts(c.accountId, null));
          if (link) await this.#persistConnection(c.storageId, { deltaLink: link });
        } catch { /* will retry on next reconcile */ }
      }
    }

    // Clear any orphan alarms (alarm exists but no matching connection).
    const existing = await browser.alarms.getAll();
    for (const a of existing) {
      if (a.name.startsWith(ALARM_PREFIX) && !wanted.has(a.name)) {
        await browser.alarms.clear(a.name);
      }
    }
  }

  /** Clean up orphan account records when all their connections are gone. */
  async gcOrphanAccounts() {
    const all = await browser.storage.local.get(null);
    const referenced = new Set(loadConnections(all).map(c => c.accountId));
    for (const key of Object.keys(all)) {
      if (!key.startsWith(ACCOUNT_PREFIX)) continue;
      const accountId = key.slice(ACCOUNT_PREFIX.length);
      if (!referenced.has(accountId)) {
        await browser.storage.local.remove(key);
      }
    }
  }
}

// ────────────────────────────────────────────────────────────────────────────
// Bootstrap — top-level so listeners re-attach on every event-page wake.
// ────────────────────────────────────────────────────────────────────────────

const provider = new OneDriveProvider();
provider.init();

browser.runtime.onInstalled.addListener(() => { provider.reconcileAlarms(); });
browser.runtime.onStartup.addListener(()   => { provider.reconcileAlarms(); });

browser.alarms.onAlarm.addListener(alarm => { provider.handleAlarm(alarm.name); });

// Strip the `Origin` header on requests to the Microsoft token endpoint.
// Firefox/Thunderbird sends an Origin header from extension fetches, which
// triggers Azure's AADSTS90023 "cross-origin token redemption" check for the
// Device Code Flow grant. Removing the header lets Azure treat the request
// as a normal confidential-free public-client token exchange.
browser.webRequest.onBeforeSendHeaders.addListener(
  (details) => ({
    requestHeaders: details.requestHeaders.filter(
      h => h.name.toLowerCase() !== 'origin'
    ),
  }),
  { urls: ['https://login.microsoftonline.com/*/oauth2/v2.0/token'] },
  ['blocking', 'requestHeaders']
);

// Reconcile whenever connection rows are added/removed/changed.
browser.storage.onChanged.addListener((changes, area) => {
  if (area !== 'local') return;
  let touched = false;
  for (const key of Object.keys(changes)) {
    if (key.startsWith(CONNECTION_PREFIX)) { touched = true; break; }
  }
  if (touched) provider.reconcileAlarms();
});

// Clean up when a client revokes a connection via the picker.
browser.runtime.onMessage.addListener(msg => {
  if (msg?.type === 'vfs-toolkit-remove-connection' && msg.storageId) {
    browser.storage.local.remove(connectionKey(msg.storageId))
      .then(() => provider.clearPollAlarm(msg.storageId))
      .then(() => provider.gcOrphanAccounts())
      .catch(() => { });
  }
});

// Kick off a reconcile on cold script load too — event pages wake via events,
// but it's harmless to run once on initial module evaluation.
provider.reconcileAlarms();
