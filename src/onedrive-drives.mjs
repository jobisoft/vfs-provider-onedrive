/**
 * Drive enumeration for the setup-page drive picker.
 *
 * Returns a deduplicated list of drive choices the user can mount: their own
 * OneDrive plus each distinct folder/item that has been shared with them
 * (`/me/drive/sharedWithMe`). Shared items are represented as pseudo-drives
 * anchored at the shared item's id so users can only navigate into the
 * subtree they actually have permission for.
 */

import { GRAPH_BASE } from './onedrive-graph.mjs';

/**
 * @param {string} accessToken
 * @param {AbortSignal} [signal]
 * @returns {Promise<Array<{
 *   kind: 'own'|'shared',
 *   driveId: string,
 *   rootItemId: string|null,
 *   driveType: string,
 *   displayName: string,
 *   owner: string|null,
 * }>>}
 */
export async function listAvailableDrives(accessToken, signal) {
  const drives = [];
  let ownErr = null;
  let sharedErr = null;

  // Own drive.
  try {
    const own = await _getJson(`${GRAPH_BASE}/me/drive?$select=id,driveType,owner`, accessToken, signal);
    if (own?.id) {
      drives.push({
        kind:        'own',
        driveId:     own.id,
        rootItemId:  null,
        driveType:   own.driveType ?? 'personal',
        displayName: browser.i18n.getMessage('driveOwnLabel'),
        owner:       own.owner?.user?.displayName ?? null,
      });
    }
  } catch (e) {
    ownErr = e;
  }

  // Shared-with-me top-level items. Each result item has `remoteItem.*`
  // pointing at the source drive/item; treat each unique (driveId, itemId)
  // pair as a separate mount choice.
  try {
    const seen = new Set();
    let url = `${GRAPH_BASE}/me/drive/sharedWithMe?$select=name,remoteItem`;
    while (url) {
      const page = await _getJson(url, accessToken, signal);
      for (const entry of page.value ?? []) {
        const r = entry.remoteItem;
        if (!r?.parentReference?.driveId || !r.id) continue;
        const key = `${r.parentReference.driveId}::${r.id}`;
        if (seen.has(key)) continue;
        seen.add(key);
        drives.push({
          kind:        'shared',
          driveId:     r.parentReference.driveId,
          rootItemId:  r.id,
          driveType:   r.parentReference.driveType ?? 'business',
          displayName: r.name ?? entry.name ?? browser.i18n.getMessage('driveSharedFallback'),
          owner:       r.shared?.owner?.user?.displayName
                    ?? r.createdBy?.user?.displayName
                    ?? null,
        });
      }
      url = page['@odata.nextLink'] ?? null;
    }
  } catch (e) {
    sharedErr = e;
  }

  // If nothing was found, surface the first error we saw so the UI can
  // show a real diagnostic instead of a generic "no drives" message.
  if (drives.length === 0 && (ownErr || sharedErr)) {
    throw ownErr || sharedErr;
  }

  return drives;
}

async function _getJson(url, accessToken, signal) {
  const resp = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
    signal,
  });
  if (!resp.ok) {
    const text = await resp.text().catch(() => '');
    throw new Error(text || `HTTP ${resp.status}`);
  }
  return resp.json();
}
