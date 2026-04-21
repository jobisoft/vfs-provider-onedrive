import * as vfs from '../vendor/vfs-provider.mjs';
import { localizeDocument } from '../vendor/i18n.mjs';
import { accountKey, connectionKey } from '../onedrive-storage.mjs';

const i18n = (key, subs) => browser.i18n.getMessage(key, subs);
const CONNECTIONS_KEY = 'vfs-toolkit-connections';

localizeDocument();

const params = new URLSearchParams(location.search);
const storageId = params.get('storageId');

const nameInput   = document.getElementById('conn-name');
const pollInput   = document.getElementById('poll-interval');
const accNameEl   = document.getElementById('acc-name');
const accUpnEl    = document.getElementById('acc-upn');
const accDriveEl  = document.getElementById('acc-drive');
const manageBtn   = document.getElementById('manage-accounts-btn');
const saveBtn     = document.getElementById('save-btn');
const cancelBtn   = document.getElementById('cancel-btn');
const statusEl    = document.getElementById('status');

const storage = await browser.storage.local.get(null);
const connection = storage[connectionKey(storageId)] ?? {};
const account    = (connection.accountId && storage[accountKey(connection.accountId)]) || {};
const conn       = (storage[CONNECTIONS_KEY] ?? []).find(c => c.storageId === storageId) ?? {};

nameInput.value = conn.name ?? '';
pollInput.value = connection.pollInterval ?? 60;
accNameEl.textContent   = account.displayName ?? account.name ?? '—';
accUpnEl.textContent    = account.userPrincipalName ?? '—';
accDriveEl.textContent  = connection.driveName ?? '—';

cancelBtn.addEventListener('click', () => window.close());
manageBtn.addEventListener('click', () => {
  browser.runtime.openOptionsPage();
  window.close();
});

function setStatus(msg, type = '') {
  statusEl.textContent = msg;
  statusEl.className = type;
}

saveBtn.addEventListener('click', async () => {
  saveBtn.disabled = true;
  try {
    const newName = nameInput.value.trim() || connection.driveName || (account.displayName ?? '');
    const poll = Math.max(0, parseInt(pollInput.value, 10) || 0);

    const needsRename = newName !== (conn.name ?? '');
    const needsPoll   = poll !== (connection.pollInterval ?? 60);

    if (needsPoll) {
      await browser.storage.local.set({
        [connectionKey(storageId)]: { ...connection, pollInterval: poll },
      });
    }

    if (needsRename && conn.addonId) {
      await vfs.reportNewConnection(conn.addonId, conn.addonName, storageId, newName, conn.capabilities);
    }

    setStatus(i18n('configStatusSaved'), 'ok');
    setTimeout(() => window.close(), 600);
  } catch {
    setStatus(i18n('configStatusSaveFailed'), 'error');
    saveBtn.disabled = false;
  }
});
