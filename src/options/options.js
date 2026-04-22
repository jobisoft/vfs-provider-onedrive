import { localizeDocument } from '../vendor/i18n.mjs';
import { runInteractiveFlow } from '../onedrive-auth.mjs';
import {
  ACCOUNT_PREFIX, CONNECTION_PREFIX,
  accountKey, connectionKey, loadAccounts, loadConnections,
} from '../onedrive-storage.mjs';

const i18n = (key, subs) => browser.i18n.getMessage(key, subs);
const CONNECTIONS_KEY = 'vfs-toolkit-connections';

localizeDocument();

// ── Render ────────────────────────────────────────────────────────────────────

async function render() {
  const storage = await browser.storage.local.get(null);
  const connections = (storage[CONNECTIONS_KEY] ?? []).filter(
    c => storage[CONNECTION_PREFIX + c.storageId] != null
  );
  renderAccounts(loadAccounts(storage), connections, storage);
  renderConnections(connections, storage);
}

function renderAccounts(accounts, connections, storage) {
  const tbody = document.getElementById('accounts-body');
  const empty = document.getElementById('accounts-empty-state');

  tbody.replaceChildren();

  if (accounts.length === 0) {
    empty.style.display = '';
    return;
  }
  empty.style.display = 'none';

  for (const account of accounts) {
    const accountConns = connections.filter(
      c => storage[CONNECTION_PREFIX + c.storageId]?.accountId === account.accountId
    );

    const tdAccount = document.createElement('td');
    tdAccount.textContent = account.displayName ?? account.name ?? '—';
    const detail = document.createElement('div');
    detail.className = 'server-url';
    detail.textContent = account.userPrincipalName ?? '';
    if (account.userPrincipalName) tdAccount.appendChild(detail);

    const editBtn = document.createElement('button');
    editBtn.className = 'edit-btn';
    editBtn.textContent = i18n('optionsBtnEdit');
    editBtn.addEventListener('click', () => openEditPopover(account));

    const btn = document.createElement('button');
    btn.className = 'revoke-btn';
    btn.textContent = i18n('btnDelete');
    btn.addEventListener('click', () => deleteAccount(account.accountId, accountConns));

    const tdAction = document.createElement('td');
    tdAction.append(editBtn, btn);

    const tr = document.createElement('tr');
    tr.append(tdAccount, tdAction);
    tbody.appendChild(tr);
  }
}

function renderConnections(connections, storage) {
  const tbody = document.getElementById('connections-body');
  const empty = document.getElementById('empty-state');

  tbody.replaceChildren();

  if (!connections.length) {
    empty.style.display = '';
    return;
  }
  empty.style.display = 'none';

  for (const conn of connections) {
    const connRow = storage[CONNECTION_PREFIX + conn.storageId];
    const accountId = connRow?.accountId;
    const accountName = accountId
      ? (storage[ACCOUNT_PREFIX + accountId]?.displayName ?? storage[ACCOUNT_PREFIX + accountId]?.name ?? '—')
      : '—';
    const driveName = connRow?.driveName ?? '—';

    const btn = document.createElement('button');
    btn.className = 'revoke-btn';
    btn.textContent = i18n('btnRevoke');
    btn.addEventListener('click', () => revokeAccess(conn.addonId, conn.storageId));

    const tdConn = document.createElement('td');
    tdConn.textContent = conn.name ?? '—';
    const driveLine = document.createElement('div');
    driveLine.className = 'addon-name';
    driveLine.textContent = `${i18n('optionsConnDrive')} ${driveName}`;
    const accountLine = document.createElement('div');
    accountLine.className = 'addon-name';
    accountLine.textContent = `${i18n('optionsConnAccount')} ${accountName}`;
    const addonLine = document.createElement('div');
    addonLine.className = 'addon-name';
    addonLine.textContent = `${i18n('optionsConnExtension')} ${conn.addonName ?? conn.addonId ?? '—'}`;
    tdConn.append(driveLine, accountLine, addonLine);

    const tdAction = document.createElement('td');
    tdAction.appendChild(btn);

    const tr = document.createElement('tr');
    tr.append(tdConn, tdAction);
    tbody.appendChild(tr);
  }
}

async function deleteAccount(accountId, connections) {
  // Remove toolkit-side connection rows for every connection using this account.
  const rv = await browser.storage.local.get({ [CONNECTIONS_KEY]: [] });
  const storageIds = connections.map(c => c.storageId);
  const updated = rv[CONNECTIONS_KEY].filter(c => !storageIds.includes(c.storageId));
  await browser.storage.local.set({ [CONNECTIONS_KEY]: updated });

  await browser.storage.local.remove([
    ...storageIds.map(id => connectionKey(id)),
    accountKey(accountId),
  ]);

  for (const conn of connections) {
    browser.runtime.sendMessage(conn.addonId, {
      type: 'vfs-toolkit-remove-connection',
      storageId: conn.storageId,
    }).catch(() => { });
  }

  render();
}

async function revokeAccess(addonId, storageId) {
  const rv = await browser.storage.local.get({ [CONNECTIONS_KEY]: [] });
  const updated = rv[CONNECTIONS_KEY].filter(
    c => !(c.addonId === addonId && c.storageId === storageId)
  );
  await browser.storage.local.set({ [CONNECTIONS_KEY]: updated });

  await browser.storage.local.remove(connectionKey(storageId));

  browser.runtime.sendMessage(addonId, {
    type: 'vfs-toolkit-remove-connection',
    storageId,
  }).catch(() => { });

  render();
}

render();

browser.storage.onChanged.addListener((changes, area) => {
  if (area !== 'local') return;
  const relevant = Object.keys(changes).some(k =>
    k === CONNECTIONS_KEY || k.startsWith(ACCOUNT_PREFIX) || k.startsWith(CONNECTION_PREFIX)
  );
  if (relevant) render();
});

// ── Add-account popover ───────────────────────────────────────────────────────

const popover          = document.getElementById('add-account-popover');
const popoverTitle     = document.getElementById('aa-title');
const nameInput        = document.getElementById('aa-name');
const clientIdInput    = document.getElementById('aa-client-id');
const advancedDetails  = clientIdInput.closest('details');
const signInBtn        = document.getElementById('aa-signin-btn');
const statusEl         = document.getElementById('aa-status');
const cancelBtn        = document.getElementById('aa-cancel-btn');

let editingAccountId = null;
/** Active sign-in abort controller; cancel closes the popover + stops sign-in. */
let signInAbort = null;

function setStatus(msg, type = '') {
  statusEl.textContent = msg;
  statusEl.className = type;
}

function resetPopover() {
  nameInput.value = '';
  clientIdInput.value = '';
  advancedDetails.open = false;
  setStatus('');
  signInBtn.disabled = false;
  editingAccountId = null;
  popoverTitle.textContent = i18n('optionsAddAccountTitle');
  signInBtn.textContent = i18n('optionsBtnSignInAndAdd');
}

function openEditPopover(account) {
  resetPopover();
  editingAccountId = account.accountId;
  nameInput.value = account.name ?? account.displayName ?? '';
  const storedCustom = account.clientId ?? '';
  clientIdInput.value = storedCustom;
  // Expand advanced section only when this account uses a custom client ID.
  // Accounts on the bundled default have no `clientId` field (or an empty
  // string) and should show the advanced section collapsed.
  advancedDetails.open = storedCustom.length > 0;
  popoverTitle.textContent = i18n('optionsEditAccountTitle');
  signInBtn.textContent = i18n('optionsBtnSaveAccount');
  popover.showPopover();
}

document.getElementById('add-account-btn').addEventListener('click', () => {
  resetPopover();
  popover.showPopover();
});

const hidePopover = () => {
  signInAbort?.abort();
  signInAbort = null;
  popover.hidePopover();
};
cancelBtn.addEventListener('click', hidePopover);

popover.addEventListener('toggle', e => {
  if (e.newState === 'open') document.addEventListener('keydown', onEsc);
  else                       document.removeEventListener('keydown', onEsc);
});
function onEsc(e) { if (e.key === 'Escape') hidePopover(); }

signInBtn.addEventListener('click', async () => {
  const customClientId = clientIdInput.value.trim();  // '' = use bundled default

  signInBtn.disabled = true;

  // Edit mode: if only the name is changing and the client ID is unchanged,
  // save without running OAuth. A changed client ID requires a fresh
  // sign-in (refresh tokens don't transfer across apps).
  if (editingAccountId) {
    const storage = await browser.storage.local.get(accountKey(editingAccountId));
    const current = storage[accountKey(editingAccountId)] ?? {};
    const currentCustom = current.clientId ?? '';
    const clientIdChanged = currentCustom !== customClientId;

    if (!clientIdChanged) {
      await browser.storage.local.set({
        [accountKey(editingAccountId)]: {
          ...current,
          name: nameInput.value.trim() || current.name,
        },
      });
      hidePopover();
      return;
    }
    // Fall through to sign-in with new client ID.
  }

  setStatus(i18n('setupStatusSigningIn'), 'info');
  signInAbort = new AbortController();

  try {
    const payload = await runInteractiveFlow(customClientId, signInAbort.signal);

    // Dedup by userPrincipalName unless editing.
    const storage = await browser.storage.local.get(null);
    const existing = loadAccounts(storage).find(a =>
      a.userPrincipalName === payload.userPrincipalName &&
      a.accountId !== editingAccountId
    );
    if (existing) {
      setStatus(i18n('optionsErrorAccountExists'), 'error');
      return;
    }

    const accountId = editingAccountId ?? crypto.randomUUID();
    const name = nameInput.value.trim() || payload.displayName || payload.userPrincipalName;

    await browser.storage.local.set({
      [accountKey(accountId)]: { ...payload, name },
    });

    hidePopover();
  } catch (e) {
    if (e.name === 'AbortError') {
      setStatus('');
    } else {
      setStatus(e.message || String(e), 'error');
    }
  } finally {
    // Always re-enable so the user can retry with a different Microsoft
    // account after any failure (sign-in error, dup-account rejection, etc.).
    signInBtn.disabled = false;
    signInAbort = null;
  }
});
