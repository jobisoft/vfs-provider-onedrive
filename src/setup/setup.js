import * as vfs from '../vendor/vfs-provider.mjs';
import { localizeDocument } from '../vendor/i18n.mjs';
import { runInteractiveFlow, DEFAULT_CLIENT_ID } from '../onedrive-auth.mjs';
import { listAvailableDrives } from '../onedrive-drives.mjs';
import { accountKey, connectionKey, loadAccounts } from '../onedrive-storage.mjs';

const i18n = (key, subs) => browser.i18n.getMessage(key, subs);

localizeDocument();

const params     = new URLSearchParams(location.search);
const addonId    = params.get('addonId');
const addonName  = params.get('addonName');
const setupToken = params.get('setupToken');

const capabilities = {
  file:   { read: true, add: true, modify: true, delete: true },
  folder: { read: true, add: true, modify: true, delete: true },
};

// ── UI refs ───────────────────────────────────────────────────────────────────

const addonNameEl    = document.getElementById('addon-name');
const accountSection = document.getElementById('account-section');
const accountSelect  = document.getElementById('account-select');
const newAccountForm = document.getElementById('new-account-form');
const clientIdInput  = document.getElementById('client-id');
const signInBtn      = document.getElementById('signin-btn');
const profileBox     = document.getElementById('profile');
const profileDisplay = document.getElementById('profile-display-name');
const profileUpn     = document.getElementById('profile-upn');
const driveSection   = document.getElementById('drive-section');
const driveSelect    = document.getElementById('drive-select');
const nameInput      = document.getElementById('conn-name');
const statusEl       = document.getElementById('status');
const connectBtn     = document.getElementById('connect-btn');
const cancelBtn      = document.getElementById('cancel-btn');

/** Active sign-in abort controller; lets the cancel button stop device-code polling. */
let signInAbort = null;

cancelBtn.addEventListener('click', () => {
  signInAbort?.abort();
  window.close();
});

addonNameEl.textContent = addonName || i18n('setupSubtitleDefaultAddon');

// ── State ─────────────────────────────────────────────────────────────────────

/**
 * The account context this setup is operating on. For reuse of an existing
 * account this is populated from storage. For a new account it's populated
 * after a successful OAuth flow. `accountId` is null until we commit.
 */
let context = null;

let availableDrives = [];

// ── Existing accounts ─────────────────────────────────────────────────────────

const all = await browser.storage.local.get(null);
const existingAccounts = loadAccounts(all);

if (existingAccounts.length > 0) {
  for (const acc of existingAccounts) {
    const opt = document.createElement('option');
    opt.value = acc.accountId;
    opt.textContent = acc.name
      ? `${acc.name} — ${acc.userPrincipalName ?? ''}`
      : (acc.userPrincipalName || acc.accountId);
    accountSelect.appendChild(opt);
  }
  accountSection.hidden = false;

  accountSelect.addEventListener('change', onAccountSelectChange);
}

async function onAccountSelectChange() {
  setStatus('');
  context = null;
  availableDrives = [];
  driveSection.hidden = true;
  driveSelect.replaceChildren();
  connectBtn.disabled = true;
  profileBox.hidden = true;

  const selectedAccountId = accountSelect.value;
  if (!selectedAccountId) {
    newAccountForm.hidden = false;
    return;
  }

  newAccountForm.hidden = true;
  const acc = existingAccounts.find(a => a.accountId === selectedAccountId);
  if (!acc) return;

  context = {
    accountId: acc.accountId,
    name: acc.name,
    displayName: acc.displayName,
    userPrincipalName: acc.userPrincipalName,
    clientId: acc.clientId,
    accessToken: acc.accessToken,
    refreshToken: acc.refreshToken,
    expiresAt: acc.expiresAt,
    scope: acc.scope,
    tokenType: acc.tokenType,
    tenant: acc.tenant,
  };

  profileDisplay.textContent = acc.displayName ?? '';
  profileUpn.textContent = acc.userPrincipalName ?? '';
  profileBox.hidden = false;

  await loadDrives();
}

// ── Sign-in (new account) ─────────────────────────────────────────────────────

signInBtn.addEventListener('click', async () => {
  const clientId = clientIdInput.value.trim() || DEFAULT_CLIENT_ID;

  signInBtn.disabled = true;
  setStatus(i18n('setupStatusSigningIn'), 'info');
  signInAbort = new AbortController();

  try {
    const payload = await runInteractiveFlow(clientId, signInAbort.signal);

    context = { ...payload, accountId: null };
    profileDisplay.textContent = payload.displayName ?? '';
    profileUpn.textContent     = payload.userPrincipalName ?? '';
    profileBox.hidden = false;

    setStatus('', '');
    await loadDrives();
  } catch (e) {
    if (e.name === 'AbortError') {
      setStatus('');
    } else {
      setStatus(e.message || String(e), 'error');
    }
  } finally {
    // Always re-enable so the user can retry with a different Microsoft
    // account after a failed sign-in, a silently-failing drive load, or
    // even after a successful sign-in if they picked the wrong account.
    signInBtn.disabled = false;
    signInAbort = null;
  }
});

// ── Drive enumeration & picker ────────────────────────────────────────────────

async function loadDrives() {
  setStatus(i18n('setupStatusLoadingDrives'), 'info');
  try {
    availableDrives = await listAvailableDrives(context.accessToken);
    if (availableDrives.length === 0) {
      setStatus(i18n('setupErrorNoDrives'), 'error');
      return;
    }

    driveSelect.replaceChildren();
    for (const d of availableDrives) {
      const opt = document.createElement('option');
      opt.value = _driveKey(d);
      const kindLabel = d.kind === 'own' ? i18n('driveKindOwn') : i18n('driveKindShared');
      const owner = d.owner ? ` (${d.owner})` : '';
      opt.textContent = `${d.displayName} — ${kindLabel}${owner}`;
      driveSelect.appendChild(opt);
    }

    driveSection.hidden = false;
    driveSelect.addEventListener('change', onDriveChange, { once: false });
    onDriveChange();
    connectBtn.disabled = false;
    setStatus('');
  } catch (e) {
    setStatus(e.message || String(e), 'error');
  }
}

function onDriveChange() {
  const selected = availableDrives.find(d => _driveKey(d) === driveSelect.value);
  if (selected && !nameInput.value.trim()) {
    nameInput.value = selected.displayName;
  }
}

// ── Connect ───────────────────────────────────────────────────────────────────

connectBtn.addEventListener('click', async () => {
  const selected = availableDrives.find(d => _driveKey(d) === driveSelect.value);
  if (!selected || !context) {
    setStatus(i18n('setupErrorNotReady'), 'error');
    return;
  }

  connectBtn.disabled = true;
  setStatus(i18n('setupStatusSaving'), 'info');

  const storageId = crypto.randomUUID();
  const connName  = nameInput.value.trim() || selected.displayName;

  try {
    // Persist account first if this is a new one.
    let accountId = context.accountId;
    if (!accountId) {
      accountId = crypto.randomUUID();
      await browser.storage.local.set({
        [accountKey(accountId)]: {
          name:              context.name ?? context.displayName ?? context.userPrincipalName,
          displayName:       context.displayName,
          userPrincipalName: context.userPrincipalName,
          tenant:            context.tenant,
          clientId:          context.clientId,
          accessToken:       context.accessToken,
          refreshToken:      context.refreshToken,
          expiresAt:         context.expiresAt,
          scope:             context.scope,
          tokenType:         context.tokenType,
        },
      });
    }

    await browser.storage.local.set({
      [connectionKey(storageId)]: {
        accountId,
        driveId:      selected.driveId,
        rootItemId:   selected.rootItemId ?? null,
        driveName:    selected.displayName,
        driveType:    selected.driveType,
        deltaLink:    null,
        pollInterval: 60,
      },
    });

    await vfs.reportNewConnection(addonId, addonName, storageId, connName, capabilities, setupToken);
    window.close();
  } catch (e) {
    setStatus(e.message || String(e), 'error');
    connectBtn.disabled = false;
  }
});

// ── Helpers ───────────────────────────────────────────────────────────────────

function _driveKey(d) {
  return `${d.driveId}::${d.rootItemId ?? ''}`;
}

function setStatus(msg, type = '') {
  statusEl.textContent = msg;
  statusEl.className = type;
}
