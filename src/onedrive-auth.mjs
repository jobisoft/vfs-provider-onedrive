/**
 * Microsoft Graph OAuth 2.0 Authorization Code + PKCE via a Thunderbird popup.
 *
 * Flow:
 *   1. Generate PKCE verifier + SHA-256 challenge + state.
 *   2. Open a popup window (browser.windows.create) pointing at Microsoft's
 *      authorize endpoint, with redirect_uri set to the well-known
 *      `oauth2/nativeclient` URL.
 *   3. Watch the popup's navigation via webRequest.onBeforeRequest. When
 *      Microsoft redirects to nativeclient?code=...&state=..., extract the
 *      code, cancel the navigation, and close the popup.
 *   4. Exchange the code + verifier at the token endpoint → tokens.
 *   5. GET /me → profile → account payload.
 *
 * Why the `nativeclient` URI: Firefox/Thunderbird's identity API assigns
 * each install a random UUID redirect URI, which would require per-install
 * Azure config. `nativeclient` is Microsoft's well-known fixed URL for
 * native/desktop clients — one Azure registration works across unlimited
 * installs and profiles.
 */

// ── Endpoints ───────────────────────────────────────────────────────────────

export const AUTH_URL     = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize';
export const TOKEN_URL    = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';
export const GRAPH_ME_URL = 'https://graph.microsoft.com/v1.0/me';
export const REDIRECT_URI = 'https://login.microsoftonline.com/common/oauth2/nativeclient';

/**
 * Bundled default — the add-on's own Azure app registration. Users who don't
 * supply a custom clientId authenticate through this app. Keep in sync with
 * the app registration under the maintainer's Azure tenant.
 *
 * Deliberately NOT exported — the UI must never see this value. Accounts
 * that use the default are stored without any `clientId` field, so if this
 * constant is ever rotated, every default-using account automatically picks
 * up the new value via `resolveClientId`.
 */
const DEFAULT_CLIENT_ID = 'c6b54396-5255-4c78-86d6-a478451f0b13';

// Files.ReadWrite.All lets us traverse shared-with-me drives.
// offline_access tells Azure AD to issue a refresh_token.
// User.Read lets /me return displayName / userPrincipalName.
export const SCOPES = ['Files.ReadWrite.All', 'offline_access', 'User.Read'];

const EXPIRY_SKEW_MS = 60_000;

// Popup dimensions (matches the size OAuth2.sys.mjs uses).
const POPUP_WIDTH  = 500;
const POPUP_HEIGHT = 750;

// ── PKCE helpers ────────────────────────────────────────────────────────────

function _base64UrlEncode(bytes) {
  let s = '';
  for (const b of bytes) s += String.fromCharCode(b);
  return btoa(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

function _randomVerifier(len = 64) {
  const bytes = new Uint8Array(len);
  crypto.getRandomValues(bytes);
  return _base64UrlEncode(bytes);
}

async function _codeChallenge(verifier) {
  const data = new TextEncoder().encode(verifier);
  const hash = await crypto.subtle.digest('SHA-256', data);
  return _base64UrlEncode(new Uint8Array(hash));
}

// ── Interactive flow ────────────────────────────────────────────────────────

/**
 * Runs the full popup-based auth-code + PKCE flow and returns a full account
 * payload ready to be persisted. Throws `E:AUTH` on failure / cancel.
 *
 * @param {string} customClientId  Possibly-empty user override. When empty,
 *                                 the bundled default is used for the OAuth
 *                                 request but not stored on the returned
 *                                 payload — so `resolveClientId` can later
 *                                 resolve absent `clientId` to whatever the
 *                                 current default is.
 * @param {AbortSignal} [signal]
 */
export async function runInteractiveFlow(customClientId, signal) {
  const trimmedCustom = customClientId?.trim?.() ?? '';
  const effectiveClientId = trimmedCustom || DEFAULT_CLIENT_ID;

  const verifier  = _randomVerifier();
  const challenge = await _codeChallenge(verifier);
  const state     = crypto.randomUUID();

  const authUrl = new URL(AUTH_URL);
  authUrl.searchParams.set('client_id',             effectiveClientId);
  authUrl.searchParams.set('response_type',         'code');
  authUrl.searchParams.set('redirect_uri',          REDIRECT_URI);
  authUrl.searchParams.set('response_mode',         'query');
  authUrl.searchParams.set('scope',                 SCOPES.join(' '));
  authUrl.searchParams.set('code_challenge',        challenge);
  authUrl.searchParams.set('code_challenge_method', 'S256');
  authUrl.searchParams.set('state',                 state);
  authUrl.searchParams.set('prompt',                'select_account');

  const win = await browser.windows.create({
    url:    authUrl.toString(),
    type:   'popup',
    width:  POPUP_WIDTH,
    height: POPUP_HEIGHT,
  });

  try {
    const code    = await _waitForRedirect(win.id, state, signal);
    const tokens  = await _exchangeCodeForToken(effectiveClientId, code, verifier);
    const profile = await _fetchProfile(tokens.accessToken, signal);
    return _buildAccountPayload(trimmedCustom, tokens, profile);
  } finally {
    browser.windows.remove(win.id).catch(() => { /* already gone */ });
  }
}

/**
 * Watches the popup for a navigation to `REDIRECT_URI?code=...&state=...`,
 * returns the `code` on success, rejects on state mismatch / user-close /
 * signal abort.
 */
function _waitForRedirect(windowId, expectedState, signal) {
  return new Promise((resolve, reject) => {
    let settled = false;
    const settle = (fn, val) => {
      if (settled) return;
      settled = true;
      cleanup();
      fn(val);
    };

    const onBeforeRequest = (details) => {
      try {
        const url           = new URL(details.url);
        const code          = url.searchParams.get('code');
        const returnedState = url.searchParams.get('state');
        const error         = url.searchParams.get('error');
        const errorDesc     = url.searchParams.get('error_description');

        if (error) {
          settle(reject, Object.assign(new Error(errorDesc || error), { code: 'E:AUTH' }));
        } else if (returnedState !== expectedState) {
          settle(reject, Object.assign(
            new Error(browser.i18n.getMessage('setupErrorStateMismatch')),
            { code: 'E:AUTH' }
          ));
        } else if (!code) {
          settle(reject, Object.assign(
            new Error(browser.i18n.getMessage('setupErrorOauthNoCode')),
            { code: 'E:AUTH' }
          ));
        } else {
          settle(resolve, code);
        }
      } catch (e) {
        settle(reject, e);
      }
      // Short-circuit the navigation so users don't see Microsoft's blank
      // nativeclient page flash before we close the popup.
      return { cancel: true };
    };

    const onRemoved = (closedId) => {
      if (closedId !== windowId) return;
      // Use a plain Error (not DOMException) — DOMException's `code` is a
      // read-only legacy getter, so Object.assign({ code: ... }) throws
      // TypeError synchronously, the promise never settles, and the caller's
      // finally block never re-enables the UI.
      const err = new Error(browser.i18n.getMessage('setupErrorPopupClosed'));
      err.name = 'AbortError';
      err.code = 'E:AUTH';
      settle(reject, err);
    };

    const onAbort = () => {
      settle(reject, new DOMException('Aborted', 'AbortError'));
    };

    const cleanup = () => {
      browser.webRequest.onBeforeRequest.removeListener(onBeforeRequest);
      browser.windows.onRemoved.removeListener(onRemoved);
      signal?.removeEventListener('abort', onAbort);
    };

    browser.webRequest.onBeforeRequest.addListener(
      onBeforeRequest,
      { urls: [`${REDIRECT_URI}*`] },
      ['blocking']
    );
    browser.windows.onRemoved.addListener(onRemoved);

    if (signal) {
      if (signal.aborted) { onAbort(); return; }
      signal.addEventListener('abort', onAbort, { once: true });
    }
  });
}

async function _exchangeCodeForToken(clientId, code, verifier) {
  const resp = await _postForm(TOKEN_URL, {
    client_id:     clientId,
    grant_type:    'authorization_code',
    code,
    redirect_uri:  REDIRECT_URI,
    code_verifier: verifier,
    scope:         SCOPES.join(' '),
  });

  return {
    accessToken:  resp.access_token,
    refreshToken: resp.refresh_token,
    expiresAt:    Date.now() + Math.max(0, (resp.expires_in ?? 0) * 1000 - EXPIRY_SKEW_MS),
    scope:        resp.scope ?? SCOPES.join(' '),
    tokenType:    resp.token_type ?? 'Bearer',
  };
}

// ── Refresh ─────────────────────────────────────────────────────────────────

/**
 * Exchanges a refresh token for a fresh access token. Azure AD may rotate
 * the refresh token; callers must always re-persist both fields from the
 * returned bundle.
 */
export async function refreshAccessToken(clientId, refreshToken, signal) {
  const tokenResp = await _postForm(TOKEN_URL, {
    client_id:     clientId,
    grant_type:    'refresh_token',
    refresh_token: refreshToken,
    scope:         SCOPES.join(' '),
  }, signal);

  return {
    accessToken:  tokenResp.access_token,
    refreshToken: tokenResp.refresh_token || refreshToken,
    expiresAt:    Date.now() + Math.max(0, (tokenResp.expires_in ?? 0) * 1000 - EXPIRY_SKEW_MS),
    scope:        tokenResp.scope ?? SCOPES.join(' '),
    tokenType:    tokenResp.token_type ?? 'Bearer',
  };
}

/**
 * Returns the effective client ID for an account — the account's own stored
 * value, or the bundled default if the account didn't specify one.
 */
export function resolveClientId(account) {
  return account?.clientId?.trim() || DEFAULT_CLIENT_ID;
}

// ── Internal ────────────────────────────────────────────────────────────────

async function _postForm(url, params, signal) {
  const body = new URLSearchParams(params).toString();
  const resp = await fetch(url, {
    method:  'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body,
    signal,
  });

  let payload;
  try { payload = await resp.json(); } catch { payload = null; }

  if (!resp.ok) {
    const oerr = payload?.error;
    const desc = payload?.error_description ?? payload?.error ?? `HTTP ${resp.status}`;
    const code = (oerr === 'invalid_grant' || resp.status === 400 || resp.status === 401) ? 'E:AUTH' : 'E:PROVIDER';
    throw Object.assign(new Error(desc), { code });
  }

  return payload ?? {};
}

async function _fetchProfile(accessToken, signal) {
  const resp = await fetch(GRAPH_ME_URL, {
    headers: { Authorization: `Bearer ${accessToken}` },
    signal,
  });
  if (!resp.ok) {
    throw Object.assign(new Error(`HTTP ${resp.status} fetching /me`), { code: 'E:PROVIDER' });
  }
  return resp.json();
}

function _buildAccountPayload(customClientId, tokenResp, profile) {
  const payload = {
    displayName:       profile.displayName       ?? profile.userPrincipalName ?? '',
    userPrincipalName: profile.userPrincipalName ?? profile.mail              ?? '',
    tenant:            'common',
    accessToken:       tokenResp.accessToken,
    refreshToken:      tokenResp.refreshToken,
    expiresAt:         tokenResp.expiresAt,
    scope:             tokenResp.scope,
    tokenType:         tokenResp.tokenType,
  };
  const trimmed = customClientId?.trim?.() ?? '';
  if (trimmed) payload.clientId = trimmed;
  return payload;
}
