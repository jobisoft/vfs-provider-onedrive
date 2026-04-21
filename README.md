# VFS Toolkit OneDrive Provider

A [VFS Toolkit](https://github.com/thunderbird/webext-support/tree/master/modules/vfs-toolkit) storage provider for Microsoft OneDrive. Lets consumer add-ons browse, read, and write files stored in a user's OneDrive (personal or work/school) and drives shared with them, through the VFS Toolkit picker and client API.

## Features

- OAuth 2.0 **Authorization Code + PKCE** via a popup window (nice sign-in UX, no passwords stored).
- Supports personal Microsoft accounts and work/school (AAD) accounts.
- Connects to the user's own OneDrive, or to any drive/folder that has been shared with them (`/me/drive/sharedWithMe`).
- Change detection via Microsoft Graph `/delta`; updates propagate to all connected add-ons.
- Multiple connections per Microsoft account (one sign-in, many drives).
- Full read/write: list, read, write, rename, move, copy, delete for files and folders.
- Resumable uploads for files larger than 4 MiB (Graph upload sessions).
- Deploys cleanly to many users: one Azure app registration backs unlimited installs, and the add-on ships with a bundled default so most users don't even see an Azure step.

## Quick start

1. Install the add-on in Thunderbird.
2. Open the add-on's **Options** → **Add Account** → click **Sign in**.
3. A popup window opens with Microsoft's sign-in page. Sign in with any Microsoft account (personal or work/school). Grant consent on first use.
4. Pick a drive, name the connection, connect.

No Azure setup is required for end users. The add-on ships with a bundled Azure app registration that handles authentication.

## Advanced: register your own Azure app

Only needed if your organization's policy requires each add-on to use an org-registered Azure app, or you prefer not to authenticate through the add-on's default app. Setup takes ~2 minutes on <https://portal.azure.com>.

1. Sign in at <https://portal.azure.com> → **Microsoft Entra ID** (formerly Azure AD) → **App registrations** → **New registration**.
2. **Name**: anything.
3. **Supported account types**: "Accounts in any organizational directory (Any Microsoft Entra ID tenant – Multitenant) **and personal Microsoft accounts**".
4. Leave redirect URI blank. **Register**.
5. Open **Authentication** → **Add a platform** → **Mobile and desktop applications** → tick the suggested `https://login.microsoftonline.com/common/oauth2/nativeclient` → **Configure**.
6. Still on Authentication, scroll to **Advanced settings** → set **"Allow public client flows"** to **Yes** → Save.
7. Open **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions** → tick `Files.ReadWrite.All`, `User.Read`, `offline_access` → Add.
8. **Overview** → copy the **Application (client) ID**.
9. In the add-on, **Options** → **Add Account** → expand **"Use a custom Azure app (advanced)"** → paste the client ID → **Sign in**.

## Enterprise deployment

One Azure app registration / one client ID is sufficient for any number of users. The popup + PKCE flow uses the well-known `nativeclient` redirect URI, which is fixed across all installs — no per-user or per-install registration needed.

## Storage layout

```
onedrive-account-{accountId}  →  { clientId, displayName, userPrincipalName,
                                   accessToken, refreshToken, expiresAt, ... }
onedrive-conn-{storageId}     →  { accountId, driveId, rootItemId, driveName,
                                   driveType, deltaLink, pollInterval }
```

Multiple `onedrive-conn-*` rows can point at the same `onedrive-account-*`, so users can mount several drives (own + shared folders) without re-authenticating.

## Change detection

The provider uses Microsoft Graph's `/delta` endpoint. A `browser.alarms` alarm is registered per connection; each tick fetches the stored `@odata.deltaLink`, maps the returned items to VFS `StorageChangeEntry` objects, and broadcasts them to all connected clients. The `deltaLink` is persisted in `storage.local` so the sync state survives Thunderbird restarts and MV3 event-page unloads. 410 Gone responses trigger a silent re-prime.

Alarms have a 1-minute minimum period. User-configured poll intervals below 60 s are rounded up to 60 s.

## Build

```sh
npm run build
```

Produces `dist/vfs-toolkit-onedrive-provider_<version>.xpi` and refreshes the vendored toolkit files in `src/vendor/`.

## License

MPL 2.0 — see [LICENSE](LICENSE).
