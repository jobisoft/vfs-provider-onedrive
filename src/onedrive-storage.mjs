export const ACCOUNT_PREFIX    = 'onedrive-account-';
export const CONNECTION_PREFIX = 'onedrive-conn-';
export const accountKey    = id => ACCOUNT_PREFIX + id;
export const connectionKey = id => CONNECTION_PREFIX + id;

/**
 * Returns all stored accounts from a full storage snapshot.
 * Each entry is the stored object extended with an `accountId` field.
 */
export function loadAccounts(storage) {
  return Object.entries(storage)
    .filter(([k]) => k.startsWith(ACCOUNT_PREFIX))
    .map(([k, v]) => ({ accountId: k.slice(ACCOUNT_PREFIX.length), ...v }));
}

/**
 * Returns all stored connections from a full storage snapshot.
 * Each entry is the stored object extended with a `storageId` field.
 */
export function loadConnections(storage) {
  return Object.entries(storage)
    .filter(([k]) => k.startsWith(CONNECTION_PREFIX))
    .map(([k, v]) => ({ storageId: k.slice(CONNECTION_PREFIX.length), ...v }));
}
