const GRAPH = 'https://graph.microsoft.com/v1.0';
const PREFIX = 'household-finances-connection-test-';
const TEST_KIND = 'household-finances-connection-test';
const MESSAGE = 'Household Finances connection test';
const MAX_BYTES = 8192;

export class ConnectionError extends Error {
  constructor(message, status = 0) { super(message); this.name = 'ConnectionError'; this.status = status; }
}

export function makeSample(now = new Date(), id = crypto.randomUUID()) {
  return { schemaVersion: 1, kind: TEST_KIND, message: MESSAGE, code: id.replaceAll('-', '').slice(0, 12).toUpperCase(), savedAt: now.toISOString() };
}

export function validateSample(data) {
  if (!data || data.schemaVersion !== 1 || data.kind !== TEST_KIND || data.message !== MESSAGE ||
      typeof data.code !== 'string' || !/^[A-F0-9]{12}$/.test(data.code) ||
      typeof data.savedAt !== 'string' || !Number.isFinite(Date.parse(data.savedAt))) {
    throw new ConnectionError('This file is not a valid connection-test sample. Save a new sample and try again.');
  }
  // Return only the expected fields, never arbitrary data from a modified file.
  return { schemaVersion: 1, kind: TEST_KIND, message: MESSAGE, code: data.code, savedAt: data.savedAt };
}

function responseError(status) {
  if (status === 401) return new ConnectionError('Your Microsoft session needs refreshing. Sign out and sign in again.', status);
  if (status === 403) return new ConnectionError('Microsoft denied access to the app folder. Check that Files.ReadWrite.AppFolder is a delegated permission and sign in again to approve it.', status);
  if (status === 404) return new ConnectionError('The OneDrive item could not be found. Check your account and try again.', status);
  if (status === 409 || status === 412) return new ConnectionError('OneDrive reported a file conflict. No existing financial file was changed. Try saving a new sample.', status);
  if (status === 429) return new ConnectionError('OneDrive is receiving too many requests. Wait a moment and try again.', status);
  if (status === 507) return new ConnectionError('OneDrive is out of storage. Free some space and try again.', status);
  return new ConnectionError(`OneDrive could not complete the request (HTTP ${status}). Please try again.`, status);
}

export class OneDriveTest {
  constructor(getAccessToken, fetcher = fetch) { this.getAccessToken = getAccessToken; this.fetcher = fetcher; }

  async request(path, options = {}) {
    const url = new URL(path.startsWith('https:') ? path : `${GRAPH}${path}`);
    if (url.origin !== 'https://graph.microsoft.com' || !url.pathname.startsWith('/v1.0/')) throw new ConnectionError('Invalid OneDrive request address.');
    const token = await this.getAccessToken();
    let response;
    try {
      response = await this.fetcher(url.href, {
        ...options, credentials: 'omit', cache: 'no-store', referrerPolicy: 'no-referrer',
        signal: AbortSignal.timeout(25000),
        headers: { ...options.headers, Authorization: `Bearer ${token}` },
      });
    } catch {
      throw new ConnectionError('Could not reach OneDrive. Check your connection and try again.');
    }
    if (!response.ok) throw responseError(response.status);
    return response.json();
  }

  async folder() {
    const folder = await this.request('/me/drive/special/approot');
    if (!folder?.id || !folder.folder) throw new ConnectionError('OneDrive did not return an application folder.');
    return folder;
  }

  async save() {
    const folder = await this.folder();
    const sample = makeSample();
    // Independent sample files avoid overwriting a prior sample from another device.
    const name = `${PREFIX}${sample.savedAt.replaceAll(':', '-')}-${crypto.randomUUID()}.json`;
    const file = await this.request(`/me/drive/items/${encodeURIComponent(folder.id)}:/${encodeURIComponent(name)}:/content`, {
      method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(sample),
    });
    if (!file?.id) throw new ConnectionError('OneDrive did not confirm a saved file. Load the latest sample before retrying.');
    return { sample, name };
  }

  async load() {
    const folder = await this.folder();
    let path = `/me/drive/items/${encodeURIComponent(folder.id)}/children?$select=id,name,size,lastModifiedDateTime&$top=100`;
    const candidates = [];
    let pages = 0;
    while (path) {
      if (++pages > 20) throw new ConnectionError('The app folder contains too many files for this initial test.');
      const result = await this.request(path);
      if (!Array.isArray(result.value)) throw new ConnectionError('OneDrive returned an unexpected folder listing.');
      candidates.push(...result.value.filter(file => typeof file.name === 'string' && file.name.startsWith(PREFIX) && file.name.endsWith('.json')));
      path = result['@odata.nextLink'];
    }
    if (!candidates.length) return null;
    candidates.sort((a, b) => String(b.lastModifiedDateTime).localeCompare(String(a.lastModifiedDateTime)) || b.name.localeCompare(a.name));
    const latest = candidates[0];
    if (latest.size > MAX_BYTES) throw new ConnectionError('The latest sample is larger than expected. Save a new test sample.');
    const metadata = await this.request(`/me/drive/items/${encodeURIComponent(latest.id)}?$select=id,size,@microsoft.graph.downloadUrl`);
    const address = metadata['@microsoft.graph.downloadUrl'];
    if (typeof address !== 'string' || !address.startsWith('https://') || metadata.size > MAX_BYTES) throw new ConnectionError('OneDrive did not return a valid sample download.');
    // Graph /content redirects fail CORS preflight in browsers. Use its signed URL,
    // with no bearer token, cookies, logging, or storage of that URL.
    let response;
    try {
      response = await this.fetcher(address, { credentials: 'omit', cache: 'no-store', referrerPolicy: 'no-referrer', signal: AbortSignal.timeout(25000) });
    } catch { throw new ConnectionError('Could not download the sample. Check your connection and try loading again.'); }
    if (!response.ok) throw new ConnectionError('The sample download expired or failed. Click Load latest sample again.');
    const text = await response.text();
    if (text.length > MAX_BYTES) throw new ConnectionError('The sample is larger than expected.');
    try { return validateSample(JSON.parse(text)); }
    catch (error) { if (error instanceof ConnectionError) throw error; throw new ConnectionError('The sample file could not be read. Save a new sample and try again.'); }
  }
}
