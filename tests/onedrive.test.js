import test from 'node:test';
import assert from 'node:assert/strict';
import { OneDriveTest, makeSample, validateSample, ConnectionError } from '../src/onedrive.js';

const fixed = makeSample(new Date('2026-09-06T12:00:00.000Z'), '01234567-89ab-4def-8123-456789abcdef');
const json = value => new Response(JSON.stringify(value), { status: 200, headers: {'Content-Type':'application/json'} });

test('sample contains only synthetic fields and validates before display', () => {
  assert.equal(fixed.code, '0123456789AB');
  assert.deepEqual(validateSample({...fixed, other: '<script>bad</script>'}), fixed);
  for (const change of [{schemaVersion:2}, {message:'Changed'}, {code:'<script>'}, {savedAt:'invalid'}, {kind:'financial-data'}]) {
    assert.throws(() => validateSample({...fixed, ...change}), ConnectionError);
  }
});

test('save uses an isolated app folder and a unique file on each call', async () => {
  const writes = [];
  const drive = new OneDriveTest(async () => 'test-token', async (url, options) => {
    assert.equal(options.headers.Authorization, 'Bearer test-token');
    assert.equal(options.method, 'PUT');
    assert.match(url, /\/special\/approot:\/household-finances-connection-test-/);
    assert.match(url, /\.json:\/content$/);
    validateSample(JSON.parse(options.body));
    writes.push(url); return json({id:'new-file'});
  });
  await drive.save(); await drive.save();
  assert.equal(writes.length, 2);
  assert.notEqual(writes[0], writes[1]);
});

test('load returns null only for a successful folder listing with no samples', async () => {
  const drive = new OneDriveTest(async () => 'test-token', async () => json({value:[{id:'budget',name:'budget.json'}]}));
  assert.equal(await drive.load(), null);
});

test('load follows pagination and downloads without sending Graph credentials', async () => {
  let downloads = 0;
  const drive = new OneDriveTest(async () => 'test-token', async (url, options) => {
    if (url.includes('/children?')) return json({value:[{id:'older',name:'household-finances-connection-test-old.json',lastModifiedDateTime:'2026-09-01T00:00:00Z'}], '@odata.nextLink':'https://graph.microsoft.com/v1.0/page-two'});
    if (url.endsWith('/page-two')) return json({value:[{id:'newest',name:'household-finances-connection-test-new.json',lastModifiedDateTime:'2026-09-06T00:00:00Z'}]});
    if (url.includes('/approot:')) return json({id:'newest',size:200,'@microsoft.graph.downloadUrl':'https://download.example.test/signed-sample'});
    assert.equal(url, 'https://download.example.test/signed-sample');
    assert.equal(options.headers, undefined);
    assert.equal(options.credentials, 'omit');
    assert.equal(options.referrerPolicy, 'no-referrer');
    downloads++; return json(fixed);
  });
  assert.deepEqual(await drive.load(), fixed);
  assert.equal(downloads, 1);
});

test('refuses a pagination address outside Microsoft Graph before obtaining a token', async () => {
  let tokenRequests = 0;
  const drive = new OneDriveTest(async () => {tokenRequests++; return 'test-token';});
  await assert.rejects(drive.request('https://attacker.example/v1.0/steal'), /Invalid OneDrive/);
  assert.equal(tokenRequests, 0);
});

test('permission failures are errors, not empty folders or successful saves', async () => {
  const drive = new OneDriveTest(async () => 'test-token', async () => new Response('', {status:403}));
  await assert.rejects(drive.load(), /denied access/);
  await assert.rejects(drive.save(), /denied access/);
});

test('network failures do not claim a sample was saved', async () => {
  const drive = new OneDriveTest(async () => 'test-token', async () => {throw new TypeError('offline');});
  await assert.rejects(drive.save(), /Could not reach OneDrive/);
});

test('uses the browser context when calling the native fetch function', async () => {
  const browserContext = {};
  let receivedContext;
  const nativeLikeFetch = function () { receivedContext = this; return json({id: 'new-file'}); };
  const drive = new OneDriveTest(async () => 'test-token', nativeLikeFetch);
  const originalGlobalThis = globalThis;
  // The function only needs a browser-style receiver; in Node it is globalThis.
  await drive.save();
  assert.equal(receivedContext, originalGlobalThis);
});

test('rejects an oversized sample without downloading its contents', async () => {
  const drive = new OneDriveTest(async () => 'test-token', async url => {
    return json({value:[{id:'large',name:'household-finances-connection-test-big.json',size:99999}]});
  });
  await assert.rejects(drive.load(), /larger than expected/);
});

test('ignores a credential-bearing download scheme', async () => {
  const drive = new OneDriveTest(async () => 'test-token', async url => {
    if (url.includes('/children')) return json({value:[{id:'sample',name:'household-finances-connection-test-x.json'}]});
    return json({'@microsoft.graph.downloadUrl':'http://unsafe.example/test'});
  });
  await assert.rejects(drive.load(), /valid sample download/);
});
