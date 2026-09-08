import { PublicClientApplication, InteractionRequiredAuthError, CacheLookupPolicy } from '@azure/msal-browser';
import { CLIENT_ID, SITE_URL, AUTHORITY, SCOPES } from './config.js';
import { OneDriveTest, ConnectionError } from './onedrive.js';
import { parseCsv, normalizeTransactions } from './transactions.js';

const $ = id => document.getElementById(id);
const production = location.origin === new URL(SITE_URL).origin && location.pathname.startsWith(new URL(SITE_URL).pathname);
let busy = false;
let ready = false;
let account = null;
let transactions = [];
const msal = new PublicClientApplication({
  auth: { clientId: CLIENT_ID, authority: AUTHORITY, redirectUri: SITE_URL, postLogoutRedirectUri: SITE_URL, navigateToLoginRequestUrl: false },
  cache: { cacheLocation: 'sessionStorage' },
});

function notice(text, kind = '') { $('notice').textContent = text; $('notice').className = `notice ${kind}`; }
function displayAccount() {
  $('account-name').textContent = account?.username || account?.name || 'No account connected';
  $('account-state').textContent = account ? 'Signed in' : 'Not signed in';
  $('connection-badge').textContent = account ? 'Microsoft connected' : 'Not connected';
  $('connection-badge').className = account ? 'badge connected' : 'badge';
  $('sign-in').hidden = !!account;
  $('sign-out').hidden = !account;
  $('sign-in').disabled = busy || !ready || !production;
  $('sign-out').disabled = busy;
  $('save-sample').disabled = busy || !ready || !account || !production;
  $('load-sample').disabled = busy || !ready || !account || !production;
  $('budget-workspace').hidden = !account;
  $('command-center').hidden = !account;
  $('transactions-workspace').hidden = !account;
  if (account) {
    document.querySelector('.heading h1').textContent = 'Your financial command center';
    document.querySelector('.heading .eyebrow').textContent = 'MONTHLY OVERVIEW';
    document.querySelector('.heading .intro').textContent = 'Your private budget and financial records, saved to OneDrive.';
    document.querySelector('.content').insertBefore($('command-center'), document.querySelector('.connection-card'));
    document.querySelector('.test-grid').hidden = true;
    document.querySelector('.device-check').hidden = true;
  }
  document.querySelector('main').setAttribute('aria-busy', String(busy));
}

function clearResults() {
  $('sample-code').textContent = 'No sample loaded yet'; $('sample-code').className = '';
  $('sample-time').textContent = 'Your saved test code will appear here.';
  $('loaded-sample').classList.add('empty');
  $('save-status').textContent = 'Sign in to enable saving.';
  $('load-status').textContent = 'Sign in to enable loading.';
}

function friendlyError(error) {
  if (error instanceof ConnectionError) return error.message;
  const code = String(error?.errorCode || '');
  if (code.includes('interaction_in_progress')) return 'A Microsoft sign-in is already in progress. Complete it, or reload this page and try again.';
  if (code.includes('user_cancelled') || code.includes('access_denied')) return 'Sign-in was cancelled or permission was declined. You can try again when ready.';
  if (code.includes('network') || code.includes('timeout')) return 'Microsoft sign-in could not connect. Check your internet connection and try again.';
  return 'Microsoft sign-in could not finish. Check the app’s SPA redirect address, personal-account support, and delegated OneDrive app-folder permission, then try again.';
}

async function act(work) {
  if (busy) return;
  busy = true; displayAccount();
  try { await work(); }
  catch (error) { notice(friendlyError(error), 'error'); }
  finally { busy = false; displayAccount(); }
}

async function token() {
  if (!account) throw new ConnectionError('Sign in before accessing OneDrive.');
  try {
    const result = await msal.acquireTokenSilent({ account, scopes: SCOPES, cacheLookupPolicy: CacheLookupPolicy.AccessTokenAndRefreshToken });
    return result.accessToken;
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError || ['token_refresh_required', 'refresh_token_expired', 'no_tokens_found'].includes(error?.errorCode)) {
      notice('Microsoft needs you to sign in again. After returning, repeat your save or load.');
      await msal.acquireTokenRedirect({ account, scopes: SCOPES });
    }
    throw error;
  }
}
const drive = new OneDriveTest(token);

function updateBudgetTotals() {
  const expenses = [...document.querySelectorAll('[data-budget]')].reduce((sum, input) => sum + (Number(input.value) || 0), 0);
  const income = Number($('budget-income').value) || 0;
  $('planned-expenses').textContent = expenses.toLocaleString('en-US', { style: 'currency', currency: 'USD' });
  $('planned-surplus').textContent = (income - expenses).toLocaleString('en-US', { style: 'currency', currency: 'USD' });
}
function displayBudget(saved) {
  if (!saved) return;
  const { month, plannedIncome, categories = {} } = saved.budget;
  transactions = Array.isArray(saved.budget.transactions) ? saved.budget.transactions : [];
  $('transaction-status').textContent = transactions.length ? `${transactions.length} saved transactions loaded from OneDrive.` : 'No transactions saved yet.';
  $('budget-month').value = month || $('budget-month').value;
  $('budget-income').value = Number.isFinite(plannedIncome) ? plannedIncome : '';
  document.querySelectorAll('[data-budget]').forEach(input => { input.value = Number.isFinite(categories[input.dataset.budget]) ? categories[input.dataset.budget] : ''; });
  updateBudgetTotals();
  const income = Number(plannedIncome) || 0;
  const expenses = Object.values(categories).reduce((sum, value) => sum + (Number(value) || 0), 0);
  const money = value => value.toLocaleString('en-US', { style: 'currency', currency: 'USD' });
  $('command-month').textContent = new Date(`${month || $('budget-month').value}-01T12:00:00`).toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
  $('metric-income').textContent = money(income); $('metric-expenses').textContent = money(expenses); $('metric-surplus').textContent = money(income - expenses);
  const summary = $('category-summary'); summary.replaceChildren();
  const entries = Object.entries(categories).filter(([, value]) => Number(value));
  if (!entries.length) { const empty = document.createElement('p'); empty.className = 'muted'; empty.textContent = 'No planned expenses yet.'; summary.append(empty); }
  entries.forEach(([name, value]) => { const row = document.createElement('div'); row.className = 'category-row'; const label = document.createElement('span'); label.textContent = name; const amount = document.createElement('strong'); amount.textContent = money(Number(value)); row.append(label, amount); summary.append(row); });
  $('budget-status').textContent = `Loaded your saved budget from ${new Date(saved.savedAt).toLocaleString()}.`;
}
document.querySelectorAll('#budget-form input').forEach(input => input.addEventListener('input', updateBudgetTotals));
$('budget-month').value = new Date().toISOString().slice(0, 7);
$('budget-form').addEventListener('submit', event => act(async () => {
  event.preventDefault(); updateBudgetTotals();
  const categories = Object.fromEntries([...document.querySelectorAll('[data-budget]')].map(input => [input.dataset.budget, Number(input.value) || 0]));
  $('budget-status').textContent = 'Saving your budget to OneDrive…';
  const saved = await drive.saveBudget({ month: $('budget-month').value, plannedIncome: Number($('budget-income').value) || 0, categories, transactions });
  displayBudget(saved);
  $('budget-status').textContent = `Saved to OneDrive ${new Date(saved.savedAt).toLocaleString()}.`;
  notice('Your monthly budget is saved privately in OneDrive.', 'success');
}));
$('transaction-file').addEventListener('change', async event => {
  const file = event.target.files?.[0];
  if (!file) return;
  try {
    transactions = normalizeTransactions(parseCsv(await file.text()));
    $('transaction-status').textContent = `${transactions.length} transactions ready. Save your budget to store them in OneDrive.`;
  } catch { $('transaction-status').textContent = 'That CSV could not be read. Export a standard bank CSV and try again.'; }
});

$('sign-in').addEventListener('click', () => act(async () => {
  notice('Opening Microsoft sign-in…');
  await msal.loginRedirect({ scopes: SCOPES, prompt: 'select_account' });
}));
$('sign-out').addEventListener('click', () => act(async () => {
  const signedIn = account;
  account = null; clearResults(); displayAccount();
  notice('Signing out of Microsoft…');
  await msal.logoutRedirect({ account: signedIn, postLogoutRedirectUri: SITE_URL });
}));
$('save-sample').addEventListener('click', () => act(async () => {
  notice('Saving a new sample to your OneDrive app folder…');
  const { sample } = await drive.save();
  $('save-status').textContent = `Saved test code: ${sample.code}`;
  notice('Sample saved to OneDrive. Select Load latest sample here or on your other device to compare the test code.', 'success');
}));
$('load-sample').addEventListener('click', () => act(async () => {
  notice('Reading the latest sample from OneDrive…');
  const sample = await drive.load();
  if (!sample) {
    $('sample-code').textContent = 'No sample found'; $('sample-code').className = '';
    $('sample-time').textContent = 'Save a sample with this Microsoft account first.';
    $('loaded-sample').classList.add('empty');
    $('load-status').textContent = 'OneDrive connected; no test sample found.';
    notice('Your OneDrive connection worked. No sample exists yet—save one first.');
    return;
  }
  $('sample-code').textContent = sample.code; $('sample-code').className = 'loaded';
  $('sample-time').textContent = `Saved ${new Date(sample.savedAt).toLocaleString()}`;
  $('loaded-sample').classList.remove('empty');
  $('load-status').textContent = 'Loaded directly from OneDrive.';
  notice('Sample loaded successfully. A matching test code on your second device confirms access to the same saved file.', 'success');
}));

async function start() {
  try {
    if (!production) {
      notice('Local preview. Microsoft sign-in is enabled at the published GitHub Pages address.');
      return;
    }
    await msal.initialize();
    const response = await msal.handleRedirectPromise();
    if (response?.account) msal.setActiveAccount(response.account);
    account = msal.getActiveAccount();
    ready = true;
    notice(account ? 'Microsoft sign-in is complete. Save or load a sample to check your OneDrive connection.' : 'Ready to connect. Sign in with your personal Microsoft account.');
    if (account) {
      $('save-status').textContent = 'Ready to save a sample.';
      $('load-status').textContent = 'Ready to read from OneDrive.';
      try { displayBudget(await drive.loadBudget()); }
      catch (error) { $('budget-status').textContent = 'Your saved budget could not be loaded automatically. You can still enter and save a new one.'; }
    }
  } catch (error) {
    // Initialization failures cannot be recovered by calling login on an uninitialized instance.
    ready = false; notice(friendlyError(error) + ' Reload the page after correcting the setting.', 'error');
  } finally { displayAccount(); }
}
start();
