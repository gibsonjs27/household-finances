import test from 'node:test';
import assert from 'node:assert/strict';
import { normalizeTransactions, parseCsv, recategorizeTransactions } from '../src/transactions.js';

test('imports a Chase CSV and applies matching categories', () => {
  const csv = 'Details,Posting Date,Description,Amount,Type\nDEBIT,09/01/2026,T-MOBILE AUTOPAY,-243.23,ACH_DEBIT\nCREDIT,09/02/2026,GDMS PAYROLL,4242.00,ACH_CREDIT';
  assert.deepEqual(normalizeTransactions(parseCsv(csv)), [
    { date: '2026-09-01', description: 'T-MOBILE AUTOPAY', amount: -243.23, category: 'Utilities' },
    { date: '2026-09-02', description: 'GDMS PAYROLL', amount: 4242, category: 'Income' },
  ]);
});

test('preserves commas within quoted bank descriptions', () => {
  const rows = parseCsv('Posting Date,Description,Amount\n09/03/2026,"KROGER, STORE 123",-52.14');
  assert.equal(rows[0].Description, 'KROGER, STORE 123');
  assert.equal(normalizeTransactions(rows)[0].category, 'Living');
});

test('reclassifies previously saved uncategorized transactions', () => {
  assert.deepEqual(recategorizeTransactions([{ date: '2026-09-01', description: 'NETFLIX.COM', amount: -15.49, category: 'Other' }]), [{ date: '2026-09-01', description: 'NETFLIX.COM', amount: -15.49, category: 'Subscriptions' }]);
});

test('keeps a category selected during review', () => {
  assert.equal(recategorizeTransactions([{ date: '2026-09-01', description: 'MERCHANT', amount: -15.49, category: 'Housing' }])[0].category, 'Housing');
});
