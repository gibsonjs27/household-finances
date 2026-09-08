export function parseCsv(text) {
  const [header, ...lines] = text.replace(/^\uFEFF/, '').trim().split(/\r?\n/);
  if (!header) return [];
  const keys = header.split(',').map(value => value.trim());
  return lines.map(line => Object.fromEntries(line.split(',').map(value => value.trim()).map((value, index) => [keys[index], value])));
}

export function normalizeTransactions(rows) {
  return rows.map(row => ({
    date: (() => { const value = row['Posting Date'] || row['Post Date'] || row['Transaction Date'] || ''; const parsed = new Date(value); return Number.isNaN(parsed.valueOf()) ? value : parsed.toISOString().slice(0, 10); })(),
    description: row.Description || row.Details || '',
    amount: Number(String(row.Amount || '').replace(/[$,]/g, '')) || 0,
    category: row.Category || 'Other',
  })).filter(row => row.date && row.description);
}
