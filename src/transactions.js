function parseAmount(value) {
  const cleaned = String(value ?? '').replace(/[$,\s]/g, '').replace(/^\((.*)\)$/, '-$1');
  const amount = Number(cleaned);
  return Number.isFinite(amount) ? amount : 0;
}

function isoDate(value) {
  const trimmed = String(value ?? '').trim();
  const usDate = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (usDate) return `${usDate[3]}-${usDate[1].padStart(2, '0')}-${usDate[2].padStart(2, '0')}`;
  const parsed = new Date(trimmed);
  return Number.isNaN(parsed.valueOf()) ? trimmed : parsed.toISOString().slice(0, 10);
}

export function suggestedCategory({ description = '', category = '', type = '' }) {
  const desc = String(description).toLowerCase();
  const sourceCategory = String(category).toLowerCase();
  const sourceType = String(type).toLowerCase();
  if (sourceType.includes('transfer') || sourceType.includes('payment') || ['credit card payments', 'transfers'].includes(sourceCategory)) return 'Transfer';
  if (/\b(gdms|payroll|salary)\b/.test(desc) || (/\b(va|treas)\b/.test(desc) && /benefit|disability|deposit/.test(desc))) return 'Income';
  if (/t-?mobile|frontier|electric|electricity|water|sewer|gas bill|utility|comcast|xfinity|spectrum|at&t internet/.test(desc) || /utilit/.test(sourceCategory)) return 'Utilities';
  if (/disney|youtube tv|netflix|hulu|spotify|apple\.com\/bill|texas law shield|subscription/.test(desc) || /entertainment|subscription/.test(sourceCategory)) return 'Subscriptions';
  if (/usaa.*insurance|insurance|geico|progressive|state farm|allstate/.test(desc) || /insurance/.test(sourceCategory)) return 'Insurance';
  if (/td auto|auto loan|car payment|loan payment|mortgage|rent\b|landlord/.test(desc)) return /mortgage|rent\b|landlord/.test(desc) ? 'Housing' : 'Debt payments';
  if (/gas station|\bfuel\b|\bgas\b|shell|chevron|exxon|mobil|uber|lyft|parking|toll/.test(desc) || /gas|fuel/.test(sourceCategory)) return 'Transportation';
  if (/grocery|groceries|market|kroger|heb|h-e-b|walmart|target|costco|restaurant|dining|cafe|coffee|doordash|grubhub|ubereats/.test(desc) || /grocery|restaurant|dining|food/.test(sourceCategory)) return 'Living';
  if (/housing|home/.test(sourceCategory)) return 'Housing';
  if (/debt|loan/.test(sourceCategory)) return 'Debt payments';
  if (/transport/.test(sourceCategory)) return 'Transportation';
  if (/deposit|income|payroll/.test(sourceCategory)) return 'Income';
  return 'Other';
}

export function parseCsv(text) {
  const rows = [];
  let row = [], field = '', quoted = false;
  const input = String(text).replace(/^\uFEFF/, '');
  for (let index = 0; index < input.length; index += 1) {
    const character = input[index];
    if (character === '"') {
      if (quoted && input[index + 1] === '"') { field += '"'; index += 1; } else quoted = !quoted;
    } else if (character === ',' && !quoted) { row.push(field.trim()); field = ''; }
    else if ((character === '\n' || character === '\r') && !quoted) {
      if (character === '\r' && input[index + 1] === '\n') index += 1;
      row.push(field.trim()); field = '';
      if (row.some(value => value)) rows.push(row);
      row = [];
    } else field += character;
  }
  row.push(field.trim());
  if (row.some(value => value)) rows.push(row);
  const [headers = [], ...data] = rows;
  return data.map(values => Object.fromEntries(headers.map((key, index) => [key, values[index] ?? ''])));
}

export function normalizeTransactions(rows) {
  return rows.map(row => {
    const description = row.Description || row.Details || row.Memo || '';
    const type = row.Type || row.type || row['Type Group'] || '';
    const details = String(row.Details || '').toLowerCase();
    let amount = parseAmount(row.Amount);
    if (details === 'debit') amount = -Math.abs(amount);
    if (details === 'credit') amount = Math.abs(amount);
    return { date: isoDate(row['Posting Date'] || row['Post Date'] || row['Transaction Date'] || row.Date), description, amount, category: suggestedCategory({ description, category: row.Category, type }) };
  }).filter(row => row.date && row.description);
}

export function recategorizeTransactions(rows) {
  return rows.map(row => ({ ...row, category: row.category && row.category !== 'Other' ? row.category : suggestedCategory(row) }));
}
