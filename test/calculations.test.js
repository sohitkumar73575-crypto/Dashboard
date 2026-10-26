import test from 'node:test';
import assert from 'node:assert/strict';

// Test parseDate logic
function parseDate(dateStr) {
  if (!dateStr || dateStr === '-' || String(dateStr).trim() === '') return null;
  if (dateStr instanceof Date && !isNaN(dateStr.getTime())) {
    return new Date(dateStr.getFullYear(), dateStr.getMonth(), dateStr.getDate());
  }
  const text = String(dateStr).trim();
  if (/^\d+(\.\d+)?$/.test(text)) {
    const serial = Number(text);
    const utc = Date.UTC(1899, 11, 30) + Math.round(serial) * 86400000;
    const d = new Date(utc);
    return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
  }
  let match = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (match) return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  match = text.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/);
  if (match) {
    const year = Number(match[3].length === 2 ? '20' + match[3] : match[3]);
    return new Date(year, Number(match[2]) - 1, Number(match[1]));
  }
  return null;
}

// Test Additional PBG calculation formula
function calculateAdditionalBgAmount(dnitAmt, quotedAmt) {
  if (!dnitAmt || !quotedAmt) return 0;
  const diffPercent = ((dnitAmt - quotedAmt) / dnitAmt) * 100;
  const extraPercent = diffPercent > 10 ? diffPercent - 10 : 0;
  return Math.ceil((quotedAmt * (extraPercent / 100)) / 10) * 10;
}

test('parseDate handles Indian DD/MM/YYYY dates', () => {
  const d = parseDate('15/08/2024');
  assert.ok(d instanceof Date);
  assert.equal(d.getDate(), 15);
  assert.equal(d.getMonth(), 7); // August is index 7
  assert.equal(d.getFullYear(), 2024);
});

test('parseDate handles YYYY-MM-DD and Excel serial numbers', () => {
  const iso = parseDate('2024-08-15');
  assert.equal(iso.getDate(), 15);
  assert.equal(iso.getFullYear(), 2024);

  const empty = parseDate('-');
  assert.equal(empty, null);
});

test('calculateAdditionalBgAmount correctly computes additional PBG when discount > 10%', () => {
  // DNIT: 10,00,000, Quoted: 8,00,000 (20% below, extra = 10%)
  // Extra PBG: 8,00,000 * 10% = 80,000
  const result = calculateAdditionalBgAmount(1000000, 800000);
  assert.equal(result, 80000);

  // DNIT: 10,00,000, Quoted: 9,50,000 (5% below, under 10% threshold -> 0)
  const noExtra = calculateAdditionalBgAmount(1000000, 950000);
  assert.equal(noExtra, 0);
});
