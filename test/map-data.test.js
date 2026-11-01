import test from 'node:test';
import assert from 'node:assert/strict';

// 1. Test isAmountKey & getComparableValue
function isAmountKey(key) {
  const k = String(key || '').toLowerCase().replace(/[^a-z0-9]/g, '');
  return k.includes('cost') || k.includes('amount') || k.includes('amt') || k.includes('expenditure') || k.includes('price');
}

function parseNumericAmount(value) {
  if (value === undefined || value === null || value === '' || value === '-') return 0;
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  const cleaned = String(value).replace(/,/g, '').replace(/[^\d.-]/g, '').trim();
  const num = parseFloat(cleaned);
  return Number.isFinite(num) ? num : 0;
}

function getMapDataComparableValue(key, value) {
  if (value === undefined || value === null || value === '' || value === '-') return '';
  if (isAmountKey(key)) {
    const num = parseNumericAmount(value);
    if (Number.isFinite(num)) return `num:${Math.round(num * 100) / 100}`;
  }
  return String(value).trim().replace(/\s+/g, ' ').toLowerCase();
}

function mapDataValuesMatch(key, existingValue, incomingValue) {
  return getMapDataComparableValue(key, existingValue) === getMapDataComparableValue(key, incomingValue);
}

test('Amount comparison eliminates false conflicts across Indian formatted and raw numbers', () => {
  // Existing: "12,50,000", Incoming: 1250000 -> Match
  assert.equal(mapDataValuesMatch('DNIT_Cost', '12,50,000', 1250000), true);
  // Existing: "150.00", Incoming: "150" -> Match
  assert.equal(mapDataValuesMatch('Agreement Amount', '150.00', '150'), true);
  // Existing: "50000", Incoming: "60000" -> Differ
  assert.equal(mapDataValuesMatch('Estimate_Cost', '50000', '60000'), false);
});

// 2. Test Header row detection and deduplication
function detectHeaderRowAndDeduplicate(raw) {
  let headerRowIndex = 0;
  for (let r = 0; r < Math.min(raw.length, 10); r++) {
    const rowCells = (raw[r] || []).map(c => String(c ?? '').trim());
    const hasKeyword = rowCells.some(c => {
      const n = c.toLowerCase().replace(/[^a-z0-9]/g, '');
      return n.includes('work') || n.includes('tender') || n.includes('dnit') || n.includes('name') || n.includes('sr');
    });
    if (hasKeyword && rowCells.filter(Boolean).length >= 2) {
      headerRowIndex = r;
      break;
    }
  }

  const seenHeaders = {};
  const headerRow = (raw[headerRowIndex] || []).map(cell => String(cell ?? '').trim());
  const headers = headerRow.map((header, index) => {
    const base = header || `Column_${index + 1}`;
    if (!seenHeaders[base]) {
      seenHeaders[base] = 1;
      return base;
    }
    seenHeaders[base]++;
    return `${base} (${seenHeaders[base]})`;
  });

  return { headerRowIndex, headers };
}

test('Detects header row past title banners and deduplicates repeated headers', () => {
  const rawSheet = [
    ['MUNICIPAL CORPORATION DADRI', ''],
    ['PROGRESS REPORT OF WORKS AS ON 01-11-2024', ''],
    ['Sr. No.', 'Work Name', 'Tender ID', 'Date', 'Date', 'Amount']
  ];
  const { headerRowIndex, headers } = detectHeaderRowAndDeduplicate(rawSheet);
  assert.equal(headerRowIndex, 2);
  assert.deepEqual(headers, ['Sr. No.', 'Work Name', 'Tender ID', 'Date', 'Date (2)', 'Amount']);
});

// 3. Test Primary Tender ID matching vs fuzzy name matching
function resolvePreviewMatch(excelWorkName, incoming, masterData) {
  const incomingTenderId = String(
    incoming?.E_TENDERID ||
    incoming?.E_TenderID ||
    incoming?.tenderId ||
    ''
  ).trim();

  if (incomingTenderId) {
    const tenderMatch = masterData.find(row => String(row.E_TENDERID || row.E_TenderID || '').trim() === incomingTenderId);
    if (tenderMatch) {
      return { item: tenderMatch, score: 100, layer: 'tender_id' };
    }
  }

  const normalized = excelWorkName.trim().toLowerCase();
  const nameMatch = masterData.find(row => (row.Work_Name || '').trim().toLowerCase() === normalized);
  if (nameMatch) {
    return { item: nameMatch, score: 100, layer: 'exact_name' };
  }

  return null;
}

test('Prioritizes Tender ID match when present even if work name differs', () => {
  const masterData = [
    { E_TENDERID: 'T101', Work_Name: 'Special Repair of Road in Ward 5' },
    { E_TENDERID: 'T102', Work_Name: 'Construction of RCC Drain in Ward 6' }
  ];

  // Incoming has T101 but work name is abbreviated "SR Road W5"
  const match = resolvePreviewMatch('SR Road W5', { E_TENDERID: 'T101' }, masterData);
  assert.ok(match);
  assert.equal(match.layer, 'tender_id');
  assert.equal(match.item.E_TENDERID, 'T101');
});
