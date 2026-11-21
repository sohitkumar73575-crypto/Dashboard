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

// 4. Test Column Mapping Alignment for Estimate Cost (Col B), Admin Approval Date (Col C), Tech Approval Date (Col D)
const ALIAS_MAP = {
  Estimate_Cost: ['Estimate Cost', 'Estimated Cost', 'Estimate_Cost', 'Estimated_Cost', 'Estimate Amount', 'Estimated Amount', 'Est Cost', 'Est Amount'],
  'Admin Approval_Date': ['Admin_Approval_Date', 'Admin Approval Date', 'Admin Approved Date', 'Admin Approval', 'Admin Date', 'AA Date', 'Admin Sanction Date', 'Administrative Approval Date'],
  Tech_Approval_Date: ['Tech_Approval_Date', 'Technical Approval Date', 'Tech Approval Date', 'Technical Approval', 'Tech Approval', 'Technical Sanction Date', 'Tech Sanction Date', 'TS Date', 'Tech Date', 'Technical Date'],
  Expenditure: ['Expenditure', 'Expenditure (Rs.)', 'Expenditure_Amount', 'Expenditure Amount', 'Expenditure_Cost', 'Expenditure Cost', 'Final Expenditure', 'EXPENDITURE', 'Expendeture', 'Expendeture (Rs.)']
};

function normalizeKey(str) {
  return String(str || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function resolveFieldFromHeader(header, dashboardKeys) {
  const norm = normalizeKey(header);
  for (const key of dashboardKeys) {
    if (normalizeKey(key) === norm) return key;
  }
  for (const [canonical, aliases] of Object.entries(ALIAS_MAP)) {
    if (aliases.some(a => normalizeKey(a) === norm)) {
      return dashboardKeys.find(k => normalizeKey(k) === normalizeKey(canonical)) || canonical;
    }
  }
  return '';
}

test('Header suggestion accurately maps Estimated Cost, Admin Approval, and Tech Approval without cross-contamination', () => {
  const sheetKeys = ['Work_Name', 'Estimate_Cost', 'Admin Approval_Date', 'Tech_Approval_Date', 'DNIT_Cost'];

  // Test Estimated Cost mappings (Col B)
  assert.equal(resolveFieldFromHeader('Estimated Cost', sheetKeys), 'Estimate_Cost');
  assert.equal(resolveFieldFromHeader('Estimate Cost', sheetKeys), 'Estimate_Cost');
  assert.equal(resolveFieldFromHeader('Estimate Amount', sheetKeys), 'Estimate_Cost');

  // Test Admin Approval mappings (Col C) - MUST NEVER map to Estimate_Cost
  assert.equal(resolveFieldFromHeader('Admin Approval Date', sheetKeys), 'Admin Approval_Date');
  assert.equal(resolveFieldFromHeader('Admin Approval', sheetKeys), 'Admin Approval_Date');
  assert.equal(resolveFieldFromHeader('AA Date', sheetKeys), 'Admin Approval_Date');
  assert.notEqual(resolveFieldFromHeader('Admin Approval Date', sheetKeys), 'Estimate_Cost');

  // Test Technical Approval mappings (Col D)
  assert.equal(resolveFieldFromHeader('Technical Approval Date', sheetKeys), 'Tech_Approval_Date');
  assert.equal(resolveFieldFromHeader('Tech Approval Date', sheetKeys), 'Tech_Approval_Date');
  assert.equal(resolveFieldFromHeader('Technical Sanction Date', sheetKeys), 'Tech_Approval_Date');
  assert.equal(resolveFieldFromHeader('TS Date', sheetKeys), 'Tech_Approval_Date');
});

test('Sheet column updater routes values to correct column indices (Col B: Estimate, Col C: Admin, Col D: Tech)', () => {
  const sheetHeaders = ['Work_Name', 'Estimate_Cost', 'Admin Approval_Date', 'Tech_Approval_Date', 'DNIT_Cost'];
  const incomingDetails = {
    'Estimated Cost': '1500000',
    'Admin Approval Date': '15/08/2024',
    'Technical Approval Date': '20/08/2024'
  };

  const updates = [];
  sheetHeaders.forEach((header, index) => {
    let matchedVal = incomingDetails[header];
    if (matchedVal === undefined) {
      for (const k of Object.keys(incomingDetails)) {
        if (resolveFieldFromHeader(k, sheetHeaders) === header) {
          matchedVal = incomingDetails[k];
          break;
        }
      }
    }
    if (matchedVal !== undefined) {
      updates.push({ column: index + 1, header, value: matchedVal });
    }
  });

  // Verify Col 2 (B) is Estimate_Cost with '1500000'
  assert.deepEqual(updates.find(u => u.column === 2), { column: 2, header: 'Estimate_Cost', value: '1500000' });
  // Verify Col 3 (C) is Admin Approval_Date with '15/08/2024'
  assert.deepEqual(updates.find(u => u.column === 3), { column: 3, header: 'Admin Approval_Date', value: '15/08/2024' });
  // Verify Col 4 (D) is Tech_Approval_Date with '20/08/2024'
  assert.deepEqual(updates.find(u => u.column === 4), { column: 4, header: 'Tech_Approval_Date', value: '20/08/2024' });
});

test('Expenditure header aliases accurately resolve and route to Expenditure column', () => {
  const sheetKeys = ['Work_Name', 'Completion_Date', 'DLP', 'Expenditure'];

  // Test aliases resolve to Expenditure
  assert.equal(resolveFieldFromHeader('Expenditure', sheetKeys), 'Expenditure');
  assert.equal(resolveFieldFromHeader('Expenditure (Rs.)', sheetKeys), 'Expenditure');
  assert.equal(resolveFieldFromHeader('Expenditure Amount', sheetKeys), 'Expenditure');
  assert.equal(resolveFieldFromHeader('Expendeture', sheetKeys), 'Expenditure');

  // Test updater routes Expenditure value to correct column index (Col 4)
  const incoming = {
    'Expenditure (Rs.)': '850000',
    'Completion Date': '10/11/2026'
  };

  const updates = [];
  sheetKeys.forEach((header, index) => {
    let matchedVal = incoming[header];
    if (matchedVal === undefined) {
      for (const k of Object.keys(incoming)) {
        if (resolveFieldFromHeader(k, sheetKeys) === header) {
          matchedVal = incoming[k];
          break;
        }
      }
    }
    if (matchedVal !== undefined) {
      updates.push({ column: index + 1, header, value: matchedVal });
    }
  });

  assert.deepEqual(updates.find(u => u.column === 4), { column: 4, header: 'Expenditure', value: '850000' });
  assert.deepEqual(updates.find(u => u.column === 2), { column: 2, header: 'Completion_Date', value: '10/11/2026' });
});

test('Completion Certificate sync payload creates synchronized work details and sheet update payload', () => {
  const workItem = {
    E_TENDERID: 'T555',
    Work_Name: 'Construction of CC Road',
    Completion_Date: '10/11/2026',
    Expenditure: ''
  };

  const certificateFormData = {
    EXPENDITURE_RS: '450000',
    COMPLETION_DATE: '12/11/2026'
  };

  // Simulating syncCompletionDetailsFromPerforma logic
  const payload = { tenderId: workItem.E_TENDERID };
  if (certificateFormData.EXPENDITURE_RS && certificateFormData.EXPENDITURE_RS !== workItem.Expenditure) {
    payload.Expenditure = certificateFormData.EXPENDITURE_RS;
  }
  if (certificateFormData.COMPLETION_DATE && certificateFormData.COMPLETION_DATE !== workItem.Completion_Date) {
    payload.Completion_Date = certificateFormData.COMPLETION_DATE;
  }

  // Work details local update
  const updatedWork = { ...workItem, ...payload };

  assert.equal(payload.tenderId, 'T555');
  assert.equal(payload.Expenditure, '450000');
  assert.equal(payload.Completion_Date, '12/11/2026');
  assert.equal(updatedWork.Expenditure, '450000');
  assert.equal(updatedWork.Completion_Date, '12/11/2026');
});

// 11. Test Report Grand Total formatting and Development Works Report columns
const RUPEE = String.fromCharCode(8377);

function formatIndianNumber(value) {
  const amount = Number(value);
  if (!Number.isFinite(amount)) return value || "";
  const hasDecimals = Math.abs(amount % 1) > 0;
  return amount.toLocaleString('en-IN', {
    minimumFractionDigits: hasDecimals ? 2 : 0,
    maximumFractionDigits: hasDecimals ? 2 : 0
  });
}

function formatReportGrandTotal(value) {
  const amount = Number(value);
  if (!Number.isFinite(amount) || amount <= 0) return "-";
  const exact = formatIndianNumber(amount);
  const lacs = (amount / 100000).toLocaleString('en-IN', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
  return `${RUPEE} ${exact} Say ${RUPEE} ${lacs} Lacs`;
}

function normalizeReportKey(key) {
  return String(key || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function isReportAmountKey(key) {
  const normalized = normalizeReportKey(key);
  return normalized.includes('cost') || normalized.includes('amount') || normalized.includes('expenditure') || normalized.includes('expendeture');
}

const REPORT_SOURCE_COLUMNS = [
  "Work_Name",
  "Estimate_Cost",
  "Admin Approval_Date",
  "Tech_Approval_Date",
  "DNIT_Cost",
  "E_TENDERID",
  "T_PUB_DATE",
  "T_BIDSUB_START_DATE",
  "T_BIDSUB_END_DATE",
  "T_BID_OPEN_DATE",
  "Agency_Name",
  "Allotement_Date",
  "Agreement Start Date",
  "Agreement End Date",
  "Agreement Amount",
  "Revised_End Date(if any)",
  "Completion_Date",
  "DLP",
  "Expenditure"
];

function getReportColumnsThrough(endKey) {
  const endIndex = REPORT_SOURCE_COLUMNS.indexOf(endKey);
  return REPORT_SOURCE_COLUMNS.slice(0, endIndex === -1 ? REPORT_SOURCE_COLUMNS.length : endIndex + 1);
}

function getReportDefinitions(reportRows) {
  return [
    { name: "Tendered", rows: reportRows.tendered, columns: getReportColumnsThrough("T_BID_OPEN_DATE") },
    { name: "Allotted", rows: reportRows.allotted, columns: getReportColumnsThrough("Expenditure") },
    { name: "Completed", rows: reportRows.completed, columns: getReportColumnsThrough("Expenditure") },
    { name: "In Progress", rows: reportRows.inProgress, columns: getReportColumnsThrough("Expenditure") }
  ];
}

test('formatReportGrandTotal formats exact number and lacs according to user specification', () => {
  // Test case requested by user: 150687 -> ₹ 1,50,687 Say ₹ 1.51 Lacs
  assert.equal(formatReportGrandTotal(150687), `${RUPEE} 1,50,687 Say ${RUPEE} 1.51 Lacs`);
  // Additional cases
  assert.equal(formatReportGrandTotal(928421), `${RUPEE} 9,28,421 Say ${RUPEE} 9.28 Lacs`);
  assert.equal(formatReportGrandTotal(12500000), `${RUPEE} 1,25,00,000 Say ${RUPEE} 125.00 Lacs`);
  assert.equal(formatReportGrandTotal(0), "-");
  assert.equal(formatReportGrandTotal(-100), "-");
});

test('Development Works report includes Expenditure in Allotted, Completed, and In Progress definitions', () => {
  const definitions = getReportDefinitions({ tendered: [], allotted: [], completed: [], inProgress: [] });
  
  const tenderedDef = definitions.find(d => d.name === "Tendered");
  const allottedDef = definitions.find(d => d.name === "Allotted");
  const completedDef = definitions.find(d => d.name === "Completed");
  const inProgressDef = definitions.find(d => d.name === "In Progress");

  assert.equal(tenderedDef.columns.includes("Expenditure"), false);
  assert.equal(tenderedDef.columns.includes("Estimate_Cost"), true);
  assert.equal(tenderedDef.columns.includes("DNIT_Cost"), true);

  assert.equal(allottedDef.columns.includes("Expenditure"), true);
  assert.equal(completedDef.columns.includes("Expenditure"), true);
  assert.equal(inProgressDef.columns.includes("Expenditure"), true);

  // Check all amount columns in Allotted/Completed/In Progress are recognized by isReportAmountKey
  const amountCols = allottedDef.columns.filter(isReportAmountKey);
  assert.deepEqual(amountCols, ["Estimate_Cost", "DNIT_Cost", "Agreement Amount", "Expenditure"]);
});

test('Section totals sum and format all amount columns correctly', () => {
  const sampleRows = [
    { Estimate_Cost: "1,00,000", DNIT_Cost: "90,000", "Agreement Amount": "85,000", Expenditure: "80,000" },
    { Estimate_Cost: "50,687", DNIT_Cost: "45,000", "Agreement Amount": "42,000", Expenditure: "40,000" }
  ];

  const columns = ["Work_Name", "Estimate_Cost", "DNIT_Cost", "Agreement Amount", "Expenditure"];

  const totals = {};
  columns.forEach(col => {
    if (isReportAmountKey(col)) {
      const sum = sampleRows.reduce((acc, row) => acc + parseNumericAmount(row[col]), 0);
      totals[col] = formatReportGrandTotal(sum);
    }
  });

  // Estimate: 100000 + 50687 = 150687
  assert.equal(totals["Estimate_Cost"], `${RUPEE} 1,50,687 Say ${RUPEE} 1.51 Lacs`);
  // DNIT: 90000 + 45000 = 135000 -> 1.35 Lacs
  assert.equal(totals["DNIT_Cost"], `${RUPEE} 1,35,000 Say ${RUPEE} 1.35 Lacs`);
  // Agreement: 85000 + 42000 = 127000 -> 1.27 Lacs
  assert.equal(totals["Agreement Amount"], `${RUPEE} 1,27,000 Say ${RUPEE} 1.27 Lacs`);
  // Expenditure: 80000 + 40000 = 120000 -> 1.20 Lacs
  assert.equal(totals["Expenditure"], `${RUPEE} 1,20,000 Say ${RUPEE} 1.20 Lacs`);
});

