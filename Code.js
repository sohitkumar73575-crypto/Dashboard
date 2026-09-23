/**
 * Municipal Council Charkhi Dadri - Works Management System
 * Backend: Master data + BG_Register lookup/upload + PBG photo storage
 */

const MAIN_SS_ID = '15EEXLtWXxpH2SLk2qAxGViQX4vYW7lOBN4YvW1pRyLc';
const BG_SS_ID = '1zCVQWrU_rJaGjobv_Ht5HdwmlhKqJCRdIW6x9s-r7_Y';
const MASTER_SHEET_NAME = 'Sheet1';

const PBG_PHOTO_FOLDER_ID = '1J6b6mzkQtYVTs7XRyBPGaWZmDNugodUS';
// Paste your Drive file ID after uploading the blank Monitoring DOCX with placeholders (see MONITORING_TEMPLATE_README.md).
const ADMIN_KEY_HASH_PROPERTY = 'ADMIN_KEY_SHA256';
const MAX_UPLOAD_BYTES = 8 * 1024 * 1024;
const ALLOWED_UPLOAD_MIME_TYPES = ['application/pdf', 'image/jpeg', 'image/png', 'image/webp'];

function doGet(e) {
  return jsonOutput({ data: getMasterSheetData().data });
}

function doPost(e) {
  try {
    const params = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = params.method || params.action;

    if (action === 'calculateAndStore') {
      const auth = requireAdmin(params);
      if (auth) return auth;
      return withScriptLock(() => storeInBGCalculator(params.tenderId, params.calcData));
    }

    if (action === 'getMasterData') {
      return jsonOutput({ data: getMasterSheetData().data });
    }

    if (action === 'uploadToRegister') {
      return withScriptLock(() => saveToBGRegister(params.formData, params));
    }

    if (action === 'getBgRegisterDetails') {
      return getBgRegisterDetailsForWeb(params.tenderId);
    }

    if (action === 'getBgRegisterIndex') {
      return getBgRegisterIndexForWeb();
    }

    if (action === 'getBgReportIndex') {
      return getBgReportIndexForWeb();
    }

    if (action === 'releaseBg') {
      const auth = requireAdmin(params);
      if (auth) return auth;
      return withScriptLock(() => releaseBgFromWeb(params));
    }

    if (action === 'addWork') {
      const auth = requireAdmin(params);
      if (auth) return auth;
      return withScriptLock(() => {
        if (!Array.isArray(params.rowData) || !params.rowData.length) {
          return jsonOutput({ status: 'error', message: 'Work details missing' });
        }
        const ss = SpreadsheetApp.openById(MAIN_SS_ID);
        const sheet = getRequiredSheet(ss, MASTER_SHEET_NAME);
        sheet.appendRow(params.rowData);
        return jsonOutput({ status: 'success' });
      });
    }

    if (action === 'updateWork') {
      return withScriptLock(() => updateWorkDetails(params.details, params));
    }

    if (action === 'updateWorkByRowIndex') {
      return withScriptLock(() => updateWorkByRowIndexFromWeb(params.details, params));
    }

    return jsonOutput({ status: 'error', message: 'Unknown action: ' + action });
  } catch (err) {
    return jsonOutput({ status: 'error', message: err.message });
  }
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function testAuthorization() {
  SpreadsheetApp.openById(MAIN_SS_ID).getName();
  SpreadsheetApp.openById(BG_SS_ID).getName();
  DriveApp.getFolderById(PBG_PHOTO_FOLDER_ID).getName();
  return 'Authorization OK';
}

function getRequiredSheet(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Sheet not found: ' + sheetName + '. Please check tab name in spreadsheet.');
  }
  return sheet;
}

function getRequiredFolder(folderId, label) {
  try {
    return DriveApp.getFolderById(folderId);
  } catch (err) {
    throw new Error(label + ' folder permission is missing. Run testAuthorization() in the Apps Script editor, allow Drive access, then redeploy the web app. Original: ' + err.message);
  }
}

function requireAdmin(params) {
  const expectedHash = String(PropertiesService.getScriptProperties().getProperty(ADMIN_KEY_HASH_PROPERTY) || '').trim().toLowerCase();
  const key = String((params && params.adminKey) || '').trim();
  if (!expectedHash || !key) {
    return jsonOutput({ status: 'auth_error', message: 'Admin key required' });
  }

  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, key, Utilities.Charset.UTF_8);
  const actualHash = digest.map(byte => {
    const value = byte < 0 ? byte + 256 : byte;
    return ('0' + value.toString(16)).slice(-2);
  }).join('');

  return actualHash === expectedHash
    ? null
    : jsonOutput({ status: 'auth_error', message: 'Admin key required' });
}

function withScriptLock(work) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return work();
  } finally {
    lock.releaseLock();
  }
}

function assertTenderId(value) {
  const tenderId = String(value || '').trim();
  if (!tenderId) throw new Error('Tender ID missing');
  if (!/^[a-zA-Z0-9/_ .-]{1,80}$/.test(tenderId)) throw new Error('Tender ID contains invalid characters');
  return tenderId;
}

function assertDateText(value, label, required) {
  const text = String(value || '').trim();
  if (!text) {
    if (required) throw new Error(label + ' missing');
    return '';
  }
  if (!/^(\d{2}\/\d{2}\/\d{4}|\d{4}-\d{2}-\d{2})$/.test(text)) {
    throw new Error(label + ' must be a valid date');
  }
  return text;
}

function validateUploadFile(formData) {
  if (!formData.fileData) return;
  if (!String(formData.fileName || '').trim()) throw new Error('Uploaded file name missing');
  const mimeType = String(formData.fileMimeType || '').trim().toLowerCase();
  if (ALLOWED_UPLOAD_MIME_TYPES.indexOf(mimeType) === -1) {
    throw new Error('Only PDF, JPEG, PNG, or WebP files are allowed');
  }
  if (!/^[A-Za-z0-9+/=]+$/.test(String(formData.fileData))) {
    throw new Error('Uploaded file data is not valid base64');
  }
  const approxBytes = Math.floor(String(formData.fileData).length * 3 / 4);
  if (approxBytes > MAX_UPLOAD_BYTES) {
    throw new Error('Uploaded file is too large. Maximum allowed size is 8 MB.');
  }
}

function getMasterSheetData() {
  const ss = SpreadsheetApp.openById(MAIN_SS_ID);
  const sheet = getRequiredSheet(ss, MASTER_SHEET_NAME);
  const data = sheet.getDataRange().getDisplayValues();
  const headers = data[0].map(h => h.trim());
  const rows = data.slice(1);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const agreementEndCol = findHeaderIndex(headers, ['Agreement End Date', 'Agreement_End_Date', 'Agreement End_Date', 'AGREEMENT_END_DATE']);
  const revisedEndCol = findHeaderIndex(headers, ['Revised_End Date(if any)', 'Revised End Date', 'Revised_End_Date', 'Revised End_Date', 'REVISED_END_DATE', 'Revised Agreement End Date']);
  const completionCol = findHeaderIndex(headers, ['Completion_Date', 'Completion Date', 'Completed Date']);

  const formattedData = rows.map((row, index) => {
    let obj = { rowIndex: index + 2 };
    headers.forEach((header, i) => { obj[header] = row[i]; });

    const publishDate = parseDate(row[7]);
    const bidSubDate = parseDate(row[8]);
    const allotDate = parseDate(row[11]);
    const agreementEndDate = agreementEndCol === -1 ? parseDate(row[13]) : parseDate(row[agreementEndCol]);
    const revisedEndDate = revisedEndCol === -1 ? null : parseDate(row[revisedEndCol]);
    const hasActiveEndDate = Boolean(
      (agreementEndDate && agreementEndDate >= today) ||
      (revisedEndDate && revisedEndDate >= today)
    );
    const completionRaw = completionCol === -1 ? row[16] : row[completionCol];
    const completionFilled = String(completionRaw || '').trim() !== '' && String(completionRaw || '').trim() !== '-';
    const completionDate = parseDate(completionRaw);
    const dlpDays = parseInt(row[17], 10) || 0;
    const dlpEndDate = completionDate && dlpDays > 0 ? addDays(completionDate, dlpDays) : null;

    let status = 'Other';

    if (publishDate && publishDate <= today && bidSubDate && bidSubDate > today) {
      status = 'Live Tender';
    } else if (bidSubDate && bidSubDate < today && (!row[11] || row[11].trim() === '')) {
      status = 'Yet to be Allotted';
    } else if (completionDate && completionDate <= today && dlpEndDate && dlpEndDate >= today) {
      status = 'Under DLP';
    } else if (allotDate && allotDate <= today && !completionFilled && hasActiveEndDate) {
      status = 'In Progress';
    }

    obj.computedStatus = status;
    return obj;
  });

  return { data: formattedData };
}

function isUnlockedWorkValue(value) {
  const text = String(value === null || value === undefined ? '' : value).trim();
  return text === '' || text === '-';
}

function normalizeWorkEditValue(value, header) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }

  let text = String(value === null || value === undefined ? '' : value).trim();
  if (text && text.includes('-') && String(header || '').toLowerCase().includes('date')) {
    const parts = text.split('-');
    if (parts.length === 3 && parts[0].length === 4) text = `${parts[2]}/${parts[1]}/${parts[0]}`;
  }
  return text;
}

function resolveWorkTargetRow(data, headers, details) {
  details = details || {};
  const tenderIdText = String(details.tenderId || '').trim();

  if (tenderIdText) {
    const tenderId = assertTenderId(details.tenderId);
    const tenderColIndex = headers.indexOf('E_TENDERID') !== -1
      ? headers.indexOf('E_TENDERID')
      : headers.indexOf('E_TenderID');

    if (tenderColIndex === -1) return -1;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][tenderColIndex]).trim() === String(tenderId).trim()) {
        return i + 1;
      }
    }
    return -1;
  }

  const rowIndex = Number(details.rowIndex);
  if (Number.isFinite(rowIndex) && rowIndex >= 2 && rowIndex <= data.length) {
    return rowIndex;
  }

  return -1;
}

function applyWorkRowUpdates(sheet, headers, data, targetRow, details, params) {
  const targetValues = data[targetRow - 1];
  const updates = [];
  let lockedChangeRequested = false;

  headers.forEach((header, index) => {
    if (details[header] !== undefined && header !== 'rowIndex') {
      const value = normalizeWorkEditValue(details[header], header);
      const previousValue = targetValues[index];
      const previousText = normalizeWorkEditValue(previousValue, header);

      if (value !== previousText) {
        updates.push({ row: targetRow, column: index + 1, value: value });
        if (!isUnlockedWorkValue(previousValue)) lockedChangeRequested = true;
      }
    }
  });

  if (lockedChangeRequested) {
    const auth = requireAdmin(params);
    if (auth) return auth;
  }

  if (!updates.length) {
    return jsonOutput({ status: 'success', message: 'No changes needed' });
  }

  updates.forEach(update => {
    sheet.getRange(update.row, update.column).setValue(update.value);
  });

  return jsonOutput({ status: 'success', updated: updates.length, rowIndex: targetRow });
}

function updateWorkDetails(details, params) {
  details = details || {};
  const ss = SpreadsheetApp.openById(MAIN_SS_ID);
  const sheet = getRequiredSheet(ss, MASTER_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const tenderColIndex = headers.indexOf('E_TENDERID') !== -1
    ? headers.indexOf('E_TENDERID')
    : headers.indexOf('E_TenderID');

  if (tenderColIndex === -1) {
    return jsonOutput({ status: 'error', message: 'Tender ID column not found' });
  }

  const targetRow = resolveWorkTargetRow(data, headers, details);
  if (targetRow === -1) {
    const tenderIdText = String(details.tenderId || '').trim();
    return jsonOutput({
      status: 'error',
      message: tenderIdText ? 'Tender ID not found' : 'Tender ID missing'
    });
  }

  return applyWorkRowUpdates(sheet, headers, data, targetRow, details, params);
}

function updateWorkByRowIndexFromWeb(details, params) {
  details = details || {};
  const rowIndex = Number(details.rowIndex);
  if (!Number.isFinite(rowIndex) || rowIndex < 2) {
    return jsonOutput({ status: 'error', message: 'Invalid row index' });
  }

  const payload = Object.assign({}, details);
  delete payload.tenderId;
  payload.rowIndex = rowIndex;
  return updateWorkDetails(payload, params);
}

function storeInBGCalculator(tenderId, data) {
  tenderId = assertTenderId(tenderId);
  data = data || {};
  const ss = SpreadsheetApp.openById(BG_SS_ID);
  const sheet = getRequiredSheet(ss, 'BG_Calculator');
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][2]).trim() === String(tenderId).trim()) {
      sheet.getRange(i + 1, 6).setValue(data.pbg);
      sheet.getRange(i + 1, 7).setValue(data.addPbg);
      return jsonOutput({ status: 'success' });
    }
  }

  sheet.appendRow(['', '', tenderId, '', '', data.pbg, data.addPbg]);
  return jsonOutput({ status: 'success' });
}

function isUnlockedBgValue(value) {
  const text = String(value === null || value === undefined ? '' : value).trim();
  return text === '' || text === '-';
}

function normalizeBgEditValue(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  return String(value === null || value === undefined ? '' : value).trim();
}

function bgLockedChangeRequiresAdmin(existingDetails, formData) {
  if (!existingDetails) return false;

  const fields = [
    'pbgAmt',
    'pbgNo',
    'pbgSubDate',
    'pbgValid',
    'addAmt',
    'addNo',
    'addSubDate',
    'addValid',
    'bankName',
    'branchName'
  ];

  return fields.some(field => {
    const nextValue = normalizeBgEditValue(formData[field]);
    const previousValue = normalizeBgEditValue(existingDetails[field]);
    return nextValue !== previousValue && !isUnlockedBgValue(existingDetails[field]);
  });
}

function saveToBGRegister(formData, params) {
  formData = formData || {};
  formData.tenderId = assertTenderId(formData.tenderId);
  if (!String(formData.workName || '').trim()) throw new Error('Work name missing');
  if (!String(formData.agencyName || '').trim()) throw new Error('Agency name missing');
  if (!String(formData.pbgAmt || '').trim()) throw new Error('PBG amount missing');
  if (!String(formData.pbgNo || '').trim()) throw new Error('PBG number missing');
  assertDateText(formData.pbgSubDate, 'PBG submission date', true);
  assertDateText(formData.pbgValid, 'PBG valid date', true);
  assertDateText(formData.addSubDate, 'Additional PBG submission date', false);
  assertDateText(formData.addValid, 'Additional PBG valid date', false);
  validateUploadFile(formData);
  const ss = SpreadsheetApp.openById(BG_SS_ID);
  const sheet = getRequiredSheet(ss, 'BG_Register');
  ensureBgBranchHeader(sheet);
  const targetRow = findBgRegisterRowByTenderId(formData.tenderId);
  const existingDetails = targetRow ? getBgRegisterDetailsByRow(targetRow) : null;
  const additionalRequired = isAdditionalBgRequiredForTender(formData.tenderId, formData);
  let fileUrl = formData.fileName || '';

  if (existingDetails) {
    formData.workName = existingDetails.workName;
    formData.tenderId = existingDetails.tenderId;
    formData.agencyName = existingDetails.agencyName;
  }

  if (existingDetails && isReleasedValue(existingDetails.pbgReleaseStatus || existingDetails.releaseStatus)) {
    formData.pbgAmt = existingDetails.pbgAmt;
    formData.pbgNo = existingDetails.pbgNo;
    formData.pbgSubDate = existingDetails.pbgSubDate;
    formData.pbgValid = existingDetails.pbgValid;
  }

  if (existingDetails && additionalRequired && isReleasedValue(existingDetails.addReleaseStatus || existingDetails.releaseStatus)) {
    formData.addAmt = existingDetails.addAmt;
    formData.addNo = existingDetails.addNo;
    formData.addSubDate = existingDetails.addSubDate;
    formData.addValid = existingDetails.addValid;
  }

  if (existingDetails && !String(formData.branchName || '').trim()) {
    formData.branchName = existingDetails.branchName || '';
  }

  formData = normalizeBgDetailsForAdditionalRequirement(formData, formData.tenderId, formData);

  if (bgLockedChangeRequiresAdmin(existingDetails, formData)) {
    const auth = requireAdmin(params);
    if (auth) return auth;
  }

  if (formData.fileData && formData.fileName) {
    const folder = getRequiredFolder(PBG_PHOTO_FOLDER_ID, 'PBG photo');
    const bytes = Utilities.base64Decode(formData.fileData);
    const blob = Utilities.newBlob(
      bytes,
      formData.fileMimeType || 'image/jpeg',
      formData.fileName
    );

    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    fileUrl = file.getUrl();
  }

  const rowValues = [
    formData.workName || '',
    formData.tenderId || '',
    formData.agencyName || '',
    formData.pbgAmt || '',
    formData.pbgNo || '',
    formData.pbgSubDate || '',
    formData.pbgValid || '',
    formData.addAmt || '',
    formData.addNo || '',
    formData.addSubDate || '',
    formData.addValid || '',
    formData.bankName || '',
    fileUrl
  ];

  if (targetRow) {
    sheet.getRange(targetRow, 2, 1, 13).setValues([rowValues]);
    sheet.getRange(targetRow, 28).setValue(formData.branchName || '');
  } else {
    sheet.appendRow(['', ...rowValues]);
    sheet.getRange(sheet.getLastRow(), 28).setValue(formData.branchName || '');
  }

  return jsonOutput({ status: 'success', fileUrl: fileUrl });
}

function getBgRegisterDetailsForWeb(tenderId) {
  const row = findBgRegisterRowByTenderId(tenderId);
  if (!row) return jsonOutput({ status: 'not_found', details: null });

  const details = normalizeBgDetailsForAdditionalRequirement(getBgRegisterDetailsByRow(row), tenderId);
  if (!hasBgRegisterData(details, tenderId)) {
    return jsonOutput({ status: 'not_found', details: null });
  }

  return jsonOutput({ status: 'success', details: details });
}

function getBgRegisterIndexForWeb() {
  const ss = SpreadsheetApp.openById(BG_SS_ID);
  const sheet = getRequiredSheet(ss, 'BG_Register');
  const lastRow = sheet.getLastRow();
  const result = {};

  if (lastRow < 2) {
    return jsonOutput({ status: 'success', bgRegisterMap: result });
  }

  const additionalRequiredMap = getAdditionalBgRequiredMap();
  const rows = sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), 28)).getDisplayValues();

  rows.forEach(data => {
    const rawDetails = buildBgRegisterDetailsFromData(data);
    const tenderId = String(rawDetails.tenderId || '').trim();
    const details = normalizeBgDetailsForAdditionalRequirement(rawDetails, tenderId, additionalRequiredMap);
    if (tenderId && hasBgRegisterData(details, tenderId, additionalRequiredMap)) {
      result[tenderId] = details;
    }
  });

  return jsonOutput({ status: 'success', bgRegisterMap: result });
}

function getBgReportIndexForWeb() {
  const ss = SpreadsheetApp.openById(BG_SS_ID);
  const sheet = getRequiredSheet(ss, 'BG_Register');
  const lastRow = sheet.getLastRow();
  const result = {};

  if (lastRow < 2) {
    return jsonOutput({ status: 'success', bgRegisterMap: result });
  }

  const additionalRequiredMap = getAdditionalBgRequiredMap();
  const rows = sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), 28)).getDisplayValues();

  rows.forEach(data => {
    const rawDetails = buildBgRegisterDetailsFromData(data);
    const tenderId = String(rawDetails.tenderId || '').trim();
    const details = normalizeBgDetailsForAdditionalRequirement(rawDetails, tenderId, additionalRequiredMap);
    if (tenderId && hasPerformanceBgRegisterData(details)) {
      result[tenderId] = details;
    }
  });

  return jsonOutput({ status: 'success', bgRegisterMap: result });
}

function hasValue(value) {
  return String(value == null ? '' : value).trim() !== '';
}

function parseServerAmount(value) {
  return parseFloat(String(value || '').replace(/,/g, '').replace(/[^0-9.-]/g, '')) || 0;
}

function parseServerBooleanFlag(value) {
  if (value === true) return true;
  if (value === false) return false;

  const text = String(value == null ? '' : value).trim().toLowerCase();
  return text === 'true' || text === '1' || text === 'yes' || text === 'required';
}

function getRequestedAdditionalBgAmount(source) {
  if (!source) return null;

  const keys = ['expectedAdditionalAmount', 'additionalAmount', 'additionalPbg', 'addPbg', 'addAmt'];
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    if (!Object.prototype.hasOwnProperty.call(source, key)) continue;
    if (!hasValue(source[key])) return null;
    return parseServerAmount(source[key]);
  }

  return null;
}

function normalizeBgDetailsForAdditionalRequirement(details, tenderId, requirementSource) {
  if (!details) return details;

  const normalized = Object.assign({}, details);
  const additionalRequired = isAdditionalBgRequiredForTender(
    tenderId || normalized.tenderId,
    requirementSource || normalized
  );
  normalized.additionalRequired = additionalRequired;

  if (!additionalRequired) {
    normalized.addAmt = '';
    normalized.addNo = '';
    normalized.addSubDate = '';
    normalized.addValid = '';
    normalized.addReleaseStatus = '';
    normalized.addReleaseDate = '';
    normalized.addReleaseRemarks = '';
  }

  return normalized;
}

function hasBgRegisterData(details, tenderId, requirementSource) {
  if (!details || !hasValue(details.tenderId)) return false;

  const normalized = normalizeBgDetailsForAdditionalRequirement(details, tenderId, requirementSource);

  if (!hasPerformanceBgRegisterData(normalized)) return false;

  if (!normalized.additionalRequired) return true;

  return hasValue(normalized.addAmt) &&
    hasValue(normalized.addNo) &&
    hasValue(normalized.addSubDate) &&
    hasValue(normalized.addValid);
}

function hasPerformanceBgRegisterData(details) {
  return Boolean(details &&
    hasValue(details.tenderId) &&
    hasValue(details.pbgAmt) &&
    hasValue(details.pbgNo) &&
    hasValue(details.pbgSubDate) &&
    hasValue(details.pbgValid) &&
    hasValue(details.bankName));
}

function getBgRegisterDetailsByRow(row) {
  const ss = SpreadsheetApp.openById(BG_SS_ID);
  const sheet = getRequiredSheet(ss, 'BG_Register');
  const data = sheet.getRange(row, 1, 1, Math.max(sheet.getLastColumn(), 28)).getDisplayValues()[0];

  return buildBgRegisterDetailsFromData(data);
}

function buildBgRegisterDetailsFromData(data) {
  return {
    workName: data[1] || '',
    tenderId: data[2] || '',
    agencyName: data[3] || '',
    pbgAmt: data[4] || '',
    pbgNo: data[5] || '',
    pbgSubDate: data[6] || '',
    pbgValid: data[7] || '',
    addAmt: data[8] || '',
    addNo: data[9] || '',
    addSubDate: data[10] || '',
    addValid: data[11] || '',
    bankName: data[12] || '',
    branchName: data[27] || '',
    fileName: data[13] || '',
    fileUrl: data[13] || '',
    bgLetterUrl: data[14] || '',
    agreementUrl: data[17] || '',
    releaseStatus: data[18] || '',
    releaseDate: data[19] || '',
    releaseRemarks: data[20] || '',
    pbgReleaseStatus: data[21] || data[18] || '',
    pbgReleaseDate: data[22] || data[19] || '',
    pbgReleaseRemarks: data[23] || data[20] || '',
    addReleaseStatus: data[24] || data[18] || '',
    addReleaseDate: data[25] || data[19] || '',
    addReleaseRemarks: data[26] || data[20] || ''
  };
}

function ensureBgBranchHeader(sheet) {
  if (!String(sheet.getRange(1, 28).getDisplayValue() || '').trim()) {
    sheet.getRange(1, 28).setValue('Branch Name');
  }
}

function getBgBankBranchDisplay(bankName, branchName) {
  const bank = String(bankName || '').replace(/\s+/g, ' ').trim();
  const branch = String(branchName || '').replace(/\s+/g, ' ').trim();
  return [bank, branch].filter(Boolean).join(', ');
}

function ensureBgReleaseHeaders(sheet) {
  const headers = [
    'Release Status',
    'Release Date',
    'Release Remarks',
    'PBG Release Status',
    'PBG Release Date',
    'PBG Release Remarks',
    'Additional PBG Release Status',
    'Additional PBG Release Date',
    'Additional PBG Release Remarks'
  ];
  headers.forEach((header, index) => {
    const col = 19 + index;
    if (!String(sheet.getRange(1, col).getDisplayValue() || '').trim()) {
      sheet.getRange(1, col).setValue(header);
    }
  });
}

function isReleasedValue(value) {
  return String(value || '').trim().toLowerCase() === 'released';
}

function releaseBgFromWeb(body) {
  body = body || {};
  const tenderId = assertTenderId(body.tenderId);
  assertDateText(body.releaseDate, 'Release date', true);

  const row = findBgRegisterRowByTenderId(tenderId);
  if (!row) return jsonOutput({ status: 'error', message: 'BG details not found' });

  const ss = SpreadsheetApp.openById(BG_SS_ID);
  const sheet = getRequiredSheet(ss, 'BG_Register');
  ensureBgReleaseHeaders(sheet);
  ensureBgBranchHeader(sheet);

  const bgType = String(body.bgType || 'pbg').toLowerCase().indexOf('add') >= 0 ? 'additional' : 'pbg';
  if (bgType === 'additional' && !isAdditionalBgRequiredForTender(tenderId, body)) {
    return jsonOutput({ status: 'error', message: 'Additional PBG is not required for this tender.' });
  }

  const startCol = bgType === 'additional' ? 25 : 22;
  sheet.getRange(row, startCol, 1, 3).setValues([[
    'Released',
    body.releaseDate || '',
    body.releaseRemarks || ''
  ]]);

  const details = getBgRegisterDetailsByRow(row);
  return jsonOutput({ status: 'success', details: details });
}

function findBgRegisterRowByTenderId(tenderId) {
  if (!tenderId) return 0;

  const ss = SpreadsheetApp.openById(BG_SS_ID);
  const sheet = getRequiredSheet(ss, 'BG_Register');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const values = sheet.getRange(2, 3, lastRow - 1, 1).getDisplayValues();
  const wanted = String(tenderId).trim();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === wanted) return i + 2;
  }

  return 0;
}

function isAdditionalBgRequiredForTender(tenderId, additionalRequiredMap) {
  const key = String(tenderId || '').trim();

  const requestedAmount = getRequestedAdditionalBgAmount(additionalRequiredMap);
  if (requestedAmount !== null) return requestedAmount > 0;

  if (additionalRequiredMap && Object.prototype.hasOwnProperty.call(additionalRequiredMap, 'additionalRequired')) {
    return parseServerBooleanFlag(additionalRequiredMap.additionalRequired);
  }

  if (additionalRequiredMap && Object.prototype.hasOwnProperty.call(additionalRequiredMap, key)) {
    return parseServerBooleanFlag(additionalRequiredMap[key]);
  }

  const calc = calculateAdditionalBgAmountForTender(tenderId);
  return calc > 0;
}

function getAdditionalBgRequiredMap() {
  const ss = SpreadsheetApp.openById(MAIN_SS_ID);
  const sheet = getRequiredSheet(ss, MASTER_SHEET_NAME);
  const data = sheet.getDataRange().getDisplayValues();
  const result = {};
  if (data.length < 2) return result;

  const headers = data[0].map(h => String(h || '').trim());
  const tenderCol = findHeaderIndex(headers, ['E_TENDERID', 'E_TenderID', 'Tender ID', 'TenderID']);
  const dnitCol = findHeaderIndex(headers, ['DNIT Cost', 'DNIT COST', 'DNIT_Cost', 'Estimate Cost', 'Estimate_Cost', 'Estimate Amount', 'Estimated Amount']);
  const quotedCol = findHeaderIndex(headers, ['Agreement Amount', 'Agreement_Amount', 'Allotted Amount', 'Alloted Amount', 'Quoted Amount', 'Quoted_Amount']);

  if (tenderCol === -1 || dnitCol === -1 || quotedCol === -1) return result;

  for (let i = 1; i < data.length; i++) {
    const tenderId = String(data[i][tenderCol] || '').trim();
    if (!tenderId) continue;

    const dnitAmt = parseServerWorkAmount(data[i][dnitCol], headers[dnitCol]);
    const quotedAmt = parseServerWorkAmount(data[i][quotedCol], headers[quotedCol]);
    if (!dnitAmt || !quotedAmt) {
      result[tenderId] = false;
      continue;
    }

    const diffPercent = ((dnitAmt - quotedAmt) / dnitAmt) * 100;
    const extraPercent = diffPercent > 10 ? diffPercent - 10 : 0;
    result[tenderId] = Math.ceil((quotedAmt * (extraPercent / 100)) / 10) * 10 > 0;
  }

  return result;
}

function calculateAdditionalBgAmountForTender(tenderId) {
  const ss = SpreadsheetApp.openById(MAIN_SS_ID);
  const sheet = getRequiredSheet(ss, MASTER_SHEET_NAME);
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) return 0;

  const headers = data[0].map(h => String(h || '').trim());
  const tenderCol = findHeaderIndex(headers, ['E_TENDERID', 'E_TenderID', 'Tender ID', 'TenderID']);
  const dnitCol = findHeaderIndex(headers, ['DNIT Cost', 'DNIT COST', 'DNIT_Cost', 'Estimate Cost', 'Estimate_Cost', 'Estimate Amount', 'Estimated Amount']);
  const quotedCol = findHeaderIndex(headers, ['Agreement Amount', 'Agreement_Amount', 'Allotted Amount', 'Alloted Amount', 'Quoted Amount', 'Quoted_Amount']);

  if (tenderCol === -1 || dnitCol === -1 || quotedCol === -1) return 0;

  const wanted = String(tenderId || '').trim();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][tenderCol] || '').trim() !== wanted) continue;

    const dnitAmt = parseServerWorkAmount(data[i][dnitCol], headers[dnitCol]);
    const quotedAmt = parseServerWorkAmount(data[i][quotedCol], headers[quotedCol]);
    if (!dnitAmt || !quotedAmt) return 0;

    const diffPercent = ((dnitAmt - quotedAmt) / dnitAmt) * 100;
    const extraPercent = diffPercent > 10 ? diffPercent - 10 : 0;
    return Math.ceil((quotedAmt * (extraPercent / 100)) / 10) * 10;
  }

  return 0;
}

function findHeaderIndex(headers, candidates) {
  const normalizedCandidates = candidates.map(normalizeHeaderName);
  return headers.findIndex(header => normalizedCandidates.includes(normalizeHeaderName(header)));
}

function normalizeHeaderName(value) {
  return String(value || '').toLowerCase().replace(/[_\s]/g, '');
}

function parseServerWorkAmount(value, key) {
  const amount = parseServerAmount(value);
  if (!amount) return 0;

  const text = String(value || '').toLowerCase();
  if (/crore|\bcr\b/.test(text)) return amount * 10000000;
  if (/lakh|lac|lacs/.test(text)) return amount * 100000;

  const normalizedKey = normalizeHeaderName(key);
  const isLacsKey = normalizedKey.includes('dnit') ||
    normalizedKey.includes('estimate') ||
    normalizedKey.includes('agreementamount') ||
    normalizedKey.includes('allottedamount') ||
    normalizedKey.includes('allotedamount') ||
    normalizedKey.includes('quotedamount');

  return isLacsKey && Math.abs(amount) > 0 && Math.abs(amount) < 100
    ? amount * 100000
    : amount;
}

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
  match = text.match(/^(\d{1,2})[-/.](\\d{1,2})[-/.](\\d{2,4})$/);
  if (match) {
    const year = Number(match[3].length === 2 ? '20' + match[3] : match[3]);
    return new Date(year, Number(match[2]) - 1, Number(match[1]));
  }
  return null;
}

function addDays(date, days) {
  const result = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  result.setDate(result.getDate() + Number(days || 0));
  return result;
}

