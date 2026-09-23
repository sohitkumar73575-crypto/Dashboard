/**
 * Municipal Council Charkhi Dadri - Works Management System
 * Full Backend: Master data + BG_Register lookup/upload + PDF letter/agreement generation
 */

const MAIN_SS_ID = '15EEXLtWXxpH2SLk2qAxGViQX4vYW7lOBN4YvW1pRyLc';
const BG_SS_ID = '1zCVQWrU_rJaGjobv_Ht5HdwmlhKqJCRdIW6x9s-r7_Y';
const MASTER_SHEET_NAME = 'Sheet1';

const PBG_PHOTO_FOLDER_ID = '1J6b6mzkQtYVTs7XRyBPGaWZmDNugodUS';
const BG_LETTER_FOLDER_ID = '1wuIrGsDVWpKct0ZoeTkMvkfCqbCA-DlE';
const AGREEMENT_FOLDER_ID = '1i0SQ7CXfsYktG9c1riEMnjlQdaE9-U30';

const BG_LETTER_TEMPLATE_SINGLE_ID = '1SW-FEAwDn33JJ3ObZDL9G-s6Y2seQKNvdyhG6s-oqr8';
const BG_LETTER_TEMPLATE_DOUBLE_ID = '1lCg8JNnPEmTD0TV8EnGAyOWXoOPtTr13zNEa_xtxTjQ';
const AGREEMENT_TEMPLATE_ID = '1Sqw2xMrHX1otwsjvlO0kYGlwo7N9u9BjjGDFqYsHVZk';
// Paste your Drive file ID after uploading the blank Monitoring DOCX with placeholders (see MONITORING_TEMPLATE_README.md).
const MONITORING_TEMPLATE_ID = '1Hg-XJCXByjGwxzKkRSxN24OEMUgATMYr91HESAFTfDA';
const COMPLETION_CERTIFICATE_TEMPLATE_ID = '1yBJS18CANaE6YN5P5x6ILlVIlEmxeELC9pFNLqwKjOQ';
const SAMPLING_TEMPLATE_ID = '1D5DlT4Bc-nc_3ahpIo_H-gCWPW0DYj8oabColMn5Jxc';
const MONITORING_OUTPUT_FOLDER_ID = BG_LETTER_FOLDER_ID;
const COMPLETION_CERTIFICATE_OUTPUT_FOLDER_ID = BG_LETTER_FOLDER_ID;
const SAMPLING_OUTPUT_FOLDER_ID = BG_LETTER_FOLDER_ID;
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

    if (action === 'generateAgreementPdf' || action === 'generateAgreementDocx' || action === 'generateAgreement') {
      return withScriptLock(() => generateAgreementPdfFromWeb(params));
    }

    if (action === 'generateMonitoringPdf') {
      return withScriptLock(() => generateMonitoringPdfFromWeb(params));
    }

    if (action === 'generateCompletionCertificatePdf') {
      return withScriptLock(() => generateCompletionCertificatePdfFromWeb(params));
    }

    if (action === 'generatePerformaPdf') {
      return withScriptLock(() => generatePerformaPdfFromWeb(params));
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
  DriveApp.getFolderById(BG_LETTER_FOLDER_ID).getName();
  DriveApp.getFolderById(AGREEMENT_FOLDER_ID).getName();
  DriveApp.getFolderById(PBG_PHOTO_FOLDER_ID).getName();
  DriveApp.getFileById(COMPLETION_CERTIFICATE_TEMPLATE_ID).getName();
  DriveApp.getFileById(SAMPLING_TEMPLATE_ID).getName();
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

function getRequiredFile(fileId, label) {
  try {
    return DriveApp.getFileById(fileId);
  } catch (err) {
    throw new Error(label + ' file permission is missing. Please check template sharing/access. Original: ' + err.message);
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

function shouldReuseExistingPdf(options) {
  return Boolean(
    options &&
    Object.prototype.hasOwnProperty.call(options, 'forceRegenerate') &&
    !parseServerBooleanFlag(options.forceRegenerate)
  );
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

function generateBgLetterPdfFromWeb(body) {
  body = body || {};
  const tenderId = assertTenderId(body.tenderId);

  const row = findBgRegisterRowByTenderId(tenderId);
  if (!row) return jsonOutput({ status: 'error', message: 'Please upload PBG details first' });

  const details = normalizeBgDetailsForAdditionalRequirement(getBgRegisterDetailsByRow(row), tenderId, body);
  if (!hasPerformanceBgRegisterData(details)) {
    return jsonOutput({ status: 'error', message: 'Please upload PBG details first' });
  }

  // A BG letter is generated for the BGs actually entered in BG_Register.
  // One BG -> single format; two BGs -> double format.
  const result = generateLetterAuto(row, body);
  const pdfUrl = typeof result === 'object' ? (result.pdfUrl || '') : result;
  const docxUrl = typeof result === 'object' ? (result.docxUrl || '') : '';
  const docId = typeof result === 'object' ? (result.docId || '') : '';
  return jsonOutput({ status: 'success', docxUrl: docxUrl, pdfUrl: pdfUrl, docId: docId, tenderId: tenderId });
}

function generateAgreementPdfFromWeb(body) {
  body = body || {};
  const tenderId = assertTenderId(body.tenderId);

  const row = findBgRegisterRowByTenderId(tenderId);
  if (!row) return jsonOutput({ status: 'error', message: 'Please upload PBG details first' });

  const details = normalizeBgDetailsForAdditionalRequirement(getBgRegisterDetailsByRow(row), tenderId, body);
  if (!hasPerformanceBgRegisterData(details)) {
    return jsonOutput({ status: 'error', message: 'Please upload PBG details first' });
  }

  if (!hasBgRegisterData(details, tenderId, body)) {
    return jsonOutput({ status: 'error', message: 'Please complete Additional PBG details first' });
  }

  const result = generateAgreementAuto(row, body);
  const pdfUrl = typeof result === 'object' ? (result.pdfUrl || '') : result;
  const docxUrl = typeof result === 'object' ? (result.docxUrl || '') : '';
  const docId = typeof result === 'object' ? (result.docId || '') : '';
  return jsonOutput({ status: 'success', docxUrl: docxUrl, pdfUrl: pdfUrl, docId: docId, tenderId: tenderId });
}

function generateMonitoringPdfFromWeb(body) {
  body = body || {};
  const tenderId = assertTenderId(body.tenderId);

  if (!String(MONITORING_TEMPLATE_ID || '').trim()) {
    return jsonOutput({
      status: 'error',
      message: 'Monitoring template is not configured. Set MONITORING_TEMPLATE_ID in Apps Script (see MONITORING_TEMPLATE_README.md).'
    });
  }

  const grossAmtLacs = String(body.grossAmtLacs || '').trim();
  if (!grossAmtLacs) {
    return jsonOutput({ status: 'error', message: 'Gross Amount (in Lacs) is required.' });
  }

  const sampleReport = normalizeMonitoringSampleReport(body.sampleReport);
  const comments = String(body.comments || '').trim();
  if (!comments) {
    return jsonOutput({ status: 'error', message: 'Bill details are required for Comments.' });
  }

  const master = getMasterWorkFieldsByTenderId(tenderId);
  const payload = {
    tenderId: tenderId,
    workName: String(body.workName || master.workName || '').trim(),
    agencyName: String(body.agencyName || master.agencyName || '').trim(),
    estCostLacs: String(body.estCostLacs || master.estCostLacs || '').trim(),
    grossAmtLacs: grossAmtLacs,
    netAmtLacs: String(body.netAmtLacs || '').trim(),
    sampleReport: sampleReport,
    comments: comments
  };

  const pdfUrl = generateMonitoringPdfAuto(payload);
  if (!pdfUrl) {
    return jsonOutput({ status: 'error', message: 'Monitoring PDF could not be generated.' });
  }

  return jsonOutput({ status: 'success', pdfUrl: pdfUrl, tenderId: tenderId });
}

function normalizeMonitoringSampleReport(value) {
  const text = String(value || '').trim();
  if (/^attached$/i.test(text)) return 'Attached';
  if (/^not\s*attached$/i.test(text)) return 'Not Attached';
  return text || 'Not Attached';
}

function getMasterWorkFieldsByTenderId(tenderId) {
  const ss = SpreadsheetApp.openById(MAIN_SS_ID);
  const sheet = getRequiredSheet(ss, MASTER_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0] || [];

  const tenderCol = findHeaderIndex(headers, ['E_TENDERID', 'E_TenderID', 'TenderID', 'Tender ID']);
  const workCol = findHeaderIndex(headers, ['Work_Name', 'Work Name', 'Name of Work']);
  const agencyCol = findHeaderIndex(headers, ['Agency_Name', 'Agency Name', 'AGENCY_NAME']);
  const estimateCol = findHeaderIndex(headers, [
    'Estimate_Cost', 'Estimate Cost', 'Estimated Cost', 'DNIT_Cost', 'DNIT Cost', 'Estimate Amount'
  ]);

  if (tenderCol === -1) {
    return { workName: '', agencyName: '', estCostLacs: '' };
  }

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][tenderCol] || '').trim() !== String(tenderId).trim()) continue;

    const workName = workCol === -1 ? '' : String(data[i][workCol] || '').trim();
    const agencyName = agencyCol === -1 ? '' : String(data[i][agencyCol] || '').trim();
    const estRaw = estimateCol === -1 ? '' : data[i][estimateCol];
    const estCostLacs = formatAmountInLacs(estRaw, headers[estimateCol]);

    return { workName: workName, agencyName: agencyName, estCostLacs: estCostLacs };
  }

  return { workName: '', agencyName: '', estCostLacs: '' };
}

function formatAmountInLacs(value, key) {
  const amount = parseServerWorkAmount(value, key);
  if (!amount) return '';
  return (amount / 100000).toLocaleString('en-IN', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function generateMonitoringPdfAuto(payload) {
  payload = payload || {};
  const tenderId = String(payload.tenderId || '').trim();
  if (!tenderId) return '';

  const folder = getRequiredFolder(MONITORING_OUTPUT_FOLDER_ID, 'Monitoring output');
  const file = getRequiredFile(MONITORING_TEMPLATE_ID, 'Monitoring template')
    .makeCopy('Monitoring_' + tenderId, folder);

  const doc = DocumentApp.openById(file.getId());
  const body = doc.getBody();

  body.replaceText('{{WORK_NAME}}', safeText(payload.workName));
  body.replaceText('{{AGENCY_NAME}}', safeText(payload.agencyName));
  body.replaceText('{{EST_COST_LACS}}', safeText(payload.estCostLacs));
  body.replaceText('{{GROSS_AMT_LACS}}', safeText(payload.grossAmtLacs));
  body.replaceText('{{NET_AMT_LACS}}', safeText(payload.netAmtLacs));
  body.replaceText('{{SAMPLE_REPORT}}', safeText(payload.sampleReport));
  body.replaceText('{{COMMENTS}}', safeText(payload.comments));

  doc.saveAndClose();

  const pdfFile = folder.createFile(file.getAs(MimeType.PDF)).setName('Monitoring_' + tenderId + '.pdf');
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  try {
    file.setTrashed(true);
  } catch (e) {
    // Keep doc copy if trash fails.
  }

  return getDriveDownloadUrl(pdfFile);
}

function generateCompletionCertificatePdfFromWeb(body) {
  body = body || {};
  const tenderId = assertTenderId(body.tenderId);

  if (!String(COMPLETION_CERTIFICATE_TEMPLATE_ID || '').trim()) {
    return jsonOutput({
      status: 'error',
      message: 'Completion Certificate template is not configured.'
    });
  }

  const master = getMasterWorkFieldsByTenderId(tenderId);
  const expenditureRs = validateCompletionExpenditure(body.expenditureRs);
  const commencementDate = normalizePerformaDateText(body.commencementDate, 'Date of Commencement', true);
  const completionDate = normalizePerformaDateText(body.completionDate, 'Date of Completion', true);
  const certificateDate = normalizePerformaDateText(body.certificateDate, 'Certificate Date', true);

  const payload = {
    tenderId: tenderId,
    workName: String(body.workName || master.workName || '').trim(),
    estimateRs: normalizeCompletionAmountText(body.estimateRs),
    expenditureRs: expenditureRs,
    commencementDate: commencementDate,
    completionDate: completionDate,
    periodFrom: commencementDate,
    periodTo: completionDate,
    agencyName: String(body.agencyName || master.agencyName || '').trim(),
    certificateDate: certificateDate
  };

  const pdfUrl = generateCompletionCertificatePdfAuto(payload);
  if (!pdfUrl) {
    return jsonOutput({ status: 'error', message: 'Completion Certificate PDF could not be generated.' });
  }

  return jsonOutput({ status: 'success', pdfUrl: pdfUrl, tenderId: tenderId });
}

function normalizeCompletionAmountText(value) {
  const amount = parseServerAmount(value);
  if (!amount) return '';
  return amount.toLocaleString('en-IN', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function validateCompletionExpenditure(value) {
  const text = String(value || '').trim();
  if (!text) throw new Error('Expenditure (Rs.) is required.');
  if (!/^[0-9,]+(\.[0-9]{1,2})?$/.test(text)) {
    throw new Error('Expenditure (Rs.) must contain amount only.');
  }

  const amount = parseServerAmount(text);
  if (!amount || amount <= 0) throw new Error('Expenditure (Rs.) must be greater than zero.');
  return amount.toLocaleString('en-IN', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function normalizePerformaDateText(value, label, required) {
  const text = assertDateText(value, label, required);
  if (!text) return '';

  let date = parseDate(text);
  if (!date) throw new Error(label + ' must be a valid date');

  return Utilities.formatDate(date, 'Asia/Kolkata', 'dd.MM.yyyy');
}

function generateCompletionCertificatePdfAuto(payload) {
  payload = payload || {};
  const tenderId = String(payload.tenderId || '').trim();
  if (!tenderId) return '';

  const folder = getRequiredFolder(COMPLETION_CERTIFICATE_OUTPUT_FOLDER_ID, 'Completion Certificate output');
  const template = getRequiredFile(COMPLETION_CERTIFICATE_TEMPLATE_ID, 'Completion Certificate template');
  if (template.getMimeType() !== MimeType.GOOGLE_DOCS) {
    throw new Error('Completion Certificate template must be saved as Google Docs, not DOCX.');
  }

  const file = template.makeCopy('Completion_Certificate_' + tenderId, folder);
  const doc = DocumentApp.openById(file.getId());
  const body = doc.getBody();

  body.replaceText('{{WORK_NAME}}', safeText(payload.workName));
  body.replaceText('{{ESTIMATE_RS}}', safeText(payload.estimateRs));
  body.replaceText('{{EXPENDITURE_RS}}', safeText(payload.expenditureRs));
  body.replaceText('{{COMMENCEMENT_DATE}}', safeText(payload.commencementDate));
  body.replaceText('{{COMPLETION_DATE}}', safeText(payload.completionDate));
  body.replaceText('{{PERIOD_FROM}}', safeText(payload.periodFrom));
  body.replaceText('{{PERIOD_TO}}', safeText(payload.periodTo));
  body.replaceText('{{AGENCY_NAME}}', safeText(payload.agencyName));
  body.replaceText('{{CERTIFICATE_DATE}}', safeText(payload.certificateDate));

  doc.saveAndClose();

  const pdfFile = folder.createFile(file.getAs(MimeType.PDF)).setName('Completion_Certificate_' + tenderId + '.pdf');
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  try {
    file.setTrashed(true);
  } catch (e) {
    // Keep doc copy if trash fails.
  }

  return getDriveDownloadUrl(pdfFile);
}

function generatePerformaPdfFromWeb(body) {
  body = body || {};
  const type = String(body.performaType || '').trim();

  if (type !== 'sampling') {
    return jsonOutput({ status: 'error', message: 'Unsupported performa type: ' + type });
  }

  return generateSamplingPdfFromWeb(body);
}

function generateSamplingPdfFromWeb(body) {
  body = body || {};
  const tenderId = assertTenderId(body.tenderId);

  if (!String(SAMPLING_TEMPLATE_ID || '').trim()) {
    return jsonOutput({
      status: 'error',
      message: 'Sampling template is not configured.'
    });
  }

  const data = body.templateData || {};
  const master = getMasterWorkFieldsByTenderId(tenderId);
  const workName = String(data.NAME_OF_WORK || data.WORK_NAME || body.workName || master.workName || '').trim();
  const sampleType = normalizeSamplingSampleType(data.SAMPLE_TYPE || body.sampleType);

  if (!workName) {
    return jsonOutput({ status: 'error', message: 'Name of Work is required for Sampling Performa.' });
  }
  if (!sampleType) {
    return jsonOutput({ status: 'error', message: 'Type of Sample is required for Sampling Performa.' });
  }

  const pdfUrl = generateSamplingPdfAuto({
    tenderId: tenderId,
    workName: workName,
    sampleType: sampleType
  });

  if (!pdfUrl) {
    return jsonOutput({ status: 'error', message: 'Sampling PDF could not be generated.' });
  }

  return jsonOutput({ status: 'success', pdfUrl: pdfUrl, tenderId: tenderId });
}

function normalizeSamplingSampleType(value) {
  const text = String(value || '').trim();
  if (!text) return '';

  const lines = text.split(/\r?\n/).map(line => line.trim()).filter(Boolean);
  if (lines.length > 1) return lines.join('\n');
  if (/^\d+\.\s+/.test(text)) return text;

  return text.split(',').map(part => part.trim()).filter(Boolean).map((part, index) => {
    return (index + 1) + '. ' + part.replace(/\bmm\b/gi, 'MM');
  }).join('\n');
}

function generateSamplingPdfAuto(payload) {
  payload = payload || {};
  const tenderId = String(payload.tenderId || '').trim();
  if (!tenderId) return '';

  const folder = getRequiredFolder(SAMPLING_OUTPUT_FOLDER_ID, 'Sampling output');
  const template = getRequiredFile(SAMPLING_TEMPLATE_ID, 'Sampling template');
  if (template.getMimeType() !== MimeType.GOOGLE_DOCS) {
    throw new Error('Sampling template must be saved as Google Docs, not DOCX.');
  }

  const file = template.makeCopy('Sampling_' + tenderId, folder);
  const doc = DocumentApp.openById(file.getId());
  const body = doc.getBody();

  body.replaceText('{{NAME_OF_WORK}}', safeText(payload.workName));
  body.replaceText('{{WORK_NAME}}', safeText(payload.workName));
  body.replaceText('{{SAMPLE_TYPE}}', safeText(payload.sampleType));

  doc.saveAndClose();

  const pdfFile = folder.createFile(file.getAs(MimeType.PDF)).setName('Sampling_' + tenderId + '.pdf');
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  try {
    file.setTrashed(true);
  } catch (e) {
    // Keep doc copy if trash fails.
  }

  return getDriveDownloadUrl(pdfFile);
}

function getDriveDownloadUrl(fileOrUrl) {
  if (!fileOrUrl) return '';

  if (typeof fileOrUrl.getId === 'function') {
    return 'https://drive.google.com/uc?export=download&id=' + encodeURIComponent(fileOrUrl.getId());
  }

  const text = String(fileOrUrl).trim();
  const idMatch = text.match(/\/d\/([^/]+)/) || text.match(/[?&]id=([^&]+)/);
  return idMatch
    ? 'https://drive.google.com/uc?export=download&id=' + encodeURIComponent(idMatch[1])
    : text;
}

function generateLetterAuto(row, options) {
  options = options || {};
  const ss = SpreadsheetApp.openById(BG_SS_ID);
  const sheet = getRequiredSheet(ss, 'BG_Register');
  const data = sheet.getRange(row, 1, 1, Math.max(sheet.getLastColumn(), 28)).getValues()[0];
  const existingPdfUrl = sheet.getRange(row, 15).getDisplayValue();
  if (existingPdfUrl && shouldReuseExistingPdf(options)) return getDriveDownloadUrl(existingPdfUrl);

  const tenderId = data[2];
  if (!tenderId) return '';

  const workName = sheet.getRange(row, 2).getDisplayValue();
  const agencyName = sheet.getRange(row, 4).getDisplayValue();

  const bg1Amount = formatIndianCurrency(data[4]);
  const bg2Amount = formatIndianCurrency(data[8]);

  const bg1No = data[5];
  const bg1Start = data[6];
  const bg1End = data[7];

  const bg2No = data[9];
  const bg2Start = data[10];
  const bg2End = data[11];

  const bankName = getBgBankBranchDisplay(data[12], data[27]);
  const bg1Valid = formatDateForLetter(bg1Start) + ' to ' + formatDateForLetter(bg1End);
  const bg2Valid = formatDateForLetter(bg2Start) + ' to ' + formatDateForLetter(bg2End);
  const today = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd MMM yyyy');
  const hasBG2 = Boolean(
    String(bg2No || '').trim() &&
    String(data[8] || '').trim()
  );

  const templateId = hasBG2 ? BG_LETTER_TEMPLATE_DOUBLE_ID : BG_LETTER_TEMPLATE_SINGLE_ID;
  const folder = getRequiredFolder(BG_LETTER_FOLDER_ID, 'BG letter output');
  const file = getRequiredFile(templateId, 'BG letter template').makeCopy('BG_Letter_' + tenderId, folder);

  const doc = DocumentApp.openById(file.getId());
  const body = doc.getBody();

  body.replaceText('{{WORK_NAME}}', safeText(workName));
  body.replaceText('{{TENDER_ID}}', safeText(tenderId));
  body.replaceText('{{AGENCY_NAME}}', safeText(agencyName));
  body.replaceText('{{DATED}}', safeText(today));
  body.replaceText('{{BANK_NAME}}', safeText(bankName));
  body.replaceText('{{BG1_NO}}', safeText(bg1No));
  body.replaceText('{{BG1_AMOUNT}}', safeText(bg1Amount));
  body.replaceText('{{BG1_VALID}}', safeText(bg1Valid));

  if (hasBG2) {
    body.replaceText('{{BG2_NO}}', safeText(bg2No));
    body.replaceText('{{BG2_AMOUNT}}', safeText(bg2Amount));
    body.replaceText('{{BG2_VALID}}', safeText(bg2Valid));
  }

  doc.saveAndClose();

  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  let docxDownloadUrl = 'https://docs.google.com/document/d/' + file.getId() + '/export?format=docx';
  try {
    const exportUrl = 'https://docs.google.com/feeds/download/documents/export/Export?id=' + file.getId() + '&exportFormat=docx';
    const docxResp = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (docxResp.getResponseCode() === 200) {
      const docxFile = folder.createFile(docxResp.getBlob().setName('BG_Letter_' + tenderId + '.docx'));
      docxFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      docxDownloadUrl = getDriveDownloadUrl(docxFile);
    }
  } catch (err) {
    console.warn('DOCX file creation in Drive folder failed, using direct export link: ' + err);
  }

  const pdfFile = folder.createFile(file.getAs(MimeType.PDF)).setName('BG_Letter_' + tenderId + '.pdf');
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  sheet.getRange(row, 15).setValue(pdfFile.getUrl());

  return {
    docxUrl: docxDownloadUrl,
    docId: file.getId(),
    pdfUrl: getDriveDownloadUrl(pdfFile)
  };
}

function generateAgreementAuto(row, options) {
  options = options || {};
  const ss = SpreadsheetApp.openById(BG_SS_ID);
  const sheet = getRequiredSheet(ss, 'BG_Register');
  const data = sheet.getRange(row, 1, 1, Math.max(sheet.getLastColumn(), 28)).getDisplayValues()[0];
  const existingPdfUrl = sheet.getRange(row, 18).getDisplayValue();
  if (existingPdfUrl && shouldReuseExistingPdf(options)) return getDriveDownloadUrl(existingPdfUrl);

  const tenderId = data[2];
  if (!tenderId) return '';

  const workName = sheet.getRange(row, 2).getDisplayValue();
  const agency = sheet.getRange(row, 4).getDisplayValue();
  const raw = getAmountFromMainSheet(tenderId);
  const clean = String(raw || '').replace(/[^0-9.]/g, '');
  const rawAmount = Math.round(parseFloat(clean) || 0);
  const amount = formatIndianCurrency(rawAmount);
  const today = getFormattedDate();

  const folder = getRequiredFolder(AGREEMENT_FOLDER_ID, 'Agreement output');
  const file = getRequiredFile(AGREEMENT_TEMPLATE_ID, 'Agreement template').makeCopy('Agreement_' + tenderId, folder);

  const doc = DocumentApp.openById(file.getId());
  const body = doc.getBody();

  body.replaceText('{{FULL_DATE}}', safeText(today));
  body.replaceText('{{CONTRACTOR}}', safeText(agency));
  body.replaceText('{{WORK}}', safeText(workName));
  body.replaceText('{{TENDER_ID}}', safeText(tenderId));
  body.replaceText('{{AMOUNT}}', safeText(amount));
  body.replaceText('{{AMOUNT_WORDS}}', safeText(numberToWords(rawAmount)));

  doc.saveAndClose();

  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  let docxDownloadUrl = 'https://docs.google.com/document/d/' + file.getId() + '/export?format=docx';
  try {
    const exportUrl = 'https://docs.google.com/feeds/download/documents/export/Export?id=' + file.getId() + '&exportFormat=docx';
    const docxResp = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (docxResp.getResponseCode() === 200) {
      const docxFile = folder.createFile(docxResp.getBlob().setName('Agreement_' + tenderId + '.docx'));
      docxFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      docxDownloadUrl = getDriveDownloadUrl(docxFile);
    }
  } catch (err) {
    console.warn('Agreement DOCX file creation failed, using direct export link: ' + err);
  }

  const pdfFile = folder.createFile(file.getAs(MimeType.PDF)).setName('Agreement_' + tenderId + '.pdf');
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  sheet.getRange(row, 18).setValue(pdfFile.getUrl());

  return {
    docxUrl: docxDownloadUrl,
    docId: file.getId(),
    pdfUrl: getDriveDownloadUrl(pdfFile)
  };
}

function getAmountFromCalculator(tenderId) {
  const ss = SpreadsheetApp.openById(BG_SS_ID);
  const sheet = getRequiredSheet(ss, 'BG_Calculator');
  const data = sheet.getDataRange().getDisplayValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).trim() === String(tenderId).trim()) {
      return data[i][4];
    }
  }

  return '';
}

function formatDateForLetter(d) {
  if (!d) return '';

  if (Object.prototype.toString.call(d) === '[object Date]' && !isNaN(d.getTime())) {
    return Utilities.formatDate(d, 'Asia/Kolkata', 'dd MMM yyyy');
  }

  const text = String(d).trim();
  const parts = text.split(/[-/.]/);
  if (parts.length === 3) {
    const day = Number(parts[0]);
    const month = Number(parts[1]);
    const year = Number(parts[2]);
    const parsed = new Date(year, month - 1, day);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, 'Asia/Kolkata', 'dd MMM yyyy');
    }
  }

  const parsed = new Date(text);
  return isNaN(parsed.getTime()) ? text : Utilities.formatDate(parsed, 'Asia/Kolkata', 'dd MMM yyyy');
}

function getFormattedDate() {
  const d = new Date();
  const day = d.getDate();
  const suffix = (day === 1 || day === 21 || day === 31) ? 'st' :
                 (day === 2 || day === 22) ? 'nd' :
                 (day === 3 || day === 23) ? 'rd' : 'th';
  const month = Utilities.formatDate(d, 'Asia/Kolkata', 'MMMM');
  const year = d.getFullYear();

  return day + suffix + ' day of ' + month + ' ' + year;
}

function numberToWords(num) {
  num = parseInt(String(num).replace(/,/g, ''), 10);
  if (isNaN(num)) return '';
  if (num === 0) return 'Zero Only';

  const a = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten',
    'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
  const b = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];

  function inWords(n) {
    n = parseInt(n, 10) || 0;
    if (n === 0) return '';
    if (n < 20) return a[n];
    if (n < 100) return (b[Math.floor(n / 10)] + ' ' + a[n % 10]).trim();
    if (n < 1000) return (a[Math.floor(n / 100)] + ' Hundred ' + inWords(n % 100)).trim();
    if (n < 100000) return (inWords(Math.floor(n / 1000)) + ' Thousand ' + inWords(n % 1000)).trim();
    if (n < 10000000) return (inWords(Math.floor(n / 100000)) + ' Lakh ' + inWords(n % 100000)).trim();
    return (inWords(Math.floor(n / 10000000)) + ' Crore ' + inWords(n % 10000000)).trim();
  }

  return inWords(num).replace(/\s+/g, ' ').trim() + ' Only';
}

function formatIndianCurrency(num) {
  num = String(num || '').replace(/,/g, '').replace(/[^0-9.]/g, '').trim();
  if (!num || isNaN(num)) return num;

  const x = num.split('.');
  let lastThree = x[0].substring(x[0].length - 3);
  const otherNumbers = x[0].substring(0, x[0].length - 3);

  if (otherNumbers !== '') lastThree = ',' + lastThree;

  let result = otherNumbers.replace(/\B(?=(\d{2})+(?!\d))/g, ',') + lastThree;
  if (x.length > 1) result += '.' + x[1];

  return String.fromCharCode(8377) + ' ' + result;
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

function safeText(value) {
  return String(value == null ? '' : value).replace(/\$/g, '$$$$');
}

function getAmountFromMainSheet(tenderId) {
  const ss = SpreadsheetApp.openById(MAIN_SS_ID);
  const sheet = getRequiredSheet(ss, MASTER_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const tenderCol = headers.findIndex(h => 
    String(h).trim() === 'E_TENDERID' || String(h).trim() === 'E_TenderID'
  );
  const amountCol = headers.findIndex(h => 
    String(h).trim() === 'Agreement Amount'
  );
  
  if (tenderCol === -1 || amountCol === -1) return '';
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][tenderCol]).trim() === String(tenderId).trim()) {
      return data[i][amountCol];
    }
  }
  return '';
}
