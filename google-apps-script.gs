const DRIVE_FOLDER_ID = '1G5p90HMxgaJIaq4uolLjGNdW23clx-D2';
const SHEET_NAME = 'Vehicle Inspections';

function doGet() {
  return jsonResponse({ ok: true, message: 'F&C Vehicle Acceptance Form endpoint is live.' });
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || '{}');
    const fields = payload.fields || {};
    const photos = payload.photos || [];

    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const inspectionFolder = folder.createFolder(folderName(fields));
    const photoLinks = {};

    photos.forEach(function(photo) {
      if (!photo.dataUrl) return;
      const parts = photo.dataUrl.split(',');
      const mimeMatch = parts[0].match(/data:(.*);base64/);
      const mimeType = mimeMatch ? mimeMatch[1] : 'image/jpeg';
      const bytes = Utilities.base64Decode(parts[1]);
      const safeName = safeFileName(photo.label || photo.name || 'photo');
      const extension = extensionForMime(mimeType);
      const blob = Utilities.newBlob(bytes, mimeType, safeName + extension);
      const file = inspectionFolder.createFile(blob);
      photoLinks[photo.label || photo.name] = file.getUrl();
    });

    appendInspectionRow(fields, photoLinks, inspectionFolder.getUrl());
    return jsonResponse({ ok: true, folderUrl: inspectionFolder.getUrl(), photoLinks: photoLinks });
  } catch (error) {
    return jsonResponse({ ok: false, error: String(error) });
  }
}

function appendInspectionRow(fields, photoLinks, folderUrl) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);

  const headers = [
    'Timestamp',
    'Van License Plate',
    'Last 6 of VIN',
    'Driver Name',
    'Orientation Date',
    'Date',
    'Mileage',
    'Fuel Level',
    'Additional Equipment Present',
    'Equipment Present',
    'Equipment Notes',
    'Exterior Driver Side Photo',
    'Exterior Front Photo',
    'Exterior Passenger Side Photo',
    'Exterior Rear Photo',
    'Interior Front Photo',
    'Interior Rear Photo',
    'Photo Folder',
    'Driver Signature',
    'Driver Printed Name',
    'Signature Date',
    'Van Picked Up From',
    'Pickup Location',
    'Keys Handover Date',
    'Number of Keys',
    'Key Fob',
    'All Form Data JSON'
  ];

  ensureHeaders(sheet, headers);

  const row = headers.map(function(header) {
    if (header === 'Timestamp') return new Date();
    if (header === 'Photo Folder') return folderUrl || '';
    if (header === 'All Form Data JSON') return JSON.stringify(fields);
    if (photoLinks[header]) return photoLinks[header];
    const value = fields[header];
    return Array.isArray(value) ? value.join(', ') : (value || '');
  });

  sheet.appendRow(row);
}

function ensureHeaders(sheet, headers) {
  const firstRow = sheet.getRange(1, 1, 1, Math.max(headers.length, sheet.getLastColumn() || 1)).getValues()[0];
  const hasHeaders = firstRow.some(function(value) { return String(value || '').trim() !== ''; });
  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return;
  }

  const existing = firstRow.map(function(value) { return String(value || '').trim(); });
  headers.forEach(function(header) {
    if (existing.indexOf(header) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
    }
  });
}

function folderName(fields) {
  const plate = safeFileName(fields['Van License Plate'] || 'Unknown Van');
  const driver = safeFileName(fields['Driver Name'] || 'Unknown Driver');
  const date = safeFileName(fields['Date'] || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
  return date + ' - ' + plate + ' - ' + driver;
}

function safeFileName(value) {
  return String(value || 'file').replace(/[^A-Za-z0-9._ -]+/g, '-').trim().slice(0, 90) || 'file';
}

function extensionForMime(mimeType) {
  if (mimeType === 'image/png') return '.png';
  if (mimeType === 'image/webp') return '.webp';
  if (mimeType === 'image/gif') return '.gif';
  return '.jpg';
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
