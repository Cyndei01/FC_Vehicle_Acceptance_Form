const DRIVE_FOLDER_ID = '1G5p90HMxgaJIaq4uolLjGNdW23clx-D2';
const SPREADSHEET_ID = '';
const SHEET_NAME = 'Vehicle Inspections';
const ALERT_EMAILS = ['fandcpackaginginc@gmail.com'];

function doGet(e) {
  if (e && e.parameter && e.parameter.health === '1') {
    return jsonResponse(healthCheck());
  }
  return jsonResponse({
    ok: true,
    message: 'F&C Vehicle Acceptance Form endpoint is live.',
    healthCheckUrl: 'Add ?health=1 to this URL to verify Drive and Sheet access.'
  });
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('No form payload was received.');
    }
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
    const emailResult = sendInspectionEmail(fields, photoLinks, inspectionFolder.getUrl());
    return jsonResponse({ ok: true, folderUrl: inspectionFolder.getUrl(), photoLinks: photoLinks, email: emailResult });
  } catch (error) {
    return jsonResponse({ ok: false, error: String(error) });
  }
}

function appendInspectionRow(fields, photoLinks, folderUrl) {
  const spreadsheet = getSpreadsheet();
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

function sendInspectionEmail(fields, photoLinks, folderUrl) {
  if (!ALERT_EMAILS || ALERT_EMAILS.length === 0) {
    return { sent: false, reason: 'No alert email recipients configured.' };
  }

  try {
    const subject = 'Van Inspection Submitted - ' + (fields['Van License Plate'] || 'Unknown Van');
    const lines = [
      'A new van inspection form was submitted.',
      '',
      'Driver: ' + (fields['Driver Name'] || 'Not provided'),
      'Van License Plate: ' + (fields['Van License Plate'] || 'Not provided'),
      'Last 6 of VIN: ' + (fields['Last 6 of VIN'] || 'Not provided'),
      'Inspection Date: ' + (fields['Date'] || 'Not provided'),
      'Mileage: ' + (fields['Mileage'] || 'Not provided'),
      'Pickup Location: ' + (fields['Pickup Location'] || 'Not provided'),
      'Additional Equipment Present: ' + (fields['Additional Equipment Present'] || 'Not provided'),
      'Equipment: ' + formatValue(fields['Equipment Present']),
      '',
      'Photo Folder:',
      folderUrl || 'Not available',
      '',
      'Photo Links:'
    ];

    Object.keys(photoLinks || {}).forEach(function(label) {
      lines.push(label + ': ' + photoLinks[label]);
    });

    MailApp.sendEmail({
      to: ALERT_EMAILS.join(','),
      subject: subject,
      body: lines.join('\n')
    });

    return { sent: true, recipients: ALERT_EMAILS };
  } catch (error) {
    return { sent: false, error: String(error) };
  }
}

function getSpreadsheet() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('No spreadsheet is connected. Open this script from the target Google Sheet, or set SPREADSHEET_ID to the Sheet ID.');
  }
  return spreadsheet;
}

function healthCheck() {
  const result = {
    ok: true,
    driveFolderId: DRIVE_FOLDER_ID,
    spreadsheetIdConfigured: Boolean(SPREADSHEET_ID),
    sheetName: SHEET_NAME
  };

  try {
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    result.driveFolderName = folder.getName();
  } catch (error) {
    result.ok = false;
    result.driveError = String(error);
  }

  try {
    const spreadsheet = getSpreadsheet();
    result.spreadsheetName = spreadsheet.getName();
    result.spreadsheetId = spreadsheet.getId();
    result.targetSheetExists = Boolean(spreadsheet.getSheetByName(SHEET_NAME));
  } catch (error) {
    result.ok = false;
    result.sheetError = String(error);
  }

  return result;
}

function authorizeSetup() {
  const result = healthCheck();
  Logger.log(JSON.stringify(result, null, 2));
  return result;
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

function formatValue(value) {
  if (Array.isArray(value)) return value.join(', ');
  return value || 'Not provided';
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
