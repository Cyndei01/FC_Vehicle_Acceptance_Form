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
    const body = buildPlainTextSubmission(fields, photoLinks, folderUrl);
    const htmlBody = buildHtmlSubmission(fields, photoLinks, folderUrl);

    MailApp.sendEmail({
      to: ALERT_EMAILS.join(','),
      subject: subject,
      body: body,
      htmlBody: htmlBody
    });

    return { sent: true, recipients: ALERT_EMAILS };
  } catch (error) {
    return { sent: false, error: String(error) };
  }
}

function buildPlainTextSubmission(fields, photoLinks, folderUrl) {
  const lines = [
    'F&C Packaging Inc. - Cargo Van Inspection & Acceptance Form',
    '',
    'Photo Folder: ' + (folderUrl || 'Not available'),
    ''
  ];

  getSubmissionSections().forEach(function(section) {
    lines.push(section.title);
    section.fields.forEach(function(name) {
      lines.push(name + ': ' + formatValue(fields[name]));
    });
    lines.push('');
  });

  lines.push('Photo Links');
  Object.keys(photoLinks || {}).forEach(function(label) {
    lines.push(label + ': ' + photoLinks[label]);
  });

  return lines.join('\n');
}

function buildHtmlSubmission(fields, photoLinks, folderUrl) {
  const sectionHtml = getSubmissionSections().map(function(section) {
    const rows = section.fields.map(function(name) {
      return '<tr><th>' + escapeHtml(name) + '</th><td>' + escapeHtml(formatValue(fields[name])) + '</td></tr>';
    }).join('');
    return '<h2>' + escapeHtml(section.title) + '</h2><table>' + rows + '</table>';
  }).join('');

  const photoRows = Object.keys(photoLinks || {}).map(function(label) {
    const url = photoLinks[label];
    return '<tr><th>' + escapeHtml(label) + '</th><td><a href="' + escapeAttr(url) + '">' + escapeHtml(url) + '</a></td></tr>';
  }).join('');

  return '<!doctype html><html><body>' +
    '<div style="font-family:Arial,sans-serif;color:#111;max-width:860px;margin:0 auto;">' +
    '<div style="background:#1a3a6b;color:#fff;padding:20px;border-bottom:5px solid #f4c430;">' +
    '<div style="color:#f4c430;font-weight:bold;letter-spacing:.08em;">F&C PACKAGING</div>' +
    '<h1 style="margin:8px 0 0;font-size:24px;">Cargo Van Inspection & Acceptance Form</h1>' +
    '</div>' +
    '<p><strong>Photo Folder:</strong> <a href="' + escapeAttr(folderUrl || '') + '">' + escapeHtml(folderUrl || 'Not available') + '</a></p>' +
    sectionHtml +
    '<h2>Uploaded Photos</h2><table>' + (photoRows || '<tr><td>No photo links provided.</td></tr>') + '</table>' +
    '<p style="font-size:12px;color:#555;">This email is a submitted copy of the inspection form. Google Drive and Google Sheets remain the system of record.</p>' +
    '</div>' +
    '<style>h2{font-size:16px;margin:22px 0 8px;color:#1a3a6b;}table{border-collapse:collapse;width:100%;margin-bottom:10px;}th,td{border:1px solid #d5dbe5;padding:8px;text-align:left;vertical-align:top;}th{width:34%;background:#f4f6f8;color:#111;}a{color:#1a3a6b;}</style>' +
    '</body></html>';
}

function getSubmissionSections() {
  return [
    { title: 'Vehicle / Driver Information', fields: ['Van License Plate','Last 6 of VIN','Driver Name','Orientation Date','Date','Mileage','Fuel Level'] },
    { title: 'Exterior Inspection', fields: ['Body and Panels','Body and Panels Notes','Glass and Mirrors','Glass and Mirrors Notes','Tires','Tires Notes','Lights','Lights Notes'] },
    { title: 'Cargo Area Inspection', fields: ['Cargo Doors','Cargo Doors Notes','Cargo Floor','Cargo Floor Notes','Bulkhead','Bulkhead Notes','Tie-Downs','Tie-Downs Notes'] },
    { title: 'Interior / Safety Inspection', fields: ['Seats','Seats Notes','Seat Belts','Seat Belts Notes','Fire Extinguisher','Fire Extinguisher Notes'] },
    { title: 'Equipment Verification', fields: ['Additional Equipment Present','Equipment Present','Equipment Notes'] },
    { title: 'Cleanliness Standards', fields: ['Cabin Floor','Cabin Floor Notes','Surfaces and Dash','Surfaces and Dash Notes','Air Quality','Air Quality Notes','Cargo Cleanliness','Cargo Cleanliness Notes'] },
    { title: 'Driver Sign-Off', fields: ['Driver Signature','Driver Printed Name','Signature Date','Van Picked Up From','Pickup Location','Keys Handover Date','Number of Keys','Key Fob'] }
  ];
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

function escapeHtml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function escapeAttr(value) {
  return escapeHtml(value);
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
