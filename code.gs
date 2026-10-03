const SHEET_NAME = 'Du_Lieu';
const DRIVE_FOLDER_NAME = 'Ho_So_Dat_Dai';
const HEADER_ROW = [
  'createdAt',
  'hoTenCSD',
  'soCCCD',
  'soThua',
  'soTo',
  'diaChi',
  'tenHuyen',
  'tenXa',
  'ocrText',
  'fileName',
  'fileId',
  'fileUrl',
  'pageCount',
  'source'
];

function doGet() {
  return jsonResponse({
    success: true,
    message: 'Google Apps Script API đang hoạt động.',
    sheet: SHEET_NAME,
    folder: DRIVE_FOLDER_NAME
  });
}

function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    const data = JSON.parse(raw);
    const action = String(data.action || '').toUpperCase();
    const payload = data.payload || data;

    switch (action) {
      case 'PING':
        return jsonResponse({ success: true, message: 'API OK' });
      case 'SEARCH':
        return jsonResponse(searchRecords(payload || {}));
      case 'SAVE_SCAN':
        return jsonResponse(saveScanRecord(payload || {}));
      case 'LIST':
        return jsonResponse(listRecords(payload || {}));
      default:
        return jsonResponse({ success: false, error: 'Action không hợp lệ: ' + action });
    }
  } catch (error) {
    return jsonResponse({ success: false, error: error && error.toString ? error.toString() : String(error) });
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function ensureSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADER_ROW);
  } else if (sheet.getRange(1, 1, 1, HEADER_ROW.length).getValues()[0].join('|') !== HEADER_ROW.join('|')) {
    sheet.getRange(1, 1, 1, HEADER_ROW.length).setValues([HEADER_ROW]);
  }
  return sheet;
}

function ensureFolder() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(DRIVE_FOLDER_NAME);
}

function searchRecords(payload) {
  const query = (payload && payload.query ? String(payload.query).trim() : '').toLowerCase();
  const limit = Number(payload && payload.limit ? payload.limit : 50);
  const sheet = ensureSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return { success: true, count: 0, data: [] };
  }

  const values = sheet.getRange(2, 1, lastRow - 1, HEADER_ROW.length).getValues();
  const rows = values.map((row, index) => ({
    rowIndex: index + 2,
    createdAt: row[0] || '',
    hoTenCSD: row[1] || '',
    soCCCD: row[2] || '',
    soThua: row[3] || '',
    soTo: row[4] || '',
    diaChi: row[5] || '',
    tenHuyen: row[6] || '',
    tenXa: row[7] || '',
    ocrText: row[8] || '',
    fileName: row[9] || '',
    fileId: row[10] || '',
    fileUrl: row[11] || '',
    pageCount: row[12] || 0,
    source: row[13] || ''
  }));

  const filtered = !query ? rows.slice(-limit) : rows.filter((row) => {
    const haystack = [
      row.hoTenCSD,
      row.soCCCD,
      row.soThua,
      row.soTo,
      row.diaChi,
      row.tenHuyen,
      row.tenXa,
      row.ocrText,
      row.fileName
    ].join(' ').toLowerCase();
    return haystack.indexOf(query) >= 0;
  }).slice(0, limit);

  return { success: true, count: filtered.length, data: filtered };
}

function listRecords(payload) {
  return searchRecords({ query: '', limit: payload && payload.limit ? payload.limit : 50 });
}

function saveScanRecord(payload) {
  const sheet = ensureSheet();
  const folder = ensureFolder();

  const pdfBase64 = payload && payload.pdfBase64 ? String(payload.pdfBase64) : '';
  const fileName = (payload && payload.fileName ? String(payload.fileName) : 'Ho_So_Dat_Dai_' + Date.now() + '.pdf').trim();
  const hoTenCSD = (payload && payload.hoTenCSD ? String(payload.hoTenCSD) : '').trim();
  const soCCCD = (payload && payload.soCCCD ? String(payload.soCCCD) : '').trim();
  const soThua = (payload && payload.soThua ? String(payload.soThua) : '').trim();
  const soTo = (payload && payload.soTo ? String(payload.soTo) : '').trim();
  const diaChi = (payload && payload.diaChi ? String(payload.diaChi) : '').trim();
  const tenHuyen = (payload && payload.tenHuyen ? String(payload.tenHuyen) : '').trim();
  const tenXa = (payload && payload.tenXa ? String(payload.tenXa) : '').trim();
  const ocrText = (payload && payload.ocrText ? String(payload.ocrText) : '').trim();
  const pageCount = Number(payload && payload.pageCount ? payload.pageCount : 0);
  const createdAt = (payload && payload.createdAt) ? new Date(payload.createdAt).toISOString() : new Date().toISOString();

  let fileId = '';
  let fileUrl = '';

  if (pdfBase64) {
    const fileInfo = savePdfToDrive(folder, pdfBase64, fileName);
    fileId = fileInfo.fileId || '';
    fileUrl = fileInfo.fileUrl || '';
  }

  const row = [
    createdAt,
    hoTenCSD,
    soCCCD,
    soThua,
    soTo,
    diaChi,
    tenHuyen,
    tenXa,
    ocrText,
    fileName,
    fileId,
    fileUrl,
    String(pageCount),
    'mobile-app'
  ];

  sheet.appendRow(row);

  return {
    success: true,
    message: 'Dữ liệu đã được lưu vào Google Sheet.',
    fileId: fileId,
    fileUrl: fileUrl,
    rowCount: sheet.getLastRow()
  };
}

function savePdfToDrive(folder, pdfBase64, fileName) {
  try {
    const cleaned = String(pdfBase64 || '').replace(/^data:application\/pdf;base64,/, '').replace(/^data:.*;base64,/, '');
    const data = cleaned || '';
    if (!data) {
      throw new Error('PDF base64 rỗng.');
    }

    const blob = Utilities.newBlob(Utilities.base64Decode(data), 'application/pdf', fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { fileId: file.getId(), fileUrl: file.getUrl() };
  } catch (error) {
    throw new Error('Lỗi lưu PDF lên Drive: ' + error.toString());
  }
}

function resetSheetIfNeeded() {
  const sheet = ensureSheet();
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADER_ROW);
  }
}
