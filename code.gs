const SHEET_NAME = 'Du_Lieu';
const DRIVE_FOLDER_NAME = 'Ho_So_Dat_Dai';

const HEADER_ROW = [
  'STT',
  'Tên huyện',
  'Tên xã cũ',
  'Tên xã mới',
  'Họ Tên CSD',
  'Số hiệu tờ bản đồ mới',
  'Số thứ tự thửa mới',
  'Diện tích',
  'Loại mục đích sử dụng',
  'Số hiệu tờ bản đồ cũ',
  'Số thứ tự thửa cũ',
  'Ấp',
  'Thế chấp',
  'File GCN',
  'File CCCD',
  'Số CCCD',
  'fileName',
  'fileId',
  'fileUrl',
  'pageCount',
  'createdAt',
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
    const payload = data.payload || data || {};

    switch (action) {
      case 'PING':
        return jsonResponse({ success: true, message: 'API OK' });
      case 'SEARCH':
        return jsonResponse(searchRecords(payload));
      case 'SAVE_SCAN':
        return jsonResponse(saveScanRecord(payload));
      case 'LIST':
        return jsonResponse(listRecords(payload));
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
    return sheet;
  }

  const headerValues = sheet.getRange(1, 1, 1, HEADER_ROW.length).getValues()[0];
  if (headerValues.join('|') !== HEADER_ROW.join('|')) {
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

function normalizeValue(value) {
  return value === undefined || value === null ? '' : String(value).trim();
}

function normalizeSearchText(value) {
  return normalizeValue(value)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/\s+/g, ' ')
    .trim();
}

function rowToObject(row, rowIndex) {
  return {
    rowIndex: rowIndex,
    STT: row[0] || rowIndex,
    'Tên huyện': row[1] || '',
    'Tên xã cũ': row[2] || '',
    'Tên xã mới': row[3] || '',
    'Họ Tên CSD': row[4] || '',
    'Số hiệu tờ bản đồ mới': row[5] || '',
    'Số thứ tự thửa mới': row[6] || '',
    'Diện tích': row[7] || '',
    'Loại mục đích sử dụng': row[8] || '',
    'Số hiệu tờ bản đồ cũ': row[9] || '',
    'Số thứ tự thửa cũ': row[10] || '',
    'Ấp': row[11] || '',
    'Thế chấp': row[12] || '',
    'File GCN': row[13] || '',
    'File CCCD': row[14] || '',
    'Số CCCD': row[15] || '',
    fileName: row[16] || '',
    fileId: row[17] || '',
    fileUrl: row[18] || '',
    pageCount: row[19] || 0,
    createdAt: row[20] || '',
    source: row[21] || ''
  };
}

function searchRecords(payload) {
  const query = normalizeSearchText(payload && payload.query);
  const name = normalizeSearchText(payload && payload.name);
  const mapSheet = normalizeSearchText(payload && payload.soHieuToBanDoMoi);
  const plot = normalizeSearchText(payload && payload.soThuTuThuaMoi);
  const requestedLimit = Number(payload && payload.limit ? payload.limit : 20);
  const limit = Math.max(1, Math.min(100, isFinite(requestedLimit) ? requestedLimit : 20));
  const sheet = ensureSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return { success: true, count: 0, data: [] };
  }

  const values = sheet.getRange(2, 1, lastRow - 1, HEADER_ROW.length).getValues();
  const rows = values.map((row, index) => rowToObject(row, index + 2));

  const filtered = rows.filter((row) => {
    const normalizedName = normalizeSearchText(row['Họ Tên CSD']);
    const normalizedMapSheet = normalizeSearchText(row['Số hiệu tờ bản đồ mới']);
    const normalizedPlot = normalizeSearchText(row['Số thứ tự thửa mới']);
    const haystack = normalizeSearchText([
      row['Họ Tên CSD'],
      row['Số CCCD'],
      row['Tên huyện'],
      row['Tên xã mới'],
      row['Tên xã cũ'],
      row['Số hiệu tờ bản đồ mới'],
      row['Số thứ tự thửa mới'],
      row['Số hiệu tờ bản đồ cũ'],
      row['Số thứ tự thửa cũ'],
      row['Ấp'],
      row['File GCN'],
      row['File CCCD'],
      row['Loại mục đích sử dụng'],
      row['Diện tích']
    ].join(' '));

    return (!query || haystack.indexOf(query) >= 0) &&
      (!name || normalizedName.indexOf(name) >= 0) &&
      (!mapSheet || normalizedMapSheet.indexOf(mapSheet) >= 0) &&
      (!plot || normalizedPlot.indexOf(plot) >= 0);
  }).slice(-limit).reverse();

  return { success: true, count: filtered.length, data: filtered };
}

function listRecords(payload) {
  return searchRecords({ query: '', limit: payload && payload.limit ? payload.limit : 20 });
}

function saveScanRecord(payload) {
  const sheet = ensureSheet();
  const folder = ensureFolder();
  const fileName = normalizeValue(payload && payload.fileName ? payload.fileName : `Ho_So_Dat_Dai_${Date.now()}.pdf`);
  const rowIndex = Number(payload && payload.rowIndex ? payload.rowIndex : 0);
  const isUpdate = rowIndex >= 2 && rowIndex <= sheet.getLastRow();
  const existingRow = isUpdate ? sheet.getRange(rowIndex, 1, 1, HEADER_ROW.length).getValues()[0] : [];
  const pdfGCNBase64 = normalizeValue(payload && (payload.pdfGCNBase64 || payload.pdfBase64));
  const pdfCCCDBase64 = normalizeValue(payload && payload.pdfCCCDBase64);

  const record = {
    STT: sheet.getLastRow(),
    'Tên huyện': normalizeValue(payload && payload.tenHuyen),
    'Tên xã cũ': normalizeValue(payload && payload.tenXaCu),
    'Tên xã mới': normalizeValue(payload && payload.tenXaMoi),
    'Họ Tên CSD': normalizeValue(payload && payload.hoTenCSD),
    'Số hiệu tờ bản đồ mới': normalizeValue(payload && payload.soHieuToBanDoMoi),
    'Số thứ tự thửa mới': normalizeValue(payload && payload.soThuTuThuaMoi),
    'Diện tích': normalizeValue(payload && payload.dienTich),
    'Loại mục đích sử dụng': normalizeValue(payload && payload.loaiMucDichSuDung),
    'Số hiệu tờ bản đồ cũ': normalizeValue(payload && payload.soHieuToBanDoCu),
    'Số thứ tự thửa cũ': normalizeValue(payload && payload.soThuTuThuaCu),
    'Ấp': normalizeValue(payload && payload.ap),
    'Thế chấp': normalizeValue(payload && payload.theChap),
    'File GCN': normalizeValue(payload && payload.fileGCN) || (isUpdate ? normalizeValue(existingRow[13]) : ''),
    'File CCCD': normalizeValue(payload && payload.fileCCCD) || (isUpdate ? normalizeValue(existingRow[14]) : ''),
    'Số CCCD': normalizeValue(payload && payload.soCCCD),
    fileName,
    fileId: isUpdate ? normalizeValue(existingRow[17]) : '',
    fileUrl: isUpdate ? normalizeValue(existingRow[18]) : '',
    pageCount: Number(payload && payload.pageCount ? payload.pageCount : 0),
    createdAt: normalizeValue(payload && payload.createdAt ? new Date(payload.createdAt).toISOString() : new Date().toISOString()),
    source: 'mobile-app'
  };

  const baseName = fileName.replace(/\.pdf$/i, '');
  if (pdfGCNBase64) {
    const saved = savePdfToDrive(folder, pdfGCNBase64, `${baseName}_GCN.pdf`);
    record['File GCN'] = saved.fileUrl;
    if (!record.fileUrl) record.fileUrl = saved.fileUrl || '';
    if (!record.fileId) record.fileId = saved.fileId || '';
  }
  if (pdfCCCDBase64) {
    const saved = savePdfToDrive(folder, pdfCCCDBase64, `${baseName}_CCCD.pdf`);
    record['File CCCD'] = saved.fileUrl;
    if (!record.fileUrl) record.fileUrl = saved.fileUrl || '';
    if (!record.fileId) record.fileId = saved.fileId || '';
  }

  let row = [
    record.STT,
    record['Tên huyện'],
    record['Tên xã cũ'],
    record['Tên xã mới'],
    record['Họ Tên CSD'],
    record['Số hiệu tờ bản đồ mới'],
    record['Số thứ tự thửa mới'],
    record['Diện tích'],
    record['Loại mục đích sử dụng'],
    record['Số hiệu tờ bản đồ cũ'],
    record['Số thứ tự thửa cũ'],
    record['Ấp'],
    record['Thế chấp'],
    record['File GCN'],
    record['File CCCD'],
    record['Số CCCD'],
    record.fileName,
    record.fileId,
    record.fileUrl,
    String(record.pageCount),
    record.createdAt,
    record.source
  ];

  if (isUpdate) {
    const existingRow = sheet.getRange(rowIndex, 1, 1, HEADER_ROW.length).getValues()[0];
    row[0] = existingRow[0] || rowIndex - 1;
    sheet.getRange(rowIndex, 1, 1, HEADER_ROW.length).setValues([row]);
  } else {
    row[0] = sheet.getLastRow();
    sheet.appendRow(row);
  }

  return {
    success: true,
    message: isUpdate
      ? 'Dữ liệu hồ sơ đã được cập nhật.'
      : 'Dữ liệu đã được lưu vào Google Sheet.',
    fileId: record.fileId,
    fileUrl: record.fileUrl,
    fileGCN: record['File GCN'],
    fileCCCD: record['File CCCD'],
    updated: isUpdate,
    rowIndex: isUpdate ? rowIndex : sheet.getLastRow(),
    rowCount: sheet.getLastRow(),
    row
  };
}

function savePdfToDrive(folder, pdfBase64, fileName) {
  try {
    const raw = String(pdfBase64 || '');
    const clean = raw.replace(/^data:.*;base64,/, '').trim();

    if (!clean) {
      throw new Error('PDF base64 rỗng.');
    }

    const blob = Utilities.newBlob(Utilities.base64Decode(clean), 'application/pdf', fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return {
      fileId: file.getId(),
      fileUrl: file.getUrl()
    };
  } catch (error) {
    throw new Error('Lỗi lưu PDF lên Drive: ' + error.toString());
  }
}
