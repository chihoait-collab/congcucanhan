/**
 * GOOGLE APPS SCRIPT API TỐI ƯU CHO 20.000 DÒNG DỮ LIỆU
 * Cấu trúc các cột (16 cột - A đến P):
 * 1: STT | 2: Tên huyện | 3: Tên xã cũ | 4: Tên xã mới | 5: Họ Tên CSD | 6: Số hiệu tờ bản đồ mới
 * 7: Số thứ tự thửa mới | 8: Diện tích | 9: Loại mục đích sử dụng | 10: Số hiệu tờ bản đồ cũ
 * 11: Số thứ tự thửa cũ | 12: Ấp | 13: Thế chấp | 14: File GCN | 15: File CCCD | 16: Số CCCD
 */

const SHEET_NAME = "Du_Lieu"; // Đổi tên này nếu tab Sheet của bạn đặt tên khác
const FOLDER_GCN_NAME = "HOSO_GCN_DAT_DAI";
const FOLDER_CCCD_NAME = "HOSO_CCCD_DAT_DAI";

// Hàm xử lý chính nhận Yêu cầu POST từ Frontend GitHub Pages
function doPost(e) {
  var lock = LockService.getScriptLock();
  // Chờ tối đa 10 giây nếu nhiều request trùng nhau
  if (!lock.tryLock(10000)) {
    return createJsonResponse({ success: false, error: "Hệ thống đang bận, vui lòng thử lại!" });
  }

  try {
    var contents = JSON.parse(e.postData.contents);
    var action = contents.action;
    var payload = contents.payload;
    var result;

    switch (action) {
      case "SEARCH":
        result = searchData(payload);
        break;
      case "ADD_BATCH":
        result = addBatchData(payload);
        break;
      case "UPDATE_BATCH":
        result = updateBatchData(payload);
        break;
      case "DELETE_BATCH":
        result = deleteBatchData(payload);
        break;
      default:
        result = { success: false, error: "Hành động (action) không hợp lệ!" };
    }

    return createJsonResponse(result);

  } catch (error) {
    return createJsonResponse({ success: false, error: error.toString() });
  } finally {
    lock.releaseLock();
  }
}

// Bắt buộc thêm doGet để tránh lỗi khi test link hoặc khởi tạo
function doGet(e) {
  return createJsonResponse({ success: true, message: "API Hồ Sơ Đất Đai đang hoạt động!" });
}

// Hàm bổ trợ tạo Response chuẩn JSON tránh lỗi CORS
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// -----------------------------------------------------------------------------
// 1. CHỨC NĂNG TÌM KIẾM CỰC NHANH VỚI TEXTFINDER (~20.000 DÒNG)
// -----------------------------------------------------------------------------
function searchData(payload) {
  var query = payload.query ? payload.query.toString().trim() : "";
  var limit = payload.limit || 100; // Giới hạn số lượng trả về để tối ưu tốc độ
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  if (!query) {
    // Nếu không có từ khóa, lấy top 50 dòng mới nhất
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, data: [] };
    var startRow = Math.max(2, lastRow - limit + 1);
    var rangeData = sheet.getRange(startRow, 1, lastRow - startRow + 1, 16).getValues();
    return { success: true, count: rangeData.length, data: mapRowsToObjects(rangeData, startRow) };
  }

  // Sử dụng TextFinder quét trực tiếp trên toàn bộ bảng tính
  var finder = sheet.getRange("A2:P" + sheet.getLastRow())
                    .createTextFinder(query)
                    .matchCase(false)
                    .matchEntireCell(false);

  var results = finder.findAll();
  if (results.length === 0) {
    return { success: true, count: 0, data: [] };
  }

  // Lọc ra danh sách các dòng duy nhất (tránh 1 dòng khớp nhiều ô)
  var rowIndices = [];
  for (var i = 0; i < results.length; i++) {
    var rowIndex = results[i].getRow();
    if (rowIndices.indexOf(rowIndex) === -1) {
      rowIndices.push(rowIndex);
      if (rowIndices.length >= limit) break; // Đạt giới hạn thì dừng
    }
  }

  // Đọc dữ liệu các dòng tìm thấy
  var matchedData = [];
  for (var j = 0; j < rowIndices.length; j++) {
    var rIndex = rowIndices[j];
    var rowValues = sheet.getRange(rIndex, 1, 1, 16).getValues()[0];
    matchedData.push(mapSingleRowToObject(rowValues, rIndex));
  }

  return { success: true, count: matchedData.length, data: matchedData };
}

// -----------------------------------------------------------------------------
// 2. CHỨC NĂNG THÊM NHIỀU DÒNG CÙNG LÚC + LƯU FILE DRIVE
// -----------------------------------------------------------------------------
function addBatchData(payload) {
  var items = payload.items; // Mảng các đối tượng cần thêm
  if (!items || items.length === 0) {
    return { success: false, error: "Danh sách thêm trống!" };
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var folderGCN = getOrCreateFolder(FOLDER_GCN_NAME);
  var folderCCCD = getOrCreateFolder(FOLDER_CCCD_NAME);

  var lastRow = sheet.getLastRow();
  var startSTT = lastRow; // STT nối tiếp

  var rowsToInsert = [];

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    
    // Xử lý upload file Drive nếu có chuỗi Base64
    var linkGCN = item.fileGCNBase64 ? saveFileToDrive(folderGCN, item.fileGCNBase64, "GCN_" + (item.soThuaMoi || i)) : (item.fileGCN || "");
    var linkCCCD = item.fileCCCDBase64 ? saveFileToDrive(folderCCCD, item.fileCCCDBase64, "CCCD_" + (item.soCCCD || i)) : (item.fileCCCD || "");

    var row = [
      startSTT + i,                  // STT
      item.tenHuyen || "",          // Tên huyện
      item.tenXaCu || "",           // Tên xã cũ
      item.tenXaMoi || "",          // Tên xã mới
      item.hoTenCSD || "",          // Họ Tên CSD
      item.soToMoi || "",           // Số hiệu tờ bản đồ mới
      item.soThuaMoi || "",         // Số thứ tự thửa mới
      item.dienTich || "",          // Diện tích
      item.loaiMucDich || "",       // Loại mục đích sử dụng
      item.soToCu || "",            // Số hiệu tờ bản đồ cũ
      item.soThuaCu || "",          // Số thứ tự thửa cũ
      item.ap || "",                // Ấp
      item.theChap || "",           // Thế chấp
      linkGCN,                      // File GCN
      linkCCCD,                     // File CCCD
      item.soCCCD || ""             // Số CCCD
    ];
    rowsToInsert.push(row);
  }

  // Ghi 1 lần duy nhất xuống Google Sheet (Batch Write)
  sheet.getRange(lastRow + 1, 1, rowsToInsert.length, 16).setValues(rowsToInsert);

  return { success: true, message: "Đã thêm thành công " + rowsToInsert.length + " dòng!" };
}

// -----------------------------------------------------------------------------
// 3. CHỨC NĂNG CẬP NHẬT NHIỀU DÒNG CÙNG LÚC (BATCH UPDATE)
// -----------------------------------------------------------------------------
function updateBatchData(payload) {
  var items = payload.items; // Mảng chứa thông tin cập nhật (cần có rowIndex)
  if (!items || items.length === 0) {
    return { success: false, error: "Danh sách cập nhật trống!" };
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var folderGCN = getOrCreateFolder(FOLDER_GCN_NAME);
  var folderCCCD = getOrCreateFolder(FOLDER_CCCD_NAME);

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var rIndex = parseInt(item.rowIndex);

    if (!rIndex || rIndex < 2) continue;

    // Lấy dữ liệu cũ để giữ lại link file nếu không upload file mới
    var oldValues = sheet.getRange(rIndex, 1, 1, 16).getValues()[0];

    var linkGCN = item.fileGCNBase64 ? saveFileToDrive(folderGCN, item.fileGCNBase64, "GCN_" + item.soThuaMoi) : (item.fileGCN || oldValues[13]);
    var linkCCCD = item.fileCCCDBase64 ? saveFileToDrive(folderCCCD, item.fileCCCDBase64, "CCCD_" + item.soCCCD) : (item.fileCCCD || oldValues[14]);

    var updatedRow = [
      oldValues[0],                 // Giữ nguyên STT cũ
      item.tenHuyen !== undefined ? item.tenHuyen : oldValues[1],
      item.tenXaCu !== undefined ? item.tenXaCu : oldValues[2],
      item.tenXaMoi !== undefined ? item.tenXaMoi : oldValues[3],
      item.hoTenCSD !== undefined ? item.hoTenCSD : oldValues[4],
      item.soToMoi !== undefined ? item.soToMoi : oldValues[5],
      item.soThuaMoi !== undefined ? item.soThuaMoi : oldValues[6],
      item.dienTich !== undefined ? item.dienTich : oldValues[7],
      item.loaiMucDich !== undefined ? item.loaiMucDich : oldValues[8],
      item.soToCu !== undefined ? item.soToCu : oldValues[9],
      item.soThuaCu !== undefined ? item.soThuaCu : oldValues[10],
      item.ap !== undefined ? item.ap : oldValues[11],
      item.theChap !== undefined ? item.theChap : oldValues[12],
      linkGCN,
      linkCCCD,
      item.soCCCD !== undefined ? item.soCCCD : oldValues[15]
    ];

    sheet.getRange(rIndex, 1, 1, 16).setValues([updatedRow]);
  }

  return { success: true, message: "Đã cập nhật thành công " + items.length + " dòng!" };
}

// -----------------------------------------------------------------------------
// 4. CHỨC NĂNG XÓA NHIỀU DÒNG CÙNG LÚC (BATCH DELETE)
// -----------------------------------------------------------------------------
function deleteBatchData(payload) {
  var rowIndices = payload.rowIndices; // Mảng chứa các số dòng cần xóa [10, 25, 100]
  if (!rowIndices || rowIndices.length === 0) {
    return { success: false, error: "Danh sách dòng cần xóa trống!" };
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  // Sắp xếp thứ tự dòng từ LỚN đến BÉ để khi xóa không bị lệch chỉ số các dòng bên trên
  rowIndices.sort(function(a, b) { return b - a; });

  for (var i = 0; i < rowIndices.length; i++) {
    var rIndex = parseInt(rowIndices[i]);
    if (rIndex >= 2 && rIndex <= sheet.getLastRow()) {
      sheet.deleteRow(rIndex);
    }
  }

  return { success: true, message: "Đã xóa thành công " + rowIndices.length + " dòng!" };
}

// -----------------------------------------------------------------------------
// CÁC HÀM TIỆN ÍCH TRỢ GIÚP (HELPERS)
// -----------------------------------------------------------------------------
function mapSingleRowToObject(row, rowIndex) {
  return {
    rowIndex: rowIndex,
    stt: row[0],
    tenHuyen: row[1],
    tenXaCu: row[2],
    tenXaMoi: row[3],
    hoTenCSD: row[4],
    soToMoi: row[5],
    soThuaMoi: row[6],
    dienTich: row[7],
    loaiMucDich: row[8],
    soToCu: row[9],
    soThuaCu: row[10],
    ap: row[11],
    theChap: row[12],
    fileGCN: row[13],
    fileCCCD: row[14],
    soCCCD: row[15]
  };
}

function mapRowsToObjects(rows, startRowIndex) {
  var list = [];
  for (var i = 0; i < rows.length; i++) {
    list.push(mapSingleRowToObject(rows[i], startRowIndex + i));
  }
  return list;
}

function getOrCreateFolder(folderName) {
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

function saveFileToDrive(folder, base64Data, fileNamePrefix) {
  try {
    // Tách phần định dạng base64 nếu có
    var parts = base64Data.split(",");
    var contentType = parts[0].split(";")[0].replace("data:", "") || "application/pdf";
    var base64Content = parts.length > 1 ? parts[1] : parts[0];

    var bytes = Utilities.base64Decode(base64Content);
    var blob = Utilities.newBlob(bytes, contentType, fileNamePrefix + "_" + Date.now() + ".pdf");
    var file = folder.createFile(blob);
    
    // Cấp quyền xem link công khai
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    return "Lỗi upload file: " + e.toString();
  }
}