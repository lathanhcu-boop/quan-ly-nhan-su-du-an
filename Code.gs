const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const FOLDER_ID = '1GBqXdt2b_VPgjhyH0frq-YnGF9aMOa8e';

// ==========================================
// 1. CÁC HÀM API GIAO TIẾP VỚI GITHUB (FETCH)
// ==========================================

// Hàm helper để trả về kết quả dưới dạng JSON
function responseJson(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Xử lý các yêu cầu lấy dữ liệu (GET)
function doGet(e) {
  // Nếu không có tham số truyền vào, báo lỗi hoặc có thể trả về thông báo
  if (!e || !e.parameter || !e.parameter.action) {
    return responseJson({ error: 'Vui lòng cung cấp tham số action.' });
  }

  const action = e.parameter.action;

  if (action === 'getDuAn') {
    const result = getDanhSachDuAn();
    return responseJson(result);
  }

  if (action === 'getNhanSu') {
    const idDuAn = e.parameter.idDuAn;
    const result = getNhanSuTheoDuAn(idDuAn);
    return responseJson(result);
  }

  return responseJson({ error: 'Hành động GET không hợp lệ.' });
}

// Xử lý các yêu cầu gửi dữ liệu/upload (POST)
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const payload = body.payload;

    if (action === 'themNhanSu') {
      const result = themNhanSuMoiVoiAnh(payload);
      return responseJson(result);
    }

    if (action === 'capNhatNhanSu') {
      const result = capNhatThongTinNhanSu(payload);
      return responseJson(result);
    }

    if (action === 'uploadFile') {
      const result = uploadFileAndLink(payload.idNhanVien, payload.columnName, payload.fileData, payload.fileName);
      return responseJson(result);
    }

    return responseJson({ success: false, error: 'Hành động POST không hợp lệ.' });
  } catch (error) {
    return responseJson({ success: false, error: 'Lỗi xử lý POST: ' + error.toString() });
  }
}

// ==========================================
// 2. CÁC HÀM XỬ LÝ LOGIC CHÍNH (GIỮ NGUYÊN)
// ==========================================

// Hàm thêm nhân sự mới kèm ảnh đại diện (Link xem trực tuyến)
function themNhanSuMoiVoiAnh(data) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('NhanSu');
    const folder = DriveApp.getFolderById(FOLDER_ID);
    
    let fileUrl = '';
    if (data.fileData) {
      const contentType = data.fileData.substring(5, data.fileData.indexOf(';base64'));
      const bytes = Utilities.base64Decode(data.fileData.split(',')[1]);
      const blob = Utilities.newBlob(bytes, contentType, data.fileName);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = 'https://drive.google.com/file/d/' + file.getId() + '/view?usp=drivesdk';
    }
    
    const idMoi = 'NS_' + new Date().getTime();
    const rowData = [
      idMoi,          // 1. Mã nhân sự (A)
      data.idDuAn,    // 2. Mã dự án (B)
      data.ten,       // 3. Họ tên (C)
      data.viTri,     // 4. Vị trí (D)
      data.tinhTrang, // 5. Ngày hết hạn HD (E)
      fileUrl,        // 6. Link ảnh đại diện (F)
      data.ngaySinh,  // 7. Ngày sinh (G)
      data.cccd,      // 8. CCCD (H)
      data.thuongTru, // 9. Thường trú (I)
      data.tamTru,    // 10. Tạm trú (J)
      data.nhaThau,   // 11. Nhà thầu (K)
      '',             // 12. SyLL (L)
      '',             // 13. LLTP (M)
      '',             // 14. ScanCCCD (N)
      '',             // 15. GKSK (O)
      '',             // 16. ChungChi (P)
      data.sdt        // 17. Số điện thoại (Q)
    ];
    sheet.appendRow(rowData);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Hàm cập nhật thông tin nhân sự
function capNhatThongTinNhanSu(data) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('NhanSu');
    const values = sheet.getDataRange().getValues();
    
    let targetRow = -1;
    for (let i = 1; i < values.length; i++) {
      if (values[i][0].toString() === data.idNhanSu.toString()) {
        targetRow = i + 1;
        break;
      }
    }
    
    if (targetRow === -1) return { success: false, error: 'Không tìm thấy mã nhân viên' };
    
    sheet.getRange(targetRow, 3).setValue(data.ten);
    sheet.getRange(targetRow, 4).setValue(data.viTri);
    sheet.getRange(targetRow, 5).setValue(data.tinhTrang);
    sheet.getRange(targetRow, 7).setValue(data.ngaySinh);
    sheet.getRange(targetRow, 8).setValue(data.cccd);
    sheet.getRange(targetRow, 9).setValue(data.thuongTru);
    sheet.getRange(targetRow, 10).setValue(data.tamTru);
    sheet.getRange(targetRow, 11).setValue(data.nhaThau);
    sheet.getRange(targetRow, 17).setValue(data.sdt);
    
    if (data.fileData) {
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const contentType = data.fileData.substring(5, data.fileData.indexOf(';base64'));
      const bytes = Utilities.base64Decode(data.fileData.split(',')[1]);
      const blob = Utilities.newBlob(bytes, contentType, data.fileName);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const fileUrl = 'https://drive.google.com/file/d/' + file.getId() + '/view?usp=drivesdk';
      sheet.getRange(targetRow, 6).setValue(fileUrl);
    }
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// Hàm chuẩn hóa link ảnh hiển thị trên giao diện web
function chuyểnĐổiLinkẢnh(url) {
  if (!url || url === '') return 'https://via.placeholder.com/150';
  if (url.includes('thumbnail?id=')) return url;
  
  let match = url.match(/\/file\/d\/([^\/]+)/);
  if (match && match[1]) return 'https://drive.google.com/thumbnail?id=' + match[1] + '&sz=w1000';

  match = url.match(/id=([^&]+)/);
  if (match && match[1]) return 'https://drive.google.com/thumbnail?id=' + match[1] + '&sz=w1000';

  return url;
}

// Hàm lấy danh sách dự án
function getDanhSachDuAn() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('DuAn');
    const data = sheet.getDataRange().getValues();
    data.shift(); 
    return data.map(row => ({
      id: row[0] ? row[0].toString() : '',
      ten: row[1] ? row[1].toString() : '',
      anh: chuyểnĐổiLinkẢnh(row[2])
    }));
  } catch(e) {
    throw new Error("Lỗi đọc bảng DuAn: " + e.message);
  }
}

// Hàm lấy danh sách nhân sự
function getNhanSuTheoDuAn(idDuAn) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('NhanSu');
    const data = sheet.getDataRange().getValues();
    data.shift();
    
    return data.filter(row => row[1].toString() === idDuAn.toString()).map(row => ({
      id: row[0] ? row[0].toString() : '',
      ten: row[2] ? row[2].toString() : '',
      viTri: row[3] ? row[3].toString() : '',
      tinhTrang: row[4] ? (row[4] instanceof Date ? Utilities.formatDate(row[4], "GMT+7", "yyyy-MM-dd") : row[4].toString()) : 'Chưa cập nhật',
      anh: chuyểnĐổiLinkẢnh(row[5]),
      ngaySinh: row[6] ? (row[6] instanceof Date ? Utilities.formatDate(row[6], "GMT+7", "dd/MM/yyyy") : row[6].toString()) : 'Chưa cập nhật',
      cccd: row[7] ? row[7].toString() : '',
      thuongTru: row[8] ? row[8].toString() : '',
      tamTru: row[9] ? row[9].toString() : '',
      nhaThau: row[10] ? row[10].toString() : 'Chưa chọn',
      sdt: row[16] ? row[16].toString() : '', 
      hoSo: { 
        syll: row[11] ? row[11].toString() : '',       
        lltp: row[12] ? row[12].toString() : '',       
        cccdScan: row[13] ? row[13].toString() : '',   
        gksk: row[14] ? row[14].toString() : '',       
        chungChi: row[15] ? row[15].toString() : ''    
      }
    }));
  } catch(e) {
    throw new Error("Lỗi đọc bảng NhanSu: " + e.message);
  }
}

// Hàm tải tài liệu scan lên, trả về link MỞ XEM TRỰC TUYẾN
function uploadFileAndLink(idNhanVien, columnName, fileData, fileName) {
  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const contentType = fileData.substring(5, fileData.indexOf(';base64'));
    const bytes = Utilities.base64Decode(fileData.split(',')[1]);
    const blob = Utilities.newBlob(bytes, contentType, fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const fileUrl = 'https://drive.google.com/file/d/' + file.getId() + '/view?usp=drivesdk';
    
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('NhanSu');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const colIndex = headers.indexOf(columnName);
    
    if (colIndex === -1) return { success: false, error: 'Không tìm thấy tên cột "' + columnName + '" trên dòng tiêu đề Google Sheet.' };
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() == idNhanVien.toString()) {
        sheet.getRange(i + 1, colIndex + 1).setValue(fileUrl);
        return { success: true, url: fileUrl };
      }
    }
    return { success: false, error: 'Không tìm thấy mã nhân viên' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}
