const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const FOLDER_ID = '1GBqXdt2b_VPgjhyH0frq-YnGF9aMOa8e';

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Quản Lý Nhân Sự Dự Án')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 1. Hàm thêm nhân sự mới kèm ảnh, Nhà thầu và Số điện thoại
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
      fileUrl = 'https://lh3.googleusercontent.com/d/' + file.getId();
    }
    
    const idMoi = 'NS_' + new Date().getTime();
    
    // Tạo một mảng dòng mới gồm 17 cột, đặt SĐT ở vị trí thứ 17 (index 16)
    // Các cột hồ sơ scan (12-16) để trống ban đầu
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

// 2. Hàm cập nhật thông tin nhân sự (Bao gồm Số điện thoại)
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
    sheet.getRange(targetRow, 17).setValue(data.sdt); // Cập nhật Số điện thoại ở cột 17 (Q)
    
    if (data.fileData) {
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const contentType = data.fileData.substring(5, data.fileData.indexOf(';base64'));
      const bytes = Utilities.base64Decode(data.fileData.split(',')[1]);
      const blob = Utilities.newBlob(bytes, contentType, data.fileName);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const fileUrl = 'https://lh3.googleusercontent.com/d/' + file.getId();
      sheet.getRange(targetRow, 6).setValue(fileUrl);
    }
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// 3. Hàm chuẩn hóa link ảnh hiển thị trên web
function chuyểnĐổiLinkẢnh(url) {
  if (!url || url === '') return 'https://via.placeholder.com/150';
  if (url.includes('googleusercontent.com')) return url;

  let match = url.match(/id=([^&]+)/);
  if (match && match[1]) return 'https://lh3.googleusercontent.com/d/' + match[1];

  match = url.match(/\/file\/d\/([^\/]+)/);
  if (match && match[1]) return 'https://lh3.googleusercontent.com/d/' + match[1];

  return url;
}

// 4. Hàm lấy danh sách dự án
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

// 5. Hàm lấy danh sách nhân sự theo dự án (Lấy thêm Số điện thoại ở cột 17)
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
      sdt: row[16] ? row[16].toString() : '', // Lấy dữ liệu Số điện thoại ở cột số 17 (index 16)
      hoSo: { 
        syll: row[11] ? row[11].toString() : '', 
        lltp: row[12] ? row[12].toString() : '', 
        cccdScan: row[13] ? row[13].toString() : '', 
        gksk: row[14] ? row[14].toString() : '', 
        chungChi: row[15] ? row[15].toString() : '', 
        camKet: row[16] ? row[16].toString() : '' 
      }
    }));
  } catch(e) {
    throw new Error("Lỗi đọc bảng NhanSu: " + e.message);
  }
}

// 6. Hàm tải tài liệu scan lên hệ thống
function uploadFileAndLink(idNhanVien, columnName, fileData, fileName) {
  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const contentType = fileData.substring(5, fileData.indexOf(';base64'));
    const bytes = Utilities.base64Decode(fileData.split(',')[1]);
    const blob = Utilities.newBlob(bytes, contentType, fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const fileUrl = 'https://lh3.googleusercontent.com/d/' + file.getId();
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('NhanSu');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const colIndex = headers.indexOf(columnName);
    
    if (colIndex === -1) return { success: false, error: 'Không tìm thấy cột: ' + columnName };
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

// 7. Hàm bảo trì quét và sửa lỗi link ảnh hàng loạt trên sheet
function suaTatCaLinkAnh() {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName("NhanSu");
    const data = sheet.getDataRange().getValues();

    for(let i = 1; i < data.length; i++){
      let url = data[i][5] ? data[i][5].toString() : "";
      let match = url.match(/id=([^&]+)/);
      if(match){
        sheet.getRange(i + 1, 6).setValue(
          "https://lh3.googleusercontent.com/d/" + match[1]
        );
      }
    }
  } catch(e) {
    Logger.log("Lỗi chạy hàm sửa link ảnh: " + e.toString());
  }
}
