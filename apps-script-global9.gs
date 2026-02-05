/**
 * Google Apps Script cho Quiz Global 9 - Unit 7 & 8
 * Tự động tạo header và ghi kết quả đầy đủ
 */

// Tạo header nếu chưa có
function createHeaders() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var headers = sheet.getRange(1, 1, 1, 8).getValues()[0];
  
  // Kiểm tra nếu chưa có header
  if (headers[0] !== 'Thời gian') {
    sheet.getRange(1, 1, 1, 8).setValues([[
      'Thời gian',
      'Họ tên',
      'Lớp', 
      'SĐT Phụ huynh',
      'Điểm',
      'Số câu đúng',
      'Tổng câu',
      'Thời gian làm'
    ]]);
    
    // Format header
    var headerRange = sheet.getRange(1, 1, 1, 8);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('white');
    headerRange.setHorizontalAlignment('center');
    
    // Freeze header row
    sheet.setFrozenRows(1);
  }
}

// Xử lý POST request từ quiz
function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    createHeaders();
    
    var data = JSON.parse(e.postData.contents);
    
    // Tạo timestamp theo múi giờ Việt Nam
    var now = new Date();
    var vietnamTime = Utilities.formatDate(now, 'Asia/Ho_Chi_Minh', 'HH:mm:ss dd/MM/yyyy');
    
    // Thêm dữ liệu vào sheet
    sheet.appendRow([
      vietnamTime,                    // Thời gian (VN timezone)
      data.name || '',                // Họ tên
      data.className || '',           // Lớp
      data.parentPhone || '',         // SĐT phụ huynh
      data.score || 0,                // Điểm
      data.correctCount || 0,         // Số câu đúng
      data.totalQuestions || 20,      // Tổng câu
      data.timeUsed || ''             // Thời gian làm
    ]);
    
    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      message: 'Đã lưu kết quả'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Xử lý GET request (for testing)
function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    createHeaders();
    
    var params = e.parameter;
    
    // Tạo timestamp theo múi giờ Việt Nam
    var now = new Date();
    var vietnamTime = Utilities.formatDate(now, 'Asia/Ho_Chi_Minh', 'HH:mm:ss dd/MM/yyyy');
    
    // Thêm dữ liệu vào sheet
    sheet.appendRow([
      vietnamTime,
      params.name || params.studentName || '',
      params.className || params.studentClass || '',
      params.parentPhone || '',
      params.score || 0,
      params.correctCount || 0,
      params.totalQuestions || 20,
      params.timeUsed || ''
    ]);
    
    return ContentService.createTextOutput('OK');
    
  } catch (error) {
    return ContentService.createTextOutput('Error: ' + error.toString());
  }
}

// Hàm test
function testScript() {
  createHeaders();
  Logger.log('Headers created successfully!');
}
