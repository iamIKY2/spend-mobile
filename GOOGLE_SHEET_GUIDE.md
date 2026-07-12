# Hướng Dẫn Kết Nối Google Sheets & Apps Script API

Tài liệu này hướng dẫn bạn chi tiết cách thiết lập Google Sheets và Google Apps Script để kết nối hoàn hảo với ứng dụng **SpendManager**.

---

## 1. Chuẩn bị Google Sheet
1. Mở [Google Sheets](https://sheets.google.com/) và tạo một trang tính mới (Ví dụ đặt tên: **Quản Lý Chi Tiêu Cá Nhân - SpendManager**).
2. Tại sheet đầu tiên (đặt tên là `Transactions` hoặc để mặc định là `Trang tính 1` / `Sheet1`), hãy tạo hàng tiêu đề (Hàng 1) với các cột sau:
   - **Cột A**: `id` (Mã giao dịch)
   - **Cột B**: `date` (Ngày tháng - định dạng YYYY-MM-DD)
   - **Cột C**: `type` (Loại: `income` cho Thu nhập, `expense` cho Chi tiêu)
   - **Cột D**: `category` (Mã danh mục, ví dụ: `food`, `shopping`, `salary`...)
   - **Cột E**: `amount` (Số tiền - số nguyên VND)
   - **Cột F**: `note` (Ghi chú)
   - **Cột G**: `createdAt` (Thời gian tạo)

---

## 2. Mã Nguồn Google Apps Script (Backend API)
1. Trong Google Sheets, chọn menu **Tiện ích mở rộng (Extensions)** > **Apps Script**.
2. Xóa toàn bộ mã cũ trong file `Mã.gs` (hoặc `Code.gs`) và dán toàn bộ đoạn mã dưới đây vào:

```javascript
/**
 * SPEND MANAGER - GOOGLE APPS SCRIPT API BACKEND
 * Hỗ trợ các phương thức GET (lấy danh sách) và POST/GET (thêm/xóa giao dịch & danh mục)
 */

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Transactions") || ss.getSheets()[0];
  return sheet;
}

function getCategoriesSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss ? ss.getSheetByName("Categories") : null;
  if (!sheet && ss) {
    sheet = ss.insertSheet("Categories");
    initializeDefaultCategories(sheet);
  }
  return sheet;
}

// Xử lý yêu cầu GET từ Web App (Lấy danh sách hoặc Thêm mới qua URL parameters)
function doGet(e) {
  var action = e.parameter.action || 'get';
  
  // 1. Lấy danh sách giao dịch
  if (action === 'get' || action === 'getTransactions') {
    return getTransactions();
  }
  
  // 2. Lấy danh sách danh mục
  if (action === 'getCategories') {
    return getCategories();
  }
  
  // 3. Thêm giao dịch qua GET (dùng cho trường hợp fallback CORS)
  if (action === 'add') {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var timezone = ss ? ss.getSpreadsheetTimeZone() : Session.getScriptTimeZone();
    var data = {
      id: e.parameter.id || ('TR_' + new Date().getTime()),
      date: e.parameter.date || Utilities.formatDate(new Date(), timezone, "yyyy-MM-dd"),
      type: e.parameter.type || 'expense',
      category: e.parameter.category || 'other',
      amount: Number(e.parameter.amount) || 0,
      note: e.parameter.note || '',
      createdAt: new Date().toISOString()
    };
    return addTransaction(data);
  }

  // 4. Thêm/Cập nhật danh mục qua GET (fallback CORS)
  if (action === 'saveCategory') {
    var catData = {
      id: e.parameter.id,
      name: e.parameter.name,
      type: e.parameter.type,
      icon: e.parameter.icon,
      color: e.parameter.color,
      bgColor: e.parameter.bgColor
    };
    return saveCategory(catData);
  }

  // 5. Xóa giao dịch qua GET
  if (action === 'delete') {
    var id = e.parameter.id;
    return deleteTransaction(id);
  }

  // 6. Xóa danh mục qua GET
  if (action === 'deleteCategory') {
    var id = e.parameter.id;
    return deleteCategory(id);
  }

  // 7. Reset danh mục qua GET
  if (action === 'resetCategories') {
    var sheet = getCategoriesSheet();
    initializeDefaultCategories(sheet);
    return createJsonResponse({ status: 'success', message: 'Đã khôi phục danh mục mặc định' });
  }
  
  return createJsonResponse({ status: 'error', message: 'Invalid action' });
}

// Xử lý yêu cầu POST từ Web App (Thêm giao dịch qua JSON / FormData)
function doPost(e) {
  try {
    var data;
    if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch(err) {
        // Nếu gửi dạng FormData / URLSearchParams
        data = e.parameter;
      }
    } else {
      data = e.parameter;
    }
    
    var action = data.action || 'add';
    if (action === 'add') {
      return addTransaction(data);
    } else if (action === 'delete') {
      return deleteTransaction(data.id);
    } else if (action === 'saveCategory') {
      return saveCategory(data);
    } else if (action === 'deleteCategory') {
      return deleteCategory(data.id);
    } else if (action === 'resetCategories') {
      var sheet = getCategoriesSheet();
      initializeDefaultCategories(sheet);
      return createJsonResponse({ status: 'success', message: 'Đã khôi phục danh mục mặc định' });
    }
    
    return createJsonResponse({ status: 'success', message: 'Received POST request', data: data });
  } catch (err) {
    return createJsonResponse({ status: 'error', message: err.toString() });
  }
}

// Hàm lấy tất cả giao dịch từ Sheet
function getTransactions() {
  try {
    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return createJsonResponse({ status: 'success', transactions: [] });
    }
    
    var headers = data[0];
    var transactions = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0] && !row[4]) continue; // Bỏ qua dòng trống
      
      var dVal = row[1];
      var dStr = '';
      if (dVal instanceof Date) {
        dStr = Utilities.formatDate(dVal, sheet.getParent().getSpreadsheetTimeZone(), "yyyy-MM-dd");
      } else {
        dStr = String(dVal).split('T')[0];
      }

      transactions.push({
        id: String(row[0]),
        date: dStr,
        type: String(row[2]),
        category: String(row[3]),
        amount: Number(row[4]) || 0,
        note: String(row[5] || ''),
        createdAt: String(row[6] || '')
      });
    }
    
    // Sắp xếp mới nhất lên đầu
    transactions.reverse();
    
    return createJsonResponse({ status: 'success', transactions: transactions });
  } catch (err) {
    return createJsonResponse({ status: 'error', message: err.toString() });
  }
}

// Hàm thêm giao dịch mới
function addTransaction(data) {
  try {
    var sheet = getSheet();
    var ss = sheet.getParent();
    var timezone = ss ? ss.getSpreadsheetTimeZone() : Session.getScriptTimeZone();
    var newRow = [
      data.id || ('TR_' + new Date().getTime()),
      data.date || Utilities.formatDate(new Date(), timezone, "yyyy-MM-dd"),
      data.type || 'expense',
      data.category || 'other',
      Number(data.amount) || 0,
      data.note || '',
      new Date().toISOString()
    ];
    
    sheet.appendRow(newRow);
    
    return createJsonResponse({ 
      status: 'success', 
      message: 'Đã thêm giao dịch vào Google Sheets',
      transaction: {
        id: newRow[0],
        date: newRow[1],
        type: newRow[2],
        category: newRow[3],
        amount: newRow[4],
        note: newRow[5],
        createdAt: newRow[6]
      }
    });
  } catch (err) {
    return createJsonResponse({ status: 'error', message: err.toString() });
  }
}

// Hàm xóa giao dịch theo ID
function deleteTransaction(id) {
  try {
    if (!id) return createJsonResponse({ status: 'error', message: 'Missing ID' });
    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return createJsonResponse({ status: 'success', message: 'Đã xóa giao dịch ID: ' + id });
      }
    }
    return createJsonResponse({ status: 'error', message: 'Không tìm thấy ID: ' + id });
  } catch (err) {
    return createJsonResponse({ status: 'error', message: err.toString() });
  }
}

// Hàm khởi tạo danh mục mặc định
function initializeDefaultCategories(sheet) {
  var defaults = [
    ['food', 'Ăn uống', 'expense', 'utensils', '#FF6B6B', 'rgba(255, 107, 107, 0.15)'],
    ['shopping', 'Mua sắm', 'expense', 'shopping-bag', '#4ECDC4', 'rgba(78, 205, 196, 0.15)'],
    ['transport', 'Di chuyển', 'expense', 'car', '#45B7D1', 'rgba(69, 183, 209, 0.15)'],
    ['bills', 'Hóa đơn & Điện nước', 'expense', 'zap', '#F9A826', 'rgba(249, 168, 38, 0.15)'],
    ['entertainment', 'Giải trí', 'expense', 'film', '#9B51E0', 'rgba(155, 81, 224, 0.15)'],
    ['health', 'Sức khỏe & Y tế', 'expense', 'heart', '#EB5757', 'rgba(235, 87, 87, 0.15)'],
    ['education', 'Học tập & Sách', 'expense', 'book-open', '#2D9CDB', 'rgba(45, 156, 219, 0.15)'],
    ['other_expense', 'Chi tiêu khác', 'expense', 'more-horizontal', '#828282', 'rgba(130, 130, 130, 0.15)'],
    ['salary', 'Tiền lương', 'income', 'dollar-sign', '#27AE60', 'rgba(39, 174, 96, 0.15)'],
    ['bonus', 'Thưởng & Quà tặng', 'income', 'gift', '#F2994A', 'rgba(242, 153, 74, 0.15)'],
    ['investment', 'Đầu tư & Kinh doanh', 'income', 'trending-up', '#2F80ED', 'rgba(47, 128, 237, 0.15)'],
    ['other_income', 'Thu nhập khác', 'income', 'plus-circle', '#6FCF97', 'rgba(111, 207, 151, 0.15)']
  ];
  
  sheet.clear();
  sheet.appendRow(["id", "name", "type", "icon", "color", "bgColor"]);
  for (var i = 0; i < defaults.length; i++) {
    sheet.appendRow(defaults[i]);
  }
}

// Hàm lấy tất cả danh mục từ Sheet
function getCategories() {
  try {
    var sheet = getCategoriesSheet();
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      initializeDefaultCategories(sheet);
      data = sheet.getDataRange().getValues();
    }
    
    var categories = {};
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      var id = String(row[0]);
      categories[id] = {
        id: id,
        name: String(row[1] || ''),
        type: String(row[2] || 'expense'),
        icon: String(row[3] || 'tag'),
        color: String(row[4] || '#6366F1'),
        bgColor: String(row[5] || 'rgba(99, 102, 241, 0.15)')
      };
    }
    return createJsonResponse({ status: 'success', categories: categories });
  } catch (err) {
    return createJsonResponse({ status: 'error', message: err.toString() });
  }
}

// Hàm thêm/cập nhật danh mục
function saveCategory(data) {
  try {
    var sheet = getCategoriesSheet();
    var values = sheet.getDataRange().getValues();
    if (values.length <= 1) {
      initializeDefaultCategories(sheet);
      values = sheet.getDataRange().getValues();
    }
    var id = data.id;
    if (!id) {
      return createJsonResponse({ status: 'error', message: 'Missing Category ID' });
    }
    
    var newRow = [
      id,
      data.name || 'Danh mục mới',
      data.type || 'expense',
      data.icon || 'tag',
      data.color || '#6366F1',
      data.bgColor || 'rgba(99, 102, 241, 0.15)'
    ];
    
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(id)) {
        var range = sheet.getRange(i + 1, 1, 1, 6);
        range.setValues([newRow]);
        return createJsonResponse({ 
          status: 'success', 
          message: 'Đã cập nhật danh mục',
          category: {
            id: newRow[0],
            name: newRow[1],
            type: newRow[2],
            icon: newRow[3],
            color: newRow[4],
            bgColor: newRow[5]
          }
        });
      }
    }
    
    sheet.appendRow(newRow);
    return createJsonResponse({ 
      status: 'success', 
      message: 'Đã thêm danh mục mới',
      category: {
        id: newRow[0],
        name: newRow[1],
        type: newRow[2],
        icon: newRow[3],
        color: newRow[4],
        bgColor: newRow[5]
      }
    });
  } catch (err) {
    return createJsonResponse({ status: 'error', message: err.toString() });
  }
}

// Hàm xóa danh mục theo ID
function deleteCategory(id) {
  try {
    if (!id) return createJsonResponse({ status: 'error', message: 'Missing ID' });
    var sheet = getCategoriesSheet();
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return createJsonResponse({ status: 'success', message: 'Đã xóa danh mục ID: ' + id });
      }
    }
    return createJsonResponse({ status: 'error', message: 'Không tìm thấy ID: ' + id });
  } catch (err) {
    return createJsonResponse({ status: 'error', message: err.toString() });
  }
}

// Helper trả về JSON có hỗ trợ CORS
function createJsonResponse(data) {
  var output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}
```

---

## 3. Cách Triển Khai (Deploy) Web App
1. Nhấp vào nút **Triển khai (Deploy)** ở góc trên bên phải > Chọn **Buổi triển khai mới (New deployment)**.
2. Nhấp vào biểu tượng bánh răng bên cạnh "Chọn loại (Select type)" > Chọn **Ứng dụng web (Web app)**.
3. Cấu hình như sau:
   - **Mô tả**: `SpendManager API v1`
   - **Thực thi dưới dạng (Execute as)**: **Tôi (Me / chủ sở hữu email của bạn)**.
   - **Ai có quyền truy cập (Who has access)**: **Bất kỳ ai (Anyone)** *(Rất quan trọng để ứng dụng web có thể gọi API mà không bị chặn bởi trang đăng nhập Google)*.
4. Nhấp vào **Triển khai (Deploy)**.
5. Sao chép **URL ứng dụng web (Web app URL)** và dán vào phần cài đặt của trang web (Mặc định trang web đã tích sẵn URL của bạn: `https://script.google.com/macros/s/AKfycbyK6MP86Twyh4UAGqACHmYmxgdxh_5Zhi7RGrqfsQozEsKZFs6ZUwquHA-yZuiv_TDZ/exec`).

> [!NOTE]
> Khi triển khai lần đầu, Google sẽ yêu cầu **Cấp quyền truy cập (Authorize access)**. Hãy nhấp vào "Xem lại quyền (Review permissions)" -> Chọn tài khoản Google -> Chọn "Nâng cao (Advanced)" -> "Đi tới dự án (Go to... unsafe)" -> Chọn "Cho phép (Allow)".

> [!IMPORTANT]
> **QUAN TRỌNG KHI CẬP NHẬT MÃ NGUỒN APPS SCRIPT:** Mỗi khi bạn chỉnh sửa hoặc cập nhật mã nguồn trong file `Mã.gs` (Ví dụ: cập nhật hàm xử lý danh mục `saveCategory`, `getCategories`...), việc chỉ ấn Lưu (`Ctrl + S`) là **CHƯA ĐỦ**. Bạn **BẮT BUỘC** phải tạo bản triển khai mới bằng cách:
> Nhấp **Triển khai (Deploy)** -> **Quản lý bản triển khai (Manage deployments)** -> Nhấp biểu tượng chiếc bút **Chỉnh sửa (Edit)** -> Ở mục **Phiên bản (Version)** chọn **Mới (New)** -> Nhấp **Triển khai (Deploy)**. Nếu không làm bước này, ứng dụng sẽ vẫn chạy trên phiên bản cũ và không thể lưu/đồng bộ danh mục giữa laptop và điện thoại!
