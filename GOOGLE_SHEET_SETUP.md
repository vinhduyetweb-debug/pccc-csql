# PCCC-CSQL Google Sheet Sync

## 1. Tạo Google Sheet

1. Tạo một Google Sheet mới.
2. Đặt tên sheet, ví dụ: `PCCC-CSQL`.
3. Có thể để trống sheet; Apps Script bên dưới sẽ tự tạo header nếu chưa có.

## 2. Mở Apps Script

1. Trong Google Sheet, chọn `Extensions` > `Apps Script`.
2. Xóa code mặc định.
3. Dán toàn bộ code bên dưới.

```javascript
const SHEET_NAME = "PCCC-CSQL";

const HEADERS = [
  "Thời gian đồng bộ",
  "Mã cơ sở",
  "Số hồ sơ",
  "Số tài khoản PCCC",
  "Tên cơ sở",
  "Địa chỉ",
  "Phường/Xã",
  "Thành phố",
  "Lĩnh vực",
  "Tính chất hoạt động",
  "Tháng năm hoạt động",
  "Số tầng",
  "Khối tích",
  "Tổng diện tích sàn",
  "Nhóm quản lý",
  "Tình trạng thẩm duyệt",
  "Người đứng đầu",
  "SĐT người đứng đầu",
  "Người thường trực",
  "SĐT người thường trực",
  "Cơ quan quản lý nhà nước về PCCC",
  "Cán bộ quản lý",
  "SĐT cán bộ quản lý",
  "Nội dung cần lưu ý",
  "Ngày tạo",
  "Ngày cập nhật",
];

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || "{}");
    if (payload.action !== "UPSERT_FACILITY" || !payload.facility) {
      return jsonResponse({ success: false, error: "Invalid payload" });
    }

    const facility = payload.facility;
    const sheet = getOrCreateSheet();
    ensureHeaders(sheet);

    const rowValues = buildRow(payload.syncedAt, facility);
    const rowIndex = findRowByPcccAccount(sheet, facility.soTaiKhoanPCCC);

    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 1, 1, HEADERS.length).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }

    return jsonResponse({ success: true, action: rowIndex > 0 ? "updated" : "created" });
  } catch (error) {
    return jsonResponse({ success: false, error: error.message });
  }
}

function getOrCreateSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
}

function ensureHeaders(sheet) {
  const firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const hasHeader = firstRow.some((value) => String(value || "").trim());
  if (!hasHeader) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
}

function findRowByPcccAccount(sheet, soTaiKhoanPCCC) {
  const account = String(soTaiKhoanPCCC || "").trim();
  if (!account) return -1;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const values = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
  const index = values.findIndex((row) => String(row[0] || "").trim() === account);
  return index >= 0 ? index + 2 : -1;
}

function buildRow(syncedAt, facility) {
  return [
    syncedAt || new Date().toISOString(),
    facility.maCoSo || "",
    facility.soHoSo || "",
    facility.soTaiKhoanPCCC || "",
    facility.tenCoSo || "",
    facility.diaChi || "",
    facility.phuongXa || "",
    facility.thanhPho || "",
    facility.linhVucName || facility.linhVucCode || "",
    facility.tinhChatHoatDong || "",
    facility.thangNamHoatDong || "",
    facility.soTang === "" ? "" : facility.soTang,
    facility.khoiTich === "" ? "" : facility.khoiTich,
    facility.tongDienTichSan === "" ? "" : facility.tongDienTichSan,
    facility.nhomQuanLy || "",
    facility.tinhTrangThamDuyetNghiemThu || "",
    facility.nguoiDungDau || "",
    facility.sdtNguoiDungDau || "",
    facility.nguoiThuongTruc || "",
    facility.sdtNguoiThuongTruc || "",
    facility.donViQuanLy || "",
    facility.canBoQuanLy || "",
    facility.sdtCanBoQuanLy || "",
    facility.noiDungCanLuuY || "",
    facility.createdAt || "",
    facility.updatedAt || "",
  ];
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
```

## 3. Deploy Web App

1. Chọn `Deploy` > `New deployment`.
2. Chọn loại deployment là `Web app`.
3. `Execute as`: chọn tài khoản của bạn.
4. `Who has access`: chọn `Anyone` hoặc quyền phù hợp với môi trường sử dụng.
5. Bấm `Deploy`.
6. Copy URL kết thúc bằng `/exec`.

## 4. Dán URL vào app

Mở file `app.js`, tìm dòng:

```javascript
const GOOGLE_SHEET_ENDPOINT = "";
```

Dán URL `/exec` vào giữa dấu nháy:

```javascript
const GOOGLE_SHEET_ENDPOINT = "https://script.google.com/macros/s/DEPLOYMENT_ID/exec";
```

Sau khi cấu hình:

1. Mở lại app.
2. Tạo hoặc cập nhật một hồ sơ.
3. App sẽ lưu local trước, sau đó gửi Google Sheet.
4. Nếu gửi lỗi, dữ liệu local vẫn được giữ và danh sách sẽ hiển thị badge lỗi đồng bộ.

## 5. Ghi chú

- Frontend dùng `fetch(..., { mode: "no-cors" })`, nên không đọc response từ Apps Script.
- Nếu request không throw lỗi, app coi là đã gửi.
- Apps Script vẫn trả JSON để dễ kiểm tra bằng công cụ khác khi cần.
- Upsert dựa trên `soTaiKhoanPCCC`; nếu trường này trống, script sẽ append dòng mới.
