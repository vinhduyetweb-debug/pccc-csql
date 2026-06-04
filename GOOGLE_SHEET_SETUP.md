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
  "id",
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
  "Trạng thái bản ghi",
  "Thời gian đồng bộ",
];

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || "{}");
    const sheet = getOrCreateSheet();
    ensureHeader(sheet);

    if (payload.action === "UPSERT_FACILITY") {
      return handleUpsert(sheet, payload);
    }

    if (payload.action === "DELETE_FACILITY") {
      return handleDelete(sheet, payload);
    }

    return jsonResponse({ success: false, message: "Action không hợp lệ" });
  } catch (error) {
    return jsonResponse({ success: false, message: error.message });
  }
}

function handleUpsert(sheet, payload) {
  const facility = payload.facility || {};
  const row = findRowByKeys(sheet, facility);
  const rowValues = buildRow(payload.syncedAt, facility, "active");

  if (row > 0) {
    sheet.getRange(row, 1, 1, HEADERS.length).setValues([rowValues]);
    return jsonResponse({ success: true, action: "updated", row });
  }

  sheet.appendRow(rowValues);
  return jsonResponse({ success: true, action: "created", row: sheet.getLastRow() });
}

function handleDelete(sheet, payload) {
  const facility = payload.facility || {};
  const row = findRowByKeys(sheet, facility);

  if (row > 0) {
    sheet.deleteRow(row);
    return jsonResponse({ success: true, action: "deleted", row });
  }

  return jsonResponse({ success: false, message: "Không tìm thấy dòng cần xóa" });
}

function getOrCreateSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
}

function ensureHeader(sheet) {
  const width = Math.max(sheet.getLastColumn(), HEADERS.length);
  const current = sheet.getRange(1, 1, 1, width).getValues()[0].map((value) => String(value || "").trim());
  const currentCanonical = current.slice(0, HEADERS.length).join("||");
  const expectedCanonical = HEADERS.join("||");

  if (currentCanonical !== expectedCanonical) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
}

function findRowByKeys(sheet, facility) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const headers = getHeaders(sheet);
  const idCol = headers.indexOf("id");
  const maCoSoCol = headers.indexOf("Mã cơ sở");
  const accountCol = headers.indexOf("Số tài khoản PCCC");
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

  const keys = [
    { col: idCol, value: facility.id },
    { col: maCoSoCol, value: facility.maCoSo },
    { col: accountCol, value: facility.soTaiKhoanPCCC },
  ];

  for (const key of keys) {
    const target = String(key.value || "").trim();
    if (key.col < 0 || !target) continue;

    const index = values.findIndex((row) => String(row[key.col] || "").trim() === target);
    if (index >= 0) return index + 2;
  }

  return -1;
}

function getHeaders(sheet) {
  return sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map((value) => String(value || "").trim());
}

function buildRow(syncedAt, facility, status) {
  return [
    facility.id || "",
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
    status || "active",
    syncedAt || new Date().toISOString(),
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

## 5. Lưu ý khi cập nhật từ script cũ

- Script mới không dùng append mù.
- Script mới update/delete theo thứ tự ưu tiên: `id`, `maCoSo`, `soTaiKhoanPCCC`.
- Nếu sheet cũ chưa có cột `id` hoặc header đang theo mẫu cũ, script sẽ chuẩn hóa hàng header về đúng thứ tự mới. Dữ liệu cũ chưa có giá trị `id`, nên lần đồng bộ đầu tiên có thể cần dựa vào `maCoSo` hoặc `soTaiKhoanPCCC`.
- Nếu sheet cũ có nhiều dòng trùng từ script cũ, nên rà soát/xóa dòng trùng thủ công một lần trước khi dùng lâu dài.
- Frontend dùng `fetch(..., { mode: "no-cors" })`, nên không đọc response từ Apps Script. Nếu request không throw lỗi, app coi là đã gửi yêu cầu.
