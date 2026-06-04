const STORAGE_KEY = "pccc-csql.facilities.v1";
const DEFAULT_CITY = "TP. Hồ Chí Minh";

const WARDS = ["Phú Định", "Hưng Phú", "Bình Đông"];

const FIELDS = [
  "soHoSo",
  "soTaiKhoanPCCC",
  "tenCoSo",
  "diaChi",
  "phuongXa",
  "thanhPho",
  "linhVucCode",
  "tinhChatHoatDong",
  "thangNamHoatDong",
  "soTang",
  "khoiTich",
  "tongDienTichSan",
  "thongTinDacThu",
  "nhomQuanLy",
  "tinhTrangThamDuyetNghiemThu",
  "viPhamXayDung",
  "viPhamDangKyKinhDoanh",
  "viPhamPCCC",
  "soQuyetDinhDinhChi",
  "ngayQuyetDinhDinhChi",
  "nguoiDungDau",
  "sdtNguoiDungDau",
  "nguoiThuongTruc",
  "sdtNguoiThuongTruc",
  "donViQuanLy",
  "canBoQuanLy",
  "sdtCanBoQuanLy",
  "noiDungCanLuuY",
];

const NUMBER_FIELDS = ["soTang", "khoiTich", "tongDienTichSan"];
const CHECKBOX_FIELDS = ["viPhamXayDung", "viPhamDangKyKinhDoanh", "viPhamPCCC"];
const PHONE_FIELDS = ["sdtNguoiDungDau", "sdtNguoiThuongTruc", "sdtCanBoQuanLy"];

const SECTOR_OPTIONS = [
  ["I", "I - Lĩnh vực nhà ở, trụ sở làm việc, văn phòng, nhà đa năng"],
  ["II", "II - Lĩnh vực cơ sở giáo dục"],
  ["III", "III - Lĩnh vực cơ sở y tế"],
  ["IV", "IV - Lĩnh vực cơ sở văn hóa, thể thao, du lịch"],
  ["V", "V - Lĩnh vực cơ sở thương mại dịch vụ"],
  ["VI", "VI - Lĩnh vực cơ sở giao thông vận tải"],
  ["VII", "VII - Lĩnh vực cơ sở xăng dầu, dầu khí"],
  ["VIII", "VIII - Lĩnh vực cơ sở năng lượng"],
  ["IX", "IX - Lĩnh vực cơ sở hóa chất"],
  ["X", "X - Lĩnh vực cơ sở công nghiệp, nhà kho"],
];

const ACTIVITY_OPTIONS = ["Đang hoạt động", "Tạm ngừng", "Chưa hoạt động"];
const MANAGEMENT_GROUPS = [
  "Ngoài Phụ lục II",
  "Nhóm 1 - Phụ lục II",
  "Nhóm 2 - Phụ lục II",
  "Phụ lục III",
  "Phụ lục VII",
];
const APPROVAL_STATUS_OPTIONS = [
  "Đã thẩm duyệt, nghiệm thu",
  "Đã thẩm duyệt, chưa nghiệm thu",
  "Chưa được thẩm duyệt, nghiệm thu",
];

const EXCEL_COLUMN_ALIASES = {
  soHoSo: ["sohoso", "sohos", "so ho so", "hoso", "ho so", "so ho so doi voi co so cua pc07 quan ly"],
  soTaiKhoanPCCC: [
    "sotaikhoanpccc",
    "so tai khoan pccc",
    "so tai khoan csdl quoc gia pccc",
    "so tai khoan csdl quoc gia ve pccc",
    "so tai khoan tren he thong csdl quoc gia ve pccc",
    "tai khoan pccc",
  ],
  tenCoSo: ["tencoso", "ten co so"],
  diaChi: ["diachi", "dia chi", "so nha ten duong", "dia chi so nha ten duong"],
  phuongXa: ["phuongxa", "phuong xa", "ten phuong xa"],
  thanhPho: ["thanhpho", "thanh pho", "tp", "tinh thanh"],
  linhVucCode: ["linhvuccode", "ma linh vuc", "linh vuc code", "code linh vuc"],
  linhVucName: ["linhvucname", "linh vuc", "ten linh vuc", "thuoc linh vuc"],
  tinhChatHoatDong: ["tinhchathoatdong", "tinh chat hoat dong"],
  thangNamHoatDong: ["thangnamhoatdong", "thang nam hoat dong", "thoi gian hoat dong", "thang nam dua vao hoat dong"],
  soTang: ["sotang", "so tang", "quy mo so tang"],
  khoiTich: ["khoitich", "khoi tich"],
  tongDienTichSan: ["tongdientichsan", "tong dien tich san", "dien tich san"],
  thongTinDacThu: ["thongtindacthu", "thong tin dac thu", "thong tin dac thu khac", "dac thu"],
  nhomQuanLy: ["nhomquanly", "nhom quan ly"],
  tinhTrangThamDuyetNghiemThu: [
    "tinhtrangthamdungnghiemthu",
    "tinhtrangthamduyetnghiemthu",
    "tinh trang tham duyet nghiem thu",
    "tinh trang tham duyet",
  ],
  viPhamXayDung: ["viphamxaydung", "vi pham xay dung", "vi pham ve xay dung"],
  viPhamDangKyKinhDoanh: [
    "viphamdangkykinhdoanh",
    "vi pham dang ky kinh doanh",
    "vi pham ve dang ky kinh doanh",
  ],
  viPhamPCCC: ["viphampccc", "vi pham pccc", "vi pham ve pccc"],
  soQuyetDinhDinhChi: ["soquyetdinhdinhchi", "so quyet dinh dinh chi", "dinh chi hoat dong so quyet dinh"],
  ngayQuyetDinhDinhChi: ["ngayquyetdinhdinhchi", "ngay quyet dinh dinh chi", "ngay ban hanh quyet dinh"],
  donViQuanLy: ["donviquanly", "don vi quan ly", "thong tin don vi quan ly don vi quan ly"],
  canBoQuanLy: ["canboquanly", "can bo quan ly", "ten can bo quan ly"],
  sdtCanBoQuanLy: ["sdtcanboquanly", "sdt can bo quan ly", "so dien thoai can bo quan ly", "so dien thoai lien he"],
  nguoiDungDau: ["nguoidungdau", "nguoi dung dau", "chu co so", "dai dien"],
  sdtNguoiDungDau: ["sdtnguoidungdau", "sdt nguoi dung dau", "sdt nguoi dung dau co so", "so dien thoai nguoi dung dau"],
  nguoiThuongTruc: ["nguoithuongtruc", "nguoi thuong truc"],
  sdtNguoiThuongTruc: [
    "sdtnguoithuongtruc",
    "sdt nguoi thuong truc",
    "so dien thoai thuong truc",
    "so dien nguoi thuong xuyen co mat",
  ],
  noiDungCanLuuY: ["noidungcanluuy", "noi dung can luu y", "ghi chu", "luu y"],
};

const form = document.querySelector("#facilityForm");
const tableBody = document.querySelector("#facilityTable");
const recordCount = document.querySelector("#recordCount");
const searchInput = document.querySelector("#searchInput");
const messageBox = document.querySelector("#messageBox");
const duplicateBox = document.querySelector("#duplicateBox");
const formTitle = document.querySelector("#formTitle");
const formModeText = document.querySelector("#formModeText");
const submitBtn = document.querySelector("#submitBtn");
const cancelEditBtn = document.querySelector("#cancelEditBtn");
const resetBtn = document.querySelector("#resetBtn");
const exportBtn = document.querySelector("#exportBtn");
const exportCsvBtn = document.querySelector("#exportCsvBtn");
const exportBackupBtn = document.querySelector("#exportBackupBtn");
const importInput = document.querySelector("#importInput");
const excelInput = document.querySelector("#excelInput");
const previewExcelBtn = document.querySelector("#previewExcelBtn");
const updateDuplicatesInput = document.querySelector("#updateDuplicatesInput");
const confirmExcelImportBtn = document.querySelector("#confirmExcelImportBtn");
const cancelExcelImportBtn = document.querySelector("#cancelExcelImportBtn");
const excelPreview = document.querySelector("#excelPreview");
const statTotal = document.querySelector("#statTotal");
const wardStats = document.querySelector("#wardStats");
const sectorStats = document.querySelector("#sectorStats");
const activityStats = document.querySelector("#activityStats");
const approvalStats = document.querySelector("#approvalStats");
const statMissing = document.querySelector("#statMissing");
const statViolation = document.querySelector("#statViolation");
const statSuspended = document.querySelector("#statSuspended");
const filterWard = document.querySelector("#filterWard");
const filterSector = document.querySelector("#filterSector");
const filterApproval = document.querySelector("#filterApproval");
const filterFlag = document.querySelector("#filterFlag");

let facilities = loadFacilities();
let editingId = null;
let selectedExcelFile = null;
let pendingImportPreview = null;

function fillSelect(selectId, options, placeholder = "") {
  const select = document.querySelector(`#${selectId}`);
  select.innerHTML = "";
  if (placeholder) {
    select.append(new Option(placeholder, ""));
  }
  options.forEach((option) => {
    if (Array.isArray(option)) {
      select.append(new Option(option[1], option[0]));
    } else {
      select.append(new Option(option, option));
    }
  });
}

function fillFilterSelect(select, options, placeholder) {
  select.innerHTML = "";
  select.append(new Option(placeholder, ""));
  options.forEach((option) => {
    if (Array.isArray(option)) {
      select.append(new Option(option[1], option[0]));
    } else {
      select.append(new Option(option, option));
    }
  });
}

function loadFacilities() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function saveFacilities() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(facilities));
}

function normalizeText(value) {
  return String(value ?? "").trim().replace(/\s+/g, " ");
}

function normalizeFacilityName(value) {
  return normalizeText(value).toLocaleUpperCase("vi-VN");
}

function normalizeHeader(value) {
  return normalizeText(value)
    .toLocaleLowerCase("vi-VN")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function toNumberOrEmpty(value) {
  if (value === "" || value === null || value === undefined) return "";
  const normalized = String(value).replace(",", ".").trim();
  if (!normalized) return "";
  return Number(normalized);
}

function normalizeDigits(value) {
  return String(value ?? "").replace(/\D/g, "");
}

function normalizeBoolean(value) {
  const text = normalizeHeader(value);
  return ["1", "x", "co", "yes", "true", "vi pham", "co vi pham"].includes(text);
}

function normalizeExcelDate(value) {
  const text = normalizeText(value);
  if (!text) return "";
  const date = new Date(text);
  if (!Number.isNaN(date.getTime())) return date.toISOString().slice(0, 10);

  const match = /^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/.exec(text);
  if (!match) return text;

  const [, day, month, year] = match;
  return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function getSectorName(code) {
  return SECTOR_OPTIONS.find(([sectorCode]) => sectorCode === code)?.[1] ?? "";
}

function getNextFacilityCode() {
  return getNextFacilityCodeFrom(facilities);
}

function getNextFacilityCodeFrom(records) {
  const maxCode = records.reduce((max, item) => {
    const match = /^CS(\d{6})$/.exec(item.maCoSo || "");
    return match ? Math.max(max, Number(match[1])) : max;
  }, 0);
  return `CS${String(maxCode + 1).padStart(6, "0")}`;
}

function createId() {
  if (crypto.randomUUID) return crypto.randomUUID();
  return `id-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function getFormData() {
  const data = {};
  FIELDS.forEach((field) => {
    const input = form.elements[field];
    if (!input) return;

    if (CHECKBOX_FIELDS.includes(field)) {
      data[field] = input.checked;
    } else if (NUMBER_FIELDS.includes(field)) {
      data[field] = toNumberOrEmpty(input.value);
    } else if (field === "tenCoSo") {
      data[field] = normalizeFacilityName(input.value);
    } else if (field === "soTaiKhoanPCCC" || PHONE_FIELDS.includes(field)) {
      data[field] = normalizeDigits(input.value);
    } else {
      data[field] = normalizeText(input.value);
    }
  });

  data.thanhPho = DEFAULT_CITY;
  data.linhVucName = getSectorName(data.linhVucCode);
  return data;
}

function validateFacility(data, currentId = null) {
  const errors = [];
  const phonePattern = /^0\d{9}$/;

  if (data.soTaiKhoanPCCC && !/^\d{13}$/.test(data.soTaiKhoanPCCC)) {
    errors.push("Số tài khoản CSDL Quốc gia về PCCC phải gồm đúng 13 chữ số.");
  }

  if (data.tenCoSo.length < 5) {
    errors.push("Tên cơ sở phải từ 5 ký tự trở lên.");
  }

  if (data.diaChi.length < 5) {
    errors.push("Địa chỉ chưa đầy đủ.");
  }

  PHONE_FIELDS.forEach((field) => {
    if (data[field] && !phonePattern.test(data[field])) {
      errors.push("Số điện thoại phải gồm đúng 10 chữ số và bắt đầu bằng 0.");
    }
  });

  if (data.soTang !== "" && (!Number.isInteger(data.soTang) || data.soTang < 0)) {
    errors.push("Số tầng phải là số nguyên không âm.");
  }

  ["khoiTich", "tongDienTichSan"].forEach((field) => {
    if (data[field] !== "" && (!Number.isFinite(data[field]) || data[field] < 0)) {
      errors.push("Diện tích và khối tích phải là số không âm.");
    }
  });

  const duplicate = data.soTaiKhoanPCCC
    ? facilities.find((item) => item.soTaiKhoanPCCC === data.soTaiKhoanPCCC && item.id !== currentId)
    : null;

  return { errors: [...new Set(errors)], duplicate };
}

function validateImportedFacility(data, rowNumber, existingRecords = facilities) {
  const errors = [];
  const phonePattern = /^0\d{9}$/;

  if (!data.tenCoSo) {
    errors.push(`Dòng ${rowNumber}: Thiếu tên cơ sở`);
  } else if (data.tenCoSo.length < 5) {
    errors.push(`Dòng ${rowNumber}: Tên cơ sở phải từ 5 ký tự trở lên`);
  }

  if (!data.diaChi || data.diaChi.length < 5) {
    errors.push(`Dòng ${rowNumber}: Địa chỉ chưa đầy đủ`);
  }

  if (data.soTaiKhoanPCCC && !/^\d{13}$/.test(data.soTaiKhoanPCCC)) {
    errors.push(`Dòng ${rowNumber}: Số tài khoản PCCC không đúng 13 chữ số`);
  }

  PHONE_FIELDS.forEach((field) => {
    if (data[field] && !phonePattern.test(data[field])) {
      const labels = {
        sdtNguoiDungDau: "SĐT người đứng đầu",
        sdtNguoiThuongTruc: "SĐT người thường trực",
        sdtCanBoQuanLy: "SĐT cán bộ quản lý",
      };
      errors.push(`Dòng ${rowNumber}: ${labels[field]} không hợp lệ`);
    }
  });

  if (data.soTang !== "" && (!Number.isInteger(data.soTang) || data.soTang < 0)) {
    errors.push(`Dòng ${rowNumber}: Số tầng phải là số nguyên không âm`);
  }

  ["khoiTich", "tongDienTichSan"].forEach((field) => {
    if (data[field] !== "" && (!Number.isFinite(data[field]) || data[field] < 0)) {
      errors.push(`Dòng ${rowNumber}: Diện tích và khối tích phải là số không âm`);
    }
  });

  const duplicate = data.soTaiKhoanPCCC
    ? existingRecords.find((record) => record.soTaiKhoanPCCC === data.soTaiKhoanPCCC)
    : null;
  return { errors: [...new Set(errors)], duplicate };
}

function setMessage(text, type = "success") {
  messageBox.textContent = text;
  messageBox.className = `message ${type}`;
}

function clearMessage() {
  messageBox.textContent = "";
  messageBox.className = "message hidden";
}

function showDuplicate(record) {
  duplicateBox.className = "duplicate";
  duplicateBox.innerHTML = `
    <h3>Đã tồn tại cơ sở có số tài khoản PCCC này.</h3>
    <p><strong>Tên cơ sở:</strong> ${escapeHtml(record.tenCoSo)}</p>
    <p><strong>Địa chỉ:</strong> ${escapeHtml(record.diaChi)}</p>
    <p><strong>Phường/Xã:</strong> ${escapeHtml(record.phuongXa)}</p>
    <div class="duplicate-actions">
      <button type="button" data-action="view-duplicate" data-id="${record.id}">Xem hồ sơ</button>
      <button type="button" data-action="edit-duplicate" data-id="${record.id}" class="secondary">Cập nhật hồ sơ</button>
      <button type="button" data-action="close-duplicate" class="secondary">Hủy</button>
    </div>
  `;
}

function hideDuplicate() {
  duplicateBox.className = "duplicate hidden";
  duplicateBox.innerHTML = "";
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function formatDateTime(value) {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  return new Intl.DateTimeFormat("vi-VN", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  }).format(date);
}

function getTimestampForFileName() {
  const now = new Date();
  const pad = (value) => String(value).padStart(2, "0");
  return [
    now.getFullYear(),
    pad(now.getMonth() + 1),
    pad(now.getDate()),
    "_",
    pad(now.getHours()),
    pad(now.getMinutes()),
  ].join("");
}

function hasMissingRequiredListData(record) {
  return !record.nguoiDungDau || !record.sdtNguoiDungDau || !record.nhomQuanLy || !record.tinhTrangThamDuyetNghiemThu;
}

function isSuspended(record) {
  return Boolean(record.soQuyetDinhDinhChi || record.ngayQuyetDinhDinhChi);
}

function renderDashboard() {
  statTotal.textContent = facilities.length;
  statMissing.textContent = facilities.filter(hasMissingRequiredListData).length;
  statViolation.textContent = facilities.filter((record) => record.viPhamPCCC).length;
  statSuspended.textContent = facilities.filter(isSuspended).length;

  wardStats.innerHTML = WARDS.map((ward) => {
    const count = facilities.filter((record) => record.phuongXa === ward).length;
    return `
      <div class="stat-card">
        <span class="stat-label">${escapeHtml(ward)}</span>
        <strong>${count}</strong>
      </div>
    `;
  }).join("");

  sectorStats.innerHTML = SECTOR_OPTIONS.map(([code]) => {
    const count = facilities.filter((record) => record.linhVucCode === code).length;
    return `
      <div class="stat-card">
        <span class="stat-label">${escapeHtml(code)}</span>
        <strong>${count}</strong>
      </div>
    `;
  }).join("");

  activityStats.innerHTML = ACTIVITY_OPTIONS.map((status) => {
    const count = facilities.filter((record) => record.tinhChatHoatDong === status).length;
    return `
      <div class="stat-card">
        <span class="stat-label">${escapeHtml(status)}</span>
        <strong>${count}</strong>
      </div>
    `;
  }).join("");

  approvalStats.innerHTML = APPROVAL_STATUS_OPTIONS.map((status) => {
    const count = facilities.filter((record) => record.tinhTrangThamDuyetNghiemThu === status).length;
    return `
      <div class="stat-card">
        <span class="stat-label">${escapeHtml(status)}</span>
        <strong>${count}</strong>
      </div>
    `;
  }).join("");
}

function resetForm() {
  form.reset();
  form.elements.thanhPho.value = DEFAULT_CITY;
  document.querySelector("#maCoSo").value = "";
  editingId = null;
  formTitle.textContent = "Thêm hồ sơ cơ sở";
  formModeText.textContent = "Mã cơ sở sẽ được tự sinh khi lưu.";
  submitBtn.textContent = "Lưu hồ sơ";
  cancelEditBtn.classList.add("hidden");
  clearMessage();
  hideDuplicate();
}

function editFacility(id) {
  const record = facilities.find((item) => item.id === id);
  if (!record) return;

  editingId = id;
  FIELDS.forEach((field) => {
    const input = form.elements[field];
    if (!input) return;
    if (CHECKBOX_FIELDS.includes(field)) {
      input.checked = Boolean(record[field]);
    } else {
      input.value = record[field] ?? "";
    }
  });

  document.querySelector("#maCoSo").value = record.maCoSo;
  formTitle.textContent = "Cập nhật hồ sơ cơ sở";
  formModeText.textContent = `Đang cập nhật ${record.maCoSo}. Giữ nguyên ngày tạo ban đầu.`;
  submitBtn.textContent = "Cập nhật hồ sơ";
  cancelEditBtn.classList.remove("hidden");
  clearMessage();
  hideDuplicate();
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function deleteFacility(id) {
  const record = facilities.find((item) => item.id === id);
  if (!record) return;

  const confirmed = confirm(`Xóa hồ sơ "${record.tenCoSo}"? Thao tác này không thể hoàn tác.`);
  if (!confirmed) return;

  facilities = facilities.filter((item) => item.id !== id);
  saveFacilities();
  renderFacilities();
  if (editingId === id) resetForm();
  setMessage("Đã xóa hồ sơ.", "success");
}

function viewFacility(id) {
  const row = document.querySelector(`[data-row-id="${id}"]`);
  if (row) {
    row.scrollIntoView({ behavior: "smooth", block: "center" });
    row.classList.add("highlight");
    setTimeout(() => row.classList.remove("highlight"), 1400);
  }
}

function getSearchableText(record) {
  return [
    record.tenCoSo,
    record.diaChi,
    record.phuongXa,
    record.soTaiKhoanPCCC,
    record.sdtNguoiDungDau,
    record.sdtNguoiThuongTruc,
    record.sdtCanBoQuanLy,
    record.nguoiDungDau,
  ]
    .join(" ")
    .toLowerCase();
}

function matchesFilters(record) {
  if (filterWard.value && record.phuongXa !== filterWard.value) return false;
  if (filterSector.value && record.linhVucCode !== filterSector.value) return false;
  if (filterApproval.value && record.tinhTrangThamDuyetNghiemThu !== filterApproval.value) return false;
  if (filterFlag.value === "viPhamPCCC" && !record.viPhamPCCC) return false;
  if (filterFlag.value === "missing" && !hasMissingRequiredListData(record)) return false;
  return true;
}

function renderFacilities() {
  const query = normalizeText(searchInput.value).toLowerCase();
  const filtered = facilities.filter((record) => getSearchableText(record).includes(query) && matchesFilters(record));

  recordCount.textContent = facilities.length;
  renderDashboard();
  tableBody.innerHTML = "";

  if (!filtered.length) {
    tableBody.innerHTML = `
      <tr>
        <td colspan="9" class="muted">Chưa có hồ sơ phù hợp.</td>
      </tr>
    `;
    return;
  }

  filtered.forEach((record) => {
    const phone = record.sdtNguoiDungDau || record.sdtNguoiThuongTruc || record.sdtCanBoQuanLy || "";
    const missingBadge = hasMissingRequiredListData(record) ? `<span class="missing-badge">⚠ Thiếu dữ liệu</span>` : "";
    const row = document.createElement("tr");
    row.dataset.rowId = record.id;
    row.innerHTML = `
      <td>${escapeHtml(record.maCoSo)}</td>
      <td>${escapeHtml(record.soTaiKhoanPCCC)}</td>
      <td>
        <strong>${escapeHtml(record.tenCoSo)}</strong><br />
        ${missingBadge}
      </td>
      <td>${escapeHtml(record.diaChi)}</td>
      <td>${escapeHtml(record.phuongXa)}</td>
      <td>${escapeHtml(record.linhVucCode)}<br /><span class="muted">${escapeHtml(record.nhomQuanLy)}</span></td>
      <td>${escapeHtml(record.nguoiDungDau)}</td>
      <td>${escapeHtml(phone)}</td>
      <td>
        <div class="row-actions">
          <button type="button" data-action="edit" data-id="${record.id}">Sửa</button>
          <button type="button" data-action="delete" data-id="${record.id}" class="danger">Xóa</button>
        </div>
      </td>
    `;
    tableBody.append(row);
  });
}

function handleSubmit(event) {
  event.preventDefault();
  clearMessage();
  hideDuplicate();

  const data = getFormData();
  const isEditing = Boolean(editingId);
  const { errors, duplicate } = validateFacility(data, editingId);

  if (errors.length) {
    setMessage(errors.join(" "), "error");
    return;
  }

  if (duplicate) {
    showDuplicate(duplicate);
    return;
  }

  const now = new Date().toISOString();
  if (isEditing) {
    facilities = facilities.map((item) =>
      item.id === editingId
        ? {
            ...item,
            ...data,
            createdAt: item.createdAt,
            updatedAt: now,
          }
        : item,
    );
    setMessage("Đã cập nhật hồ sơ.", "success");
  } else {
    facilities.push({
      id: createId(),
      maCoSo: getNextFacilityCode(),
      ...data,
      createdAt: now,
      updatedAt: now,
    });
    setMessage("Đã lưu hồ sơ mới.", "success");
  }

  saveFacilities();
  renderFacilities();
  resetForm();
  setMessage(isEditing ? "Đã cập nhật hồ sơ." : "Đã lưu hồ sơ mới.", "success");
}

function exportJson() {
  const blob = new Blob([JSON.stringify(facilities, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `pccc-csql-${new Date().toISOString().slice(0, 10)}.json`;
  link.click();
  URL.revokeObjectURL(url);
}

function exportBackupJson() {
  const backup = {
    appName: "PCCC-CSQL",
    version: "FINAL LOCAL PRO",
    exportedAt: new Date().toISOString(),
    total: facilities.length,
    facilities,
  };
  const blob = new Blob([JSON.stringify(backup, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `PCCC_CSQL_BACKUP_${getTimestampForFileName()}.json`;
  link.click();
  URL.revokeObjectURL(url);
}

function protectCsvCell(value) {
  const text = String(value ?? "");
  return /^[=+\-@]/.test(text) ? `'${text}` : text;
}

function csvEscape(value) {
  const text = protectCsvCell(value);
  return `"${text.replace(/"/g, '""')}"`;
}

function exportCsv() {
  const headers = [
    "STT",
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
    "Vi phạm xây dựng",
    "Vi phạm đăng ký kinh doanh",
    "Vi phạm PCCC",
    "Số quyết định đình chỉ",
    "Ngày quyết định đình chỉ",
    "Người đứng đầu",
    "SĐT người đứng đầu",
    "Người thường trực",
    "SĐT thường trực",
    "Đơn vị quản lý",
    "Cán bộ quản lý",
    "SĐT cán bộ quản lý",
    "Ngày tạo",
    "Ngày cập nhật",
  ];

  const rows = facilities.map((record, index) => [
    index + 1,
    record.maCoSo,
    record.soHoSo,
    record.soTaiKhoanPCCC,
    record.tenCoSo,
    record.diaChi,
    record.phuongXa,
    record.thanhPho,
    record.linhVucName || getSectorName(record.linhVucCode),
    record.tinhChatHoatDong,
    record.thangNamHoatDong,
    record.soTang,
    record.khoiTich,
    record.tongDienTichSan,
    record.nhomQuanLy,
    record.tinhTrangThamDuyetNghiemThu,
    record.viPhamXayDung ? "Có" : "",
    record.viPhamDangKyKinhDoanh ? "Có" : "",
    record.viPhamPCCC ? "Có" : "",
    record.soQuyetDinhDinhChi,
    record.ngayQuyetDinhDinhChi,
    record.nguoiDungDau,
    record.sdtNguoiDungDau,
    record.nguoiThuongTruc,
    record.sdtNguoiThuongTruc,
    record.donViQuanLy,
    record.canBoQuanLy,
    record.sdtCanBoQuanLy,
    formatDateTime(record.createdAt),
    formatDateTime(record.updatedAt),
  ]);

  const csv = [headers, ...rows].map((row) => row.map(csvEscape).join(",")).join("\r\n");
  const blob = new Blob([`\uFEFF${csv}`], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `PCCC_CSQL_${getTimestampForFileName()}.csv`;
  link.click();
  URL.revokeObjectURL(url);
}

function findHeaderKey(headers, aliases) {
  const normalizedAliases = aliases.map(normalizeHeader);
  return (
    headers.find((header) => normalizedAliases.includes(normalizeHeader(header))) ||
    headers.find((header) => {
      const normalizedHeader = normalizeHeader(header);
      return normalizedAliases.some((alias) => alias.length >= 12 && normalizedHeader.includes(alias));
    }) ||
    ""
  );
}

function getExcelCell(row, headers, field) {
  const header = findHeaderKey(headers, [field, ...(EXCEL_COLUMN_ALIASES[field] || [])]);
  return header ? row[header] : "";
}

function getMarkedOption(row, headers, options) {
  const matchedOption = options.find((option) => {
    const header = findHeaderKey(headers, [option]);
    return header && normalizeBoolean(row[header]);
  });
  return matchedOption || "";
}

function matchOption(value, options) {
  const text = normalizeHeader(value);
  if (!text) return "";
  return options.find((option) => normalizeHeader(option) === text) || "";
}

function normalizeSectorCode(codeValue, nameValue) {
  const rawCode = normalizeText(codeValue).toUpperCase();
  if (SECTOR_OPTIONS.some(([code]) => code === rawCode)) return rawCode;

  const combined = normalizeText(`${codeValue} ${nameValue}`).toUpperCase();
  const romanMatch = /\b(I|II|III|IV|V|VI|VII|VIII|IX|X)\b/.exec(combined);
  if (romanMatch && SECTOR_OPTIONS.some(([code]) => code === romanMatch[1])) {
    return romanMatch[1];
  }

  const normalizedName = normalizeHeader(nameValue);
  const option = SECTOR_OPTIONS.find(([, label]) => normalizeHeader(label).includes(normalizedName) && normalizedName);
  return option?.[0] || "";
}

function mapExcelRowToFacility(row, headers) {
  const data = {};

  FIELDS.forEach((field) => {
    const value = getExcelCell(row, headers, field);
    if (CHECKBOX_FIELDS.includes(field)) {
      data[field] = normalizeBoolean(value);
    } else if (NUMBER_FIELDS.includes(field)) {
      data[field] = toNumberOrEmpty(value);
    } else if (field === "tenCoSo") {
      data[field] = normalizeFacilityName(value);
    } else if (field === "soTaiKhoanPCCC" || PHONE_FIELDS.includes(field)) {
      data[field] = normalizeDigits(value);
    } else if (field === "ngayQuyetDinhDinhChi") {
      data[field] = normalizeExcelDate(value);
    } else {
      data[field] = normalizeText(value);
    }
  });

  data.nhomQuanLy = data.nhomQuanLy || getMarkedOption(row, headers, MANAGEMENT_GROUPS);
  data.tinhTrangThamDuyetNghiemThu =
    data.tinhTrangThamDuyetNghiemThu || getMarkedOption(row, headers, APPROVAL_STATUS_OPTIONS);

  data.phuongXa = matchOption(data.phuongXa, WARDS) || data.phuongXa;
  data.thanhPho = normalizeText(data.thanhPho) || DEFAULT_CITY;
  data.tinhChatHoatDong = matchOption(data.tinhChatHoatDong, ACTIVITY_OPTIONS) || data.tinhChatHoatDong;
  data.nhomQuanLy = matchOption(data.nhomQuanLy, MANAGEMENT_GROUPS) || data.nhomQuanLy;
  data.tinhTrangThamDuyetNghiemThu =
    matchOption(data.tinhTrangThamDuyetNghiemThu, APPROVAL_STATUS_OPTIONS) || data.tinhTrangThamDuyetNghiemThu;
  data.linhVucCode = normalizeSectorCode(data.linhVucCode, data.linhVucName);
  data.linhVucName = data.linhVucName || getSectorName(data.linhVucCode);

  return data;
}

async function parseExcelFile(file) {
  if (!window.XLSX) {
    throw new Error("Chưa tải được thư viện SheetJS/xlsx. Kiểm tra kết nối mạng rồi thử lại.");
  }

  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) {
    throw new Error("File Excel không có sheet dữ liệu.");
  }

  const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], {
    header: 1,
    defval: "",
    raw: false,
  });

  const table = resolveExcelTable(matrix);
  return table.rows.map((row, index) => ({
    rowNumber: table.dataStartRowNumber + index,
    data: mapExcelRowToFacility(row, table.headers),
  }));
}

function isNumberingRow(row) {
  const filled = row.filter((value) => normalizeText(value));
  if (filled.length < 5) return false;
  const numeric = filled.filter((value) => /^\d+$/.test(normalizeText(value)));
  return numeric.length / filled.length > 0.8;
}

function combineHeaderRows(primary = [], secondary = []) {
  const length = Math.max(primary.length, secondary.length);
  return Array.from({ length }, (_, index) => {
    const top = normalizeText(primary[index]);
    const bottom = normalizeText(secondary[index]);
    if (top && bottom && normalizeHeader(top) !== normalizeHeader(bottom)) return `${top} ${bottom}`;
    return top || bottom || `Cột ${index + 1}`;
  });
}

function resolveExcelTable(matrix) {
  const nonEmptyRows = matrix
    .map((row, index) => ({ row, index }))
    .filter(({ row }) => row.some((value) => normalizeText(value)));

  const headerIndex =
    nonEmptyRows.find(({ row }) => {
      const normalized = row.map(normalizeHeader).join(" ");
      return normalized.includes("ten co so") && normalized.includes("so tai khoan");
    })?.index ?? 0;

  const subHeaderIndex = headerIndex + 1 < matrix.length ? headerIndex + 1 : headerIndex;
  const numberingIndex = matrix.findIndex((row, index) => index > headerIndex && isNumberingRow(row));
  const dataStartIndex = numberingIndex >= 0 ? numberingIndex + 1 : subHeaderIndex + 1;
  const headers = combineHeaderRows(matrix[headerIndex] || [], matrix[subHeaderIndex] || []);

  const rows = matrix
    .slice(dataStartIndex)
    .filter((row) => row.some((value) => normalizeText(value)))
    .map((row) =>
      headers.reduce((record, header, index) => {
        record[header] = row[index] ?? "";
        return record;
      }, {}),
    );

  return {
    headers,
    rows,
    dataStartRowNumber: dataStartIndex + 1,
  };
}

function buildImportPreview(mappedRows, shouldUpdateDuplicates) {
  const seenAccounts = new Set();
  const errors = [];
  const errorRows = [];
  const validRows = [];
  const duplicateRows = [];
  let validCount = 0;

  mappedRows.forEach((item) => {
    const validation = validateImportedFacility(item.data, item.rowNumber);
    const rowErrors = [...validation.errors];

    if (item.data.soTaiKhoanPCCC && seenAccounts.has(item.data.soTaiKhoanPCCC)) {
      rowErrors.push(`Dòng ${item.rowNumber}: Số tài khoản PCCC trùng trong file import`);
    }
    if (item.data.soTaiKhoanPCCC) {
      seenAccounts.add(item.data.soTaiKhoanPCCC);
    }

    const previewItem = {
      ...item,
      duplicate: validation.duplicate,
      errors: rowErrors,
      mode: validation.duplicate ? "update" : "create",
    };

    if (validation.duplicate) {
      duplicateRows.push(previewItem);
    }
    if (rowErrors.length) {
      errors.push(...rowErrors);
      errorRows.push(previewItem);
      return;
    }
    validCount += 1;
    if (validation.duplicate && !shouldUpdateDuplicates) {
      return;
    }
    validRows.push(previewItem);
  });

  return {
    totalRows: mappedRows.length,
    validCount,
    validRows,
    errors,
    errorRows,
    duplicateRows,
    skippedDuplicateRows: shouldUpdateDuplicates ? [] : duplicateRows.filter((item) => !item.errors.length),
    shouldUpdateDuplicates,
  };
}

function previewImportResult(preview) {
  pendingImportPreview = preview;
  confirmExcelImportBtn.disabled = preview.validRows.length === 0;
  confirmExcelImportBtn.classList.remove("hidden");
  cancelExcelImportBtn.classList.remove("hidden");
  excelPreview.classList.remove("hidden");

  const errorRows = preview.errors.length
    ? preview.errors
        .map(
          (error) => `
            <tr>
              <td>${escapeHtml(error)}</td>
            </tr>
          `,
        )
        .join("")
    : `<tr><td class="muted">Không có lỗi dữ liệu.</td></tr>`;

  excelPreview.innerHTML = `
    <div class="preview-summary">
      <div class="stat-card"><span class="stat-label">Tổng số dòng đọc được</span><strong>${preview.totalRows}</strong></div>
      <div class="stat-card total"><span class="stat-label">Số dòng hợp lệ</span><strong>${preview.validCount}</strong></div>
      <div class="stat-card warning"><span class="stat-label">Số dòng lỗi</span><strong>${preview.errorRows.length}</strong></div>
      <div class="stat-card warning"><span class="stat-label">Số dòng trùng</span><strong>${preview.duplicateRows.length}</strong></div>
    </div>
    <div class="preview-errors">
      <table>
        <thead>
          <tr><th>Danh sách lỗi theo dòng</th></tr>
        </thead>
        <tbody>${errorRows}</tbody>
      </table>
    </div>
  `;
}

function resetExcelImport() {
  selectedExcelFile = null;
  pendingImportPreview = null;
  excelInput.value = "";
  previewExcelBtn.disabled = true;
  confirmExcelImportBtn.disabled = true;
  confirmExcelImportBtn.classList.add("hidden");
  cancelExcelImportBtn.classList.add("hidden");
  excelPreview.className = "excel-preview hidden";
  excelPreview.innerHTML = "";
}

function confirmImport() {
  if (!pendingImportPreview) return;

  const now = new Date().toISOString();
  let nextFacilities = [...facilities];
  let created = 0;
  let updated = 0;

  pendingImportPreview.validRows.forEach((item) => {
    if (item.mode === "update" && item.duplicate) {
      nextFacilities = nextFacilities.map((record) =>
        record.id === item.duplicate.id
          ? {
              ...record,
              ...item.data,
              id: record.id,
              maCoSo: record.maCoSo,
              createdAt: record.createdAt,
              updatedAt: now,
            }
          : record,
      );
      updated += 1;
      return;
    }

    nextFacilities.push({
      id: createId(),
      maCoSo: getNextFacilityCodeFrom(nextFacilities),
      ...item.data,
      createdAt: now,
      updatedAt: now,
    });
    created += 1;
  });

  facilities = nextFacilities;
  saveFacilities();
  renderFacilities();
  setMessage(
    `Import Excel hoàn tất. Đã thêm mới ${created} hồ sơ. Cập nhật ${updated} hồ sơ. Bỏ qua ${pendingImportPreview.errorRows.length} dòng lỗi. Bỏ qua ${pendingImportPreview.skippedDuplicateRows.length} hồ sơ trùng.`,
    "success",
  );
  resetExcelImport();
}

function sanitizeImportedRecord(record) {
  const data = {};
  FIELDS.forEach((field) => {
    if (CHECKBOX_FIELDS.includes(field)) {
      data[field] = Boolean(record[field]);
    } else if (NUMBER_FIELDS.includes(field)) {
      data[field] = toNumberOrEmpty(record[field]);
    } else if (field === "tenCoSo") {
      data[field] = normalizeFacilityName(record[field]);
    } else if (field === "soTaiKhoanPCCC" || PHONE_FIELDS.includes(field)) {
      data[field] = normalizeDigits(record[field]);
    } else {
      data[field] = normalizeText(record[field]);
    }
  });
  data.thanhPho = DEFAULT_CITY;
  data.phuongXa = WARDS.includes(data.phuongXa) ? data.phuongXa : WARDS[0];
  data.linhVucCode = SECTOR_OPTIONS.some(([code]) => code === data.linhVucCode) ? data.linhVucCode : "";
  data.linhVucName = getSectorName(data.linhVucCode);
  return data;
}

async function importJson(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  try {
    const text = await file.text();
    const parsed = JSON.parse(text);
    const records = Array.isArray(parsed) ? parsed : parsed.records || parsed.facilities;
    if (!Array.isArray(records)) {
      throw new Error("Tệp JSON phải là mảng hồ sơ hoặc có thuộc tính records là mảng.");
    }

    let imported = 0;
    let skipped = 0;
    const now = new Date().toISOString();

    records.forEach((record) => {
      const data = sanitizeImportedRecord(record);
      const { errors, duplicate } = validateFacility(data);
      if (errors.length || duplicate) {
        skipped += 1;
        return;
      }

      facilities.push({
        id: createId(),
        maCoSo: getNextFacilityCode(),
        ...data,
        createdAt: record.createdAt || now,
        updatedAt: now,
      });
      imported += 1;
    });

    saveFacilities();
    renderFacilities();
    setMessage(`Đã nhập ${imported} hồ sơ. Bỏ qua ${skipped} hồ sơ lỗi hoặc trùng số tài khoản PCCC.`, "success");
  } catch (error) {
    setMessage(error.message || "Không thể nhập tệp JSON.", "error");
  } finally {
    importInput.value = "";
  }
}

function handleExcelFileSelect(event) {
  selectedExcelFile = event.target.files?.[0] || null;
  pendingImportPreview = null;
  previewExcelBtn.disabled = !selectedExcelFile;
  confirmExcelImportBtn.disabled = true;
  confirmExcelImportBtn.classList.add("hidden");
  cancelExcelImportBtn.classList.toggle("hidden", !selectedExcelFile);
  excelPreview.className = "excel-preview hidden";
  excelPreview.innerHTML = "";
}

async function handlePreviewExcel() {
  if (!selectedExcelFile) return;

  try {
    clearMessage();
    const mappedRows = await parseExcelFile(selectedExcelFile);
    const preview = buildImportPreview(mappedRows, updateDuplicatesInput.checked);
    previewImportResult(preview);
  } catch (error) {
    setMessage(error.message || "Không thể đọc file Excel.", "error");
  }
}

function handleTableClick(event) {
  const button = event.target.closest("button");
  if (!button) return;

  const { action, id } = button.dataset;
  if (action === "edit") editFacility(id);
  if (action === "delete") deleteFacility(id);
}

function handleDuplicateClick(event) {
  const button = event.target.closest("button");
  if (!button) return;

  const { action, id } = button.dataset;
  if (action === "view-duplicate") viewFacility(id);
  if (action === "edit-duplicate") editFacility(id);
  if (action === "close-duplicate") hideDuplicate();
}

function init() {
  fillSelect("phuongXa", WARDS);
  fillSelect("linhVucCode", SECTOR_OPTIONS, "Chọn lĩnh vực");
  fillSelect("tinhChatHoatDong", ACTIVITY_OPTIONS);
  fillSelect("nhomQuanLy", MANAGEMENT_GROUPS);
  fillSelect("tinhTrangThamDuyetNghiemThu", APPROVAL_STATUS_OPTIONS);
  fillFilterSelect(filterWard, WARDS, "Tất cả");
  fillFilterSelect(filterSector, SECTOR_OPTIONS, "Tất cả");
  fillFilterSelect(filterApproval, APPROVAL_STATUS_OPTIONS, "Tất cả");
  form.elements.thanhPho.value = DEFAULT_CITY;

  form.addEventListener("submit", handleSubmit);
  tableBody.addEventListener("click", handleTableClick);
  duplicateBox.addEventListener("click", handleDuplicateClick);
  searchInput.addEventListener("input", renderFacilities);
  resetBtn.addEventListener("click", resetForm);
  cancelEditBtn.addEventListener("click", resetForm);
  exportBtn.addEventListener("click", exportJson);
  exportCsvBtn.addEventListener("click", exportCsv);
  exportBackupBtn.addEventListener("click", exportBackupJson);
  importInput.addEventListener("change", importJson);
  [filterWard, filterSector, filterApproval, filterFlag].forEach((filter) => {
    filter.addEventListener("change", renderFacilities);
  });
  excelInput.addEventListener("change", handleExcelFileSelect);
  previewExcelBtn.addEventListener("click", handlePreviewExcel);
  confirmExcelImportBtn.addEventListener("click", confirmImport);
  cancelExcelImportBtn.addEventListener("click", resetExcelImport);
  updateDuplicatesInput.addEventListener("change", () => {
    if (selectedExcelFile) {
      confirmExcelImportBtn.disabled = true;
      pendingImportPreview = null;
    }
  });

  renderFacilities();
}

init();
