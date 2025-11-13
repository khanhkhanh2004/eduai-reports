// ==============================
// 📤 UPLOAD & TẠO BÁO CÁO NHẬN XÉT + SO SÁNH TỰ ĐỘNG
// ==============================

// ⚙️ Cấu hình cơ bản
const SHEET_NAME = "Tạo nhận xét";
const DATA_START_ROW = 5; // Dòng bắt đầu dữ liệu trong bảng trung tâm
const UPLOAD_PARENT_FOLDER_ID = '1mLSVUfQlA9pLnzwXkbl2tANJ3OyIFyNn'; // Folder gốc EDUAI Reports

// ⚙️ Đảm bảo các biến toàn cục từ file khác luôn sẵn sàng
var REPORT_FOLDER_ROOT_ID   = globalThis.REPORT_FOLDER_ROOT_ID   || UPLOAD_PARENT_FOLDER_ID;
var TEMPLATE_DOC_ID         = globalThis.TEMPLATE_DOC_ID         || '1CWZ6eP2xLlsiyz-h446H2t2Q1FE_FDEd0-qG-tImck0';
var TEMPLATE_COMPARE_DOC_ID = globalThis.TEMPLATE_COMPARE_DOC_ID || '1W52gParRbWW_MHXwAyG_am0lzqAvWAe9GvoMxGMqTIA';

// ==============================
// 🚀 MỞ GIAO DIỆN UPLOAD
// ==============================
function openUploadDialog() {
  const html = HtmlService.createHtmlOutputFromFile("ui")
    .setWidth(700)
    .setHeight(760);
  SpreadsheetApp.getUi().showModalDialog(html, "Upload & Tạo nhận xét AI");
}

// ==============================
// 🚀 UPLOAD FILE GIAI ĐOẠN (được gọi từ UI)
// -> Chỉ upload + convert + ghi link; KHÔNG tạo nhận xét/so sánh ở đây
// ==============================
function uploadStageFile(region, school, stage, fileName, base64Data) {
  if (![region, school, stage, fileName, base64Data].every(Boolean)) {
    throw new Error('Thiếu dữ liệu upload. Hãy dùng hộp thoại "Upload & Tạo nhận xét AI" để tải file.');
  }
  try {
    Logger.log(`📦 Upload bắt đầu: ${region} | ${school} | ${stage}`);

    // === Chuẩn hóa ===
    const cleanSpace = s => (s || "").toString().trim().replace(/\s+/g, " ");
    const cap = s => cleanSpace(s)
      .split(" ")
      .map(w => w.charAt(0).toLocaleUpperCase("vi-VN") + w.slice(1).toLocaleLowerCase("vi-VN"))
      .join(" ");

    const regionClean = cap(region);
    const schoolClean = cap(school);
    const stageClean  = cap(stage);

    // === Tạo cấu trúc Drive ===
    const parent       = DriveApp.getFolderById(UPLOAD_PARENT_FOLDER_ID);
    const regionFolder = getOrCreateFolder(parent, regionClean);
    const schoolFolder = getOrCreateFolder(regionFolder, schoolClean);
    const stageFolder  = getOrCreateFolder(schoolFolder, stageClean);

    // === Upload file Excel gốc (tạm) ===
    const bytes = Utilities.base64Decode(base64Data);
    const blob  = Utilities.newBlob(bytes, MimeType.MICROSOFT_EXCEL, fileName);
    const xlsx  = stageFolder.createFile(blob); // file tạm
    const xlsxId = xlsx.getId();

    // === Convert sang Google Sheet & xóa XLSX gốc ===
    const gsFileId = convertExcelToGoogleSheet_(xlsxId, `${schoolClean} - ${stageClean}`);
    if (!gsFileId) throw new Error('Convert Excel → Google Sheet thất bại (không có fileId).');
    const gsUrl    = `https://docs.google.com/spreadsheets/d/${gsFileId}/edit`;
    try { xlsx.setTrashed(true); } catch (e) { Logger.log('⚠️ Không thể xóa file XLSX tạm: ' + e); }

    // === Ghi link vào sheet trung tâm (cột D/E, ngày tạo, người tạo) ===
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
    if (!sh) throw new Error(`Không tìm thấy sheet "${SHEET_NAME}"`);

    const norm = s => (s || "").toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
    const regionKey = norm(region);
    const schoolKey = norm(school);

    let row = findRowBySchool_(sh, regionKey, schoolKey);
    if (!row) {
      row = Math.max(sh.getLastRow() + 1, DATA_START_ROW);
      sh.getRange(row, 1).setValue(getNextStt_(sh)); // cột A: STT
    }

    sh.getRange(row, 2).setValue(regionClean);           // B: Khu vực
    sh.getRange(row, 3).setValue(schoolClean);           // C: Tên trường
    sh.getRange(row, 6).setValue(new Date());            // F: Ngày tạo
    sh.getRange(row,10).setValue(Session.getActiveUser().getEmail()); // J: Người tạo

    const locale  = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale();
    const sep     = locale.startsWith("en") ? "," : ";";
    const linkVal = `=HYPERLINK("${gsUrl}"${sep}"${schoolClean} - ${stageClean}")`;

    if (stageClean.toLowerCase().includes("1")) {
      sh.getRange(row, 4).setFormula(linkVal);           // D: Giai đoạn 1
      Logger.log("✅ Ghi link Giai đoạn 1");
    } else if (stageClean.toLowerCase().includes("2")) {
      sh.getRange(row, 5).setFormula(linkVal);           // E: Giai đoạn 2
      Logger.log("✅ Ghi link Giai đoạn 2");
    }

    Logger.log(`✅ Upload hoàn tất cho ${schoolClean}`);
    return { status: "success", fileId: gsFileId, gsheetUrl: gsUrl };

  } catch (err) {
    Logger.log(`❌ Lỗi uploadStageFile: ${err}`);
    throw new Error("Lỗi upload: " + err.message);
  }
}

// ==============================
// 🧠 Tạo nhận xét/so sánh thủ công hoặc từ UI
// ==============================
function generateAllReports(region, school, gd1Id, gd2Id) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`Không tìm thấy sheet "${SHEET_NAME}"`);

  const norm = s => (s || "").toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
  const row = findRowBySchool_(sh, norm(region), norm(school));
  if (!row) throw new Error(`Không tìm thấy dữ liệu cho trường ${school}`);

  const locale = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale();
  const sep    = locale.startsWith("en") ? "," : ";";

  // Nếu UI không truyền id, sẽ lấy từ cột D/E
  if (!gd1Id) {
    const gd1Url = extractUrlFromCell_(sh.getRange(row, 4).getFormula() || sh.getRange(row, 4).getValue());
    gd1Id = gd1Url ? (gd1Url.match(/[-\w]{25,}/)?.[0] || null) : null;
  }
  if (!gd2Id) {
    const gd2Url = extractUrlFromCell_(sh.getRange(row, 5).getFormula() || sh.getRange(row, 5).getValue());
    gd2Id = gd2Url ? (gd2Url.match(/[-\w]{25,}/)?.[0] || null) : null;
  }

  if (!gd1Id && !gd2Id) {
    throw new Error("⚠️ Chưa có file giai đoạn nào — hãy tải lên Excel trước.");
  }

  Logger.log(`--- 🔄 BẮT ĐẦU tạo nhận xét cho ${school} (${region}) ---`);
  const result = {};

  // GĐ1
  if (gd1Id) {
    const r1 = generateOnluyenReport(gd1Id, `${school} - Giai đoạn 1`, region);
    Utilities.sleep(2000); // 🔧 SỬA: nghỉ 2s giữa các lần copy để tránh quota
    sh.getRange(row, 7).setFormula(`=HYPERLINK("${r1}"${sep}"Nhận xét Giai đoạn 1")`);
    result.reportUrlGD1 = r1;
  }

  // GĐ2
  if (gd2Id) {
    const r2 = generateOnluyenReport(gd2Id, `${school} - Giai đoạn 2`, region);
    Utilities.sleep(2000); // 🔧 SỬA: nghỉ 2s giữa các lần copy
    sh.getRange(row, 8).setFormula(`=HYPERLINK("${r2}"${sep}"Nhận xét Giai đoạn 2")`);
    result.reportUrlGD2 = r2;
  }

  // So sánh
  if (gd1Id && gd2Id) {
    const cUrl = compareStagesAndUpdateReport_v2(region, school, gd1Id, gd2Id);
    sh.getRange(row, 9).setFormula(`=HYPERLINK("${cUrl}"${sep}"So sánh 2 Giai đoạn")`);
    result.compareUrl = cUrl;
  }

  Logger.log(`✅ Đã tạo nhận xét/so sánh cho trường ${school}`);
  return result;
}

// ==============================
// ⚙️ CÁC HÀM TIỆN ÍCH
// ==============================
function getOrCreateFolder(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

// 🔧 SỬA: thêm retry & delay để tránh lỗi “Invalid JSON payload”
function convertExcelToGoogleSheet_(fileId, newName) {
  let ready = false;
  for (let i = 0; i < 5; i++) {
    try { DriveApp.getFileById(fileId); ready = true; break; }
    catch (e) { Utilities.sleep(500); }
  }
  if (!ready) throw new Error("Không thể truy cập file vừa upload.");

  const resource = { title: newName, mimeType: MimeType.GOOGLE_SHEETS };

  for (let retry = 0; retry < 3; retry++) {
    try {
      const copied = Drive.Files.copy(resource, fileId, { convert: true });
      if (copied && copied.id) {
        Logger.log(`✅ Convert thành công sang Google Sheet: ${copied.id}`);
        return copied.id;
      }
    } catch (e) {
      if (e.message.includes("User rate limit exceeded")) {
        Logger.log(`⚠️ Drive quota tạm đầy — thử lại sau 3s (${retry + 1}/3)...`);
        Utilities.sleep(3000);
      } else {
        throw e;
      }
    }
  }
  throw new Error("Convert Excel → Google Sheet thất bại sau 3 lần thử.");
}

function findRowBySchool_(sh, regionKey, schoolKey) {
  const last = sh.getLastRow();
  const norm = s => (s || "").toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
  for (let r = DATA_START_ROW; r <= last; r++) {
    const reg = norm(sh.getRange(r, 2).getValue());
    const sch = norm(sh.getRange(r, 3).getValue());
    if (reg === regionKey && sch === schoolKey) return r;
  }
  return null;
}

function getNextStt_(sh) {
  const last = sh.getLastRow();
  if (last < DATA_START_ROW) return 1;
  const vals = sh.getRange(DATA_START_ROW, 1, last - DATA_START_ROW + 1, 1).getValues();
  const nums = vals.flat().filter(v => !isNaN(v) && v !== "");
  return nums.length ? Math.max(...nums) + 1 : 1;
}

function extractUrlFromCell_(formulaOrValue) {
  if (!formulaOrValue) return null;
  if (typeof formulaOrValue === 'string' && formulaOrValue.startsWith("=")) {
    const match = formulaOrValue.match(/HYPERLINK\("([^"]+)"/);
    return match ? match[1] : null;
  }
  const match = ('' + formulaOrValue).match(/https?:\/\/[^\s"]+/);
  return match ? match[0] : null;
}
