// ==============================
// 📊 TẠO FILE NHẬN XÉT & SO SÁNH TỰ ĐỘNG (BẢN GỘP–ỔN ĐỊNH–KHÔNG BẢN SAO)
// ==============================

// ⚙️ XÓA CACHE CÁC BIẾN GLOBAL TRÁNH XUNG ĐỘT
delete globalThis.REPORT_FOLDER_ROOT_ID;
delete globalThis.TEMPLATE_DOC_ID;
delete globalThis.TEMPLATE_COMPARE_DOC_ID;

// ⚙️ ID THỰC TẾ (KIỂM TRA KỸ)
var REPORT_FOLDER_ROOT_ID   = '1mLSVUfQlA9pLnzwXkbl2tANJ3OyIFyNn'; // Folder EDUAI Reports
var TEMPLATE_DOC_ID         = '1CWZ6eP2xLlsiyz-h446H2t2Q1FE_FDEd0-qG-tImck0'; // Template Nhận xét
var TEMPLATE_COMPARE_DOC_ID = '1W52gParRbWW_MHXwAyG_am0lzqAvWAe9GvoMxGMqTIA'; // Template So sánh

// ⚙️ CẤU HÌNH DRIVE API V2
var DRIVE_V2_OPTS = { supportsAllDrives: true, supportsTeamDrives: true };

// ==============================
// 🧠 TẠO FILE NHẬN XÉT (ONLUYEN REPORT)
// ==============================
function generateOnluyenReport(fileId, reportName, region) {
  try {
    if (!fileId || !region || !reportName)
      throw new Error('Thiếu dữ liệu (fileId, region, reportName).');

    // 1️⃣ Tạo cấu trúc thư mục: /EDUAI Reports/{Region}/{School}
    const parent = DriveApp.getFolderById(REPORT_FOLDER_ROOT_ID);
    const regionFld = getOrCreateFolder(parent, region);
    const schoolName = (reportName.split(' - ')[0] || '').trim();
    const schoolFld = getOrCreateFolder(regionFld, schoolName);

    // 2️⃣ Lấy ngày trong tên file (nếu có “từ ... đến ...”)
    const fileMeta = Drive.Files.get(fileId);
    const fileName = (fileMeta.title || '').normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    const re = /(tu|từ)\s*(\d{1,2}[-\/]\d{1,2}[-\/]\d{4})\s*(den|đến|to)\s*(\d{1,2}[-\/]\d{1,2}[-\/]\d{4})/i;
    const match = fileName.match(re);
    const dateFrom = match ? match[2].replace(/-/g, "/") : "";
    const dateTo   = match ? match[4].replace(/-/g, "/") : "";

    // 3️⃣ Xác định thư mục giai đoạn
    const normalizeFolderName = s => s.toString().trim().toLowerCase().replace(/\s+/g, " ");
    const stageName = reportName.includes("1") ? "Giai đoạn 1" : "Giai đoạn 2";
    let stageFld = null;
    const allFolders = schoolFld.getFolders();
    while (allFolders.hasNext()) {
      const f = allFolders.next();
      if (normalizeFolderName(f.getName()) === normalizeFolderName(stageName)) {
        stageFld = f;
        break;
      }
    }
    if (!stageFld) stageFld = schoolFld.createFolder(stageName);

    const targetName = `📄 Nhận xét ${reportName}`;

    // 4️⃣ Xóa file cũ trùng tên (nếu có)
    const existing = stageFld.getFilesByName(targetName);
    while (existing.hasNext()) existing.next().setTrashed(true);

    // 5️⃣ Copy template Docs vào đúng thư mục giai đoạn
    Logger.log("📂 Folder cha (Giai đoạn): " + stageFld.getName() + " | ID: " + stageFld.getId());
    Logger.log("📘 Template dùng: " + TEMPLATE_DOC_ID);

    const copyMeta = {
      title: targetName,
      parents: [{ id: stageFld.getId() }]
    };

   // ✅ Sử dụng cú pháp đúng: (resource, fileId, options)
const copied = Drive.Files.copy(copyMeta, TEMPLATE_DOC_ID, DRIVE_V2_OPTS);

// ✅ Đổi tên file để tránh "Bản sao của..."
DriveApp.getFileById(copied.id).setName(targetName);

// ✅ Ép file copy gắn vào đúng thư mục giai đoạn (tránh rơi ra thư mục tổng)
DriveApp.getFolderById(stageFld.getId()).addFile(DriveApp.getFileById(copied.id));
Logger.log("📎 Đã gắn file vào thư mục: " + stageFld.getName());

// ✅ Log đường dẫn file
Logger.log("✅ Đã tạo file nhận xét: https://docs.google.com/document/d/" + copied.id);

    // 📝 Mở file Docs vừa copy để ghi nội dung
    const doc = DocumentApp.openById(copied.id);
    const body = doc.getBody();

    // 6️⃣ Đọc dữ liệu GV/HS
    const ss = SpreadsheetApp.openById(fileId);
    const shGV = pickSheet(ss, ['gv', 'giáo viên']);
    const shHS = pickSheet(ss, ['hs', 'học sinh']);
    const gv = shGV ? readObjects(shGV) : [];
    const hs = shHS ? readObjects(shHS) : [];

    // 7️⃣ Tính toán chỉ số
    const numGV = countTeachersAssigning(gv);
    const totalTasksCreated  = sumByField(gv, ['bộ đề', 'bo de', 'bài tập đã tạo']);
    const totalTasksAssigned = sumByField(gv, ['giao đề', 'bai tap da giao']);
    const numHS = countStudentsDoing(hs);

    const usageLevel = numGV > 20 ? "đang có mức sử dụng cao"
      : numGV > 10 ? "đang ở mức trung bình" : "cần cải thiện thêm";
    const schoolLevel = numGV > 15 ? "Tốt" : numGV > 5 ? "Khá" : "Thấp";

    // 8️⃣ Điền dữ liệu vào template
    replaceAll(body, {
      '{{SCHOOL}}': schoolName,
      '{{DATE_FROM}}': dateFrom,
      '{{DATE_TO}}': dateTo,
      '{{USAGE_LEVEL}}': usageLevel,
      '{{SCHOOL_LEVEL}}': schoolLevel,
      '{{NUM_TEACHERS_ASSIGNING}}': String(numGV),
      '{{TOTAL_TASKS_CREATED}}': String(totalTasksCreated),
      '{{TOTAL_TASKS_ASSIGNED}}': String(totalTasksAssigned),
      '{{NUM_STUDENTS_DOING}}': String(numHS)
    });

    // 9️⃣ Nhận xét AI tự động
    const aiFn = (typeof globalThis.generateAISummary === 'function')
      ? globalThis.generateAISummary : generateAISummaryFallback;
    const aiComment = aiFn({ region, schoolName, numGV, numHS, totalTasksCreated, totalTasksAssigned })
      || '(Chưa có dữ liệu AI)';

    body.appendParagraph('\n🤖 Nhận xét tự động:').setBold(true);
    body.appendParagraph(aiComment);
    body.appendParagraph('\n🔗 Dữ liệu gốc: https://docs.google.com/spreadsheets/d/' + fileId);
    doc.saveAndClose();

    // 🔗 Trả về link Docs đầy đủ
    const url = `https://docs.google.com/document/d/${copied.id}/edit`;
    Logger.log('✅ Tạo xong báo cáo: ' + url);
    return url;

  } catch (err) {
    Logger.log('❌ Lỗi generateOnluyenReport: ' + err);
    throw new Error('Không thể tạo báo cáo nhận xét: ' + err.message);
  }
}

// ==============================
// ⚙️ HÀM TIỆN ÍCH
// ==============================
function getOrCreateFolder(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function normalize(s) {
  return (s || "").toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
}

function toNum(x) {
  const n = Number(x);
  return isNaN(n) ? 0 : n;
}

function replaceAll(body, map) {
  for (const [key, val] of Object.entries(map)) {
    if (val !== undefined && val !== null) body.replaceText(key, val);
  }
}

// ==============================
// 📖 ĐỌC DỮ LIỆU GV / HS
// ==============================
function pickSheet(ss, keywords) {
  const sheets = ss.getSheets();
  for (const sh of sheets) {
    const name = normalize(sh.getName());
    if (keywords.some(k => name.includes(normalize(k)))) return sh;
  }
  return null;
}

function readObjects(sh) {
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(normalize);
  return data.slice(1).map(r => {
    const o = {};
    headers.forEach((h, i) => o[h] = r[i]);
    return o;
  });
}

function pickKey(arr, candidates) {
  if (!arr.length) return null;
  const keys = Object.keys(arr[0]);
  for (const k of keys) {
    const nk = normalize(k);
    if (candidates.some(c => nk.includes(normalize(c)))) return k;
  }
  return null;
}

function countTeachersAssigning(gvData) {
  const key = pickKey(gvData, ['giao de', 'bài tập đã giao']);
  return key ? gvData.filter(o => toNum(o[key]) > 0).length : 0;
}

function countStudentsDoing(hsData) {
  const key = pickKey(hsData, ['số bài đã làm', 'bt lam']);
  return key ? hsData.filter(o => toNum(o[key]) > 0).length : 0;
}

function sumByField(arr, candidates) {
  const key = pickKey(arr, candidates);
  return key ? arr.reduce((s, o) => s + toNum(o[key]), 0) : 0;
}

// ==============================
// 🤖 NHẬN XÉT AI – DỰ PHÒNG
// ==============================
function generateAISummaryFallback({ region, schoolName, numGV, numHS, totalTasksCreated, totalTasksAssigned }) {
  try {
    const score = (numGV * 0.4) + (numHS * 0.3) + ((totalTasksCreated + totalTasksAssigned) * 0.3 / 10);
    let level = score < 5 ? '🔴 **RẤT THẤP**'
      : score < 15 ? '🟠 **THẤP**'
      : score < 30 ? '🟡 **TRUNG BÌNH**'
      : '🟢 **TỐT**';

    return `Trong kỳ báo cáo, trường **${schoolName}** (${region}) có ${numGV} GV và ${numHS} HS hoạt động. ` +
      `Đã tạo ${totalTasksCreated} bài và giao ${totalTasksAssigned} bài. => Mức độ sử dụng: ${level}.`;
  } catch (err) {
    Logger.log('❌ Lỗi generateAISummaryFallback: ' + err);
    return '(Không thể sinh nhận xét tự động)';
  }
}

// ==============================
// 📊 GHI LOG FACT_Usage
// ==============================
function logToFactUsage(region, schoolName, stage, numGV, numHS, totalTasksCreated, totalTasksAssigned) {
  try {
    const ss = SpreadsheetApp.openById("1rhsVChmwvA1tHIsGZbno9R-GU_FznTdcC7N9yf14k6Q");
    const sh = ss.getSheetByName("FACT_Usage") || ss.insertSheet("FACT_Usage");
    if (sh.getLastRow() === 0) {
      sh.appendRow(["Khu_vuc", "Trường", "Giai_đoạn", "GV_giao_bai", "HS_lam_bai", "Bai_tap_tao", "Bai_tap_giao", "Ngày_báo_cáo"]);
    }
    sh.appendRow([region, schoolName, stage, numGV, numHS, totalTasksCreated, totalTasksAssigned, new Date()]);
  } catch (err) {
    Logger.log("❌ Lỗi ghi FACT_Usage: " + err);
  }
}

// ==============================
// 🧪 TEST
// ==============================
function testGenerateOnluyen() {
  const region = "Hà Nam";
  const schoolName = "THPT Bắc Lý";
  const fileId = "ID_FILE_GIAI_DOAN_1"; // Thay ID thật
  const url = generateOnluyenReport(fileId, `${schoolName} - Giai đoạn 1`, region);
  Logger.log("📄 Test thành công: " + url);
}
