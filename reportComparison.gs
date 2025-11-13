/***** =========== CONFIG =========== *****/
// Dùng var + globalThis để tương thích giữa các file
var COMPARE_PARENT_FOLDER_ID = globalThis.REPORT_FOLDER_ROOT_ID
  || '1mLSVUfQlA9pLnzwXkbl2tANJ3OyIFyNn'; // Folder gốc EDUAI Reports

var COMPARE_TEMPLATE_ID = globalThis.TEMPLATE_COMPARE_DOC_ID
  || '1W52gParRbWW_MHXwAyG_am0lzqAvWAe9GvoMxGMqTIA'; // Template Docs So sánh

/***** =========== MAIN ENTRY =========== *****/
function compareStagesAndUpdateReport_v2(region, school, fileId1, fileId2) {
  try {
    if (!region || !school || !fileId1 || !fileId2)
      throw new Error("Thiếu dữ liệu region/school/fileId1/fileId2 khi tạo báo cáo so sánh.");

    Logger.log(`📊 So sánh: ${school} (${region})`);

    // 🗂️ Cấu trúc thư mục: EDUAI Reports/{Region}/{School}
    const parent = DriveApp.getFolderById(COMPARE_PARENT_FOLDER_ID);
    const regionFolder = getOrCreateFolder(parent, region);
    const schoolFolder = getOrCreateFolder(regionFolder, school);

    // --- Mở file dữ liệu ---
    const ss1 = SpreadsheetApp.openById(fileId1);
    const ss2 = SpreadsheetApp.openById(fileId2);
    const gv1 = readStageData_(ss1, ['gv', 'giáo viên']);
    const gv2 = readStageData_(ss2, ['gv', 'giáo viên']);
    const hs1 = readStageData_(ss1, ['hs', 'học sinh']);
    const hs2 = readStageData_(ss2, ['hs', 'học sinh']);

    // --- Tính toán ---
    const numGV1 = countTeachersAssigning(gv1);
    const numGV2 = countTeachersAssigning(gv2);
    const numHS1 = countStudentsDoing(hs1);
    const numHS2 = countStudentsDoing(hs2);
    const task1 = sumByField(gv1, ['bài tập đã tạo', 'so bai tap da tao']);
    const task2 = sumByField(gv2, ['bài tập đã tạo', 'so bai tap da tao']);

    const deltaGV = numGV2 - numGV1;
    const deltaHS = numHS2 - numHS1;
    const deltaTask = task2 - task1;

    const deltaGVPercent = numGV1 === 0 ? 0 : ((deltaGV / numGV1) * 100).toFixed(1);
    const deltaHSPercent = numHS1 === 0 ? 0 : ((deltaHS / numHS1) * 100).toFixed(1);
    const deltaTaskPercent = task1 === 0 ? 0 : ((deltaTask / task1) * 100).toFixed(1);

    // --- Tạo báo cáo Docs ---
    const docId = createComparisonDocWithAI_(
      region,
      school,
      {
        numGV1, numGV2, numHS1, numHS2,
        task1, task2,
        deltaGV, deltaHS, deltaTask,
        deltaGVPercent, deltaHSPercent, deltaTaskPercent,
        fileId1, fileId2
      },
      schoolFolder
    );

    // --- Ghi log vào FACT_Usage (nếu có)
    if (typeof logToFactUsage === 'function') {
      logToFactUsage(region, school, 'So sánh 2 giai đoạn', numGV2, numHS2, task2, task2);
    }

    Logger.log("✅ Đã tạo xong báo cáo: " + docId);
    return "https://docs.google.com/document/d/" + docId + "/edit";

  } catch (err) {
    Logger.log("❌ Lỗi compareStagesAndUpdateReport_v2: " + err);
    throw new Error("Không thể tạo báo cáo so sánh: " + err.message);
  }
}

/***** =========== CORE FUNCTION =========== *****/
function createComparisonDocWithAI_(region, school, stats, folder) {
  // 🧹 Dọn bản cũ trùng tên
  const compareName = `[So sánh] ${school} - 2 Giai đoạn`;
  const existing = folder.getFilesByName(compareName);
  while (existing.hasNext()) existing.next().setTrashed(true);

  // 📝 Copy template Docs vào thư mục trường
  Logger.log("📂 Folder trường: " + folder.getName() + " | ID: " + folder.getId());
  Logger.log("📘 Template so sánh: " + COMPARE_TEMPLATE_ID);

  const copyMeta = {
    title: compareName,
    parents: [{ id: folder.getId() }]
  };

 // ✅ Cú pháp đúng: (resource, fileId, options)
const copied = Drive.Files.copy(copyMeta, COMPARE_TEMPLATE_ID, { supportsAllDrives: true, supportsTeamDrives: true });

// ✅ Đổi lại tên file để tránh “Bản sao của…”
DriveApp.getFileById(copied.id).setName(compareName);

// ✅ Ép file copy gắn vào đúng thư mục trường (tránh rơi ra thư mục tổng)
DriveApp.getFolderById(folder.getId()).addFile(DriveApp.getFileById(copied.id));
Logger.log("📎 Đã gắn file vào thư mục: " + folder.getName());

// ✅ Log đường dẫn file
Logger.log("✅ Đã tạo báo cáo so sánh: https://docs.google.com/document/d/" + copied.id);

  // 📄 Mở file Docs và ghi nội dung
  const doc = DocumentApp.openById(copied.id);
  const body = doc.getBody();

  // 🔁 Thay các placeholder
  replaceAll(body, {
    '{{SCHOOL}}': school,
    '{{REGION}}': region,
    '{{GV_STAGE1}}': stats.numGV1,
    '{{GV_STAGE2}}': stats.numGV2,
    '{{HS_STAGE1}}': stats.numHS1,
    '{{HS_STAGE2}}': stats.numHS2,
    '{{TASK_STAGE1}}': stats.task1,
    '{{TASK_STAGE2}}': stats.task2,
    '{{DELTA_GV}}': formatDelta_(stats.deltaGV, stats.deltaGVPercent, 'giáo viên'),
    '{{DELTA_HS}}': formatDelta_(stats.deltaHS, stats.deltaHSPercent, 'học sinh'),
    '{{DELTA_TASK}}': formatDelta_(stats.deltaTask, stats.deltaTaskPercent, 'bài tập'),
    '{{DATE}}': new Date().toLocaleDateString("vi-VN")
  });

  // 🤖 Nhận xét tự động
  if (typeof generateAICompareSummary === 'function') {
    const aiComment = generateAICompareSummary(stats, school, region);
    body.appendParagraph("\n🤖 Nhận xét tự động:").setBold(true);
    body.appendParagraph(aiComment);
  }

  // 🔗 Ghi link nguồn dữ liệu
  body.appendParagraph("\n🔗 Dữ liệu nguồn:");
  body.appendParagraph("• Giai đoạn 1: https://docs.google.com/spreadsheets/d/" + stats.fileId1);
  body.appendParagraph("• Giai đoạn 2: https://docs.google.com/spreadsheets/d/" + stats.fileId2);

  // 💾 Lưu Docs
  doc.saveAndClose();
  return copied.id;
}

/***** =========== AI SUMMARY =========== *****/
function generateAICompareSummary(stats, school, region) {
  const { deltaGV, deltaGVPercent, deltaHS, deltaHSPercent, deltaTask, deltaTaskPercent } = stats;
  const comment = [];

  comment.push(`📊 **Báo cáo so sánh hai giai đoạn sử dụng Onluyen.vn tại trường ${school} (${region})**`);
  comment.push("");

  // 1️⃣ Giáo viên
  if (deltaGV > 0)
    comment.push(`• 🟢 **Giáo viên:** tăng ${deltaGV} (+${deltaGVPercent}%) – tích cực hơn trong việc giao bài.`);
  else if (deltaGV < 0)
    comment.push(`• 🔴 **Giáo viên:** giảm ${Math.abs(deltaGV)} (${deltaGVPercent}%) – cần khuyến khích thêm.`);
  else comment.push(`• ⚪ **Giáo viên:** không thay đổi.`);

  // 2️⃣ Học sinh
  if (deltaHS > 0)
    comment.push(`• 🟢 **Học sinh:** tăng ${deltaHS} (+${deltaHSPercent}%) – tương tác tốt hơn.`);
  else if (deltaHS < 0)
    comment.push(`• 🔴 **Học sinh:** giảm ${Math.abs(deltaHS)} (${deltaHSPercent}%) – cần thúc đẩy tham gia.`);
  else comment.push(`• ⚪ **Học sinh:** ổn định.`);

  // 3️⃣ Bài tập
  if (deltaTask > 0)
    comment.push(`• 🟢 **Bài tập:** tăng ${deltaTask} (+${deltaTaskPercent}%) – giáo viên tạo nội dung tích cực.`);
  else if (deltaTask < 0)
    comment.push(`• 🔴 **Bài tập:** giảm ${Math.abs(deltaTask)} (${deltaTaskPercent}%) – cần đẩy mạnh ra đề.`);
  else comment.push(`• ⚪ **Bài tập:** không đổi.`);

  // Tổng quan
  const avg = (Number(deltaGVPercent) + Number(deltaHSPercent) + Number(deltaTaskPercent)) / 3;
  comment.push("");
  comment.push("📈 **Nhận xét tổng quan:**");
  if (avg > 20)
    comment.push("🟢 Mức sử dụng tăng mạnh – duy trì đà tích cực này.");
  else if (avg > 5)
    comment.push("🟡 Mức sử dụng tăng nhẹ – ổn định, cần khích lệ thêm.");
  else if (avg > -5)
    comment.push("⚪ Ổn định – không biến động lớn, nên duy trì.");
  else
    comment.push("🔴 Giảm rõ rệt – cần hỗ trợ GV & HS khôi phục hoạt động.");

  comment.push("");
  comment.push("🧩 **Đề xuất:** Duy trì phong trào giao bài định kỳ, chia sẻ đề hay, tuyên dương GV/HS hoạt động tốt.");

  return comment.join("\n");
}

/***** =========== UTILS =========== *****/
function readStageData_(ss, keywords) {
  const sh = pickSheet(ss, keywords);
  return sh ? readObjects(sh) : [];
}

function getOrCreateFolder(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function normalize(s) {
  return (s || "").toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
}

function replaceAll(body, map) {
  for (const [k, v] of Object.entries(map)) {
    if (v !== undefined && v !== null) body.replaceText(k, v);
  }
}

function formatDelta_(delta, percent, label) {
  if (delta > 0) return `📈 Tăng ${delta} ${label} (+${percent}%)`;
  if (delta < 0) return `📉 Giảm ${Math.abs(delta)} ${label} (${percent}%)`;
  return `⚖️ Không thay đổi ${label}`;
}

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
  const headers = data[0].map(h => normalize(h));
  return data.slice(1).map(row => {
    const o = {};
    headers.forEach((h, i) => (o[h] = row[i]));
    return o;
  });
}

function toNum(x) {
  const n = Number(x);
  return isNaN(n) ? 0 : n;
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

function countTeachersAssigning(gv) {
  const key = pickKey(gv, ['bài tập đã giao', 'so bai tap da giao', 'bt giao']);
  return key ? gv.filter(o => toNum(o[key]) > 0).length : 0;
}

function countStudentsDoing(hs) {
  const key = pickKey(hs, ['số bài đã làm', 'so bai da lam', 'bt lam']);
  return key ? hs.filter(o => toNum(o[key]) > 0).length : 0;
}

function sumByField(arr, candidates) {
  const key = pickKey(arr, candidates);
  return key ? arr.reduce((s, o) => s + toNum(o[key]), 0) : 0;
}
