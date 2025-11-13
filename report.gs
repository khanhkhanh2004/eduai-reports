function createReportDoc(school, cmp, aiText) {
  // 🗂️ Lấy thư mục EDUAI Reports gốc từ Script Properties
  const rootId = PropertiesService.getScriptProperties().getProperty("REPORT_FOLDER_ID") 
    || '1mLSVUfQlA9pLnzwXkbl2tANJ3OyIFyNn'; // fallback
  const parent = DriveApp.getFolderById(rootId);

  // 🧩 Template so sánh (giống reportComparison.gs)
  const templateId = '1W52gParRbWW_MHXwAyG_am0lzqAvWAe9GvoMxGMqTIA';

  // 🏫 Tạo hoặc tìm thư mục theo tên trường
  const schoolFld = getOrCreateFolder(parent, school);

  // 📄 Tên file báo cáo
  const fileName = `[So sánh] ${school} - 2 Giai đoạn`;

  // 🔄 Xóa file cũ trùng tên
  const existing = schoolFld.getFilesByName(fileName);
  while (existing.hasNext()) existing.next().setTrashed(true);

  // 📑 Copy template vào thư mục trường
  const copied = Drive.Files.copy(
    { title: fileName, parents: [{ id: schoolFld.getId() }] },
    templateId,
    { supportsAllDrives: true, supportsTeamDrives: true }
  );
  Logger.log("✅ Tạo báo cáo Docs mới: " + copied.id);

  // ✍️ Ghi nội dung vào file
  const doc = DocumentApp.openById(copied.id);
  const b = doc.getBody();

  b.replaceText('{{SCHOOL}}', school);
  b.replaceText('{{DATE}}', new Date().toLocaleDateString("vi-VN"));

  b.appendParagraph("\n📊 TỔNG QUAN SỐ LIỆU").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  b.appendTable([
    ["Chỉ tiêu", "Giai đoạn 1", "Giai đoạn 2", "Chênh lệch"],
    ["Số HS hoàn thành", cmp.done1, cmp.done2, cmp.done2 - cmp.done1],
    ["Tỷ lệ hoàn thành (%)", cmp.percent1, cmp.percent2, `${cmp.diff}%`]
  ]);

  b.appendParagraph("\n💬 NHẬN XÉT TỰ ĐỘNG (AI)").setHeading(DocumentApp.ParagraphHeading.HEADING2);
  b.appendParagraph(aiText || "(Chưa có dữ liệu AI)");

  doc.saveAndClose();

  return `https://docs.google.com/document/d/${copied.id}/edit`;
}

// ⚙️ Hàm tiện ích (nếu chưa có)
function getOrCreateFolder(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
