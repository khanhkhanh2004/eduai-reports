/**
 * 🧠 WEB CONTROLLER – Xử lý chính logic tạo báo cáo & nhận xét AI
 * Liên kết giữa Google Sheets và giao diện web (UI.html)
 */

// ==============================
// 📊 GIAO DIỆN WEB APP
// ==============================
function doGet() {
  return HtmlService.createHtmlOutputFromFile("ui")
    .setTitle("📊 Hệ thống phân tích giáo dục AI")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==============================
// 🧩 HÀM TẠO NHẬN XÉT VÀ BÁO CÁO CHO MỘT TRƯỜNG
// ==============================
function runGenerateForSchool(schoolName) {
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName("Tạo nhận xét");
    const last = sh.getLastRow();
    const user = Session.getActiveUser().getEmail();

    for (let r = 6; r <= last; r++) {
      const school = (sh.getRange(r, 3).getValue() || "").toString().trim();
      if (!school || school !== schoolName) continue;

      const linkGD1 = sh.getRange(r, 4).getValue();
      const linkGD2 = sh.getRange(r, 5).getValue();
      if (!linkGD1 || !linkGD2)
        throw new Error("Thiếu link Giai đoạn 1/2. Hãy upload đủ 2 file.");

      Logger.log(`📂 Đang tạo báo cáo cho: ${schoolName}`);

      // 1️⃣ Đọc dữ liệu từng giai đoạn
      const gd1 = readDataFromFile(linkGD1);
      const gd2 = readDataFromFile(linkGD2);

      // 2️⃣ So sánh hai giai đoạn
      const cmp = compareStages(gd1, gd2);

      // 3️⃣ Tạo nhận xét AI thông minh (gọi từ ai.gs)
      const aiText = generateAISummary({
        region: sh.getRange(r, 2).getValue() || "",
        schoolName: schoolName,
        numGV: cmp.numGV || 0,
        numHS: cmp.numHS || 0,
        totalTasksCreated: cmp.baiTapTao || 0,
        totalTasksAssigned: cmp.baiTapGiao || 0,
      });

      // 4️⃣ Tạo file Google Docs báo cáo
      const url = createReportDoc(schoolName, cmp, aiText);

      // 5️⃣ Cập nhật lại vào Google Sheet
      sh.getRange(r, 6).setValue(new Date()); // Ngày tạo
      sh.getRange(r, 8).setValue(url); // Link báo cáo
      sh.getRange(r, 9).setValue(user); // Người tạo
      sh.getRange(r, 9).setHorizontalAlignment("left");
      sh.getRange(r, 9).setWrap(true);
      sh.getRange(r, 10).setValue("✅ Đã tạo báo cáo");

      Logger.log(`✅ Hoàn thành cho ${schoolName}: ${url}`);

      return {
        status: "success",
        school: schoolName,
        reportUrl: url,
        aiSummary: aiText,
      };
    }

    throw new Error("Không tìm thấy trường trong Sheet.");
  } catch (err) {
    Logger.log("❌ Lỗi runGenerateForSchool: " + err);
    return {
      status: "error",
      message: err.toString(),
    };
  }
}
