// ==============================
// 🤖 AI CONTROLLER / GIAO TIẾP GIỮA UI & SCRIPT
// ==============================

// Hiển thị menu trong Google Sheets
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🤖 AI Báo Cáo")
    .addItem("📤 Mở giao diện Upload & Tạo nhận xét", "openUploadDialog")
    .addToUi();
}

// Khi triển khai web app, đây là hàm khởi động giao diện
function doGet() {
  return HtmlService.createHtmlOutputFromFile("ui")
    .setTitle("📊 Hệ thống phân tích giáo dục AI")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Cho phép nhúng file HTML phụ (nếu cần)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==============================
// 📤 HÀM TRUNG GIAN: UPLOAD FILE GIAI ĐOẠN (UI → SERVER)
// ==============================
function uploadStageFile(region, school, stage, fileName, base64Data) {
  return globalThis.uploadStageFile(region, school, stage, fileName, base64Data);
}

// ==============================
// 🧠 HÀM TRUNG GIAN: TẠO NHẬN XÉT & SO SÁNH (UI → SERVER)
// ==============================
function generateAllReports(region, school, gd1Id, gd2Id) {
  return globalThis.generateAllReports(region, school, gd1Id, gd2Id);
}

// ==============================
// 🤖 PHÂN TÍCH & TẠO NHẬN XÉT AI
// ==============================
function generateAISummary({ region, schoolName, numGV, numHS, totalTasksCreated, totalTasksAssigned }) {
  try {
    let summary = [];
    const schoolType = schoolName.toLowerCase().includes("thpt") ? "THPT" :
                       schoolName.toLowerCase().includes("thcs") ? "THCS" :
                       "Trường";

    // --- Giáo viên ---
    if (numGV === 0) {
      summary.push(`Chưa có giáo viên nào của ${schoolType} ${schoolName} sử dụng Onluyen để giao bài.`);
    } else if (numGV < 5) {
      summary.push(`Số lượng giáo viên sử dụng Onluyen tại ${schoolType} ${schoolName} còn hạn chế (${numGV} GV).`);
    } else if (numGV < 15) {
      summary.push(`Khoảng ${numGV} giáo viên đang sử dụng Onluyen, mức độ tham gia ở mức trung bình.`);
    } else {
      summary.push(`Rất tích cực! Có tới ${numGV} giáo viên đã sử dụng Onluyen để giao bài tập cho học sinh.`);
    }

    // --- Học sinh ---
    if (numHS === 0) {
      summary.push("Hiện chưa có học sinh nào làm bài trên hệ thống.");
    } else if (numHS < 50) {
      summary.push(`Số học sinh tham gia làm bài còn khiêm tốn (${numHS} HS), cần đẩy mạnh hoạt động giao bài và khuyến khích HS tham gia.`);
    } else if (numHS < 200) {
      summary.push(`Khoảng ${numHS} học sinh đã tham gia làm bài, thể hiện mức độ triển khai khá ổn định.`);
    } else {
      summary.push(`Tuyệt vời! ${numHS} học sinh đã làm bài trên Onluyen, cho thấy mức độ sử dụng rộng rãi trong toàn trường.`);
    }

    // --- Bài tập ---
    if (totalTasksCreated === 0 && totalTasksAssigned === 0) {
      summary.push("Chưa có dữ liệu bài tập nào được tạo hoặc giao trong giai đoạn này.");
    } else {
      const ratio = totalTasksAssigned && totalTasksCreated
        ? (totalTasksAssigned / totalTasksCreated * 100).toFixed(1)
        : 0;
      summary.push(`Tổng cộng ${totalTasksCreated} bài tập đã được tạo, trong đó ${totalTasksAssigned} bài đã được giao (${ratio}% bài được sử dụng).`);
      if (ratio < 40) {
        summary.push("Tỷ lệ bài tập được giao còn thấp — cần khuyến khích GV tận dụng kho bài đã tạo để giao cho HS.");
      } else if (ratio < 80) {
        summary.push("Tỷ lệ bài tập được giao ở mức khá tốt, có thể tiếp tục cải thiện để tăng mức độ hoạt động của GV.");
      } else {
        summary.push("Rất hiệu quả — hầu hết các bài tập được tạo đã được giao đến học sinh.");
      }
    }

    // --- Tổng kết ---
    const overallScore = numGV * 0.4 + numHS * 0.3 + totalTasksAssigned * 0.3;
    const scoreInfo = getAIScoreLevel(overallScore);
    summary.push(`➡️ **Đánh giá tổng quan:** mức độ sử dụng Onluyen tại ${schoolType} ${schoolName} đang ở mức **${scoreInfo.level.toUpperCase()}**.`);

    return summary.join("<br><br>");
  } catch (err) {
    Logger.log("⚠️ Lỗi generateAISummary: " + err);
    return "(Không thể tạo nhận xét AI do lỗi nội bộ)";
  }
}

// ==============================
// 🎯 ĐIỂM PHÂN LOẠI MỨC ĐỘ
// ==============================
function getAIScoreLevel(score) {
  if (score > 1000) return { level: "rất cao", color: "#00C853" };
  if (score > 500) return { level: "khá tốt", color: "#64DD17" };
  if (score > 200) return { level: "trung bình", color: "#FFD600" };
  return { level: "thấp", color: "#FF3D00" };
}

// ==============================
// 📤 HÀM PHỤ TRẢ VỀ KẾT QUẢ CHO UI
// ==============================
function getAISummaryFromData(data) {
  try {
    const summaryText = generateAISummary(data);
    return {
      status: "success",
      summary: summaryText,
      timestamp: new Date().toLocaleString("vi-VN"),
    };
  } catch (e) {
    return {
      status: "error",
      summary: "Không thể tạo nhận xét. Vui lòng kiểm tra lại dữ liệu đầu vào.",
      error: e.toString(),
    };
  }
}
