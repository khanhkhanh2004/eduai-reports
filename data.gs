// Đọc bảng dữ liệu từ link Google Sheet (đã convert), mặc định lấy sheet tên "HocSinh" nếu có, nếu không lấy sheet đầu.
function readDataFromFile(url) {
  const id = (url || "").match(/[-\w]{25,}/)[0];
  const ss = SpreadsheetApp.openById(id);
  const sht = ss.getSheetByName("HocSinh") || ss.getSheets()[0];
  const values = sht.getDataRange().getValues();
  const header = values[0].map(h => (h || "").toString().trim().toLowerCase());
  const rows = values.slice(1);

  // Tìm các cột “bài làm” / “hoàn thành” một cách linh hoạt
  const colDoneIdx = header.findIndex(h => /bài.*làm|hoàn.*thành|completed/.test(h));
  // Nếu không có, tạm coi “đã có dữ liệu” = 1
  const doneCount = rows.filter(r => {
    if (colDoneIdx >= 0) return Number(r[colDoneIdx]) > 0;
    return r.some(x => x !== "" && x !== null); // fallback
  }).length;

  return { rows, header, total: rows.length, done: doneCount };
}

// So sánh hai giai đoạn → trả về object tổng hợp
function compareStages(gd1, gd2) {
  const total1 = gd1.total || 0;
  const total2 = gd2.total || 0;
  const done1  = gd1.done  || 0;
  const done2  = gd2.done  || 0;

  const p1 = total1 ? (done1 / total1) * 100 : 0;
  const p2 = total2 ? (done2 / total2) * 100 : 0;

  return {
    total1, total2, done1, done2,
    percent1: p1.toFixed(1),
    percent2: p2.toFixed(1),
    diff: (p2 - p1).toFixed(1)
  };
}
