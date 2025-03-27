const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

async function updateHeExcel() {
  const jsonPath = path.join(__dirname, "he-usage-backup.json");
  const excelPath = path.join(__dirname, "assets", "He.xlsx");

  if (!fs.existsSync(jsonPath)) {
    console.error("❌ he-usage-backup.json 파일이 없습니다.");
    return;
  }

  const jsonData = JSON.parse(fs.readFileSync(jsonPath, "utf-8"));
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(excelPath);

  const sheet1 = workbook.getWorksheet("일정");
  const sheet2 = workbook.getWorksheet("기록");

  const headerRow1 = sheet2.getRow(1); // 고객사
  const headerRow2 = sheet2.getRow(2); // 지역
  const headerRow3 = sheet2.getRow(3); // Magnet

  for (const record of jsonData) {
    const { 고객사, 지역, Magnet, 충진일, 다음충진일, "충진주기(개월)": 주기 } = record;

    // ✅ 일정 시트 업데이트
    const rowToUpdate = sheet1.findRow((row) => {
      const [cust, reg, mag] = [
        String(row.getCell(1).value || "").trim(),
        String(row.getCell(2).value || "").trim(),
        String(row.getCell(3).value || "").trim()
      ];
      return cust === 고객사 && reg === 지역 && mag === Magnet;
    });

    if (rowToUpdate) {
      rowToUpdate.getCell(4).value = 충진일;
      rowToUpdate.getCell(5).value = 다음충진일;
      rowToUpdate.getCell(6).value = 주기;
      console.log(`✅ 일정 업데이트 완료: ${고객사} / ${지역} / ${Magnet}`);
    } else {
      console.warn(`❌ 일정 시트에서 못 찾음: ${고객사}, ${지역}, ${Magnet}`);
    }

    // ✅ 기록 시트 업데이트
    let targetCol = -1;
    for (let col = 2; col <= headerRow1.cellCount; col++) {
      const cName = String(headerRow1.getCell(col).value || "").trim();
      const cRegion = String(headerRow2.getCell(col).value || "").trim();
      const cMagnet = String(headerRow3.getCell(col).value || "").trim();
      if (cName === 고객사 && cRegion === 지역 && cMagnet === Magnet) {
        targetCol = col;
        break;
      }
    }

    if (targetCol !== -1) {
      let rowIndex = 4;
      while (sheet2.getCell(rowIndex, targetCol).value) rowIndex++;
      sheet2.getCell(rowIndex, targetCol).value = 충진일;
      console.log(`📌 기록 추가: ${고객사} / ${지역} / ${Magnet} → ${rowIndex}행`);
    } else {
      console.warn(`❌ 기록 시트에서 못 찾음: ${고객사}, ${지역}, ${Magnet}`);
    }
  }

  await workbook.xlsx.writeFile(excelPath);
  console.log("🚀 He.xlsx 업데이트 완료!");
}

updateHeExcel();
