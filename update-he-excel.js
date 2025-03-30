
const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

async function updateHeExcel() {
  const backupPath = path.join(__dirname, "he-usage-backup.json");
  const filePath = path.join(__dirname, "assets", "He.xlsx");

  if (!fs.existsSync(backupPath)) {
    console.error("❌ he-usage-backup.json not found.");
    return;
  }

  const raw = fs.readFileSync(backupPath, "utf-8").trim();
  const records = raw ? JSON.parse(raw) : [];

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const sheet1 = workbook.getWorksheet("일정");
  const sheet2 = workbook.getWorksheet("기록");

  const headerRow1 = sheet2.getRow(1);
  const headerRow2 = sheet2.getRow(2);
  const headerRow3 = sheet2.getRow(3);

  const usageByKey = {};

  records.forEach((record) => {
    const key = `${record["고객사"]}||${record["지역"]}||${record["Magnet"]}`;
    if (!usageByKey[key]) usageByKey[key] = [];
    usageByKey[key].push(record["충진일"]);
  });

  Object.entries(usageByKey).forEach(([key, dateList]) => {
    const [customer, region, magnet] = key.split("||");
    let col = -1;

    for (let i = 2; i <= sheet2.columnCount; i++) {
      const name = String(headerRow1.getCell(i).value ?? "").trim();
      const reg = String(headerRow2.getCell(i).value ?? "").trim();
      const mag = String(headerRow3.getCell(i).value ?? "").trim();

      if (name === customer && reg === region && mag === magnet) {
        col = i;
        break;
      }
    }

    if (col !== -1) {
      let rowIndex = 4;
      for (let date of dateList) {
        sheet2.getCell(rowIndex++, col).value = date;
      }
      console.log(`📌 기록 시트 업데이트: ${customer}, ${region}, ${magnet}`);
    } else {
      console.warn(`❗ 기록 시트에 ${customer} / ${region} / ${magnet} 찾을 수 없음`);
    }
  });

  await workbook.xlsx.writeFile(filePath);
  console.log("✅ He.xlsx 업데이트 완료!");
}

updateHeExcel();
