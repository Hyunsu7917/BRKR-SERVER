// 📦 필요한 모듈 불러오기
const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

// 🛠 실제 작업 함수
async function updateHeExcel() {
  try {
    // 1. 백업 파일 읽기
    const data = JSON.parse(fs.readFileSync("he-usage-backup.json", "utf-8"));
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile("assets/He.xlsx");

    // 2. 일정 시트
    const sheet1 = workbook.getWorksheet("일정");
    const rows = sheet1.getRows(2, sheet1.rowCount - 1);

    data.forEach((record) => {
      const customer = record["고객사"]?.trim();
      const region = record["지역"]?.trim();
      const magnet = record["Magnet"]?.trim();
      const chargeDate = record["충진일"];
      const nextCharge = record["다음충진일"];
      const cycle = record["충진주기(개월)"];

      const matched = rows.find((row) => {
        const c = row.getCell(1).value?.toString().trim();
        const r = row.getCell(2).value?.toString().trim();
        const m = row.getCell(3).value?.toString().trim();
        return c === customer && r === region && m === magnet;
      });

      if (matched) {
        matched.getCell(4).value = chargeDate;
        matched.getCell(5).value = nextCharge;
        matched.getCell(6).value = cycle;
        console.log(`✅ 일정 시트 업데이트: ${customer}, ${region}, ${magnet}`);
      } else {
        console.log(`❌ 일정 시트에서 못 찾음: ${customer}, ${region}, ${magnet}`);
      }
    });

    // 3. 기록 시트
    const sheet2 = workbook.getWorksheet("기록");
    const headerRow1 = sheet2.getRow(1);
    const headerRow2 = sheet2.getRow(2);
    const headerRow3 = sheet2.getRow(3);

    data.forEach((record) => {
      const customer = record["고객사"]?.trim();
      const region = record["지역"]?.trim();
      const magnet = record["Magnet"]?.trim();
      const chargeDate = record["충진일"];

      let targetCol = -1;
      for (let i = 2; i <= sheet2.columnCount; i++) {
        const c = String(headerRow1.getCell(i).value || "").trim();
        const r = String(headerRow2.getCell(i).value || "").trim();
        const m = String(headerRow3.getCell(i).value || "").trim();
        if (c === customer && r === region && m === magnet) {
          targetCol = i;
          break;
        }
      }

      if (targetCol !== -1) {
        let rowIndex = 4;
        while (sheet2.getCell(rowIndex, targetCol).value) {
          rowIndex++;
        }

        // ✅ 중복 방지
        const lastValue = sheet2.getCell(rowIndex - 1, targetCol).value;
        if (lastValue !== chargeDate) {
          sheet2.getCell(rowIndex, targetCol).value = chargeDate;
          console.log(`✅ 기록 추가: ${customer} / ${region} / ${magnet} ➜ ${rowIndex}행`);
        } else {
          console.log(`⚠️ 기록 중복 생략: ${customer} / ${region} / ${magnet}`);
        }
      } else {
        console.log(`❗ 기록 시트에서 못 찾음: ${customer} / ${region} / ${magnet}`);
      }
    });

    // 4. 저장
    await workbook.xlsx.writeFile("assets/He.xlsx");
    console.log("🚀 He.xlsx 업데이트 완료!");
  } catch (err) {
    console.error("💥 업데이트 중 오류 발생:", err);
  }
}

// ✅ 함수 실행
updateHeExcel();
