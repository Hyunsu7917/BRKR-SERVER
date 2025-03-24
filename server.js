// ✅ server.js — Part.xlsx를 단일 원본으로 사용하는 버전
const express = require("express");
const fs = require("fs");
const path = require("path");
const cors = require("cors");
const basicAuth = require("express-basic-auth");
const ExcelJS = require("exceljs");
const xlsx = require("xlsx");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());
app.use(
  basicAuth({
    users: { BBIOK: "Bruker_2025" },
    challenge: true,
  })
);

// ✅ 엑셀 → JSON 변환 API (전체)
app.get("/excel/part/all", async (req, res) => {
  try {
    const filePath = path.join(__dirname, "assets", "Part.xlsx");
    if (!fs.existsSync(filePath)) return res.status(404).json({ error: "파일 없음" });

    const workbook = xlsx.readFile(filePath);
    const worksheet = workbook.Sheets["part"];

    const jsonData = xlsx.utils.sheet_to_json(worksheet, {
      range: 1, // A2부터 시작
      defval: "",
      header: ["Part#", "Serial #", "PartName", "Remark", "사용처", "Rack", "Count"],
    });

    res.json(jsonData);
  } catch (e) {
    console.error("❌ /excel/part/all 실패:", e.message);
    res.status(500).json({ error: "서버 오류" });
  }
});

// ✅ 특정 부품명 검색
// ✅ 특정 부품명 검색 (단일 또는 다중 결과 반환 구분)
app.get("/excel/part/:value", async (req, res) => {
  try {
    const value = decodeURIComponent(req.params.value);
    const filePath = path.join(__dirname, "assets", "Part.xlsx");
    if (!fs.existsSync(filePath)) return res.status(404).json({ error: "파일 없음" });

    const workbook = xlsx.readFile(filePath);
    const worksheet = workbook.Sheets["part"];

    const jsonData = xlsx.utils.sheet_to_json(worksheet, {
      range: 1,
      defval: "",
      header: ["Part#", "Serial #", "PartName", "Remark", "사용처", "Rack", "Count"],
    });

    const matchedRow = jsonData.filter(
      (row) =>
        row["Part#"]?.toLowerCase() === value.toLowerCase() ||
        row["PartName"]?.toLowerCase() === value.toLowerCase()
    );

    if (matchedRow.length === 1) {
      return res.json(matchedRow[0]); // 단일 객체 반환 (사이트플랜 등)
    } else {
      return res.json(matchedRow); // 배열 전체 반환 (리스트용)
    }
  } catch (e) {
    console.error("❌ /excel/part/:value 실패:", e.message);
    res.status(500).json({ error: "서버 오류" });
  }
});


// ✅ 사용 기록 저장 → Part.xlsx 직접 반영
app.post("/api/save-usage", async (req, res) => {
  try {
    const { Part, Serial, Remark, UsageNote } = req.body;
    const filePath = path.join(__dirname, "assets", "Part.xlsx");

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet("part");
    if (!sheet) return res.status(404).json({ error: "시트 없음" });

    // 헤더 인덱스 찾기
    const headerRow = sheet.getRow(1);
    const headers = headerRow.values.map((v) => (typeof v === "string" ? v.trim() : v));
    const partIdx = headers.indexOf("Part#");
    const serialIdx = headers.indexOf("Serial #");
    const remarkIdx = headers.indexOf("Remark");
    const usageIdx = headers.indexOf("사용처");

    if (partIdx === -1 || serialIdx === -1 || remarkIdx === -1 || usageIdx === -1) {
      return res.status(400).json({ error: "헤더 오류" });
    }

    let updated = false;
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;

      const part = String(row.getCell(partIdx + 1).value || "").trim();
      const serial = String(row.getCell(serialIdx + 1).value || "").trim();

      if (part === Part && serial === Serial) {
        row.getCell(remarkIdx + 1).value = Remark || "";
        row.getCell(usageIdx + 1).value = UsageNote || "";
        updated = true;
      }
    });

    if (!updated) {
      return res.status(404).json({ error: "일치하는 항목 없음" });
    }

    await workbook.xlsx.writeFile(filePath);
    res.json({ success: true });
  } catch (err) {
    console.error("❌ 사용기록 저장 실패:", err.message);
    res.status(500).json({ error: "서버 오류" });
  }
});

// ✅ 서버 실행
app.listen(PORT, () => {
  console.log(`🚀 Server running on port ${PORT}`);
});
