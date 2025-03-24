const express = require("express");
const basicAuth = require("express-basic-auth");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");

const app = express();
const PORT = process.env.PORT || 3001;

app.use(cors());
app.use(express.json());

// 🔐 Basic Auth 설정
const basicAuthMiddleware = basicAuth({
  users: { BBIOK: "Bruker_2025" },
  challenge: true,
});

// ✅ 국내 재고 전체 조회 (Part.xlsx)
app.get("/excel/part/all", basicAuthMiddleware, (req, res) => {
  const filePath = path.join(__dirname, "assets", "Part.xlsx");
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: "파일 없음" });

  const workbook = xlsx.readFile(filePath);
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const jsonData = xlsx.utils.sheet_to_json(worksheet, { defval: "" });
  res.json(jsonData);
});

// ✅ 국내 재고 Part# 검색
app.get("/excel/part/value/:value", basicAuthMiddleware, (req, res) => {
  const { value } = req.params;
  const filePath = path.join(__dirname, "assets", "Part.xlsx");
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: "파일 없음" });

  const workbook = xlsx.readFile(filePath);
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const jsonData = xlsx.utils.sheet_to_json(worksheet, { defval: "" });
  const matchedRow = jsonData.filter(row => String(row["Part#"]).toLowerCase() === value.toLowerCase());

  if (matchedRow.length === 1) {
    return res.json(matchedRow[0]);
  } else {
    return res.json(matchedRow);
  }
});

// ✅ 항목별 정리 (site.xlsx - Magnet, Console 등)
app.get("/excel/:sheet/value/:value", basicAuthMiddleware, (req, res) => {
  const { sheet, value } = req.params;
  const filePath = path.join(__dirname, "assets", "site.xlsx");
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: "파일 없음" });

  const workbook = xlsx.readFile(filePath);
  const worksheet = workbook.Sheets[sheet];
  if (!worksheet) return res.status(404).json({ error: `시트 ${sheet} 없음` });

  const jsonData = xlsx.utils.sheet_to_json(worksheet, { defval: "" });
  const firstCol = Object.keys(jsonData[0])[0]; // ✅ 첫 번째 열 이름 가져오기
  const matchedRow = jsonData.filter(row => String(row[firstCol]).toLowerCase() === value.toLowerCase());


  if (matchedRow.length === 1) {
    return res.json(matchedRow[0]);
  } else {
    return res.json(matchedRow);
  }
});

// ✅ 서버 시작
app.listen(PORT, () => {
  console.log(`🚀 Server running on port ${PORT}`);
});
