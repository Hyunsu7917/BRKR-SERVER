// 수정 및 보완된 server.js 라우팅 포함 버전
const express = require("express");
const path = require("path");
const fs = require("fs");
const xlsx = require("xlsx");
const basicAuth = require("express-basic-auth");
const cors = require("cors");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());

app.use(
  basicAuth({
    users: { BBI0K: "Bruker_2025" },
    challenge: true,
  })
);

// ✅ Part 전체 데이터 (리스트용)
app.get("/excel/part/all", (req, res) => {
  const filePath = path.join(__dirname, "assets", "Part.xlsx");
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: "파일 없음" });

  const workbook = xlsx.readFile(filePath);
  const worksheet = workbook.Sheets["part"];
  const jsonData = xlsx.utils.sheet_to_json(worksheet, { defval: "" });
  return res.json(jsonData);
});

// ✅ Part 특정 항목 조회 (value)
app.get("/excel/part/value", (req, res) => {
  const { part, serial } = req.query;
  const filePath = path.join(__dirname, "assets", "Part.xlsx");
  const workbook = xlsx.readFile(filePath);
  const sheet = workbook.Sheets["part"];
  const data = xlsx.utils.sheet_to_json(sheet, { defval: "" });
  const matched = data.filter(
    (row) => String(row["Part#"]).trim() === part && String(row["Serial #"]).trim() === serial
  );
  if (matched.length === 1) return res.json(matched[0]);
  else return res.json(matched);
});

// ✅ 항목별 정보 (Magnet 등)
app.get("/excel/:sheet/:value", (req, res) => {
  const { sheet, value } = req.params;
  const filePath = path.join(__dirname, "assets", "site.xlsx");
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: "site.xlsx 없음" });

  const workbook = xlsx.readFile(filePath);
  const worksheet = workbook.Sheets[sheet];
  if (!worksheet) return res.status(404).json({ error: `시트 ${sheet} 없음` });

  const jsonData = xlsx.utils.sheet_to_json(worksheet, { defval: "" });
  const matchedRow = jsonData.filter((row) => Object.values(row).includes(value));
  if (matchedRow.length === 1) return res.json(matchedRow[0]);
  else return res.json(matchedRow);
});

// ✅ usage.json 조회
app.get("/api/usage-json", (req, res) => {
  const usageFilePath = path.join(__dirname, "assets", "usage.json");
  if (!fs.existsSync(usageFilePath)) return res.json([]);
  const data = fs.readFileSync(usageFilePath, "utf-8");
  res.json(JSON.parse(data));
});

// ✅ usage 저장
app.post("/api/save-usage", (req, res) => {
  const usageFilePath = path.join(__dirname, "assets", "usage.json");
  const newRecord = req.body;
  let usageData = [];

  if (fs.existsSync(usageFilePath)) {
    usageData = JSON.parse(fs.readFileSync(usageFilePath, "utf-8"));
  }

  const idx = usageData.findIndex(
    (u) => u["Part#"] === newRecord["Part#"] && u["Serial #"] === newRecord["Serial #"]
  );
  if (idx !== -1) usageData[idx] = newRecord;
  else usageData.push(newRecord);

  fs.writeFileSync(usageFilePath, JSON.stringify(usageData, null, 2));
  res.json({ success: true });
});

app.listen(PORT, () => {
  console.log(`🚀 Server running on port ${PORT}`);
});