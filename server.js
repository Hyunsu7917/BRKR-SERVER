const express = require("express");
const basicAuth = require("express-basic-auth");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");
//require("dotenv").config();

const app = express();
//const PORT = process.env.PORT || 3001;
const PORT = 3001;

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
  const matchedRow = jsonData.filter(row => String(row["Part#"]).toLowerCase() === value.toLowerCase());

  if (matchedRow.length === 1) {
    return res.json(matchedRow[0]);
  } else {
    return res.json(matchedRow);
  }
});

// ✅ usage.json 전체 조회
app.get("/api/usage", basicAuthMiddleware, (req, res) => {
  const usageFilePath = path.join(__dirname, "assets", "usage.json");
  if (!fs.existsSync(usageFilePath)) return res.json([]);

  const data = fs.readFileSync(usageFilePath, "utf-8");
  res.json(JSON.parse(data));
});

// ✅ usage.json 저장 (append or update)
app.post("/api/save-usage", basicAuthMiddleware, (req, res) => {
  const usageFilePath = path.join(__dirname, "assets", "usage.json");
  const newUsage = req.body;

  let current = [];
  if (fs.existsSync(usageFilePath)) {
    current = JSON.parse(fs.readFileSync(usageFilePath, "utf-8"));
  }

  const updated = current.filter(
    (item) => !(item["Part#"] === newUsage["Part#"] && item["Serial #"] === newUsage["Serial #"])
  );
  updated.push(newUsage);
  fs.writeFileSync(usageFilePath, JSON.stringify(updated, null, 2));
  res.json({ success: true });
});

// ✅ usage.json → Part.xlsx 반영
app.post("/api/sync-usage-to-excel", basicAuthMiddleware, (req, res) => {
  const partPath = path.join(__dirname, "assets", "Part.xlsx");
  const usagePath = path.join(__dirname, "assets", "usage.json");

  if (!fs.existsSync(partPath)) return res.status(404).json({ error: "Part.xlsx 없음" });
  if (!fs.existsSync(usagePath)) return res.status(404).json({ error: "usage.json 없음" });

  const workbook = xlsx.readFile(partPath);
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  const usageList = JSON.parse(fs.readFileSync(usagePath, "utf-8"));
  const data = xlsx.utils.sheet_to_json(worksheet, { defval: "" });

  usageList.forEach((usage) => {
    const matchIndex = data.findIndex(
      (row) => row["Part#"] === usage["Part#"] && row["Serial #"] === usage["Serial #"]
    );
    if (matchIndex !== -1) {
      data[matchIndex]["Remark"] = usage.Remark;
      data[matchIndex]["사용처"] = usage.UsageNote;
    }
  });

  const newWorksheet = xlsx.utils.json_to_sheet(data);
  workbook.Sheets[workbook.SheetNames[0]] = newWorksheet;
  xlsx.writeFile(workbook, partPath);

  res.json({ success: true });
});

// ✅ 서버 시작
app.listen(PORT, () => {
  console.log(`🚀 Server running on port ${PORT}`);
});