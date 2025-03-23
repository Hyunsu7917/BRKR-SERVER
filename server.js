// server.js
const express = require("express");
const cors = require("cors");
const xlsx = require("xlsx");
const path = require("path");
const fs = require("fs");
const basicAuth = require("basic-auth");

const app = express();
const PORT = process.env.PORT || 8080;

// ----------------------------
// 🧠 설정 파일에서 버전 자동 로드
// ----------------------------
const versionFilePath = path.join(__dirname, "version.json");
let versionData = { version: "1.0.0", apkUrl: "" };

if (fs.existsSync(versionFilePath)) {
  try {
    versionData = JSON.parse(fs.readFileSync(versionFilePath, "utf-8"));
  } catch (err) {
    console.error("Failed to parse version.json:", err);
  }
}

// ----------------------------
// 🌐 CORS 허용
// ----------------------------
app.use(cors());

// ----------------------------
// 🔐 인증 미들웨어 설정
// ----------------------------
const auth = (req, res, next) => {
  const user = basicAuth(req);
  const isAuthorized =
    user && user.name === "BBIOK" && user.pass === "Bruker_2025";

  if (!isAuthorized) {
    res.set("WWW-Authenticate", 'Basic realm="Authorization Required"');
    return res.status(401).send("Access denied");
  }
  next();
};

// ----------------------------
// 📦 정적 파일 제공 및 인증
// ----------------------------
app.use(auth);
app.use("/assets", express.static(path.join(__dirname, "assets")));

// ----------------------------
// 📤 최신 버전 정보 제공 API
// ----------------------------
app.get("/latest-version.json", (req, res) => {
  res.json(versionData);
});

// ----------------------------
// 📊 site.xlsx 불러오기
// ----------------------------
const siteWorkbook = xlsx.readFile(path.join(__dirname, "assets/site.xlsx"));

// ----------------------------
// 📊 Part.xlsx 불러오기
// ----------------------------
const partWorkbook = xlsx.readFile(path.join(__dirname, "assets/Part.xlsx"));

// ----------------------------
// 📊 Excel 데이터 조회 API
// ----------------------------
app.get("/excel/:sheet/:value", (req, res) => {
  const { sheet, value } = req.params;

  const filePath =
    sheet.toLowerCase() === "part"
      ? path.join(__dirname, "assets", "Part.xlsx")
      : path.join(__dirname, "assets", "site.xlsx");

  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ error: `File not found.` });
  }

  const workbook = xlsx.readFile(filePath);
  const worksheet = workbook.Sheets[sheet];

  if (!worksheet) {
    return res.status(404).json({ error: `Sheet '${sheet}' not found.` });
  }

  const jsonData = xlsx.utils.sheet_to_json(worksheet, { defval: "" });

  // 🔍 공백 제거 + 부분 매칭 필터
  const matchedRow = jsonData.filter((row) => {
    return Object.values(row).some((v) =>
      String(v).trim().toLowerCase().includes(decodeURIComponent(value).toLowerCase())
    );
  });

  if (!matchedRow || matchedRow.length === 0) {
    return res.status(404).json({ error: `'${value}' not found in sheet '${sheet}'.` });
  }

  // ✅ 파일 기준으로 분리 처리
  if (filePath.includes("Part.xlsx")) {
    res.json(matchedRow); // 국내 재고 → 여러 개
  } else {
    res.json(matchedRow[0]); // 사이트플랜 → 첫 번째만
  }
});
