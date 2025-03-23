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
  let worksheet = siteWorkbook.Sheets[sheet];

  // 🔄 site.xlsx에서 못 찾으면 Part.xlsx에서 찾기
  if (!worksheet) {
    worksheet = partWorkbook.Sheets[sheet];
  }

  if (!worksheet) {
    return res.status(404).json({ error: `Sheet '{sheet}' not found.` });
  }

  const jsonData = xlsx.utils.sheet_to_json(worksheet, { defval: "" });
  const matchedRow = jsonData.find((row) =>
    Object.values(row).some((cell) =>
      String(cell).trim().includes(decodeURIComponent(value))
    )
  );

  if (!matchedRow) {
    return res.status(404).json({ error: `'{value}' not found in sheet '{sheet}'.` });
  }

  res.json(matchedRow);
});

// ----------------------------
// 🚀 서버 시작
// ----------------------------
app.listen(PORT, () => {
  console.log(`🛰️  Server running on http://localhost:${PORT}`);
});