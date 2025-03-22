const express = require("express");
const cors = require("cors");
const xlsx = require("xlsx");
const path = require("path");
const basicAuth = require("basic-auth");

const app = express();
const PORT = process.env.PORT || 8080;

// CORS 허용
app.use(cors());

// 인증 미들웨어 설정
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

// 정적 파일 제공 및 인증 적용
app.use(auth);
app.use("/assets", express.static(path.join(__dirname, "assets")));

// 최신 버전 JSON 응답
app.get("/latest-version.json", (req, res) => {
  res.json({
    version: "1.0.1", // 최신 앱 버전
    apkUrl: "https://expo.dev/artifacts/eas/k3rjaa7axiyRhmqwGuH5Zb.apk" // 실제 APK 주소로 교체
  });
});

// Excel 파일 로딩
const workbook = xlsx.readFile(path.join(__dirname, "assets/site.xlsx"));

// Excel 데이터 API
app.get("/excel/:sheet/:value", (req, res) => {
  const { sheet, value } = req.params;
  const worksheet = workbook.Sheets[sheet];

  if (!worksheet) {
    return res.status(404).json({ error: `Sheet '${sheet}' not found.` });
  }

  const jsonData = xlsx.utils.sheet_to_json(worksheet, { defval: "" });

  const matchedRow = jsonData.find((row) => {
    const firstKey = Object.keys(row)[0];
    return String(row[firstKey]).trim() === decodeURIComponent(value);
  });

  if (!matchedRow) {
    return res
      .status(404)
      .json({ error: `'${value}' not found in sheet '${sheet}'.` });
  }

  res.json(matchedRow);
});

// 서버 시작
app.listen(PORT, () => {
  console.log(`🛰️  Server running on http://localhost:${PORT}`);
});
