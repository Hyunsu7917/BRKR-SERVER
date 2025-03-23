const express = require("express");
const cors = require("cors");
const xlsx = require("xlsx");
const path = require("path");
const fs = require("fs");
const basicAuth = require("basic-auth");

const app = express();
const PORT = process.env.PORT || 8080;

app.use(cors());
app.use(express.json()); // ✅ JSON 파싱 추가

// 버전 정보
const versionFilePath = path.join(__dirname, "version.json");
let versionData = { version: "1.0.0", apkUrl: "" };

if (fs.existsSync(versionFilePath)) {
  try {
    versionData = JSON.parse(fs.readFileSync(versionFilePath, "utf-8"));
  } catch (err) {
    console.error("Failed to parse version.json:", err);
  }
}

// 인증
const auth = (req, res, next) => {
  const user = basicAuth(req);
  const isAuthorized = user && user.name === "BBIOK" && user.pass === "Bruker_2025";
  if (!isAuthorized) {
    res.set("WWW-Authenticate", 'Basic realm="Authorization Required"');
    return res.status(401).send("Access denied");
  }
  next();
};
app.use(auth);

// 정적 파일
app.use("/assets", express.static(path.join(__dirname, "assets")));

// 버전 정보
app.get("/latest-version.json", (req, res) => {
  res.json(versionData);
});

// 엑셀 조회 API
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

  // 🔍 검색 필터
  const matchedRow = jsonData.filter((row) =>
    Object.values(row).some((v) =>
      String(v).toLowerCase().includes(decodeURIComponent(value).toLowerCase())
    )
  );

  if (!matchedRow || matchedRow.length === 0) {
    return res.status(404).json({ error: `'${value}' not found in sheet '${sheet}'.` });
  }

  // ✅ usage.json에서 Remark 덮어쓰기 (Part.xlsx 전용)
  if (filePath.includes("Part.xlsx")) {
    try {
      const usageData = JSON.parse(
        fs.readFileSync(path.join(__dirname, "assets", "usage.json"), "utf-8")
      );

      matchedRow.forEach((row) => {
        const match = usageData.find(
          (u) => u.Part === row["Part#"] && u.Serial === row["Serial #"]
        );
        if (match) {
          row["Remark"] = match.Remark;
        }
      });
    } catch (e) {
      console.warn("⚠️ usage.json 불러오기 실패:", e.message);
    }

    return res.json(matchedRow); // 배열 전체 반환
  } else {
    return res.json(matchedRow[0]); // 사이트플랜은 단일
  }
});

// ✅ 사용 기록 저장 API
// ✅ usage.json 저장 API (Part + Serial 기준 병합 저장)
app.post("/api/save-usage", express.json(), (req, res) => {
  const newRecord = req.body;
  const usageFilePath = path.join(__dirname, "assets", "usage.json");

  try {
    let existingData = [];

    if (fs.existsSync(usageFilePath)) {
      const raw = fs.readFileSync(usageFilePath, "utf-8");
      existingData = JSON.parse(raw);
    }

    // 동일 Part + Serial이 있다면 덮어쓰기
    const updatedData = [
      ...existingData.filter(
        (item) => !(item.Part === newRecord.Part && item.Serial === newRecord.Serial)
      ),
      newRecord,
    ];

    fs.writeFileSync(usageFilePath, JSON.stringify(updatedData, null, 2), "utf-8");
    console.log("✅ usage.json 저장 완료:", newRecord);
    res.json({ success: true, message: "사용 기록 저장 완료" });
  } catch (err) {
    console.error("❌ usage.json 저장 실패:", err);
    res.status(500).json({ success: false, error: "서버 저장 중 오류 발생" });
  }
});

// 서버 실행
app.listen(PORT, () => {
  console.log(`🚀 Server running on port ${PORT}`);
});