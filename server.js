const express = require("express");
const cors = require("cors");
const xlsx = require("xlsx");
const path = require("path");
const fs = require("fs");
const basicAuth = require("basic-auth");
const { execSync } = require("child_process"); // ✅ Git 커맨드용

const app = express();
const PORT = process.env.PORT || 8080;

app.use(cors());
app.use(express.json());

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

  const matchedRow = jsonData.filter((row) =>
    Object.values(row).some((v) =>
      String(v).toLowerCase().includes(decodeURIComponent(value).toLowerCase())
    )
  );

  if (!matchedRow || matchedRow.length === 0) {
    return res.status(404).json({ error: `'${value}' not found in sheet '${sheet}'.` });
  }

  // ✅ usage.json에서 Remark 덮어쓰기
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
    return res.json(matchedRow[0]); // 단일
  }
});

// ✅ 사용 기록 저장 API (Git 커밋 포함)
app.post("/api/save-usage", express.json(), (req, res) => {
  const newRecord = req.body;
  const usageFilePath = path.join(__dirname, "assets", "usage.json");

  try {
    let existingData = [];

    if (fs.existsSync(usageFilePath)) {
      const raw = fs.readFileSync(usageFilePath, "utf-8");
      existingData = JSON.parse(raw);
    }

    const updatedData = [
      ...existingData.filter(
        (item) => !(item.Part === newRecord.Part && item.Serial === newRecord.Serial)
      ),
      newRecord,
    ];

    fs.writeFileSync(usageFilePath, JSON.stringify(updatedData, null, 2), "utf-8");
    console.log("✅ usage.json 저장 완료:", newRecord);
    // ✅ Git 사용자 정보 자동 설정
    try {
      execSync('git config user.email "keyower159@gmail.com"');
      execSync('git config user.name "BBIOK-SERVER"');
    } catch (err) {
      console.error("❌ Git 사용자 정보 설정 실패:", err.message);
    }

    // ✅ 원격 저장소 origin 등록 (이미 등록된 경우 무시)
    try {
      execSync('git remote add origin https://github.com/Hyunsu7917/BRKR-SERVER.git');
    } catch (err) {
      if (!err.message.includes("remote origin already exists")) {
        console.error("❌ Git remote 설정 실패:", err.message);
      }
    }

    // ✅ 변경 사항 커밋 및 푸시
    try {
      execSync('git add assets/usage.json');
      execSync('git commit -m "📝 usage 기록: ' + newRecord.Part + ' ' + newRecord.Serial + '"');
      execSync('git push origin main');
      console.log("🚀 Git push 완료");
    } catch (pushErr) {
      console.error("❌ Git push 실패:", pushErr.message);
    }


    // ✅ Git commit & push
    execSync("git add assets/usage.json");
    execSync(`git commit -m "📝 usage 기록: ${newRecord.Part} ${newRecord.Serial}"`);
    execSync("git push");

    res.json({ success: true, message: "사용 기록 저장 및 커밋 완료" });
  } catch (err) {
    console.error("❌ usage.json 저장 또는 커밋 실패:", err);
    res.status(500).json({ success: false, error: "서버 저장 또는 Git 오류 발생" });
  }
});

// usage.json 조회용
app.get("/api/usage", (req, res) => {
  const usageFilePath = path.join(__dirname, "assets", "usage.json");
  if (!fs.existsSync(usageFilePath)) {
    return res.json([]);
  }
  const data = fs.readFileSync(usageFilePath, "utf-8");
  res.json(JSON.parse(data));
});

// 서버 실행
app.listen(PORT, () => {
  console.log(`🚀 Server running on port ${PORT}`);
});
