const express = require("express");
const cors = require("cors");
const xlsx = require("xlsx");
const path = require("path");
const fs = require("fs");
const basicAuth = require("basic-auth");
const { execSync } = require("child_process");
const ExcelJS = require("exceljs");

app.post("/api/sync-usage-to-excel", async (req, res) => {
  try {
    const usagePath = path.join(__dirname, "assets", "usage.json");
    const excelPath = path.join(__dirname, "assets", "Part.xlsx");

    if (!fs.existsSync(usagePath) || !fs.existsSync(excelPath)) {
      return res.status(404).json({ error: "파일이 존재하지 않습니다." });
    }

    const usageData = JSON.parse(fs.readFileSync(usagePath, "utf-8"));

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelPath);

    const worksheet = workbook.getWorksheet("part");
    if (!worksheet) {
      return res.status(404).json({ error: "시트 'part'를 찾을 수 없습니다." });
    }

    // 헤더 인식
    const headerRow = worksheet.getRow(1);
    const headers = headerRow.values.map((v) => (typeof v === "string" ? v.trim() : v));
    const partIdx = headers.indexOf("Part#");
    const serialIdx = headers.indexOf("Serial #");
    const remarkIdx = headers.indexOf("Remark");
    const usageIdx = headers.indexOf("사용처");

    if (partIdx === -1 || serialIdx === -1 || remarkIdx === -1 || usageIdx === -1) {
      return res.status(400).json({ error: "필수 열이 누락되었습니다." });
    }

    // 데이터 반영
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // skip header

      const part = row.getCell(partIdx + 1).value;
      const serial = row.getCell(serialIdx + 1).value;

      const match = usageData.find(
        (u) => u.Part == part && u.Serial == serial
      );

      if (match) {
        row.getCell(remarkIdx + 1).value = match.Remark || "";
        row.getCell(usageIdx + 1).value = match.UsageNote || "";
      }
    });

    await workbook.xlsx.writeFile(excelPath);
    console.log("✅ usage.json → Part.xlsx 반영 완료");
    res.json({ success: true, message: "Part.xlsx 업데이트 완료" });
  } catch (err) {
    console.error("❌ Part.xlsx 업데이트 실패:", err);
    res.status(500).json({ success: false, error: "업데이트 중 오류 발생" });
  }
});

const app = express();
const PORT = process.env.PORT || 8080;

// ✅ SSH 키 등록 (환경변수에서 가져와서 등록)
if (process.env.SSH_PRIVATE_KEY) {
  const sshDir = path.join(__dirname, ".ssh");
  const privateKeyPath = path.join(sshDir, "id_ed25519");

  fs.mkdirSync(sshDir, { recursive: true });
  fs.writeFileSync(privateKeyPath, process.env.SSH_PRIVATE_KEY + "\n", { mode: 0o600 });

  execSync("mkdir -p ~/.ssh && cp ./.ssh/id_ed25519 ~/.ssh/id_ed25519");

  // ✅ GitHub 호스트 키 등록
  const knownHostsPath = path.join(sshDir, "known_hosts");
  execSync("ssh-keyscan github.com >> " + knownHostsPath);
  execSync("cp ./.ssh/known_hosts ~/.ssh/known_hosts");
}

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

app.use(cors());

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
// server.js
app.get("/excel/part/all", (req, res) => {
  const filePath = path.join(__dirname, "assets", "Part.xlsx");
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: "파일 없음" });

  const workbook = xlsx.readFile(filePath);
  const worksheet = workbook.Sheets["part"];
  const jsonData = xlsx.utils.sheet_to_json(worksheet, { defval: "" });

  try {
    const usageData = JSON.parse(
      fs.readFileSync(path.join(__dirname, "assets", "usage.json"), "utf-8")
    );
    jsonData.forEach((row) => {
      const match = usageData.find(
        (u) => u.Part === row["Part#"] && u.Serial === row["Serial #"]
      );
      if (match) {
        row["Remark"] = match.Remark;
        row["사용처"] = match.UsageNote;
      }
    });
  } catch (e) {
    console.warn("⚠️ usage.json 불러오기 실패:", e.message);
  }

  return res.json(jsonData);
});
// 엑셀 조회 API
app.get("/excel/:sheet/:value", (req, res) => {
  const { sheet, value } = req.params;

  const filePath =
    sheet.toLowerCase() === "part"
      ? path.join(__dirname, "assets", "Part.xlsx")
      : path.join(__dirname, "assets", "site.xlsx");

  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ error: "File not found." });
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

  // ✅ usage.json 덮어쓰기 (Part 전용)
  if (filePath.includes("Part.xlsx")) {
    try {
      const usageData = JSON.parse(
        fs.readFileSync(path.join(__dirname, "assets", "usage.json"), "utf-8")
      );
      matchedRow.forEach((row) => {
        const match = usageData.find(
          (u) => u["Part#"] === row["Part#"] && u["Serial #"] === row["Serial #"]
        );
        if (match) {
          row["Remark"] = match.Remark;
        }
      });
    } catch (e) {
      console.warn("⚠️ usage.json 불러오기 실패:", e.message);
    }
    return res.json(matchedRow);
  } else {
    return res.json(matchedRow[0]);
  }
});
// ✅ usage.json 저장 및 Git 푸시
app.post("/api/save-usage", express.json(), (req, res) => {
  const newRecord = req.body;
  const usageFilePath = path.join(__dirname, "assets", "usage.json");

  let existingData = [];
  try {
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

    // ✅ Git 자동 푸시
    try {
      const timestamp = new Date().toISOString();
    
      execSync("git config user.email 'keyower1591@gmail.com'");
      execSync("git config user.name 'BRKR-SERVER'");
    
      // ✅ origin remote 없으면 등록 (이미 있으면 무시)
      try {
        execSync("git remote add origin git@github.com:Hyunsu7917/BRKR-SERVER.git");
        console.log("✅ origin remote 추가 완료");
      } catch (e) {
        console.log("ℹ️ origin remote 이미 존재하거나 무시:", e.message);
      }
    
      execSync("git add assets/usage.json");
      execSync(`git commit -m '💾 usage 기록: ${timestamp}'`);
      execSync("git push origin HEAD:main");
    
      console.log("✅ usage.json Git push 성공");
    } catch (e) {
      console.error("❌ usage.json Git push 실패:", e.message);
    }
    

    res.json({ success: true, message: "사용 기록 저장 완료" });
  } catch (err) {
    console.error("❌ usage 저장 실패:", err);
    res.status(500).json({ success: false, error: "서버 저장 오류 발생" });
  }
});

// ✅ usage.json 조회용
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