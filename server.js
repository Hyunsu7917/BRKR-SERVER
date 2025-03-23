const express = require("express");
const cors = require("cors");
const xlsx = require("xlsx");
const path = require("path");
const fs = require("fs");
const basicAuth = require("basic-auth");
const { execSync } = require("child_process");

const app = express();
const PORT = process.env.PORT || 8080;

// ✅ SSH 키 등록 (환경변수에서 가져와서 등록)
if (process.env.SSH_PRIVATE_KEY) {
  const sshDir = path.join(__dirname, ".ssh");
  const privateKeyPath = path.join(sshDir, "id_ed25519");

  fs.mkdirSync(sshDir, { recursive: true });
  fs.writeFileSync(privateKeyPath, process.env.SSH_PRIVATE_KEY + "\n", { mode: 0o600 });

  execSync("mkdir -p ~/.ssh && cp ./.ssh/id_ed25519 ~/.ssh/id_ed25519");
  //execSync("eval $(ssh-agent -s)");
  //execSync("ssh-add ~/.ssh/id_ed25519");
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
          (u) => u.Part === row["Part#"] && u.Serial === row["Serial #"]
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
      execSync("git remote add origin git@github.com:Hyunsu7917/BRKR-SERVER.git");
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
