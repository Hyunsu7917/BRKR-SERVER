const express = require("express");
const basicAuth = require("express-basic-auth");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");
const ExcelJS = require("exceljs");


const app = express();
const PORT = process.env.PORT || 3001;
// ✅ SSH 키 저장
const sshKeyPath = "/opt/render/.ssh/render_deploy_key";
if (process.env.SSH_PRIVATE_KEY && !fs.existsSync(sshKeyPath)) {
  fs.mkdirSync("/opt/render/.ssh", { recursive: true });
  fs.writeFileSync(sshKeyPath, process.env.SSH_PRIVATE_KEY + '\n', { mode: 0o600 });
  console.log("✅ SSH 키 파일 저장 완료");
}
const { exec, execSync } = require("child_process");
// ✅ GitHub 호스트 등록
try {
  execSync("ssh-keyscan github.com >> ~/.ssh/known_hosts", { stdio: "inherit" });
  console.log("🔐 GitHub 호스트 키 등록 완료");
} catch (err) {
  console.error("❌ 호스트 키 등록 실패:", err.message);
}
// ✅ Git 환경 설정
try {
  const gitEnv = {
    ...process.env,
    GIT_SSH_COMMAND: 'ssh -i ~/.ssh/render_deploy_key -o StrictHostKeyChecking=no',
  };

  execSync("git init", { cwd: process.cwd(), env: gitEnv });

  try {
    execSync("git remote remove origin", { cwd: process.cwd(), env: gitEnv });
    console.log("🧹 기존 origin 제거 완료");
  } catch {
    console.log("ℹ️ origin 없음 → 제거 생략");
  }

  execSync("git remote add origin git@github.com:Hyunsu7917/BRKR-SERVER.git", {
    cwd: process.cwd(),
    env: gitEnv,
  });

  execSync("git pull origin main", { cwd: process.cwd(), env: gitEnv });
  console.log("✅ Git init & origin 등록 + 최신 내용 pull 완료");
} catch (err) {
  console.error("⚠️ Git init/pull 오류:", err.message);
}


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
// ✅ 국내 재고 엑셀에 사용 기록 반영하기
app.post("/api/update-part-excel", basicAuthMiddleware, (req, res) => {
  console.log("📩 Received update request", req.body);
  const filePath = path.join(__dirname, "assets", "Part.xlsx");
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: "파일 없음" });

  const { ["Part#"]: Part, ["Serial #"]: Serial, PartName, Remark, UsageNote } = req.body;

  try {
    // ✅ 엑셀 업데이트
    const workbook = xlsx.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    const rowIndex = jsonData.findIndex(row =>
      String(row["Part#"]).toLowerCase() === String(Part).toLowerCase() &&
      String(row["Serial #"]) === String(Serial)
    );

    if (rowIndex === -1) return res.status(404).json({ error: "해당 부품을 찾을 수 없습니다." });

    jsonData[rowIndex]["Remark"] = Remark;
    jsonData[rowIndex]["사용처"] = UsageNote;

    const newSheet = xlsx.utils.json_to_sheet(jsonData);
    workbook.Sheets[workbook.SheetNames[0]] = newSheet;
    xlsx.writeFile(workbook, filePath);
    console.log("📁 로컬 Part.xlsx 저장 완료:", filePath);

    // ✅ 백업 파일 저장 + 500개 초과 시 정리
    const backupPath = path.join(__dirname, "assets", "usage-backup.json");
    const currentBackup = fs.existsSync(backupPath)
      ? JSON.parse(fs.readFileSync(backupPath, "utf-8"))
      : [];

    // 🔥 500개 초과 시 오래된 기록 제거
    if (currentBackup.length >= 500) {
      const removeCount = currentBackup.length - 499;
      currentBackup.splice(0, removeCount); // 앞에서 오래된 것부터 제거
    }

    currentBackup.push({
      "Part#": Part,
      "Serial #": Serial,
      PartName,
      Remark,
      UsageNote,
      Timestamp: new Date().toISOString(),
    });

    fs.writeFileSync(backupPath, JSON.stringify(currentBackup, null, 2), "utf-8");

    const { execSync } = require("child_process");

    try {
      const branch = execSync("git branch", {
        cwd: process.cwd(),
        env: {
          ...process.env,
          GIT_SSH_COMMAND: 'ssh -i ~/.ssh/render_deploy_key -o StrictHostKeyChecking=no',
        },
      }).toString();

      const status = execSync("git status", {
        cwd: process.cwd(),
        env: {
          ...process.env,
          GIT_SSH_COMMAND: 'ssh -i ~/.ssh/render_deploy_key -o StrictHostKeyChecking=no',
        },
      }).toString();

      console.log("📂 현재 브랜치 상태:\n", branch);
      console.log("📋 Git 상태:\n", status);
    } catch (err) {
      console.error("❌ Git 상태 확인 실패:", err.message);
    }


    const diffStatus = execSync('git status --short').toString();
    console.log("🧪 Git 변경 감지 상태:\n", diffStatus);

    // ✅ Git push만 수행
    try {
      execSync('git config user.name "brkr-server"', { cwd: process.cwd() });      
      execSync('git config user.email "kc7917@naver.com"', { cwd: process.cwd() });
      execSync(`git add assets/Part.xlsx assets/usage-backup.json`, {
        cwd: process.cwd(),
        env: {
          ...process.env,
          GIT_SSH_COMMAND: 'ssh -i ~/.ssh/render_deploy_key -o StrictHostKeyChecking=no',
        },
      });
      console.log("깃에드 실행함!")
      const now = new Date().toISOString();
      execSync(`git commit -m "backup update: ${now}" --allow-empty`, {
        cwd: process.cwd(),
        env: {
          ...process.env,
          GIT_SSH_COMMAND: 'ssh -i ~/.ssh/render_deploy_key -o StrictHostKeyChecking=no',
        },
      });
      const log = execSync('git log --oneline -n 5').toString();
      console.log("📜 최근 커밋 로그:\n", log);
      execSync(`git push origin main`, {
        cwd: process.cwd(),
        env: {
          ...process.env,
          GIT_SSH_COMMAND: 'ssh -i ~/.ssh/render_deploy_key -o StrictHostKeyChecking=no',
        },
      });
      console.log("✅ Git push 성공!");
    } catch (err) {
      console.error("❌ Git push 실패:", err.message);
    }

    return res.json({ success: true });
  } catch (err) {
    console.error("엑셀 저장 실패:", err);
    return res.status(500).json({ error: "엑셀 저장 중 오류 발생" });
  }
});

app.get("/api/sync-usage-to-excel", async (req, res) => {
  try {
    const backupPath = path.join(__dirname, "assets", "usage-backup.json");
    const filePath = path.join(__dirname, "assets", "Part.xlsx");

    // 백업 파일 존재 확인
    if (!fs.existsSync(backupPath)) {
      return res.status(404).json({ error: "백업 파일이 존재하지 않습니다." });
    }

    // 파일 불러오기
    const backupRaw = fs.readFileSync(backupPath, "utf-8").trim();
    const backupData = backupRaw ? JSON.parse(backupRaw) : [];
    const workbook = xlsx.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    // 백업 내용을 엑셀 데이터에 반영
    backupData.forEach(backup => {
      const rowIndex = jsonData.findIndex(row =>
        String(row["Part#"]).toLowerCase() === String(backup["Part#"]).toLowerCase() &&
        String(row["Serial #"]) === String(backup["Serial #"])
      );

      if (rowIndex !== -1) {
        jsonData[rowIndex]["Remark"] = backup.Remark || "";
        jsonData[rowIndex]["사용처"] = backup.UsageNote || "";
      }
    });

    // 다시 저장
    const newSheet = xlsx.utils.json_to_sheet(jsonData);
    workbook.Sheets[workbook.SheetNames[0]] = newSheet;

    console.log("🟡 Buffer 생성 완료");
    fs.writeFileSync(filePath, xlsx.write(workbook, { type: "buffer", bookType: "xlsx" }));

    console.log("✅ 로컬 Part.xlsx 덮어쓰기 완료!");

    return res.json({ success: true, message: "사용기록이 엑셀에 반영되었습니다." });
  } catch (err) {
    console.error("⛔️ 동기화 오류:", err);
    return res.status(500).json({ error: "사용기록 반영 중 오류 발생" });
  }
});

// 🔁 서버 부팅 시 백업 데이터를 엑셀에 자동 반영
const restoreExcelFromBackup = () => {
  try {
    console.log("🟠 restoreExcelFromBackup 시작");
    const filePath = path.join(__dirname, "assets", "Part.xlsx");
    const backupPath = path.join(__dirname, "assets", "usage-backup.json");
    if (!fs.existsSync(backupPath)) return;

    const backupData = JSON.parse(fs.readFileSync(backupPath, "utf-8"));
    const workbook = xlsx.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = xlsx.utils.sheet_to_json(sheet, { defval: "" });

    console.log("🟡 백업 데이터 개수:", backupData.length);
    console.log("🟡 백업 내용 미리보기:", JSON.stringify(backupData[0], null, 2));
    console.log("🟡 엑셀 행 수:", jsonData.length);

    for (const backup of backupData) {
      const rowIndex = jsonData.findIndex(
        row =>
          String(row["Part#"]).toLowerCase() === String(backup.Part).toLowerCase() &&
          String(row["Serial #"]) === String(backup.Serial)
      );
      if (rowIndex !== -1) {
        jsonData[rowIndex]["Remark"] = backup.Remark || "";
        jsonData[rowIndex]["사용처"] = backup.UsageNote || "";
      }
    }

    const newSheet = xlsx.utils.json_to_sheet(jsonData);
    workbook.Sheets[workbook.SheetNames[0]] = newSheet;
    fs.writeFileSync(filePath, xlsx.write(workbook, { type: "buffer", bookType: "xlsx" }));
    console.log("🛠 서버 부팅 시 백업 데이터로 Part.xlsx 복구 완료!");
  } catch (err) {
    console.error("❌ 복구 실패:", err);
  }
};
app.get("/api/show-backup", (req, res) => {
  try {
    const backupPath = path.join(__dirname, "assets", "usage-backup.json");

    if (!fs.existsSync(backupPath)) {
      return res.status(404).json({ error: "백업 파일이 존재하지 않습니다." });
    }

    const backupData = JSON.parse(fs.readFileSync(backupPath, "utf-8"));
    return res.json({ success: true, data: backupData });
  } catch (err) {
    console.error("❌ 백업 파일 조회 오류:", err);
    return res.status(500).json({ error: "백업 파일을 불러오는 중 오류 발생" });
  }
});

restoreExcelFromBackup(); // 💡 서버 실행 시 바로 동작!

// 🧠 Render 서버가 detached 상태일 경우 main 브랜치로 강제 이동
try {
  execSync("git checkout main", {
    cwd: process.cwd(),
    env: {
      ...process.env,
      GIT_SSH_COMMAND: 'ssh -i ~/.ssh/render_deploy_key -o StrictHostKeyChecking=no',
    },
  });
  console.log("🔁 Git 브랜치 → main 체크아웃 완료");
} catch (err) {
  console.error("❌ Git 브랜치 체크아웃 실패:", err.message);
}
app.get("/excel/part/download", (req, res) => {
  const filePath = path.join(__dirname, "assets", "Part.xlsx");
  res.download(filePath, "Part.xlsx", (err) => {
    if (err) {
      console.error("❌ Part.xlsx 전송 실패:", err.message);
      res.status(500).send("Download failed.");
    } else {
      console.log("📦 Part.xlsx 파일 전송 완료!");
    }
  });
});
app.post("/api/trigger-local-update", (req, res) => {
  try {
    execSync("node update-local-excel.js");
    console.log("✅ 로컬 엑셀 자동 업데이트 완료");
    res.status(200).json({ success: true, message: "Local Excel updated" });
  } catch (err) {
    console.error("❌ 로컬 엑셀 업데이트 실패:", err.message);
    res.status(500).json({ success: false, error: err.message });
  }
});
app.get("/excel/he/schedule", async (req, res) => {
  try {
    const filePath = path.join(__dirname, "assets", "He.xlsx");
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet("일정");

    if (!sheet) {
      return res.status(404).json({ error: "시트 '일정'을 찾을 수 없습니다." });
    }

    const rows = [];
    const headers = sheet.getRow(1).values.slice(1); // ✅ A열 비어있을 경우 대응

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;

      const rowData = {};
      row.eachCell((cell, colNumber) => {
        const key = headers[colNumber - 1]; // ✅ colNumber는 1부터 시작하므로 -1
        rowData[key] = cell.value !== undefined ? cell.value : "";
      });

      rows.push(rowData);
      console.log("✅ 서버 JSON 예시:", rows[0]);
    });

    res.json(rows);
  } catch (err) {
    console.error("엑셀 파싱 에러:", err);
    res.status(500).json({ error: "서버 에러" });
  }
});



// ✅ 서버 시작
app.listen(PORT, () => {
  console.log(`🚀 Server running on port ${PORT}`);
});
