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
    GIT_SSH_COMMAND: 'ssh -i /opt/render/.ssh/render_deploy_key -o StrictHostKeyChecking=no',
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
try {
  execSync(`git config --global user.email "keyower159@gmail.com"`);
  execSync(`git config --global user.name "BRKR-HELIUM-BOT"`);
  console.log("✅ Git 사용자 정보 설정 완료");
} catch (err) {
  console.error("❌ Git 사용자 설정 실패:", err.message);
}

function pushToGit() {
  return new Promise((resolve, reject) => {
    exec(
      `git add . && git commit -m "auto: helium update" && git push --set-upstream origin main`,    
      {
        cwd: __dirname,
        env: {
          ...process.env,
          GIT_SSH_COMMAND: `ssh -i /opt/render/.ssh/render_deploy_key`,
          GIT_AUTHOR_NAME: "BRKR-AUTO",
          GIT_AUTHOR_EMAIL: "keyower159@gmail.com",
          GIT_COMMITTER_NAME: "BRKR-AUTO",
          GIT_COMMITTER_EMAIL: "keyower159@gmail.com",
        },
      },
      (err, stdout, stderr) => {
        if (err) {
          console.error("Git push 실패:", stderr);
          return reject(stderr);
        }
        console.log("✅ Git push 성공:", stdout);
        resolve(stdout);
      }
    );
  });
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

    // ✅ He.xlsx도 함께 업데이트
    try {
      const heFilePath = path.join(__dirname, "assets", "He.xlsx");
      const heWorkbook = xlsx.readFile(heFilePath);
      const heSheet = heWorkbook.Sheets[heWorkbook.SheetNames[0]];
      const heJsonData = xlsx.utils.sheet_to_json(heSheet, { defval: "" });

      const heBackupPath = path.join(__dirname, "he-usage-backup.json");
      const heBackup = fs.existsSync(heBackupPath)
        ? JSON.parse(fs.readFileSync(heBackupPath, "utf-8"))
        : [];

      // 일정 시트 업데이트
      heJsonData.forEach(record => {
        const { 고객사, 지역, Magnet, 충진일, 다음충진일, "충진주기(개월)": 주기 } = record;
        const rowIndex = heJsonData.findIndex(row =>
          row["고객사"] === 고객사 && row["지역"] === 지역 && row["Magnet"] === Magnet
        );
        if (rowIndex !== -1) {
          heJsonData[rowIndex]["충진일"] = 충진일;
          heJsonData[rowIndex]["다음충진일"] = 다음충진일;
          heJsonData[rowIndex]["충진주기(개월)"] = 주기;
        }
      });

      const newHeSheet = xlsx.utils.json_to_sheet(heJsonData);
      heWorkbook.Sheets[heWorkbook.SheetNames[0]] = newHeSheet;
      xlsx.writeFile(heWorkbook, heFilePath);
      console.log("📁 로컬 He.xlsx 일정 시트 저장 완료:", heFilePath);

      // 🔄 he-usage-backup.json 기록
      heBackup.push(...heJsonData);
      fs.writeFileSync(heBackupPath, JSON.stringify(heBackup, null, 2), "utf-8");
      console.log("📄 he-usage-backup.json 기록 완료");

    } catch (err) {
      console.error("⚠️ He.xlsx 업데이트 실패:", err.message);
    }
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
    // === 1. Part.xlsx 업데이트 ===
    const backupPath = path.join(__dirname, "assets", "usage-backup.json");
    const filePath = path.join(__dirname, "assets", "Part.xlsx");

    if (fs.existsSync(backupPath)) {
      const backupRaw = fs.readFileSync(backupPath, "utf-8").trim();
      const backupData = backupRaw ? JSON.parse(backupRaw) : [];
      const workbook = xlsx.readFile(filePath);
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = xlsx.utils.sheet_to_json(sheet, { defval: "" });

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

      const newSheet = xlsx.utils.json_to_sheet(jsonData);
      workbook.Sheets[workbook.SheetNames[0]] = newSheet;
      fs.writeFileSync(filePath, xlsx.write(workbook, { type: "buffer", bookType: "xlsx" }));
      console.log("✅ 로컬 Part.xlsx 덮어쓰기 완료!");
    } else {
      console.warn("⚠️ usage-backup.json 파일 없음. Part 업데이트 생략");
    }

    // === 2. He.xlsx 업데이트 ===
    const heBackupPath = path.join(__dirname, "he-usage-backup.json");
    const heFilePath = path.join(__dirname, "assets", "He.xlsx");

    if (fs.existsSync(heBackupPath)) {
      const heRaw = fs.readFileSync(heBackupPath, "utf-8").trim();
      const heData = heRaw ? JSON.parse(heRaw) : [];
      const heWorkbook = xlsx.readFile(heFilePath);

      // "일정" 시트 업데이트
      const scheduleSheet = heWorkbook.Sheets["일정"];
      const scheduleJson = xlsx.utils.sheet_to_json(scheduleSheet, { defval: "" });

      heData.forEach(record => {
        const idx = scheduleJson.findIndex(row =>
          row["고객사"] === record["고객사"] &&
          row["지역"] === record["지역"] &&
          String(row["Magnet"]) === String(record["Magnet"])
        );
        if (idx !== -1) {
          scheduleJson[idx]["충진일"] = record["충진일"];
          scheduleJson[idx]["다음충진일"] = record["다음충진일"];
          scheduleJson[idx]["충진주기(개월)"] = record["충진주기(개월)"];
        }
      });

      const newScheduleSheet = xlsx.utils.json_to_sheet(scheduleJson);
      heWorkbook.Sheets["일정"] = newScheduleSheet;

      // "기록" 시트 업데이트 (고객사별 열 구조)
      const recordSheet = heWorkbook.Sheets["기록"];
      const customerRow = recordSheet["1"];
      const regionRow = recordSheet["2"];
      const magnetRow = recordSheet["3"];

      heData.forEach(record => {
        const { 고객사, 지역, Magnet, 충진일 } = record;
        const refSheet = heWorkbook.getWorksheet("기록");
        const customerNames = refSheet.getRow(1).values;
        let colIndex = -1;

        for (let i = 2; i < customerNames.length; i++) {
          const name = customerNames[i];
          const region = refSheet.getRow(2).getCell(i).value;
          const magnet = refSheet.getRow(3).getCell(i).value;

          if (name === 고객사 && region === 지역 && magnet == Magnet) {
            colIndex = i;
            break;
          }
        }

        if (colIndex !== -1) {
          let row = 4;
          while (refSheet.getRow(row).getCell(colIndex).value) {
            row++;
          }
          refSheet.getRow(row).getCell(colIndex).value = 충진일;
        }
      });

      // 저장
      fs.writeFileSync(heFilePath, xlsx.write(heWorkbook, { type: "buffer", bookType: "xlsx" }));
      console.log("✅ 로컬 He.xlsx 덮어쓰기 완료!");
    } else {
      console.warn("⚠️ he-usage-backup.json 파일 없음. He 업데이트 생략");
    }

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

    // ✅ Part.xlsx 복구
    const partFilePath = path.join(__dirname, "assets", "Part.xlsx");
    const partBackupPath = path.join(__dirname, "assets", "usage-backup.json");
    if (fs.existsSync(partBackupPath)) {
      const backupData = JSON.parse(fs.readFileSync(partBackupPath, "utf-8"));
      const workbook = xlsx.readFile(partFilePath);
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = xlsx.utils.sheet_to_json(sheet, { defval: "" });

      console.log("🟡 Part 백업 개수:", backupData.length);
      console.log("🟡 Part 백업 미리보기:", JSON.stringify(backupData[0], null, 2));

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
      fs.writeFileSync(partFilePath, xlsx.write(workbook, { type: "buffer", bookType: "xlsx" }));
      console.log("📁 Part.xlsx 복구 완료!");
    } else {
      console.log("⚠️ Part 백업 파일 없음");
    }

    // ✅ He.xlsx 복구
    const heFilePath = path.join(__dirname, "assets", "He.xlsx");
    const heBackupPath = path.join(__dirname, "he-usage-backup.json");
    if (fs.existsSync(heBackupPath)) {
      const heBackup = JSON.parse(fs.readFileSync(heBackupPath, "utf-8"));
      const heWorkbook = xlsx.readFile(heFilePath);

      // === 일정 시트 복구 ===
      const scheduleSheet = heWorkbook.Sheets["일정"];
      const scheduleJson = xlsx.utils.sheet_to_json(scheduleSheet, { defval: "" });

      for (const item of heBackup) {
        const idx = scheduleJson.findIndex(row =>
          row["고객사"] === item["고객사"] &&
          row["지역"] === item["지역"] &&
          String(row["Magnet"]) === String(item["Magnet"])
        );
        if (idx !== -1) {
          scheduleJson[idx]["충진일"] = item["충진일"];
          scheduleJson[idx]["다음충진일"] = item["다음충진일"];
          scheduleJson[idx]["충진주기(개월)"] = item["충진주기(개월)"];
        }
      }

      const newSchedule = xlsx.utils.json_to_sheet(scheduleJson);
      heWorkbook.Sheets["일정"] = newSchedule;

      // === 기록 시트 복구 ===
      const recordSheet = heWorkbook.Sheets["기록"];
      const range = xlsx.utils.decode_range(recordSheet["!ref"]);
      const customerNames = [];
      for (let col = 1; col <= range.e.c; col++) {
        const cell = recordSheet[xlsx.utils.encode_cell({ r: 0, c: col })];
        customerNames.push(cell ? cell.v : "");
      }

      heBackup.forEach(item => {
        const { 고객사, 지역, Magnet, 충진일 } = item;
        let colIndex = -1;

        for (let col = 1; col < customerNames.length; col++) {
          const name = customerNames[col];
          const region = recordSheet[xlsx.utils.encode_cell({ r: 1, c: col })]?.v;
          const magnet = recordSheet[xlsx.utils.encode_cell({ r: 2, c: col })]?.v;
          if (name === 고객사 && region === 지역 && String(magnet) === String(Magnet)) {
            colIndex = col;
            break;
          }
        }

        if (colIndex !== -1) {
          let row = 3;
          while (true) {
            const cellRef = xlsx.utils.encode_cell({ r: row, c: colIndex });
            if (!recordSheet[cellRef] || !recordSheet[cellRef].v) {
              recordSheet[cellRef] = { t: "s", v: 충진일 };
              break;
            }
            row++;
          }
        }
      });

      fs.writeFileSync(heFilePath, xlsx.write(heWorkbook, { type: "buffer", bookType: "xlsx" }));
      console.log("📁 He.xlsx 복구 완료!");
    } else {
      console.log("⚠️ He 백업 파일 없음");
    }

    console.log("✅ restoreExcelFromBackup 완료");
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
// ✅ He 백업 JSON 조회
app.get("/api/show-he-backup", (req, res) => {
  try {
    const heBackupPath = path.join(__dirname, "he-usage-backup.json");

    if (!fs.existsSync(heBackupPath)) {
      return res.status(404).json({ error: "He 백업 파일이 존재하지 않습니다." });
    }

    const backupData = JSON.parse(fs.readFileSync(heBackupPath, "utf-8"));
    return res.json({ success: true, data: backupData });
  } catch (err) {
    console.error("❌ He 백업 조회 오류:", err);
    return res.status(500).json({ error: "He 백업 파일 조회 중 오류 발생" });
  }
});

// ✅ 서버 실행 시 Excel 자동 복원
restoreExcelFromBackup(); // He & Part 자동 복원 수행

// ✅ Render 서버가 detached 상태일 경우 main 브랜치로 강제 이동
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

// ✅ Part 엑셀 다운로드 API
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

// ✅ He 엑셀 다운로드 API
app.get("/excel/he/download", (req, res) => {
  const filePath = path.join(__dirname, "assets", "He.xlsx");
  res.download(filePath, "He.xlsx", (err) => {
    if (err) {
      console.error("❌ He.xlsx 전송 실패:", err.message);
      res.status(500).send("Download failed.");
    } else {
      console.log("📦 He.xlsx 파일 전송 완료!");
    }
  });
});

app.post("/api/trigger-local-update", (req, res) => {
  let localResult = "❌ 실패";
  let heResult = "❌ 실패";

  try {
    execSync("node update-local-excel.js");
    localResult = "✅ 성공";
  } catch (err) {
    console.error("❌ update-local-excel.js 실패:", err.message);
  }

  try {
    execSync("node update-he-excel.js");
    heResult = "✅ 성공";
  } catch (err) {
    console.error("❌ update-he-excel.js 실패:", err.message);
  }

  const success = localResult === "✅ 성공" && heResult === "✅ 성공";

  console.log(`📦 결과 요약 → Part: ${localResult} / He: ${heResult}`);

  res.status(success ? 200 : 207).json({
    success,
    message: `Part: ${localResult}, He: ${heResult}`,
  });
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
    const headers = sheet.getRow(1).values.slice(1); // ✅ A열부터 정확하게

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      const rowData = {};
      row.eachCell((cell, colNumber) => {
        const key = headers[colNumber - 1]; // ✅ headers와 정렬
        rowData[key] = cell.value !== undefined ? cell.value : "";
      });
      rows.push(rowData);
    });


    res.json(rows);
  } catch (err) {
    console.error("엑셀 파싱 에러:", err);
    res.status(500).json({ error: "서버 에러" });
  }
});
// ✅ Helium Excel 저장 + Git 반영
app.post("/api/he/save", async (req, res) => {
  const records = req.body;
  const filePath = path.join(__dirname, "he-usage-backup.json");

  if (!Array.isArray(records)) {
    return res.status(400).json({ success: false, message: "데이터 형식이 배열이 아닙니다." });
  }

  try {
    // ✅ 1. 기존 백업 불러오기 + 중첩 배열 방지
    let backup = [];
    if (fs.existsSync(filePath)) {
      const raw = fs.readFileSync(filePath, "utf8");
      const json = JSON.parse(raw);
      backup = Array.isArray(json[0]) ? json.flat() : json;
    }

    backup.push(...records);
    fs.writeFileSync(filePath, JSON.stringify(backup, null, 2));

    // ✅ 2. 엑셀 파일 로드
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile("assets/He.xlsx");

    const sheet1 = workbook.getWorksheet("일정");
    const sheet2 = workbook.getWorksheet("기록");

    // ✅ G열 이후 불필요한 열 제거 (깨짐 방지)
    if (sheet1.columnCount > 6) {
      sheet1.spliceColumns(7, sheet1.columnCount - 6);
    }

    const rows = sheet1.getRows(2, sheet1.rowCount - 1);
    const headerRow1 = sheet2.getRow(1);
    const headerRow2 = sheet2.getRow(2);
    const headerRow3 = sheet2.getRow(3);

    // ✅ 3. 일정 시트 업데이트
    records.forEach((record) => {
      const customer = String(record["고객사"] ?? "").trim();
      const region = String(record["지역"] ?? "").trim();
      const magnet = String(record["Magnet"] ?? "").trim();
      const chargeDate = record["충진일"];
      const nextChargeDate = record["다음충진일"];
      const cycle = record["충진주기(개월)"];

      const matchedRow = rows.find((row) => {
        const rowCustomer = String(row.getCell(1).value ?? "").trim();
        const rowRegion = String(row.getCell(2).value ?? "").trim();
        const rowMagnet = String(row.getCell(3).value ?? "").trim();
        return rowCustomer === customer && rowRegion === region && rowMagnet === magnet;
      });

      if (matchedRow) {
        matchedRow.getCell(4).value = chargeDate;
        matchedRow.getCell(5).value = nextChargeDate;
        matchedRow.getCell(6).value = cycle;
        console.log(`✅ 일정 업데이트: ${customer} / ${region} / ${magnet}`);
      } else {
        console.warn(`❌ 일정 시트에서 ${customer} / ${region} / ${magnet} 찾지 못함`);
      }
    });

    // ✅ 4. 기록 시트 업데이트
    records.forEach((record) => {
      const newCustomer = String(record["고객사"] ?? "").trim();
      const newRegion = String(record["지역"] ?? "").trim();
      const newMagnet = String(record["Magnet"] ?? "").trim();
      const chargeDate = record["충진일"];

      let targetCol = -1;
      for (let i = 2; i <= sheet2.columnCount; i++) {
        const customer = String(headerRow1.getCell(i).value ?? "").trim();
        const region = String(headerRow2.getCell(i).value ?? "").trim();
        const magnet = String(headerRow3.getCell(i).value ?? "").trim();

        if (customer === newCustomer && region === newRegion && magnet === newMagnet) {
          targetCol = i;
          break;
        }
      }

      if (targetCol !== -1) {
        let rowIndex = 4;
        while (sheet2.getCell(rowIndex, targetCol).value) rowIndex++;
        sheet2.getCell(rowIndex, targetCol).value = chargeDate;
        console.log(`✅ ${newCustomer} (${newRegion} / ${newMagnet}) → ${rowIndex}행 기록됨`);
      } else {
        console.warn(`❗ 기록 시트에 ${newCustomer} (${newRegion} / ${newMagnet}) 찾을 수 없음`);
      }
    });

    // ✅ 5. 저장 → He.xlsx로 저장 (안전하게)
    
    // 🔒 G열 이후 불필요한 열 제거 (파일 깨짐 방지)
    if (sheet1.columnCount > 6) {
      sheet1.spliceColumns(7, sheet1.columnCount - 6);
    }

    // 💡 저장 옵션 설정
    workbook.calcProperties.fullCalcOnLoad = true;

    // ✅ 저장
    await workbook.xlsx.writeFile("assets/He.xlsx");

    // 🕒 저장 완료까지 0.5초 대기 (flush 보장)
    await new Promise(resolve => setTimeout(resolve, 500));

    // ✅ 6. Git 푸시
    await pushToGit();


    res.json({ success: true });
  } catch (err) {
    console.error("💥 저장 실패:", err);
    res.status(500).json({ success: false, error: err.message });
  }
});




app.post('/api/set-helium-reservation', async (req, res) => {
  const { 고객사, 지역, Magnet, 충진일, 예약여부 } = req.body;

  try {
    // 1. 기존 백업 파일 로드
    const usagePath = path.join(__dirname, 'he-usage-backup.json');
    let usageData = [];
    if (fs.existsSync(usagePath)) {
      usageData = JSON.parse(fs.readFileSync(usagePath, 'utf-8'));
    }

    // 2. 동일한 고객사+지역+Magnet+충진일 항목 찾기
    let found = false;
    usageData = usageData.map(entry => {
      if (
        entry['고객사'] === 고객사 &&
        entry['지역'] === 지역 &&
        entry['Magnet'] === Magnet &&
        entry['충진일'] === 충진일
      ) {
        found = true;
        return { ...entry, 예약여부 };
      }
      return entry;
    });

    // 없으면 새 항목 추가
    if (!found) {
      usageData.push({ 고객사, 지역, Magnet, 충진일, 예약여부 });
    }

    // 3. 백업 파일 저장
    fs.writeFileSync(usagePath, JSON.stringify(usageData, null, 2), 'utf-8');

    // 4. Git commit + push
    const exec = require('child_process').exec;
    exec(`git add ${usagePath} && git commit -m "Update He reservation for ${고객사}" && git push`, {
      cwd: __dirname,
      env: {
        ...process.env,
        GIT_SSH_COMMAND: `ssh -i ${process.env.SSH_PRIVATE_KEY_PATH}`,
      },
    });

    // 5. Excel 업데이트 스크립트 실행
    const { execSync } = require('child_process');
    execSync('node update-he-excel.js', { stdio: 'inherit' });

    res.status(200).json({ success: true, message: '예약 정보 저장 완료' });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, message: '예약 처리 중 오류 발생' });
  }
});

app.get('/api/check-manual-mode', (req, res) => {
  const lockPath = path.join(__dirname, 'manual-mode.txt');
  const isLocked = fs.existsSync(lockPath);
  res.json({ manual: isLocked });
});

app.post('/api/lock', (req, res) => {
  fs.writeFileSync(path.join(__dirname, 'manual-mode.txt'), 'LOCKED');
  res.json({ success: true });
});

app.post('/api/unlock', (req, res) => {
  const filePath = path.join(__dirname, 'manual-mode.txt');
  if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
  res.json({ success: true });
});
// ✅ 서버 시작
app.listen(PORT, () => {
  console.log(`🚀 Server running on port ${PORT}`);
});
