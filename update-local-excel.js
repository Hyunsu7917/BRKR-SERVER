// update-local-excel.js
const https = require('https');
const fs = require('fs');
const path = require('path');

const fileUrl = "https://brkr-server.onrender.com/excel/part/download";
const localPath = path.join(__dirname, "assets", "Part.xlsx");
const manualModePath = path.join(__dirname, "manual-mode.txt");

const usageBackupUrl = "https://raw.githubusercontent.com/Hyunsu7917/BRKR-SERVER/main/assets/usage-backup.json";
const localUsageBackupPath = path.join(__dirname, "assets", "usage-backup.json");

// 엑셀 파일이 열려있는지 확인
function isFileLocked(filePath) {
  try {
    fs.renameSync(filePath, filePath); // 락되었으면 오류
    return false;
  } catch {
    return true;
  }
}

// 수동모드 여부 확인
function isManualMode() {
  return fs.existsSync(manualModePath);
}

// 최근 수정 여부 (5분 내)
function wasRecentlyModified(filePath, minutes = 5) {
  const stat = fs.existsSync(filePath) ? fs.statSync(filePath) : null;
  if (!stat) return false;
  const modifiedTime = new Date(stat.mtime);
  const now = new Date();
  const diffMinutes = (now - modifiedTime) / 60000;
  return diffMinutes < minutes;
}

// usage-backup.json 다운로드
function downloadJSON(url, dest, cb) {
  const file = fs.createWriteStream(dest);
  https.get(url, (res) => {
    if (res.statusCode !== 200) {
      console.error("❌ usage-backup.json 다운로드 실패:", res.statusCode);
      cb(new Error("Failed to download"));
      return;
    }
    res.pipe(file);
    file.on("finish", () => {
      file.close(cb);
    });
  }).on("error", (err) => {
    fs.unlink(dest, () => {});
    cb(err);
  });
}

// 엑셀 다운로드
function downloadExcel() {
  if (isManualMode()) {
    return console.log("⚠️ 수동 모드: 동기화 중단됨.");
  }

  if (isFileLocked(localPath)) {
    return console.log("⚠️ Part.xlsx가 열려 있어서 동기화 건너뜀.");
  }

  if (wasRecentlyModified(localPath, 5)) {
    return console.log("⚠️ 최근 수정된 파일: 동기화 건너뜀.");
  }

  const file = fs.createWriteStream(localPath);
  https.get(fileUrl, (response) => {
    if (response.statusCode !== 200) {
      console.error("❌ 다운로드 실패:", response.statusCode);
      return;
    }

    response.pipe(file);
    file.on("finish", () => {
      file.close(() => {
        console.log("✅ 최신 Part.xlsx 파일이 로컬에 저장되었습니다!");
        console.log("📂 저장 위치:", localPath);
      });
    });
  }).on("error", (err) => {
    console.error("❌ 요청 실패:", err.message);
  });
}

// 실행 순서: usage-backup.json → Part.xlsx
downloadJSON(usageBackupUrl, localUsageBackupPath, (err) => {
  if (err) {
    console.error("❌ usage-backup.json 다운로드 에러:", err.message);
    return;
  }
  console.log("✅ 최신 usage-backup.json 다운로드 완료");
  downloadExcel();
});