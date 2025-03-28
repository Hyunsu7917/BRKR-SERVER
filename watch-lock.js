// watch-lock.js

const fs = require('fs');
const path = require('path');
const axios = require('axios');

const partPath = path.join(__dirname, 'assets', 'Part.xlsx');
const hePath = path.join(__dirname, 'assets', 'He.xlsx');

let lastState = null;

function isFileLocked(filePath) {
  try {
    fs.openSync(filePath, 'r+');
    return false;
  } catch {
    return true;
  }
}

async function syncLockState(locked) {
  const url = locked
    ? 'https://brkr-server.onrender.com/api/lock'
    : 'https://brkr-server.onrender.com/api/unlock';

  try {
    await axios.post(url);
    console.log(`🔁 서버에 ${locked ? 'LOCK' : 'UNLOCK'} 요청 완료`);
  } catch (err) {
    console.error('❌ 서버 통신 실패:', err.message);
  }
}

setInterval(async () => {
  const partLocked = isFileLocked(partPath);
  const heLocked = isFileLocked(hePath);
  const nowLocked = partLocked || heLocked;

  if (nowLocked !== lastState) {
    lastState = nowLocked;
    await syncLockState(nowLocked);
  }
}, 5000); // 5초 간격 체크

console.log('👀 엑셀 잠금 감시 시작 (5초 간격)... Ctrl+C로 종료 가능');
