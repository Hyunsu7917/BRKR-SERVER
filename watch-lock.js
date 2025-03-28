const fs = require('fs');
const path = require('path');
const axios = require('axios');

const filesToWatch = [
  { name: 'Part.xlsx', path: path.join(__dirname, 'assets', 'Part.xlsx') },
  { name: 'He.xlsx', path: path.join(__dirname, 'assets', 'He.xlsx') },
];

const serverUrl = 'https://brkr-server.onrender.com';
const lockStates = {}; // 파일별 상태 기억

function isFileOpenSync(filePath) {
  try {
    const fd = fs.openSync(filePath, 'r+');
    fs.closeSync(fd);
    return false;
  } catch {
    return true;
  }
}

console.log('📡 엑셀 잠금 감시 시작 (5초 간격)... Ctrl+C로 종료 가능');

setInterval(async () => {
  for (const file of filesToWatch) {
    const locked = isFileOpenSync(file.path);

    if (lockStates[file.name] !== locked) {
      lockStates[file.name] = locked;

      console.log(
        locked
          ? `🔒 ${file.name} 잠김 감지됨`
          : `🔓 ${file.name} 닫힘 감지됨`
      );

      try {
        const url = locked ? `${serverUrl}/api/lock` : `${serverUrl}/api/unlock`;
        await axios.post(url);
        console.log(`✅ 서버에 ${locked ? 'LOCK' : 'UNLOCK'} 요청 완료`);
      } catch (err) {
        console.error(`❌ 서버 통신 실패 (${file.name}):`, err.message);
      }
    }
  }
}, 5000);
