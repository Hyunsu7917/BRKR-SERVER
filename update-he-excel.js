const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");

const heBackupPath = path.join(__dirname, "he-usage-backup.json");
const heFilePath = path.join(__dirname, "assets", "He.xlsx");

// JSON이 없으면 중단
if (!fs.existsSync(heBackupPath)) {
  console.warn("⚠️ he-usage-backup.json 파일 없음. He 업데이트 생략");
  process.exit(0);
}

const raw = fs.readFileSync(heBackupPath, "utf-8").trim();
const json = raw ? JSON.parse(raw) : [];

// 최신 기록만 추리는 Map 생성
const latestMap = new Map();
json.forEach((record) => {
  const key = `${record["고객사"]}_${record["지역"]}_${record["Magnet"]}`;
  const existing = latestMap.get(key);
  if (!existing || new Date(record["충진일"]) > new Date(existing["충진일"])) {
    latestMap.set(key, record);
  }
});

// 엑셀 불러오기
const workbook = xlsx.readFile(heFilePath);
const scheduleSheetName = "일정";
const recordSheetName = "기록";

// =============== 일정 시트 처리 ===============
const scheduleSheet = xlsx.utils.sheet_to_json(workbook.Sheets[scheduleSheetName], { defval: "" });
const updatedSchedule = [...scheduleSheet];

latestMap.forEach((record) => {
  const { 고객사, 지역, Magnet, 충진일, 다음충진일, "충진주기(개월)": 주기 } = record;
  const matchIndex = updatedSchedule.findIndex(
    (row) => row["고객사"] === 고객사 && row["지역"] === 지역 && row["Magnet"] === Magnet
  );

  const newRow = { 고객사, 지역, Magnet, 충진일, 다음충진일, "충진주기(개월)": 주기 };

  if (matchIndex !== -1) {
    updatedSchedule[matchIndex] = newRow; // 기존 업데이트
  } else {
    updatedSchedule.push(newRow); // 새로 추가
  }
});

workbook.Sheets[scheduleSheetName] = xlsx.utils.json_to_sheet(updatedSchedule);

// =============== 기록 시트 처리 ===============
// 1. 기존 시트 읽기
const recordSheet = xlsx.utils.sheet_to_json(workbook.Sheets[recordSheetName], { header: 1 });
const headers = recordSheet.slice(0, 3); // 고객사, 지역, Magnet
const dataRows = recordSheet.slice(3);   // 실제 기록 행

// 고객사+지역+Magnet 조합 → 열 번호 매핑
const columnIndexMap = new Map();
headers[0].forEach((customer, idx) => {
  if (idx === 0 || !customer) return;
  const key = `${customer}_${headers[1][idx]}_${headers[2][idx]}`;
  columnIndexMap.set(key, idx);
});

// 각 셀 값 개수 기록 (중복 포함)
const existingCounts = new Map();
for (let row of dataRows) {
  for (let colIdx = 1; colIdx < row.length; colIdx++) {
    const customer = headers[0][colIdx];
    const region = headers[1][colIdx];
    const magnet = headers[2][colIdx];
    const value = row[colIdx];
    if (value) {
      const key = `${customer}_${region}_${magnet}_${value}`;
      existingCounts.set(key, (existingCounts.get(key) || 0) + 1);
    }
  }
}

// JSON 기록 삽입 (중복 허용하지만 중복횟수 비교)
const newCounts = new Map();
json.forEach(({ 고객사, 지역, Magnet, 충진일 }) => {
  const key = `${고객사}_${지역}_${Magnet}`;
  const fullKey = `${key}_${충진일}`;
  const colIdx = columnIndexMap.get(key);
  if (colIdx === undefined) return;

  const existing = existingCounts.get(fullKey) || 0;
  const used = newCounts.get(fullKey) || 0;

  if (used < (existing ? existing : 0)) {
    // 이미 다 들어간 기록이면 패스
    newCounts.set(fullKey, used + 1);
    return;
  }

  // 새로 삽입
  let inserted = false;
  for (let row of dataRows) {
    if (!row[colIdx]) {
      row[colIdx] = 충진일;
      inserted = true;
      break;
    }
  }

  if (!inserted) {
    const newRow = new Array(headers[0].length).fill("");
    newRow[colIdx] = 충진일;
    dataRows.push(newRow);
  }

  newCounts.set(fullKey, used + 1);
});

// 기록 시트 덮어쓰기
workbook.Sheets[recordSheetName] = xlsx.utils.aoa_to_sheet([...headers, ...dataRows]);

// 저장
xlsx.writeFile(workbook, heFilePath);
console.log("✅ He.xlsx 업데이트 완료!");
