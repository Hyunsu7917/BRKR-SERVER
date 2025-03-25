const express = require("express");
const basicAuth = require("express-basic-auth");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");

const app = express();
const PORT = process.env.PORT || 3001;

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

    // ✅ 백업 파일도 이 위치에서 만들어줌
    const backupPath = path.join(__dirname, "usage-backup.json");
    const currentBackup = fs.existsSync(backupPath)
      ? JSON.parse(fs.readFileSync(backupPath, "utf-8"))
      : [];

    currentBackup.push({
      "Part#": Part,
      "Serial #": Serial,
      PartName,
      Remark,
      UsageNote,
      Timestamp: new Date().toISOString(),
    });

    fs.writeFileSync(backupPath, JSON.stringify(currentBackup, null, 2), "utf-8");

    fs.writeFileSync(filePath, xlsx.write(workbook, { type: "buffer", bookType: "xlsx" }));
    console.log("📁 로컬 Part.xlsx 저장 완료:", filePath);

    return res.json({ success: true });
  } catch (err) {
    console.error("엑셀 저장 실패:", err);
    return res.status(500).json({ error: "엑셀 저장 중 오류 발생" });
  }
});

// ✅ 서버 시작
app.listen(PORT, () => {
  console.log(`🚀 Server running on port ${PORT}`);
});
