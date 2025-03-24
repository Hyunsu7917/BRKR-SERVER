// server.js - 개수 복구된 버전 (site.xlsx + 그룹 각종 확인 API + usage 저장 API 정사)

const express = require("express");
const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");
const basicAuth = require("express-basic-auth");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());

app.use(
  basicAuth({
    users: { BBIOK: "Bruker_2025" },
    challenge: true,
  })
);

// 항목별 정리 (site.xlsx)
app.get("/excel/:sheet/:value", (req, res) => {
  const filePath = path.join(__dirname, "assets", "site.xlsx");
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: "site.xlsx not found" });

  const { sheet, value } = req.params;
  const workbook = xlsx.readFile(filePath);
  const worksheet = workbook.Sheets[sheet];
  if (!worksheet) return res.status(404).json({ error: `Sheet '${sheet}' not found` });

  const jsonData = xlsx.utils.sheet_to_json(worksheet, { defval: "" });
  const matchedRow = jsonData.filter(row => {
    return Object.values(row).some(cell => String(cell).toLowerCase().includes(value.toLowerCase()));
  });

  if (matchedRow.length === 1) {
    return res.json(matchedRow[0]);
  } else {
    return res.json(matchedRow);
  }
});

// 국내 재고 조회 (Part.xlsx)
app.get("/excel/part/all", (req, res) => {
  const filePath = path.join(__dirname, "assets", "Part.xlsx");
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: "Part.xlsx not found" });

  const workbook = xlsx.readFile(filePath);
  const worksheet = workbook.Sheets["part"];
  const jsonData = xlsx.utils.sheet_to_json(worksheet, { defval: "" });

  try {
    const usageData = JSON.parse(
      fs.readFileSync(path.join(__dirname, "assets", "usage.json"), "utf-8")
    );

    jsonData.forEach((row) => {
      const match = usageData.find((u) => {
        const part = String(row["Part#"] || "").trim();
        const serial = String(row["Serial #"] || "").trim();
        return String(u.Part).trim() === part && String(u.Serial).trim() === serial;
      });
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

// usage 저장 API
app.post("/api/save-usage", (req, res) => {
  const usagePath = path.join(__dirname, "assets", "usage.json");
  const newRecord = req.body;

  let usageData = [];
  if (fs.existsSync(usagePath)) {
    usageData = JSON.parse(fs.readFileSync(usagePath, "utf-8"));
  }

  const index = usageData.findIndex(
    (u) => u.Part === newRecord.Part && u.Serial === newRecord.Serial
  );

  if (index !== -1) {
    usageData[index] = newRecord;
  } else {
    usageData.push(newRecord);
  }

  fs.writeFileSync(usagePath, JSON.stringify(usageData, null, 2), "utf-8");
  res.json({ success: true });
});

// usage 조회
app.get("/api/usage", (req, res) => {
  const usageFilePath = path.join(__dirname, "assets", "usage.json");
  if (!fs.existsSync(usageFilePath)) return res.json([]);
  const data = fs.readFileSync(usageFilePath, "utf-8");
  res.json(JSON.parse(data));
});

app.listen(PORT, () => {
  console.log(`🚀 Server running on port ${PORT}`);
});