const express = require("express");
const cors = require("cors");
const xlsx = require("xlsx");
const path = require("path");
const basicAuth = require("basic-auth");

const app = express();
const PORT = process.env.PORT || 8080;

const auth = (req, res, next) => {
    const user = basicAuth(req);
    const isAuthorized =
      user && user.name === "BBIOK" && user.pass === "Bruker_2025";
  
    if (!isAuthorized) {
      res.set("WWW-Authenticate", 'Basic realm="Authorization Required"');
      return res.status(401).send("Access denied");
    }
    next();
  };

app.use(cors());
app.use(auth); 
app.use("/assets", express.static(path.join(__dirname, "assets")));

// 엑셀 파일 경로
const workbook = xlsx.readFile(path.join(__dirname, "assets/site.xlsx"));

app.get("/excel/:sheet/:value", (req, res) => {
  const { sheet, value } = req.params;
  const worksheet = workbook.Sheets[sheet];

  if (!worksheet) {
    return res.status(404).json({ error: `Sheet '${sheet}' not found.` });
  }

  const jsonData = xlsx.utils.sheet_to_json(worksheet, { defval: "" });

  // 첫 번째 열 기준으로 검색 (A열)
  const matchedRow = jsonData.find((row) => {
    const firstKey = Object.keys(row)[0];
    return String(row[firstKey]).trim() === decodeURIComponent(value);
  });

  if (!matchedRow) {
    return res.status(404).json({ error: `'${value}' not found in sheet '${sheet}'.` });
  }

  res.json(matchedRow);
});

app.listen(PORT, () => {
  console.log(`🛰️  Server running on http://localhost:${PORT}`);
});
