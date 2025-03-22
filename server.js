const express = require("express");
const cors = require("cors");
const xlsx = require("xlsx");
const path = require("path");
const basicAuth = require("basic-auth");

const app = express();
const PORT = process.env.PORT || 8080;

app.get("/latest-version.json", (req, res) => {
  res.json({
    version: "1.0.1", // 최신 앱 버전
    apkUrl: "https://your-cdn.com/your-app.apk" // 최신 APK 파일 경로
  });
});

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});

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


const workbook = xlsx.readFile(path.join(__dirname, "assets/site.xlsx"));

app.get("/excel/:sheet/:value", (req, res) => {
  const { sheet, value } = req.params;
  const worksheet = workbook.Sheets[sheet];

  if (!worksheet) {
    return res.status(404).json({ error: `Sheet '${sheet}' not found.` });
  }

  const jsonData = xlsx.utils.sheet_to_json(worksheet, { defval: "" });


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
