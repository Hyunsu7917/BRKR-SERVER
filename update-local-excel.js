const https = require("https");
const fs = require("fs");
const path = require("path");

const fileUrl = "https://brkr-server.onrender.com/excel/part/download";
const localPath = path.join(__dirname, "assets", "Part.xlsx");
const file = fs.createWriteStream(localPath);

https.get(fileUrl, (res) => {
  if (res.statusCode !== 200) {
    console.error("❌ 다운로드 실패. 응답 코드:", res.statusCode);
    return;
  }

  res.pipe(file);

  file.on("finish", () => {
    file.close(() => {
      console.log("✅ 최신 Part.xlsx 파일이 로컬에 저장되었습니다!");
      console.log("📁 저장 위치:", localPath);
    });
  });
}).on("error", (err) => {
  console.error("❌ 요청 중 오류 발생:", err.message);
});
