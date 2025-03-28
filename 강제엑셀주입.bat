@echo off
echo 📁 파일 수정시간 갱신 중...
powershell -Command "Get-Item assets\He.xlsx | % { $_.LastWriteTime = Get-Date }"

echo ✅ Git add 시작...
git add assets\He.xlsx

echo ✅ Git 커밋...
git commit -m "Force update He.xlsx"

echo 🚀 Git push 중...
git push

pause
