@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo 🔁 자동 동기화 시작 (3분마다 실행)

:loop
echo [%DATE% %TIME%] ⏳ 동기화 시도 중...

REM ✅ 수동 모드 여부 확인
if exist manual-mode.txt (
  echo ⚠️ 수동 모드: 동기화 중단됨.
) else (
  REM ✅ Git fetch로 최신 정보 확인
  git fetch origin

  REM ✅ he-usage-backup.json 강제 덮어쓰기
  git show origin/main:he-usage-backup.json > he-usage-backup.json
  echo ✅ 최신 he-usage-backup.json 다운로드 완료!

  REM ✅ usage-backup.json 강제 덮어쓰기
  git show origin/main:assets/usage-backup.json > assets/usage-backup.json
  echo ✅ 최신 usage-backup.json 다운로드 완료!

  REM ✅ 엑셀 파일들 업데이트
  node update-local-excel.js
  node update-he-excel.js

  echo 📗 Part.xlsx 업데이트 완료!
  echo 📘 He.xlsx 업데이트 완료!
)

timeout /t 180 >nul
goto loop
