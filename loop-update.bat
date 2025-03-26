@echo off
chcp 65001 >nul
echo 🔁 자동 동기화 시작 (3분마다 실행)

:loop
  echo [%date% %time%] ⏳ 동기화 시도 중...
  node update-local-excel.js
  timeout /t 180 >nul
goto loop
