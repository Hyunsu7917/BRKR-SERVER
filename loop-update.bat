@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 💠 자동 동기화 시작 (3분마다 실행)

:loop
echo [%DATE% %TIME%] ⏳ 동기화 시도 중...

if exist manual-mode.txt (
  echo ⚠️ 수동 모드: 동기화 중단됨.
) else (
  node update-local-excel.js
)

timeout /t 180 >nul
goto loop
