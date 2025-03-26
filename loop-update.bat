@echo off
chcp 65001 >nul
:loop
echo 💠 자동 동기화 시작 (3분마다 실행)
echo [%DATE% %TIME%] ⏳ 동기화 시도 중...

IF EXIST manual-mode.txt (
  echo ⚠️ 수동 모드: 동기화 중단됨.
) ELSE (
  node update-local-excel.js
)

timeout /t 10
goto loop
