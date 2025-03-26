@echo off
:loop
node update-local-excel.js
echo 🔁 다음 실행까지 180초 대기합니다...
timeout /t 180 > nul
goto loop
