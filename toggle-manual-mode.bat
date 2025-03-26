@echo off
chcp 65001 >nul
set FILE=manual-mode.txt

if exist %FILE% (
    del %FILE%
    echo ✅ 수동 모드 OFF: 자동 동기화 재개됨!
) else (
    echo.>%FILE%
    echo 🔒 수동 모드 ON: 자동 동기화 중지됨!
)
pause
