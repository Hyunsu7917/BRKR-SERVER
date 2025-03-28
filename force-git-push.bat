@echo off
chcp 65001 >nul
echo 🔄 엑셀 파일 Git 강제 Push 시작...

REM Git 설정 (옵션)
git config --global core.autocrlf false

REM 최신 상태 확인
git status

REM 엑셀 파일만 추가
git add assets/*.xlsx

REM 커밋 메시지 (시간 자동 포함)
set TIME_STR=%DATE% %TIME%
git commit -m "엑셀 파일 강제 푸시 - %TIME_STR%"

REM 강제 푸시
git push origin main --force

echo ✅ 완료되었습니다!
pause
