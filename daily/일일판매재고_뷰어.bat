@echo off
chcp 65001 >nul
REM 배포 없이 로컬에서만 미리보기용. 배포 후 확인은 coupang-po.pages.dev/daily/ 사용.
setlocal
pushd "%~dp0"
netstat -ano | findstr ":8765 " >nul
if not errorlevel 1 (
  start "" "http://localhost:8765/"
  popd
  exit /b
)
start "일일판매재고 로컬 서버 (창 닫으면 종료)" cmd /k "chcp 65001 >nul & python -m http.server 8765"
timeout /t 2 /nobreak >nul
start "" "http://localhost:8765/"
popd
