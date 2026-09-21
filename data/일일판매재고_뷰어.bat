@echo off
chcp 65001 >nul
setlocal
pushd "%~dp0"
REM 이미 8765 포트 사용 중이면 그대로 브라우저만 열기
netstat -ano | findstr ":8765 " >nul
if not errorlevel 1 (
  start "" "http://localhost:8765/일일판매재고_뷰어.html"
  popd
  exit /b
)
REM 로컬 서버 백그라운드 실행 후 브라우저 오픈
start "일일판매재고 서버 (창 닫으면 종료)" cmd /k "chcp 65001 >nul & python -m http.server 8765"
REM 서버가 뜰 시간 확보
timeout /t 2 /nobreak >nul
start "" "http://localhost:8765/일일판매재고_뷰어.html"
popd
