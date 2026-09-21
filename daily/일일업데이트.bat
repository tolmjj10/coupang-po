@echo off
chcp 65001 >nul
setlocal
pushd "%~dp0"
echo === 일일판매재고 HTML 재생성 + 배포 ===
echo.

echo [1/3] HTML 생성 중...
python xlsb_to_luckysheet.py
if errorlevel 1 (
  echo !! HTML 생성 실패
  popd
  pause
  exit /b 1
)

echo.
echo [2/3] Git 커밋 + 푸시...
cd /d "%~dp0.."
git pull --rebase 2>nul
git add daily/index.html
git commit -m "daily: 일일판매재고 뷰어 업데이트" 2>nul
git push
if errorlevel 1 (
  echo !! 푸시 실패 - 나중에 다시 시도
  pause
  exit /b 1
)

echo.
echo [3/3] Cloudflare Pages 자동 배포 대기 (약 2-3분)
echo.
echo 배포 완료 후 확인: https://coupang-po.pages.dev/daily/
echo.
timeout /t 5 /nobreak >nul
start "" "https://coupang-po.pages.dev/daily/"
popd
