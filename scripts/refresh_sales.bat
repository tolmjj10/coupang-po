@echo off
REM 작업 스케줄러에서 호출: 매일 09:00
REM 로그는 scripts\refresh_sales.log 에 날짜별 추가
cd /d "C:\Users\admin\Desktop\17. claude\po"
echo. >> scripts\refresh_sales.log
echo === %date% %time% === >> scripts\refresh_sales.log
python scripts\refresh_sales.py >> scripts\refresh_sales.log 2>&1
