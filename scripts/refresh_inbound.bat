@echo off
chcp 65001 >nul
cd /d "C:\Users\admin\Desktop\17. claude\po"
echo. >> scripts\refresh_sales.log
echo === %date% %time% [inbound only] === >> scripts\refresh_sales.log
python scripts\refresh_sales.py --mode=inbound >> scripts\refresh_sales.log 2>&1
