@echo off
chcp 65001 > nul
cd /d "%~dp0"
python refresh_price_matrix.py >> refresh_price_matrix.log 2>&1
