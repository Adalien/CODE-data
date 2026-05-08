@echo off
chcp 65001 > nul
cd /d "C:\Users\admin\OneDrive\桌面\CODE資料"
echo 正在執行缺料分析，請稍候...
python 缺料分析.py
echo.
echo 完成！請到 CODE資料 資料夾查看今天的分析結果。
pause
