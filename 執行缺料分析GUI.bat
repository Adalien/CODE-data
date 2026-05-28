@echo off
cd /d "%~dp0"

:: 執行缺料分析 GUI，等待視窗關閉後才繼續
start /wait "" "C:\Users\admin\AppData\Local\Programs\Python\Python314\pythonw.exe" run_gui.py

:: GUI 關閉後，自動更新網頁查詢（顯示進度視窗）
"C:\Users\admin\AppData\Local\Programs\Python\Python314\python.exe" gen_invoice_html.py

echo.
echo 網頁更新完成！
timeout /t 3 >nul
