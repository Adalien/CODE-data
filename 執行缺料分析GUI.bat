@echo off
cd /d "%~dp0"

:: 執行缺料分析 GUI（等待視窗關閉）
"C:\Users\admin\AppData\Local\Programs\Python\Python314\python.exe" run_gui.py

:: GUI 關閉後，自動更新網頁查詢
echo.
echo 正在更新進貨明細查詢網頁...
"C:\Users\admin\AppData\Local\Programs\Python\Python314\python.exe" gen_invoice_html.py

echo.
echo 網頁更新完成！
timeout /t 3 >nul
