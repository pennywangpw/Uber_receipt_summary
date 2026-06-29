@echo off
chcp 65001 >nul
cd /d "%~dp0"

title Uber Eats 自動記帳轉 Excel 工具

echo ===================================================
echo     歡迎使用 Uber Eats 收據下載暨 Excel 自動化工具
echo ===================================================
echo.

set /p start_date="📌 請輸入【開始日期】(格式範例 2026-06-24): "
set /p end_date="📌 請輸入【結束日期】(格式範例 2026-06-24): "

:: 🎯 新增：讓使用者輸入既有的 Excel 檔名
set /p excel_name="📌 請輸入【既有 Excel 檔案名稱】(例如 expense.xlsx): "

echo.
echo 🚀 正在啟動瀏覽器與後台 PDF/JSON 解析引擎，請稍候...
echo ---------------------------------------------------

:: 🎯 修改：將 %excel_name% 一併傳給 Python 腳本
python main.py --start %start_date% --end %end_date% --excel "%excel_name%"

echo ---------------------------------------------------
echo.
echo 🎉 執行程序已結束！
echo 💡 提示：請檢查資料夾內是否已順利生成對應日期的 Excel 分頁。
echo.

pause