@echo off
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"

echo ============================================================
echo   正在解析「每日报表」并生成今日通报预览...
echo ============================================================
echo.
python main.py daily-preview
echo.
if errorlevel 1 (
    echo [错误] 预览失败，请检查「每日报表」文件或上方报错信息。
) else (
    echo [完成] 以上为本次预览效果（尚未发送到钉钉群）。
)
echo.
pause
