@echo off
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"

echo ============================================================
echo   准备推送每日通报到钉钉群（将发送主卡片 + @提醒）
echo ============================================================
echo.
set "CONFIRM="
set /p "CONFIRM=确认正式推送到钉钉群吗？输入 Y 回车发送，其他键取消: "
echo.
if /i "%CONFIRM%"=="Y" (
    python main.py daily-push
    echo.
    if errorlevel 1 (
        echo [错误] 发送失败，请检查上方报错信息或网络。
    ) else (
        echo [完成] 已成功发送到钉钉群！
    )
) else (
    echo 已取消发送。
)
echo.
pause
