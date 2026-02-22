@echo off
chcp 65001 >nul
echo 正在安装 calculate-profit 命令...

set TARGET=%USERPROFILE%\.config\opencode\commands

if not exist "%TARGET%" (
    mkdir "%TARGET%"
)

copy /Y "%~dp0calculate_profit.py" "%TARGET%\calculate_profit.py" >nul
copy /Y "%~dp0calculate-profit.md" "%TARGET%\calculate-profit.md" >nul

echo.
echo 安装完成！
echo   命令文件: %TARGET%\calculate-profit.md
echo   脚本文件: %TARGET%\calculate_profit.py
echo.
echo 请重启 OpenCode，然后使用:
echo   /calculate-profit "文件路径1.xlsx" "文件路径2.xlsx"
pause
