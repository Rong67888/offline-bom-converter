@echo off
setlocal
chcp 65001 >nul
title 上传到GitHub私有仓库

set "SCRIPT_DIR=%~dp0"
set "PS_SCRIPT=%SCRIPT_DIR%上传到GitHub私有仓库.ps1"

if not exist "%PS_SCRIPT%" (
    echo 找不到PowerShell上传脚本：
    echo %PS_SCRIPT%
    echo 请确认BAT和PS1位于同一个文件夹。
    echo.
    pause
    exit /b 1
)

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%"
set "UPLOAD_EXIT_CODE=%ERRORLEVEL%"

echo.
if "%UPLOAD_EXIT_CODE%"=="0" (
    echo 操作已完成。详细结果请查看 github_upload_result.txt。
) else (
    echo 操作失败。请阅读上方中文提示和 github_upload_result.txt。
)
echo.
pause
exit /b %UPLOAD_EXIT_CODE%
