@echo off
cd /d "%~dp0"

echo.
echo  ╔══════════════════════════════════╗
echo  ║  党员微信群成员核对工具  v1.1  ║
echo  ╚══════════════════════════════════╝
echo.

:: ── 文件路径配置（改这里就行）──────────────────
set "EXCEL=D:\xwechat_files\wxid_ca3oc4kna5gz22_f609\msg\file\2026-06\2026届毕业生党员名单-在群.xlsx"
set "OUTPUT=D:\核对结果_党员微信群.xlsx"
:: ──────────────────────────────────────────────

:: 如果拖了文件进来，就用拖进来的
if not "%~1"=="" (
    set "WECHAT_FILE=%~1"
    echo  拖入: %~1
)

:: 构建命令行
set "PYTHON=C:\Users\HONOR\AppData\Local\Programs\Python\Python37\python.exe"
set "CMD=%PYTHON% party_wechat_check.py"
if exist "%EXCEL%" (
    set "CMD=%CMD% --excel "%EXCEL%""
    echo  Excel: %EXCEL%
) else (
    echo  Excel: （未设置，稍后手动输入）
)
if defined WECHAT_FILE (
    if exist "%WECHAT_FILE%" (
        set "CMD=%CMD% --wechat-file "%WECHAT_FILE%""
        echo  微信群: %WECHAT_FILE%
    )
)
if exist "%OUTPUT%\.." (
    set "CMD=%CMD% --output "%OUTPUT%""
    echo  输出: %OUTPUT%
)
echo.
echo  按 Enter 开始...
pause >nul

%CMD%

echo.
echo  核对完成！按任意键打开 D 盘...
pause >nul
start "" "D:\"
