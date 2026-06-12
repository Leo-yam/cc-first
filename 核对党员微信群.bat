@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo.
echo  ╔══════════════════════════════════╗
echo  ║  党员微信群成员核对工具  v1.1  ║
echo  ╚══════════════════════════════════╝
echo.

:: ── 文件路径配置（改这里就行）──────────────────
set "EXCEL=D:\xwechat_files\wxid_ca3oc4kna5gz22_f609\msg\file\2026-06\2026届毕业生党员名单-在群.xlsx"
set "OUTPUT=D:\核对结果_党员微信群.xlsx"

:: 微信群文本：优先用拖拽进来的文件，否则用默认路径
if not "%~1"=="" (
    set "WECHAT=%~1"
    echo  拖入文件: %~1
) else (
    set "WECHAT=C:\Users\HONOR\Desktop\微信群成员.txt"
)
:: ──────────────────────────────────────────────

:: 检查 Excel 在不在
if not exist "%EXCEL%" (
    echo  ❌ 找不到 Excel 文件！
    echo     %EXCEL%
    echo.
    echo  右键编辑这个 bat 文件，修改 EXCEL 的路径。
    pause
    exit /b 1
)

:: 检查微信群文本在不在
if not exist "%WECHAT%" (
    echo  ❌ 找不到微信群文本！
    echo     %WECHAT%
    echo.
    echo  用法 1: 把微信群OCR文本保存为 "微信群成员.txt" 放在桌面
    echo  用法 2: 把 txt 文件直接拖到这个 bat 图标上
    pause
    exit /b 1
)

echo   Excel: %EXCEL%
echo   微信群: %WECHAT%
echo   输出:   %OUTPUT%
echo.
echo  按 Enter 开始核对...
pause >nul

python party_wechat_check.py ^
  --excel "%EXCEL%" ^
  --wechat-file "%WECHAT%" ^
  --output "%OUTPUT%"

echo.
echo  ┌─────────────────────────────┐
echo  │  核对完成！                │
echo  └─────────────────────────────┘
echo.
echo  结果文件: %OUTPUT%
echo.
echo  按任意键打开结果文件夹...
pause >nul
start "" "D:\"
