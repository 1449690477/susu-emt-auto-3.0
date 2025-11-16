@echo off
:: ========================================
:: 🔧 苏苏自动化 一键打包工具
:: 使用方法：直接把 py 文件拖到这个 bat 上即可
:: ========================================

:: 检查参数
if "%~1"=="" (
    echo 请将要打包的 .py 文件拖到此 bat 上！
    pause
    exit /b
)

:: 提取文件信息
set FILEPATH=%~1
set FILENAME=%~n1
set FILEDIR=%~dp1

:: 切换到文件所在目录
cd /d "%FILEDIR%"

echo ----------------------------------------
echo 🚀 正在打包：%FILENAME%.py
echo ----------------------------------------

:: 调用 Python 3.11 的 PyInstaller 打包命令
py -3.11 -m PyInstaller "%FILENAME%.py" ^
    -n "%FILENAME%" ^
    --onefile --windowed --clean --noconfirm ^
    --distpath . ^
    --workpath build ^
    --specpath . ^
    --add-data "templates;templates" ^
    --add-data "templates_letters;templates_letters" ^
    --add-data "templates_drops;templates_drops" ^
    --add-data "scripts;scripts" ^
    --add-data "SP;SP" ^
    --collect-all cv2 ^
    --collect-all numpy ^
    --collect-all pyautogui ^
    --collect-all Pillow ^
    --collect-submodules keyboard ^
    --hidden-import pygetwindow

echo.
echo ✅ 打包完成！
echo 输出文件位置：
echo %FILEDIR%%FILENAME%.exe
echo.
pause
