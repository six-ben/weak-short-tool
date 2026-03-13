@echo off
REM Windows 打包脚本 - 生成 .exe
cd /d "%~dp0"

echo === 安装依赖 ===
pip install -r requirements.txt

echo === 开始打包 Windows 版本 ===
pyinstaller ^
    --name "WeakShortTool" ^
    --windowed ^
    --onefile ^
    --noconfirm ^
    --clean ^
    main.py

echo === 打包完成 ===
echo 输出路径: dist\WeakShortTool.exe
explorer dist
pause
