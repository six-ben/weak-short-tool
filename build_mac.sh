#!/bin/bash
# Mac 打包脚本 - 生成 .app
cd "$(dirname "$0")"

echo "=== 安装依赖 ==="
pip3 install -r requirements.txt

echo "=== 开始打包 Mac 版本 ==="
pyinstaller \
    --name "WeakShortTool" \
    --windowed \
    --onefile \
    --noconfirm \
    --clean \
    main.py

echo "=== 打包完成 ==="
echo "输出路径: dist/WeakShortTool.app (Mac) 或 dist/WeakShortTool"
open dist/
