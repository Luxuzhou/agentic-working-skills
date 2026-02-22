#!/bin/bash
echo "正在安装 calculate-profit 命令..."

TARGET="$HOME/.config/opencode/commands"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

mkdir -p "$TARGET"

cp "$SCRIPT_DIR/calculate_profit.py" "$TARGET/calculate_profit.py"
cp "$SCRIPT_DIR/calculate-profit.md" "$TARGET/calculate-profit.md"

echo ""
echo "安装完成！"
echo "  命令文件: $TARGET/calculate-profit.md"
echo "  脚本文件: $TARGET/calculate_profit.py"
echo ""
echo "请重启 OpenCode，然后使用:"
echo '  /calculate-profit "文件路径1.xlsx" "文件路径2.xlsx"'
