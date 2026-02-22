#!/bin/bash
# Agentic Working Skills - 一键安装脚本 (Mac/Linux)
# 用法: curl -fsSL https://raw.githubusercontent.com/Luxuzhou/agentic-working-skills/master/install.sh | bash

set -e

REPO_URL="https://github.com/Luxuzhou/agentic-working-skills/archive/refs/heads/master.tar.gz"
TEMP_DIR=$(mktemp -d)
COMMANDS_DIR="$HOME/.config/opencode/commands"

echo ""
echo "=== Agentic Working Skills Installer ==="
echo ""

# 下载并解压
echo "[1/3] 下载仓库..."
curl -fsSL "$REPO_URL" | tar -xz -C "$TEMP_DIR"

# 创建目标目录
mkdir -p "$COMMANDS_DIR"

# 安装所有 skills
echo "[2/3] 安装 skills..."
installed=0
for skill_dir in "$TEMP_DIR"/agentic-working-skills-master/*/; do
    [ -d "$skill_dir" ] || continue
    skill_name=$(basename "$skill_dir")

    # 复制 .md 命令文件（排除 README.md）
    for md_file in "$skill_dir"*.md; do
        [ -f "$md_file" ] || continue
        [ "$(basename "$md_file")" = "README.md" ] && continue
        cp "$md_file" "$COMMANDS_DIR/"
        echo "  + 命令: $(basename "${md_file%.md}")"
    done

    # 复制 .py 脚本文件
    for py_file in "$skill_dir"*.py; do
        [ -f "$py_file" ] || continue
        cp "$py_file" "$COMMANDS_DIR/"
        echo "    脚本: $(basename "$py_file")"
    done

    installed=$((installed + 1))
done

# 清理
echo "[3/3] 清理临时文件..."
rm -rf "$TEMP_DIR"

echo ""
echo "安装完成! 共安装 ${installed} 个 skill"
echo "安装位置: $COMMANDS_DIR"
echo ""
echo "请重启 OpenCode 后使用。"
echo ""
