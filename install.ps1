
# Agentic Working Skills - 一键安装脚本 (Windows PowerShell)
# 用法: irm https://raw.githubusercontent.com/Luxuzhou/agentic-working-skills/master/install.ps1 | iex

$ErrorActionPreference = "Stop"
$RepoUrl = "https://github.com/Luxuzhou/agentic-working-skills/archive/refs/heads/master.zip"
$TempZip = "$env:TEMP\agentic-working-skills.zip"
$TempDir = "$env:TEMP\agentic-working-skills-master"
$CommandsDir = "$env:USERPROFILE\.config\opencode\commands"

Write-Host ""
Write-Host "=== Agentic Working Skills Installer ===" -ForegroundColor Cyan
Write-Host ""

# 下载
Write-Host "[1/4] 下载仓库..." -ForegroundColor Yellow
if (Test-Path $TempDir) { Remove-Item $TempDir -Recurse -Force }
if (Test-Path $TempZip) { Remove-Item $TempZip -Force }
Invoke-WebRequest -Uri $RepoUrl -OutFile $TempZip

# 解压
Write-Host "[2/4] 解压文件..." -ForegroundColor Yellow
Expand-Archive -Path $TempZip -DestinationPath $env:TEMP -Force

# 创建目标目录
if (!(Test-Path $CommandsDir)) {
    New-Item -ItemType Directory -Path $CommandsDir -Force | Out-Null
}

# 安装所有 skills
Write-Host "[3/4] 安装 skills..." -ForegroundColor Yellow
$installed = 0
Get-ChildItem "$TempDir" -Directory | ForEach-Object {
    $skillDir = $_.FullName
    $skillName = $_.Name

    # 复制 .md 命令文件
    Get-ChildItem $skillDir -Filter "*.md" | Where-Object { $_.Name -ne "README.md" } | ForEach-Object {
        Copy-Item $_.FullName "$CommandsDir\$($_.Name)" -Force
        Write-Host "  + 命令: $($_.BaseName)" -ForegroundColor Green
    }

    # 复制 .py 脚本文件
    Get-ChildItem $skillDir -Filter "*.py" | ForEach-Object {
        Copy-Item $_.FullName "$CommandsDir\$($_.Name)" -Force
        Write-Host "    脚本: $($_.Name)" -ForegroundColor Gray
    }

    $installed++
}

# 清理
Write-Host "[4/4] 清理临时文件..." -ForegroundColor Yellow
Remove-Item $TempZip -Force -ErrorAction SilentlyContinue
Remove-Item $TempDir -Recurse -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "安装完成! 共安装 $installed 个 skill" -ForegroundColor Green
Write-Host "安装位置: $CommandsDir" -ForegroundColor Gray
Write-Host ""
Write-Host "请重启 OpenCode 后使用。" -ForegroundColor Cyan
Write-Host ""
