# Agentic Working Skills

面向 [OpenCode](https://opencode.ai) 的自定义命令集合。

## 一键安装

**Windows (PowerShell):**
```powershell
irm https://raw.githubusercontent.com/Luxuzhou/agentic-working-skills/master/install.ps1 | iex
```

**Mac / Linux:**
```bash
curl -fsSL https://raw.githubusercontent.com/Luxuzhou/agentic-working-skills/master/install.sh | bash
```

安装后重启 OpenCode 即可使用所有命令。

## 包含的 Skills

| 命令 | 说明 |
|------|------|
| `/calculate-profit` | 抖店利润表自动计算与回填 |

## 前置要求

- Python 3.8+
- pandas (`pip install pandas`)
- OpenCode

## 详细说明

各 skill 目录下有独立的 README：

- [doudian-profit-tools](./doudian-profit-tools/) - 抖店利润表工具
