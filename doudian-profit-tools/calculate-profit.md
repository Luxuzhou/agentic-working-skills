---
description: 计算抖店利润表并回填到Excel文件
---

# 利润表自动计算与回填

使用 calculate_profit.py 脚本处理用户指定的 Excel 利润表文件。

## 脚本路径

脚本安装在与本命令文件相同的目录下:

!`echo ~/.config/opencode/commands/calculate_profit.py`

## 执行步骤

1. 解析用户提供的文件路径: $ARGUMENTS
2. 确认所有文件存在
3. 执行命令（路径含空格需加引号）:

```bash
python ~/.config/opencode/commands/calculate_profit.py $ARGUMENTS
```

4. 检查输出结果，确认每个文件的【回填验证】均通过
5. 如有失败项，向用户报告具体错误原因

## 注意事项

- 处理前确保目标 Excel 文件未被 Excel 等程序打开，否则写入会报 PermissionError
- 文件路径含空格时需用引号包裹
- 脚本会自动跳过不存在的文件并报告
