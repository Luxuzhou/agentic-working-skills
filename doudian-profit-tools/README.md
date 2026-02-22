# doudian-profit-tools

抖店利润表自动计算与回填工具，配合 [OpenCode](https://opencode.ai) 使用。

## 功能

从 Excel 利润表文件的各工作表中自动计算以下指标并回填：

- **主营业务收入**：销售收入（已结算/待结算/退款）、收入总计
- **主营业务成本**：货品成本、退货成本、成本总计
- **毛利润 / 毛利润率**
- **销售费用**：平台服务费、佣金、运费险、发货超时扣款、小额打款

## 前置要求

- Python 3.8+
- pandas (`pip install pandas`)
- OpenCode

## 安装

```bash
git clone https://github.com/<your-username>/doudian-profit-tools.git
cd doudian-profit-tools
```

**Windows：** 双击 `install.bat`

**Mac/Linux：**
```bash
chmod +x install.sh && ./install.sh
```

## 使用

在 OpenCode 中：

```
/calculate-profit "D:\数据\店铺A利润表.xlsx" "D:\数据\店铺B利润表.xlsx"
```

也可以直接用 Python 运行：

```bash
python calculate_profit.py "店铺A利润表.xlsx" "店铺B利润表.xlsx"
```

## Excel 文件要求

文件必须包含以下工作表：

| 工作表 | 用途 |
|--------|------|
| 利润表 | 回填目标 |
| 结算账单 | 平台服务费、佣金、已结算收入/成本 |
| 明细订单 | 待结算收入 |
| 赔付明细 | 运费险、发货超时扣款 |
| 小额打款 | 小额打款金额 |

### 列名要求

- **结算账单**：结算单类型、收入合计、平台服务费、达人佣金、金额（末列，预匹配成本）
- **明细订单**：订单状态、是否结算、退款方式、售后状态、订单应付金额
- **赔付明细**：赔付场景、动账金额
- **小额打款**：打款金额（元）

## 注意事项

- 处理前请关闭 Excel 文件，否则写入会报 PermissionError
- 结算账单第 2 行为说明行，脚本会自动跳过
- 结算账单末尾如有汇总行（结算单类型为空），脚本会自动排除
