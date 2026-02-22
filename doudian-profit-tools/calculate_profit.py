"""
利润表自动计算与回填工具

用法: python calculate_profit.py <文件1.xlsx> [文件2.xlsx] ...

从每个 Excel 文件的各工作表中计算利润表指标，并回填到利润表工作表。
"""
import sys
import os
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd

sys.stdout.reconfigure(encoding='utf-8')


def read_sheets(file_path):
    """读取 Excel 各工作表，返回所需的 DataFrame 字典。"""
    df_settlement = pd.read_excel(file_path, sheet_name='结算账单', skiprows=[1])
    df_settlement = df_settlement[df_settlement['结算单类型'].notna()].reset_index(drop=True)
    df_orders = pd.read_excel(file_path, sheet_name='明细订单')
    df_compensation = pd.read_excel(file_path, sheet_name='赔付明细')
    df_small_payment = pd.read_excel(file_path, sheet_name='小额打款')

    print(f'  结算账单: {len(df_settlement)} 行')
    print(f'  明细订单: {len(df_orders)} 行')
    print(f'  赔付明细: {len(df_compensation)} 行')
    print(f'  小额打款: {len(df_small_payment)} 行')

    return df_settlement, df_orders, df_compensation, df_small_payment


def calculate(df_settlement, df_orders, df_compensation, df_small_payment):
    """计算所有利润表指标，返回结果字典。"""
    # 数值列安全转换
    for col in ['收入合计', '平台服务费', '达人佣金', '金额']:
        df_settlement[col] = pd.to_numeric(df_settlement[col], errors='coerce')
    df_compensation['动账金额'] = pd.to_numeric(df_compensation['动账金额'], errors='coerce')
    df_small_payment['打款金额（元）'] = pd.to_numeric(df_small_payment['打款金额（元）'], errors='coerce')

    # 结算类型掩码
    settled_mask = df_settlement['结算单类型'] == '已结算'
    refund_types = ['结算后退款-原路退回', '结算后退款-非原路退回']
    refund_mask = df_settlement['结算单类型'].isin(refund_types)

    # 1. 主营业务收入
    revenue_settled = df_settlement.loc[settled_mask, '收入合计'].sum()
    revenue_refund = df_settlement.loc[refund_mask, '收入合计'].abs().sum()

    df_orders['订单应付金额'] = pd.to_numeric(df_orders['订单应付金额'], errors='coerce')
    pending_mask = df_orders['订单状态'] == '已完成'
    pending_mask &= df_orders['是否结算'].isna()
    pending_mask &= df_orders['退款方式'].isna()
    exclude_statuses = ['退款成功', '待退货', '待收退货']
    pending_mask &= ~df_orders['售后状态'].isin(exclude_statuses)
    revenue_pending = df_orders.loc[pending_mask, '订单应付金额'].sum()

    revenue_total = revenue_settled - revenue_refund

    # 2. 主营业务成本
    cost_settled = df_settlement.loc[settled_mask, '金额'].sum()
    cost_refund = df_settlement.loc[refund_mask, '金额'].sum()
    cost_total = cost_settled - cost_refund

    # 3. 毛利润
    gross_profit = revenue_total - cost_total
    gross_margin = gross_profit / revenue_total if revenue_total != 0 else 0

    # 4. 销售费用
    platform_fee = abs(df_settlement['平台服务费'].sum())
    commission = abs(df_settlement['达人佣金'].sum())

    freight_scenes = ['运费争议保障运费赔付', '商达责售后运费赔付']
    freight_insurance = abs(
        df_compensation[df_compensation['赔付场景'].isin(freight_scenes)]['动账金额'].sum()
    )

    late_scenes = ['发货超时', '订单缺货无货']
    late_shipment = abs(
        df_compensation[df_compensation['赔付场景'].isin(late_scenes)]['动账金额'].sum()
    )

    small_payment = df_small_payment['打款金额（元）'].sum()

    return {
        'revenue_settled': revenue_settled,
        'revenue_refund': revenue_refund,
        'revenue_pending': revenue_pending,
        'revenue_total': revenue_total,
        'cost_settled': cost_settled,
        'cost_refund': cost_refund,
        'cost_total': cost_total,
        'gross_profit': gross_profit,
        'gross_margin': gross_margin,
        'platform_fee': platform_fee,
        'commission': commission,
        'freight_insurance': freight_insurance,
        'late_shipment': late_shipment,
        'small_payment': small_payment,
    }


def print_report(results, file_name):
    """格式化打印利润表到控制台。"""
    r = results
    print(f'\n{"=" * 60}')
    print(f'  {file_name}')
    print('=' * 60)

    print('\n【主营业务收入】')
    print(f'  销售收入-已结算          {r["revenue_settled"]:>12,.2f}')
    print(f'  销售收入-退款            {r["revenue_refund"]:>12,.2f}')
    print(f'  销售收入-待结算          {r["revenue_pending"]:>12,.2f}')
    print(f'  {"─" * 40}')
    print(f'  收入总计                 {r["revenue_total"]:>12,.2f}')

    print('\n【主营业务成本】')
    print(f'  货品成本-已结算          {r["cost_settled"]:>12,.2f}')
    print(f'  退货成本                 {r["cost_refund"]:>12,.2f}')
    print(f'  {"─" * 40}')
    print(f'  成本总计                 {r["cost_total"]:>12,.2f}')

    print('\n【毛利润】')
    print(f'  毛利润                   {r["gross_profit"]:>12,.2f}')
    print(f'  毛利润率                 {r["gross_margin"]:>11.2%}')

    print('\n【销售费用】')
    print(f'  平台服务费               {r["platform_fee"]:>12,.2f}')
    print(f'  佣金                     {r["commission"]:>12,.2f}')
    print(f'  运费险                   {r["freight_insurance"]:>12,.2f}')
    print(f'  发货超时扣款             {r["late_shipment"]:>12,.2f}')
    print(f'  小额打款                 {r["small_payment"]:>12,.2f}')
    print('=' * 60)


def find_profit_sheet_path(file_path):
    """在 xlsx ZIP 中找到利润表对应的 sheet XML 路径。"""
    with zipfile.ZipFile(file_path, 'r') as z:
        wb_xml = z.read('xl/workbook.xml')
        rels_xml = z.read('xl/_rels/workbook.xml.rels')

    ns_main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    ns_r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

    # 找到利润表的 rId
    wb_root = ET.fromstring(wb_xml)
    target_rid = None
    for sheet in wb_root.findall(f'{{{ns_main}}}sheets/{{{ns_main}}}sheet'):
        if sheet.get('name') == '利润表':
            target_rid = sheet.get(f'{{{ns_r}}}id')
            break

    if target_rid is None:
        return None

    # 通过 rId 找到文件路径
    rels_root = ET.fromstring(rels_xml)
    for rel in rels_root:
        if rel.get('Id') == target_rid:
            return 'xl/' + rel.get('Target')

    return None


def write_back(file_path, results, sheet_path):
    """通过修改 ZIP 内的 sheet XML 回填数据，不碰其他工作表。"""
    r = results
    cell_values = {
        'D6':  r['revenue_settled'],
        'D7':  r['revenue_pending'],
        'D8':  r['revenue_refund'],
        'D9':  r['revenue_total'],
        'D10': r['cost_settled'],
        'D11': r['cost_refund'],
        'D12': r['cost_total'],
        'D13': r['gross_profit'],
        'D14': r['gross_margin'],
        'D18': r['platform_fee'],
        'D19': r['commission'],
        'D20': r['freight_insurance'],
        'D21': r['late_shipment'],
        'D26': r['small_payment'],
    }

    NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

    with zipfile.ZipFile(file_path, 'r') as zin:
        sheet_xml = zin.read(sheet_path)

    ET.register_namespace('', NS)
    ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    ET.register_namespace('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
    ET.register_namespace('x14ac', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac')

    root = ET.fromstring(sheet_xml)
    ns_map = {'s': NS}

    for row in root.findall('.//s:row', ns_map):
        for cell in row.findall('s:c', ns_map):
            ref = cell.get('r', '')
            if ref in cell_values:
                if 't' in cell.attrib:
                    del cell.attrib['t']
                v_elem = cell.find('s:v', ns_map)
                if v_elem is None:
                    v_elem = ET.SubElement(cell, f'{{{NS}}}v')
                v_elem.text = str(cell_values[ref])

    temp_path = file_path + '.tmp'
    backup_path = file_path + '.bak'
    with zipfile.ZipFile(file_path, 'r') as zin:
        with zipfile.ZipFile(temp_path, 'w') as zout:
            for item in zin.infolist():
                if item.filename == sheet_path:
                    modified_xml = ET.tostring(root, encoding='unicode', xml_declaration=True)
                    zout.writestr(item, modified_xml.encode('utf-8'))
                else:
                    zout.writestr(item, zin.read(item.filename))

    # Windows 下 os.replace 可能因文件锁定失败，改用 rename 链
    if os.path.exists(backup_path):
        os.remove(backup_path)
    os.rename(file_path, backup_path)
    os.rename(temp_path, file_path)
    os.remove(backup_path)


def verify(file_path):
    """重新读取利润表验证回填结果。"""
    df = pd.read_excel(file_path, sheet_name='利润表', header=None)
    check_rows = {
        5:  '销售收入-已结算',
        6:  '销售收入-待结算',
        7:  '销售收入-退款',
        8:  '收入总计',
        9:  '货品成本-已结算',
        10: '退货成本',
        11: '成本总计',
        12: '毛利润',
        13: '毛利润率',
        17: '平台服务费',
        18: '佣金',
        19: '运费险',
        20: '发货超时扣款',
        25: '小额打款',
    }
    print('\n【回填验证】')
    for idx, label in check_rows.items():
        val = df.iloc[idx, 3]
        if isinstance(val, float) and abs(val) < 1:
            print(f'  {label:<20s} {val:.2%}')
        else:
            print(f'  {label:<20s} {val:>12,.2f}')


def process_file(file_path):
    """处理单个 Excel 文件：计算 → 打印 → 回填 → 验证。"""
    file_name = os.path.basename(file_path)
    print(f'\n{"#" * 60}')
    print(f'  处理文件: {file_name}')
    print('#' * 60)

    # 读取
    print('\n正在读取...')
    df_settlement, df_orders, df_compensation, df_small_payment = read_sheets(file_path)

    # 计算
    results = calculate(df_settlement, df_orders, df_compensation, df_small_payment)

    # 打印
    print_report(results, file_name)

    # 找到利润表 sheet 路径
    sheet_path = find_profit_sheet_path(file_path)
    if sheet_path is None:
        print('\n  错误: 未找到"利润表"工作表，跳过回填。')
        return False

    # 回填
    print('\n正在回填利润表...')
    write_back(file_path, results, sheet_path)
    print('回填完成。')

    # 验证
    verify(file_path)
    return True


def main():
    if len(sys.argv) < 2:
        print('用法: python calculate_profit.py <文件1.xlsx> [文件2.xlsx] ...')
        sys.exit(1)

    files = sys.argv[1:]
    success = 0
    failed = 0

    for f in files:
        if not os.path.isfile(f):
            print(f'\n错误: 文件不存在 - {f}')
            failed += 1
            continue
        try:
            if process_file(f):
                success += 1
            else:
                failed += 1
        except Exception as e:
            print(f'\n处理失败: {f}')
            print(f'  错误: {e}')
            failed += 1

    print(f'\n{"=" * 60}')
    print(f'  全部完成: 成功 {success} 个, 失败 {failed} 个')
    print('=' * 60)


if __name__ == '__main__':
    main()
