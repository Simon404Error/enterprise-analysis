# -*- coding: utf-8 -*-
"""
可选工具：生成演示数据（「年度报告记录（地区名）.xlsx」格式），便于开箱试用与测试。
运行：python sample_data_generator.py [输出目录]，默认输出到 ./data

演示数据包含：
  - 两个示例地区（示范镇A / 示范镇B），2022~2024 三个年度；
  - 示范镇A 拆分为两个分卷（地区名 / 地区名1），含“同企业同年多条记录”与“完全重复行”，
    用于验证加总合并与去重；
  - 个别企业缺少年度，用于验证增长率标记 NA。
"""
import os
import sys

import numpy as np
import pandas as pd

RNG = np.random.default_rng(42)

REGIONS = ['示范镇A', '示范镇B']
YEARS = [2022, 2023, 2024]
N_ENTERPRISES = {'示范镇A': 12, '示范镇B': 10}

COLUMNS = ['企业名称', '年度', '资产总额', '所有者权益合计', '销售额或营业收入',
           '利润总额', '营业总收入中主营业务收入', '净利润', '负债总额']


def gen_features():
    return dict(
        base_assets=float(np.exp(RNG.normal(8.0, 0.8))),
        turnover=float(RNG.uniform(0.5, 1.5)),
        margin=float(RNG.uniform(0.02, 0.15)),
        debt_ratio=float(RNG.uniform(0.2, 0.7)),
        focus=float(RNG.uniform(0.8, 0.98)),
        growth=float(RNG.uniform(-0.05, 0.2)),
    )


def year_row(f, y, name):
    assets = f['base_assets'] * (1 + f['growth']) ** (y - YEARS[0])
    sales = assets * f['turnover'] * (1 + RNG.normal(0, 0.04))
    main = sales * f['focus']
    profit = sales * f['margin']
    gross = profit * 1.15 if profit > 0 else profit
    debt = assets * f['debt_ratio']
    equity = assets - debt
    return [name, y, round(assets, 2), round(equity, 2), round(sales, 2),
            round(gross, 2), round(main, 2), round(profit, 2), round(debt, 2)]


def build_region(region, indices):
    """为地区生成指定编号（0 基）企业的全部年度数据。"""
    rows = []
    for i in indices:
        name = f'{region}{i + 1:02d}号企业'
        f = gen_features()
        for y in YEARS:
            # 缺年度测试：示范镇A 02号缺2023；示范镇B 03号缺2023
            if region == '示范镇A' and i == 1 and y == 2023:
                continue
            if region == '示范镇B' and i == 2 and y == 2023:
                continue
            rows.append(year_row(f, y, name))
    return rows


def main(out_dir):
    os.makedirs(out_dir, exist_ok=True)

    # 示范镇A：分卷1 = 企业 1-6；分卷2 = 企业 7-12 + 测试行
    vol1 = build_region('示范镇A', range(0, 6))
    vol2 = build_region('示范镇A', range(6, 12))
    # 测试“同企业同年多条记录加总”：01号企业 2022 在分卷2中追加一条 销售额+10 的记录
    vol2.append(year_row(gen_features(), 2022, '示范镇A01号企业'))
    vol2[-1][4] = round(vol2[-1][4] + 10, 2)  # 销售额列 +10
    # 测试“完全重复行去重”：07号企业 2023 追加一条完全相同的记录
    src = pd.DataFrame(vol1 + vol2, columns=COLUMNS)
    dup = src[(src['企业名称'] == '示范镇A07号企业') & (src['年度'] == 2023)].iloc[0].tolist()
    vol2.append(dup)

    pd.DataFrame(vol1, columns=COLUMNS).to_excel(os.path.join(out_dir, '年度报告记录（示范镇A）.xlsx'), index=False)
    pd.DataFrame(vol2, columns=COLUMNS).to_excel(os.path.join(out_dir, '年度报告记录（示范镇A1）.xlsx'), index=False)
    print('生成 年度报告记录（示范镇A）.xlsx / 年度报告记录（示范镇A1）.xlsx')

    shihua = build_region('示范镇B', range(0, N_ENTERPRISES['示范镇B']))
    pd.DataFrame(shihua, columns=COLUMNS).to_excel(os.path.join(out_dir, '年度报告记录（示范镇B）.xlsx'), index=False)
    print('生成 年度报告记录（示范镇B）.xlsx')
    print('完成。')


if __name__ == '__main__':
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')
    main(out)
