# -*- coding: utf-8 -*-
"""
示例 / 演示数据生成器 —— 用于本地验证与开箱体验（所有企业名、地区名、数值均为虚构）。

生成规则（与真实数据一致）：
  - 文件名：年度报告记录 (地区名).xlsx；分卷：年度报告记录 (地区名1).xlsx
  - data/ 下按年份分子文件夹（如 2025/、2026/），年度取自文件夹名
  - 表头含 12 项业务字段（不含「年度」列）+ 企业名称
  - 同一地区拆多卷时，同一企业同一年度可能跨卷出现（用于验证数值加总）
  - 演示公示三态（全部/部分/全部不公示）、缺年度（增长率 NA）、净利率上升趋势

运行：python sample_data_generator.py [输出目录，默认 ./data]
"""
import os
import sys

import numpy as np
import pandas as pd

RNG = np.random.default_rng(20250824)

YEARS = [2025, 2026]                                   # 年份子文件夹
# 地区 -> 分卷数（>1 表示拆卷，写多个同名地区不同分卷号的文件）
REGION_VOLUMES = {'示范镇A': 2, '示范镇B': 1}
ENTERPRISES_PER_REGION = {'示范镇A': 8, '示范镇B': 5}

# 表头（12 项业务字段 + 企业名称；不含年度，年度来自年份文件夹）
COLUMNS = ['企业名称', '是否股权转让', '是否对外投资',
           '资产总额', '所有者权益合计', '销售额或营业收入', '利润总额',
           '营业总收入中主营业务收入', '净利润', '负债总额', '纳税总额',
           '资产认缴额', '资产实缴额']
TEXT_COLS = ['是否股权转让', '是否对外投资']


def _features():
    """单个企业的恒定经营特征（跨年度一致）。含部分亏损企业。"""
    return {
        'assets': float(np.exp(RNG.normal(8.0, 0.8))),
        'turnover': float(RNG.uniform(0.5, 1.5)),      # 销售额/资产
        'margin': float(RNG.uniform(-0.05, 0.15)),      # 净利润率（可为负，制造亏损企业）
        'debt_ratio': float(RNG.uniform(0.2, 0.7)),     # 资产负债率
        'focus': float(RNG.uniform(0.85, 0.99)),        # 主营业务收入占比
        'growth': float(RNG.uniform(-0.05, 0.25)),      # 年度销售额增速
        'tax_rate': float(RNG.uniform(0.01, 0.06)),     # 税负率
    }


def _publish_mode(idx):
    """按企业序号安排公示模式：all=全部公示，part=部分公示，none=全部不公示。"""
    return ['all', 'all', 'part', 'none', 'all', 'part', 'all', 'none'][idx % 8]


def _row(name, feats, mode, year_idx):
    """生成一个年度的 12 项业务数据（数值均万元；文本列用“是/否”）。

    净利润率逐年小幅抬升（margin + 0.008×year_idx），使地区整体净利率呈上升趋势，
    以验证“净利润率总体上升 → 营收/资本/利润扩张”分析。
    """
    assets = feats['assets'] * (1 + feats['growth']) ** year_idx
    sales = assets * feats['turnover']
    main = sales * feats['focus']
    margin = feats['margin'] + 0.03 * year_idx             # 逐年抬升（使地区净利率呈上升趋势）
    profit = sales * margin
    gross = profit * 1.15 if profit > 0 else profit        # 利润总额 ≈ 净利润 + 所得税
    equity = assets * (1 - feats['debt_ratio'])
    debt = assets * feats['debt_ratio']
    tax = sales * feats['tax_rate']
    subscribed = assets * 0.6                               # 资产认缴额
    paid = subscribed * 0.85                                # 资产实缴额

    if mode == 'none':                                      # 全部不公示：全部空白/0
        return [name, '', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    if mode == 'part':                                      # 部分公示：纳税=0、文本部分空白
        return [name, '是', '', round(assets, 2), round(equity, 2), round(sales, 2),
                round(gross, 2), round(main, 2), round(profit, 2), round(debt, 2),
                0, round(subscribed, 2), round(paid, 2)]
    # 全部公示
    return [name, '是', '是', round(assets, 2), round(equity, 2), round(sales, 2),
            round(gross, 2), round(main, 2), round(profit, 2), round(debt, 2),
            round(tax, 2), round(subscribed, 2), round(paid, 2)]


def _region_rows(region, year, year_idx, vol):
    """生成某地区在某年度（某一分卷）的企业行。

    拆卷时：偶数序号企业进分卷1、奇数进分卷2；示范镇A 的 01 号（i=0）在两卷都出现，
    用于验证“同一企业同一年度跨卷出现时按数值加总”。
    """
    n = ENTERPRISES_PER_REGION[region]
    volumes = REGION_VOLUMES[region]
    rows = []
    for i in range(n):
        # 缺年度测试：示范镇A 的 01 号企业（i=0）缺 2026（其增长率记为 NA）
        if region == '示范镇A' and i == 0 and year == 2026:
            continue
        if volumes > 1:
            if vol == 1 and i % 2 != 0:            # 卷1只收偶数号
                continue
            if vol == 2 and (i % 2 == 0 and i != 0):  # 卷2收奇数号 + 01号
                continue
        feats = _features()
        # 01号（缺年度企业）：固定中值净利率、低资产权重，避免扰动地区整体净利率趋势
        if region == '示范镇A' and i == 0:
            feats['margin'] = 0.08
            feats['assets'] = np.exp(6.5)
        rows.append(_row(f'{region}{i + 1:02d}号企业', feats, _publish_mode(i), year_idx))
    return rows


def main(out_dir):
    os.makedirs(out_dir, exist_ok=True)
    for year in YEARS:
        # 企业序号 0（示范镇A 01号）在 2026 缺年度，故 2026 分卷企业数变化；以下用年份文件夹
        ydir = os.path.join(out_dir, str(year))
        os.makedirs(ydir, exist_ok=True)
        year_idx = YEARS.index(year)
        for region, volumes in REGION_VOLUMES.items():
            for v in range(1, volumes + 1):
                rows = _region_rows(region, year, year_idx, v)
                if not rows:
                    continue
                fname = f'年度报告记录 ({region}{v if volumes > 1 else ""}).xlsx'
                pd.DataFrame(rows, columns=COLUMNS).to_excel(
                    os.path.join(ydir, fname), index=False)
                print(f'生成 {ydir}/{fname}  行数={len(rows)}')
    print('完成。')


if __name__ == '__main__':
    out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')
    main(out)