# -*- coding: utf-8 -*-
"""
报告生成模块：
  1) 生成静态备查 Excel：分析结果汇总.xlsx
     - Sheet1「明细」：所有企业各年度完整指标表（百分数列为 ×100 展示）
     - Sheet2「企业汇总」：每个企业一行的汇总统计
  2) 动态生成文字经营报告（Markdown），供面板展示/下载
"""
import numpy as np
import pandas as pd

from config import (OUTPUT_XLSX, COL_REGION, COL_ENTERPRISE, COL_YEAR,
                    NUMERIC_COLS, RATIO_COLS, GROWTH_COLS, STD_COLS)
from analysis import add_metrics, enterprise_summary, metrics_for_display, safe_div


# ---------------- 静态 Excel ----------------
def write_result_excel(df, path=None):
    """将完整指标表写入 Excel。df 应为已计算指标的明细表。"""
    path = path or OUTPUT_XLSX
    detail = metrics_for_display(df)      # 百分数列 ×100
    summary = metrics_for_display(enterprise_summary(df))

    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        detail.to_excel(writer, sheet_name='明细', index=False)
        summary.to_excel(writer, sheet_name='企业汇总', index=False)

    # 调整列宽（openpyxl 直接操作）
    from openpyxl import load_workbook
    wb = load_workbook(path)
    for ws in wb.worksheets:
        for cell in ws[1]:
            ws.column_dimensions[cell.column_letter].width = 16
    wb.save(path)
    return path


# ---------------- 文本报告 ----------------
def _pct(v, digits=2):
    """比例 -> '12.34%'；NaN/None -> 'NA'。"""
    if v is None or pd.isna(v):
        return 'NA'
    return f'{v * 100:.{digits}f}%'


def _fmt(v, digits=2):
    if v is None or pd.isna(v):
        return 'NA'
    return f'{v:,.{digits}f}'


def _num_with_year(df, col):
    """数据中出现过多少个不同的年度值。"""
    return sorted(df[COL_YEAR].dropna().unique())


def generate_region_report(df, region):
    """按地区生成文字报告（Markdown）。"""
    sub = df[df[COL_REGION] == region]
    if sub.empty:
        return f'# 【{region}】\n\n（该地区无数据）'
    years = _num_with_year(sub, COL_YEAR)
    n_ent = sub[COL_ENTERPRISE].nunique()
    latest = years[-1]
    cur = sub[sub[COL_YEAR] == latest]

    L = [f'# 【{region}】经营状况报告', '']
    L.append(f'- 覆盖年度：{years[0]} ~ {years[-1]}（共 {len(years)} 个年度）')
    L.append(f'- 企业数量：{n_ent} 家；明细记录：{len(sub)} 条')
    L.append('')

    # 最新年度整体盈利
    sales_sum = cur['销售额或营业收入'].sum()
    profit_sum = cur['净利润'].sum()
    equity_sum = cur['所有者权益合计'].sum()
    L.append(f'## 盈利能力（{latest} 年）')
    L.append(f'- 销售总额：{_fmt(sales_sum)} 万元；净利润总额：{_fmt(profit_sum)} 万元')
    L.append(f'- 整体销售净利率：{_pct(safe_div(profit_sum, sales_sum))}')
    L.append(f'- 整体净资产收益率(ROE)：{_pct(safe_div(profit_sum, equity_sum))}')
    L.append('')

    # 增长趋势
    L.append('## 增长趋势')
    g = sub.groupby(COL_YEAR, as_index=False)['销售额或营业收入'].sum()
    for i in range(1, len(g)):
        y = int(g.iloc[i][COL_YEAR])
        prev = g.iloc[i - 1]['销售额或营业收入']
        c = g.iloc[i]['销售额或营业收入']
        if pd.notna(prev) and prev != 0:
            L.append(f'- {y} 年销售额 {_fmt(c)} 万元，同比 {_pct((c - prev) / prev)}')
        else:
            L.append(f'- {y} 年销售额 {_fmt(c)} 万元，同比 NA（上年数据缺失）')
    L.append('')

    # 稳定性（企业级标准差）
    L.append('## 经营稳定性')
    std_s = sub.groupby(COL_ENTERPRISE)['净利润'].std()
    avg_std = std_s.mean()
    n_unstable = (std_s > std_s.median()).sum() if len(std_s) else 0
    L.append(f'- 企业净利润标准差均值：{_fmt(avg_std)} 万元（值越大波动越大）')
    if len(std_s):
        worst = std_s.idxmax()
        L.append(f'- 波动最大企业：{worst}（标准差 {_fmt(std_s.max())} 万元）')
    L.append('')

    # 异常提示
    L.append('## 异常数据提示')
    n_first_year = len(sub[sub[COL_YEAR] == years[0]])
    tips = []
    if len(years) == 1:
        tips.append('仅有 1 个年度数据，无法计算增长率（同比/环比均为 NA）。')
    else:
        n_na = int(sub['销售额同比增长率'].isna().sum() - n_first_year)
        if n_na > 0:
            tips.append(f'有 {n_na} 条记录因上年数据缺失（或上年为 0）无法计算增长率，已标记 NA。')
    n_zero_equity = int((sub['所有者权益合计'] == 0).sum())
    if n_zero_equity > 0:
        tips.append(f'有 {n_zero_equity} 条记录所有者权益为 0，ROE 无法计算（NA）。')
    n_neg_equity = int((sub['所有者权益合计'] < 0).sum())
    if n_neg_equity > 0:
        tips.append(f'有 {n_neg_equity} 条记录所有者权益为负，ROE 为负值，注意复核数据。')
    missing_years = []
    for ent, g in sub.groupby(COL_ENTERPRISE):
        ys = set(g[COL_YEAR].unique())
        for y in range(min(ys), max(ys) + 1):
            if y not in ys:
                missing_years.append((ent, y))
    if missing_years:
        tips.append(f'有 {len(missing_years)} 个“企业-年度”存在年度断档（如企业某年缺失），'
                    f'相关增长率为 NA。示例：{missing_years[0][0]} 缺 {missing_years[0][1]} 年。')
    if not tips:
        tips.append('未发现明显异常。')
    for t in tips:
        L.append(f'- {t}')
    L.append('')
    return '\n'.join(L)


def generate_enterprise_report(df, region, enterprise):
    """按企业生成文字报告（Markdown）。"""
    sub = df[(df[COL_REGION] == region) & (df[COL_ENTERPRISE] == enterprise)]
    if sub.empty:
        return f'# 【{enterprise}】\n\n（无数据）'
    sub = sub.sort_values(COL_YEAR)
    years = list(sub[COL_YEAR])

    L = [f'# 【{enterprise}】经营状况报告', '']
    L.append(f'- 地区：{region}；覆盖年度：{years[0]} ~ {years[-1]}（共 {len(years)} 年）')
    L.append('')

    L.append('## 各年度核心指标')
    L.append('| 年度 | 销售额(万元) | 净利润(万元) | 销售净利率 | ROE | 销售额同比 | 净利润同比 |')
    L.append('|---|---|---|---|---|---|---|')
    for _, r in sub.iterrows():
        L.append(f"| {int(r[COL_YEAR])} | {_fmt(r['销售额或营业收入'])} | {_fmt(r['净利润'])} | "
                 f"{_pct(r['销售净利率'])} | {_pct(r['净资产收益率(ROE)'])} | "
                 f"{_pct(r['销售额同比增长率'])} | {_pct(r['净利润同比增长率'])} |")
    L.append('')

    L.append('## 综合评价')
    first, last = sub.iloc[0], sub.iloc[-1]
    # 盈利能力趋势
    margin0, margin1 = first['销售净利率'], last['销售净利率']
    if pd.notna(margin0) and pd.notna(margin1):
        trend = '提升' if margin1 > margin0 else ('下降' if margin1 < margin0 else '持平')
        L.append(f'- 盈利能力：销售净利率由 {_pct(margin0)}（{int(first[COL_YEAR])} 年）'
                 f"{trend}至 {_pct(margin1)}（{int(last[COL_YEAR])} 年）。")
    # 稳定性
    std_profit = sub['净利润'].std()
    std_sales = sub['销售额或营业收入'].std()
    L.append(f'- 稳定性：净利润标准差 {_fmt(std_profit)} 万元，销售额标准差 {_fmt(std_sales)} 万元'
             f"{'（仅一个年度，波动无法评估）' if len(years) < 2 else '。'}")
    # 最新增长
    g = last['销售额同比增长率']
    if pd.isna(g):
        L.append(f'- 最新年度（{int(last[COL_YEAR])}）销售额同比：NA（上年数据缺失或为 0）。')
    else:
        L.append(f"- 最新年度（{int(last[COL_YEAR])}）销售额同比 {_pct(g)}，净利润同比 {_pct(last['净利润同比增长率'])}。")
    # 异常
    probs = []
    if pd.isna(last['净资产收益率(ROE)']):
        probs.append('最新年度 ROE 无法计算（所有者权益为 0 或缺失）')
    elif last['净资产收益率(ROE)'] < 0:
        probs.append('最新年度 ROE 为负（所有者权益为负或亏损）')
    if len(years) < 2:
        probs.append('仅单年度数据，增长与波动指标不可用')
    if probs:
        L.append(f'- ⚠ 异常提示：{"；".join(probs)}。')
    else:
        L.append('- 未发现明显异常。')
    L.append('')
    return '\n'.join(L)


def generate_report(df, region=None, enterprise=None):
    """总入口：region/enterprise 均为空时生成地区级报告合集。"""
    if df is None or df.empty:
        return '# 经营状况报告\n\n（当前筛选条件下无数据）'
    if enterprise:
        return generate_enterprise_report(df, region, enterprise)
    if region:
        return generate_region_report(df, region)
    # 全部地区：逐个生成并拼接
    parts = []
    for r in sorted(df[COL_REGION].unique()):
        parts.append(generate_region_report(df, r))
    return '\n\n---\n\n'.join(parts)


if __name__ == '__main__':
    import file_loader
    d, msgs = file_loader.load_all()
    for m in msgs:
        print(m)
    if not d.empty:
        d2 = add_metrics(d)
        print(write_result_excel(d2))
        print(generate_report(d2)[:500])
