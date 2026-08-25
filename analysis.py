# -*- coding: utf-8 -*-
"""
指标计算模块：在合并后的明细表上计算全部派生指标。

指标口径：
  比率类（净利率、ROE、资产负债率、周转率、主营业务占比）：
      按 (地区, 企业名称, 年度) 行内计算，输出为小数，展示时 ×100 加 %。
  增长率（销售额/净利润 同比、环比）：
      按 (地区, 企业名称) 分组、按年度升序，与“上一个实际出现的年度”相比；
      上年缺失或上年为 0 时记为 NaN（展示为 NA），避免除零。
  标准差（销售额标准差、净利润标准差）：
      企业各年度数值的样本标准差（单位：万元）；仅一个年度时为 NaN。
"""
import numpy as np
import pandas as pd

from config import (COL_ENTERPRISE, COL_YEAR, COL_REGION,
                    NUMERIC_COLS, RATIO_COLS, GROWTH_COLS, STD_COLS)


def safe_div(a, b):
    """安全除法：除数为 0 或缺失时返回 NaN，避免 inf。兼容 Series 与标量。"""
    with np.errstate(divide='ignore', invalid='ignore'):
        r = a / b
    if isinstance(r, (pd.Series, pd.DataFrame)):
        return r.replace([np.inf, -np.inf], np.nan)
    if isinstance(r, (float, np.floating)) or np.isscalar(r):
        if pd.isna(r) or np.isinf(r):
            return np.nan
    return r


def strict_growth(df, col):
    """严格相邻年度的同比增长率：仅当上一个出现的年度 == 当前年度 - 1 时计算；
    上年缺失（年度断档）或上年为 0 时返回 NaN。df 需已按 地区+企业+年度 排序。"""
    g = df.groupby([COL_REGION, COL_ENTERPRISE])
    prev_v = g[col].shift(1)      # 上一行（同企业）的数值
    prev_y = g[COL_YEAR].shift(1)  # 上一行（同企业）的年度
    with np.errstate(divide='ignore', invalid='ignore'):
        r = (df[col] - prev_v) / prev_v
    r = r.replace([np.inf, -np.inf], np.nan)
    adjacent = (df[COL_YEAR] - prev_y) == 1   # 仅相邻年度算增长率
    return r.where(adjacent)


def add_metrics(df):
    """计算全部派生指标，返回新增列后的 DataFrame。"""
    df = df.copy()

    # ---- 1. 比率类（行内） ----
    sales = df['销售额或营业收入']
    df['销售净利率'] = safe_div(df['净利润'], sales)
    df['净利润率'] = safe_div(df['净利润'], sales)
    df['净资产收益率(ROE)'] = safe_div(df['净利润'], df['所有者权益合计'])
    df['资产负债率'] = safe_div(df['负债总额'], df['资产总额'])
    df['总资产周转率'] = safe_div(sales, df['资产总额'])
    df['主营业务占比'] = safe_div(df['营业总收入中主营业务收入'], sales)

    # ---- 2. 增长率（按 地区+企业 分组、严格相邻年度计算） ----
    df = df.sort_values([COL_REGION, COL_ENTERPRISE, COL_YEAR]).reset_index(drop=True)
    df['销售额同比增长率'] = strict_growth(df, '销售额或营业收入')
    df['销售额环比增长率'] = df['销售额同比增长率']      # 年度数据下 同比=环比
    df['净利润同比增长率'] = strict_growth(df, '净利润')
    df['净利润环比增长率'] = df['净利润同比增长率']

    # ---- 3. 企业级标准差（各年度数值的波动） ----
    for col, std_col in [('销售额或营业收入', '销售额标准差'), ('净利润', '净利润标准差')]:
        s = df.groupby([COL_REGION, COL_ENTERPRISE])[col].std().rename(std_col)
        df = df.merge(s, left_on=[COL_REGION, COL_ENTERPRISE], right_index=True, how='left')

    return df


def enterprise_summary(df):
    """企业汇总表：每个企业一行，含年度数、均值、标准差、最新年度指标。"""
    if df.empty:
        return pd.DataFrame()
    rows = []
    for (region, ent), g in df.groupby([COL_REGION, COL_ENTERPRISE]):
        g = g.sort_values(COL_YEAR)
        last = g.iloc[-1]
        rows.append({
            COL_REGION: region,
            COL_ENTERPRISE: ent,
            '覆盖年度数': len(g),
            '起始年度': g[COL_YEAR].min(),
            '结束年度': g[COL_YEAR].max(),
            '销售额均值(万元)': g['销售额或营业收入'].mean(),
            '净利润均值(万元)': g['净利润'].mean(),
            '销售额标准差(万元)': g['销售额或营业收入'].std(),
            '净利润标准差(万元)': g['净利润'].std(),
            '最新年度销售净利率': last['销售净利率'],
            '最新年度ROE': last['净资产收益率(ROE)'],
            '最新年度销售额同比': last['销售额同比增长率'],
            '最新年度净利润同比': last['净利润同比增长率'],
        })
    return pd.DataFrame(rows)


def metrics_for_display(df):
    """供面板/Excel 展示的副本：比率与增长率转百分数并保留两位。"""
    out = df.copy()
    for c in RATIO_COLS + GROWTH_COLS:
        if c in out.columns:
            out[c] = (out[c] * 100).round(2)
    return out


if __name__ == '__main__':
    import file_loader
    d, _ = file_loader.load_all()
    if d.empty:
        print('（无数据）')
    else:
        d2 = add_metrics(d)
        print(d2.head(10).to_string())
        print(enterprise_summary(d2).head())
