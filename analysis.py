# -*- coding: utf-8 -*-
"""
指标计算模块（地区维度）：把企业明细聚合到「地区 × 年度」，在地区整体上计算全部指标。

指标口径（均为地区整体，不再按单个企业）：
  比率类（销售净利率、净利润率、ROE、资产负债率、主营业务收入占比、总资产周转率）：
      - 主营业务收入占比 = 主营业务收入 / 营业收入 ×100%
      - 总资产周转率 = 营业收入 / 平均资产总额 ×100%（平均资产 = (上年末+当年末)/2）
  增长率（销售额 / 净利润 / 资产总额 / 主营业务收入 同比与环比，年度数据下同比=环比）：
      按「地区」分组、年度升序，仅与上一个实际出现且相邻的年度比较；
      上年缺失（断档）或上年为 0 时记 NaN（展示为 NA），避免除零。
  标准差（销售额、净利润）：地区各年度数值的样本标准差（万元），衡量地区经营波动。
  公示状况：按明细行（=一家企业的年度记录）的 12 项判定该企业公示状态
      （全部公示 / 部分公示 / 全部不公示），再按「地区」汇总各类别企业数。
"""
import numpy as np
import pandas as pd

from config import (COL_YEAR, COL_REGION, COL_FILE,
                    NUMERIC_COLS, TEXT_COLS, PUBLISH_COLS, PUBLISH_COL,
                    PUBLISH_ALL, PUBLISH_PART, PUBLISH_NONE,
                    RATIO_COLS, GROWTH_COLS, STD_COLS)

# 地区维度数值列（用于聚合加总）
AGG_NUMERIC = NUMERIC_COLS


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


def region_yearly(df):
    """企业明细 → 地区×年度聚合表（数值列加总，文本/来源列保留首个）。"""
    if df.empty:
        return pd.DataFrame()
    g = df.groupby([COL_REGION, COL_YEAR], as_index=False)
    agg = {c: (c, lambda s: s.sum(min_count=1)) for c in NUMERIC_COLS}
    for c in TEXT_COLS + [COL_FILE]:
        agg[c] = (c, 'first')
    out = g.agg(**agg)
    return out.sort_values([COL_REGION, COL_YEAR]).reset_index(drop=True)


def _prev_row(df, col):
    """同地区上一行（上一个实际出现年度）的值。"""
    g = df.groupby(COL_REGION, as_index=False)
    return g[col].shift(1)


def strict_growth(df, col):
    """地区严格相邻年度的同比增长率：仅当上一个出现的年度 == 当前年度-1 时计算；
    上年缺失（断档）或上年为 0 时记 NaN。df 需已按 地区+年度 排序。"""
    g = df.groupby(COL_REGION, as_index=False)
    prev_v = g[col].shift(1)
    prev_y = g[COL_YEAR].shift(1)
    with np.errstate(divide='ignore', invalid='ignore'):
        r = (df[col] - prev_v) / prev_v
    r = r.replace([np.inf, -np.inf], np.nan)
    adjacent = (df[COL_YEAR] - prev_y) == 1
    return r.where(adjacent)


def add_region_metrics(region_df):
    """在地区×年度聚合表上计算全部派生指标（地区整体口径）。返回新 DataFrame。"""
    region_df = region_df.copy()
    sales = region_df['销售额或营业收入']
    region_df['销售净利率'] = safe_div(region_df['净利润'], sales)
    region_df['净利润率'] = safe_div(region_df['净利润'], sales)
    region_df['净资产收益率(ROE)'] = safe_div(region_df['净利润'], region_df['所有者权益合计'])
    region_df['资产负债率'] = safe_div(region_df['负债总额'], region_df['资产总额'])
    region_df['主营业务收入占比'] = safe_div(region_df['营业总收入中主营业务收入'], sales)

    # 平均资产（上年末+当年末）/2，无上年用当年末
    region_df = region_df.sort_values([COL_REGION, COL_YEAR]).reset_index(drop=True)
    prev_assets = _prev_row(region_df, '资产总额')
    prev_y = _prev_row(region_df, COL_YEAR)
    adjacent = (region_df[COL_YEAR] - prev_y) == 1
    avg_asset = ((region_df['资产总额'] + prev_assets) / 2).where(
        adjacent & prev_assets.notna(), region_df['资产总额'])
    region_df['总资产周转率'] = safe_div(sales, avg_asset)

    # 增长率（年度数据下 同比=环比）
    region_df['销售额同比增长率'] = strict_growth(region_df, '销售额或营业收入')
    region_df['销售额环比增长率'] = region_df['销售额同比增长率']
    region_df['净利润同比增长率'] = strict_growth(region_df, '净利润')
    region_df['净利润环比增长率'] = region_df['净利润同比增长率']
    region_df['资产总额同比增长率'] = strict_growth(region_df, '资产总额')
    region_df['资产总额环比增长率'] = region_df['资产总额同比增长率']
    region_df['主营业务收入同比增长率'] = strict_growth(region_df, '营业总收入中主营业务收入')
    region_df['主营业务收入环比增长率'] = region_df['主营业务收入同比增长率']

    # 地区年度波动（销售额/净利润 的年度标准差）
    for col, std_col in [('销售额或营业收入', '销售额标准差'), ('净利润', '净利润标准差')]:
        s = region_df.groupby(COL_REGION)[col].std().rename(std_col)
        region_df = region_df.merge(s, left_on=COL_REGION, right_index=True, how='left')

    return region_df


def classify_publish(row):
    """按一行（一家企业年度记录）的 12 项判定公示状况。

    口径（用户明确规定）：
      数值为 0 视为空位（未公示）。
      - 全部公示：12 项均非空位（无空白、无 0）。
      - 全部不公示：12 项全为空位（全部空白或 0）。
      - 部分公示：其余情况（存在空位，但未全空）。
    """
    missing = 0
    for c in PUBLISH_COLS:
        v = row.get(c)
        if c in TEXT_COLS:
            if v is None or pd.isna(v) or (isinstance(v, str) and not v.strip()):
                missing += 1
        else:  # 数值列：空白或 0 均视为空位
            try:
                nv = float(v)
                if pd.isna(nv) or nv == 0:
                    missing += 1
            except (TypeError, ValueError):
                missing += 1
    if missing == len(PUBLISH_COLS):
        return PUBLISH_NONE
    if missing > 0:
        return PUBLISH_PART
    return PUBLISH_ALL


def publish_by_region(detail):
    """地区公示分布：对企业明细每行判公示，再按地区统计各类别企业数。返回 (明细带公示列, 地区类别计数表)。"""
    d = detail.copy()
    d[PUBLISH_COL] = d.apply(classify_publish, axis=1)
    # 地区 × 公示类别 计数（所有年度的企业记录数）
    tab = d.groupby([COL_REGION, PUBLISH_COL]).size().unstack(fill_value=0)
    for cls in (PUBLISH_ALL, PUBLISH_PART, PUBLISH_NONE):
        if cls not in tab.columns:
            tab[cls] = 0
    tab = tab[tab.columns.tolist()]  # 保持列序
    tab.columns.name = None
    tab = tab.rename_axis(COL_REGION).reset_index()
    return d, tab


def publish_by_region_year(detail):
    """地区×年度公示分布：按「地区 × 年度」统计各类别企业数，不计跨年总和。
    返回 (明细带公示列, 表：地区/年度/全部公示/部分公示/全部不公示)。"""
    d = detail.copy()
    d[PUBLISH_COL] = d.apply(classify_publish, axis=1)
    tab = d.groupby([COL_REGION, COL_YEAR, PUBLISH_COL]).size().unstack(fill_value=0)
    for cls in (PUBLISH_ALL, PUBLISH_PART, PUBLISH_NONE):
        if cls not in tab.columns:
            tab[cls] = 0
    tab = tab[tab.columns.tolist()]
    tab.columns.name = None
    tab = tab.rename_axis([COL_REGION, COL_YEAR]).reset_index()
    tab = tab.sort_values([COL_REGION, COL_YEAR]).reset_index(drop=True)
    return d, tab


def region_summary(region_df):
    """地区汇总表：每个地区一行（最新年度指标 + 均值 + 年度波动）。"""
    if region_df.empty:
        return pd.DataFrame()
    rows = []
    for region, g in region_df.groupby(COL_REGION):
        g = g.sort_values(COL_YEAR)
        last = g.iloc[-1]
        rows.append({
            COL_REGION: region,
            '覆盖年度数': len(g),
            '起始年度': g[COL_YEAR].min(),
            '结束年度': g[COL_YEAR].max(),
            '销售额总计(万元)': g['销售额或营业收入'].sum(),
            '净利润总计(万元)': g['净利润'].sum(),
            '销售额标准差(万元)': g['销售额标准差'].iloc[0],
            '净利润标准差(万元)': g['净利润标准差'].iloc[0],
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
    d, msgs = file_loader.load_all()
    for m in msgs:
        print(m)
    if d.empty:
        print('（无数据）')
    else:
        region = region_yearly(d)
        r = add_region_metrics(region)
        print(r.head(10).to_string())
        pub_d, pub_tab = publish_by_region(d)
        print('公示分布:')
        print(pub_tab.to_string() if not pub_tab.empty else '（空）')
        print(region_summary(r).to_string() if not region_summary(r).empty else '（空）')