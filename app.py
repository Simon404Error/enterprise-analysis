# -*- coding: utf-8 -*-
"""
企业生产经营分析 —— 本地 Web 交互面板（Streamlit + Plotly）
=============================================================
启动：streamlit run app.py （或双击 run.bat）
功能：
  标签页1「可视化分析」：地区/企业/年度筛选；趋势折线图、企业对比柱状图、汇总统计表。
  标签页2「文本报告」：按地区或企业动态生成的经营状况文字总结，可下载。
离线运行：不调用任何网络 API。
"""
import os

import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go

from config import (COL_REGION, COL_ENTERPRISE, COL_YEAR,
                    OUTPUT_XLSX, RATIO_COLS, GROWTH_COLS, STD_COLS, NUMERIC_COLS)
import file_loader
import analysis
import report_generator

st.set_page_config(page_title='企业生产经营分析', layout='wide')
st.title('企业生产经营分析')
st.caption('本地离线运行 · 数据不离开本机 · 图表可缩放与悬停查看数值')


@st.cache_data(show_spinner='正在加载数据并计算指标…')
def load_data():
    """加载 data/ 全部数据 → 合并 → 计算指标 → 生成静态 Excel 备查。"""
    df, msgs = file_loader.load_all()
    if df.empty:
        return df, msgs, ''
    df = analysis.add_metrics(df)
    try:
        xlsx = report_generator.write_result_excel(df)
    except Exception as e:
        xlsx = f'静态 Excel 生成失败：{e}'
    return df, msgs, xlsx


# ============================ 数据加载 ============================
df, msgs, xlsx = load_data()
if df.empty:
    st.warning('未找到可分析的数据。')
    for m in msgs:
        st.info(m)
    st.info('请将「年度报告记录（地区名）.xlsx」放入项目根目录的 data/ 文件夹后，点击左侧「重新加载数据」。')
    st.stop()

# ============================ 侧边栏：筛选器 ============================
st.sidebar.header('筛选条件')
if st.sidebar.button('🔄 重新加载数据（data/ 更新后点击）'):
    st.cache_data.clear()
    st.rerun()

regions = sorted(df[COL_REGION].unique())
sel_regions = st.sidebar.multiselect('地区', regions, default=regions)
if not sel_regions:
    st.warning('请至少选择一个地区。')
    st.stop()

filtered = df[df[COL_REGION].isin(sel_regions)]

# 年度范围
years_all = sorted(filtered[COL_YEAR].dropna().unique())
if not years_all:
    st.warning('当前筛选下无年度数据。')
    st.stop()
ymin, ymax = int(years_all[0]), int(years_all[-1])
if ymin < ymax:
    y0, y1 = st.sidebar.slider('年度范围', ymin, ymax, (ymin, ymax), step=1)
else:
    y0 = y1 = ymin
    st.sidebar.caption(f'数据仅有 {ymin} 一个年度')
filtered = filtered[(filtered[COL_YEAR] >= y0) & (filtered[COL_YEAR] <= y1)]

enterprises = sorted(filtered[COL_ENTERPRISE].unique())
sel_ent = st.sidebar.selectbox('企业名称', ['（全部企业）'] + enterprises)

if isinstance(xlsx, str) and xlsx.startswith('D:'):
    st.sidebar.caption('已生成静态备查：' + os.path.basename(OUTPUT_XLSX))
else:
    st.sidebar.caption('静态 Excel：' + str(xlsx))

# ============================ 数据视图 ============================
view = filtered  # 当前筛选后的明细（含全部指标）


def pct_col(df_, col):
    """取列并转百分数（%）。"""
    return (df_[col] * 100)


def region_yearly(reg_df):
    """地区整体按年汇总：销售额、净利润、所有者权益。"""
    g = reg_df.groupby(COL_YEAR).agg(
        销售额=('销售额或营业收入', 'sum'),
        净利润=('净利润', 'sum'),
        所有者权益=('所有者权益合计', 'sum')).reset_index()
    g['净利润率'] = g['净利润'] / g['销售额'] * 100
    g['ROE'] = g['净利润'] / g['所有者权益'] * 100
    return g


tab1, tab2 = st.tabs(['📈 可视化分析', '📄 文本报告'])

# ================= 标签页1：可视化分析 =================
with tab1:
    if view.empty:
        st.info('当前筛选条件下无数据。')
    else:
        st.subheader('趋势图：净利润率 / ROE 年度变化')
        if sel_ent == '（全部企业）':
            g = region_yearly(view)
            src = f'地区整体（{", ".join(sel_regions)}）'
        else:
            ent_view = view[view[COL_ENTERPRISE] == sel_ent].sort_values(COL_YEAR)
            g = ent_view[[COL_YEAR]].copy()
            g['净利润率'] = ent_view['净利润率'] * 100
            g['ROE'] = ent_view['净资产收益率(ROE)'] * 100
            src = f'企业：{sel_ent}'
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=g[COL_YEAR], y=g['净利润率'].round(2),
                                 mode='lines+markers', name='净利润率(%)'))
        fig.add_trace(go.Scatter(x=g[COL_YEAR], y=g['ROE'].round(2),
                                 mode='lines+markers', name='ROE(%)'))
        fig.update_layout(title=f'{src} 各年度盈利指标', xaxis_title='年度',
                          yaxis_title='%', hovermode='x unified', height=400,
                          legend=dict(orientation='h', y=1.12))
        st.plotly_chart(fig)

        st.subheader('对比图：最新年度各企业 ROE（前 20）')
        latest = int(view[COL_YEAR].max())
        bar_df = view[view[COL_YEAR] == latest][[COL_ENTERPRISE, '净资产收益率(ROE)']].dropna()
        bar_df = bar_df.nlargest(20, '净资产收益率(ROE)').sort_values('净资产收益率(ROE)')
        if bar_df.empty:
            st.info('最新年度无有效的 ROE 数据。')
        else:
            fig = go.Figure(go.Bar(x=(bar_df['净资产收益率(ROE)'] * 100).round(2),
                                   y=bar_df[COL_ENTERPRISE], orientation='h',
                                   text=(bar_df['净资产收益率(ROE)'] * 100).round(2),
                                   textposition='outside'))
            fig.update_layout(title=f'{latest} 年企业 ROE 对比（%）',
                              xaxis_title='ROE(%)', yaxis_title='企业', height=500)
            st.plotly_chart(fig)

        st.subheader('销售额与净利润总额趋势（万元）')
        g2 = region_yearly(view)
        fig = go.Figure()
        fig.add_trace(go.Bar(x=g2[COL_YEAR], y=g2['销售额'].round(2), name='销售额'))
        fig.add_trace(go.Bar(x=g2[COL_YEAR], y=g2['净利润'].round(2), name='净利润'))
        fig.update_layout(barmode='group', xaxis_title='年度',
                          yaxis_title='万元', hovermode='x unified', height=380)
        st.plotly_chart(fig)

        st.subheader('汇总统计表（当前筛选）')
        show_cols = [COL_REGION, COL_ENTERPRISE, COL_YEAR] + NUMERIC_COLS + \
                    RATIO_COLS + GROWTH_COLS + STD_COLS
        disp = analysis.metrics_for_display(view)[show_cols]
        st.dataframe(disp, height=360)

        with st.expander('企业汇总（每个企业一行）'):
            st.dataframe(analysis.metrics_for_display(analysis.enterprise_summary(view)), height=360)

# ================= 标签页2：文本报告 =================
with tab2:
    if view.empty:
        st.info('当前筛选条件下无数据。')
    else:
        c1, c2 = st.columns([1, 1])
        rep_region = c1.selectbox('报告地区', ['（全部地区）'] + regions)
        ents = sorted(view[COL_ENTERPRISE].unique()) if rep_region == '（全部地区）' else \
            sorted(view[view[COL_REGION] == rep_region][COL_ENTERPRISE].unique())
        rep_ent = c2.selectbox('报告企业', ['（不指定企业）'] + ents)

        report = report_generator.generate_report(
            view, region=None if rep_region == '（全部地区）' else rep_region,
            enterprise=None if rep_ent == '（不指定企业）' else rep_ent)
        st.markdown(report)
        fname = f'经营报告_{rep_ent if rep_ent != "（不指定企业）" else (rep_region if rep_region != "（全部地区）" else "全部")}.md'
        st.download_button('⬇ 下载报告 (Markdown)', report, file_name=fname, mime='text/markdown')
