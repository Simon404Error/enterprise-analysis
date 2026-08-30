# -*- coding: utf-8 -*-
"""
企业生产经营分析 —— 本地 Web 交互面板（Streamlit + Plotly）
============================================================
启动：streamlit run app.py （或双击 run.bat）
功能（仅按「整个地区」维度，不按单个企业）：
  标签页1「可视化分析」：地区 / 年度筛选；地区整体趋势折线图、地区间对比柱状图、
     销售与净利润趋势、汇总统计表、企业公示分布表。
  标签页2「文本报告」：按地区生成经营总结（可下载）。
支持 .xlsx 与 .xls；离线运行，不调用任何网络 API。
"""
import os

import pandas as pd
import streamlit as st
import plotly.graph_objects as go

from config import (COL_REGION, COL_YEAR, OUTPUT_XLSX,
                    RATIO_COLS, GROWTH_COLS, STD_COLS, NUMERIC_COLS)
import file_loader
import analysis
import report_generator

st.set_page_config(page_title='企业生产经营分析', layout='wide')
st.title('企业生产经营分析')
st.caption('本地离线运行（支持 .xlsx/.xls）· 按地区整体分析 · 数据不离开本机')


@st.cache_data(show_spinner='正在加载数据并计算指标…')
def load_data():
    """加载 data/ 全部数据 → 地区×年度聚合 → 计算指标 → 生成静态 Excel 备查。"""
    detail, msgs = file_loader.load_all()
    if detail.empty:
        return detail, msgs, ''
    region_df = analysis.add_region_metrics(analysis.region_yearly(detail))
    try:
        xlsx = report_generator.write_result_excel(detail, region_df)
    except Exception as e:
        xlsx = f'静态 Excel 生成失败：{e}'
    return detail, region_df, msgs, xlsx


# ============================ 数据加载 ============================
detail, region_df, msgs, xlsx = load_data()
if region_df.empty:
    st.warning('未找到可分析的数据。')
    for m in msgs:
        st.info(m)
    st.info('请将「年度报告记录 (地区名).xlsx / .xls」放入 data/（可按年份分子文件夹），再点左侧「重新加载数据」。')
    st.stop()

# ============================ 侧边栏：筛选 ============================
st.sidebar.header('筛选条件')
if st.sidebar.button('🔄 重新加载数据（data/ 更新后点击）'):
    st.cache_data.clear()
    st.rerun()

regions = sorted(region_df[COL_REGION].unique())
sel_regions = st.sidebar.multiselect('地区', regions, default=regions)
if not sel_regions:
    st.warning('请至少选择一个地区。')
    st.stop()
view = region_df[region_df[COL_REGION].isin(sel_regions)].copy()

# 年度范围
years_all = sorted(view[COL_YEAR].dropna().unique())
if not years_all:
    st.warning('当前筛选下无年度数据。')
    st.stop()
ymin, ymax = int(years_all[0]), int(years_all[-1])
if ymin < ymax:
    y0, y1 = st.sidebar.slider('年度范围', ymin, ymax, (ymin, ymax), step=1)
else:
    y0 = y1 = ymin
    st.sidebar.caption(f'数据仅有 {ymin} 一个年度')
view = view[(view[COL_YEAR] >= y0) & (view[COL_YEAR] <= y1)]

if isinstance(xlsx, str) and xlsx.startswith('D:'):
    st.sidebar.caption('已生成静态备查：' + os.path.basename(OUTPUT_XLSX))
else:
    st.sidebar.caption('静态 Excel：' + str(xlsx))


# 地区整体按年度汇总（多地区合并）：销售额、净利润、所有者权益 → 净利润率 / ROE（%）
def overall_yearly(v):
    g = v.groupby(COL_YEAR).agg(
        销售额=('销售额或营业收入', 'sum'),
        净利润=('净利润', 'sum'),
        所有者权益=('所有者权益合计', 'sum')).reset_index()
    g['净利润率'] = g['净利润'] / g['销售额'] * 100
    g['ROE'] = g['净利润'] / g['所有者权益'] * 100
    return g


tab1, tab2 = st.tabs(['📈 可视化分析', '📄 文本报告'])

# ================= 标签页1：可视化 =================
with tab1:
    if view.empty:
        st.info('当前筛选条件下无数据。')
    else:
        st.subheader('趋势图：地区整体 净利润率 / ROE 年度变化')
        oy = overall_yearly(view)
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=oy[COL_YEAR], y=oy['净利润率'].round(2),
                                 mode='lines+markers', name='净利润率(%)'))
        fig.add_trace(go.Scatter(x=oy[COL_YEAR], y=oy['ROE'].round(2),
                                 mode='lines+markers', name='ROE(%)'))
        fig.update_layout(title='所选地区整体各年度盈利指标', xaxis_title='年度',
                          yaxis_title='%', hovermode='x unified', height=400,
                          legend=dict(orientation='h', y=1.12))
        st.plotly_chart(fig)

        st.subheader('对比图：最新年度各地区 净利润率 与 ROE（%）')
        latest = int(view[COL_YEAR].max())
        bar = view[view[COL_YEAR] == latest].sort_values(COL_REGION)
        if bar.empty:
            st.info('最新年度无数据。')
        else:
            fig = go.Figure()
            fig.add_trace(go.Bar(x=bar[COL_REGION], y=bar['净利润率'].round(2), name='净利润率(%)'))
            fig.add_trace(go.Bar(x=bar[COL_REGION], y=(bar['净资产收益率(ROE)'] * 100).round(2), name='ROE(%)'))
            fig.update_layout(barmode='group', title=f'{latest} 年各地区盈利指标对比',
                              xaxis_title='地区', yaxis_title='%', hovermode='x unified', height=420)
            st.plotly_chart(fig)

        st.subheader('销售额与净利润总额趋势（万元）')
        oy2 = overall_yearly(view)
        fig = go.Figure()
        fig.add_trace(go.Bar(x=oy2[COL_YEAR], y=oy2['销售额'].round(2), name='销售额'))
        fig.add_trace(go.Bar(x=oy2[COL_YEAR], y=oy2['净利润'].round(2), name='净利润'))
        fig.update_layout(barmode='group', xaxis_title='年度',
                          yaxis_title='万元', hovermode='x unified', height=380)
        st.plotly_chart(fig)

        st.subheader('汇总统计表（地区 × 年度，比率为 %）')
        show_cols = [COL_REGION, COL_YEAR] + NUMERIC_COLS + RATIO_COLS + GROWTH_COLS + STD_COLS
        disp = analysis.metrics_for_display(view)[show_cols]
        st.dataframe(disp, height=360)

        st.subheader('企业公示分布（本地区内企业，12 项判定，按年区分）')
        _, pub_tab = analysis.publish_by_region_year(detail)
        pub_view = pub_tab[pub_tab[COL_REGION].isin(sel_regions)] if not pub_tab.empty else pub_tab
        if pub_view.empty:
            st.info('（无企业明细，无法统计公示分布。）')
        else:
            st.dataframe(pub_view, height=360)
            st.caption('判定口径（12 项）：是否股权转让、是否对外投资、资产总额、所有者权益合计、销售额或营业收入、'
                       '利润总额、营业总收入中主营业务收入、净利润、负债总额、纳税总额、资产认缴额、资产实缴额'
                       '——数值为 0 视为空位：无任何空位为「全部公示」；存在空位（未全空）为「部分公示」；'
                       '全部为空位为「全部不公示」.')

# ================= 标签页2：文本报告 =================
with tab2:
    if view.empty:
        st.info('当前筛选条件下无数据。')
    else:
        rep_region = st.selectbox('报告地区', ['（全部地区）'] + regions)
        report = report_generator.generate_report(
            detail, region_df, region=None if rep_region == '（全部地区）' else rep_region)
        st.markdown(report)
        fname = f'经营报告_{rep_region if rep_region != "（全部地区）" else "全部"}.md'
        st.download_button('⬇ 下载报告 (Markdown)', report, file_name=fname, mime='text/markdown')