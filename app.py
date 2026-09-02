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
import json

import pandas as pd
import streamlit as st
import plotly.graph_objects as go

from config import (COL_REGION, COL_YEAR, OUTPUT_XLSX,
                    RATIO_COLS, GROWTH_COLS, STD_COLS, NUMERIC_COLS,
                    REPORT_DEFAULT_DOCX, REPORT_FORM_XLSX, REPORT_SETTINGS_JSON)
import file_loader
import analysis
import report_generator
import report_docx

st.set_page_config(page_title='企业生产经营分析', layout='wide')
st.title('企业生产经营分析')
st.caption('本地离线运行（支持 .xlsx/.xls）· 按地区整体分析 · 数据不离开本机')


@st.cache_data(show_spinner='正在加载数据并计算指标…')
def load_data():
    """加载 data/ 全部数据 → 地区×年度聚合 → 计算指标（不再自动生成静态 Excel，改由「报告导出」页导出）。"""
    detail, msgs = file_loader.load_all()
    if detail.empty:
        return detail, msgs
    region_df = analysis.add_region_metrics(analysis.region_yearly(detail))
    return detail, region_df, msgs


# ============================ 数据加载 ============================
detail, region_df, msgs = load_data()
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

st.sidebar.caption('静态备查表「分析结果汇总.xlsx」请在「📤 报告导出」页导出。')


# 地区整体按年度汇总（多地区合并）：销售额、净利润、所有者权益 → 净利润率 / ROE（%）
def overall_yearly(v):
    g = v.groupby(COL_YEAR).agg(
        销售额=('销售额或营业收入', 'sum'),
        净利润=('净利润', 'sum'),
        所有者权益=('所有者权益合计', 'sum')).reset_index()
    g['净利润率'] = g['净利润'] / g['销售额'] * 100
    g['ROE'] = g['净利润'] / g['所有者权益'] * 100
    return g


tab1, tab2, tab3 = st.tabs(['📈 可视化分析', '📄 文本报告', '📤 报告导出'])

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

# ================= 标签页3：报告导出 =================
with tab3:
    st.subheader('📤 导出中心')
    if not region_df.empty:
        # 顶部：静态备查表导出（不在启动时自动生成）
        if st.button('📊 导出「分析结果汇总.xlsx」（静态备查表）'):
            try:
                _xp = report_generator.write_result_excel(detail, region_df)
                st.success('已生成：' + _xp)
                with open(_xp, 'rb') as f:
                    st.download_button('⬇ 下载 分析结果汇总.xlsx', f.read(),
                                       file_name=os.path.basename(_xp))
            except Exception as _e:
                st.error('生成失败：' + str(_e))

    st.markdown('---')
    st.markdown('**选择报告生成方式**')
    _mode = st.radio('生成方式', ['方式A：直接生成（自动）', '方式B：填写表单生成'],
                     horizontal=True, label_visibility='collapsed')

    if _mode.startswith('方式A'):
        st.caption('方式A：直接用 data/ 的全部数据自动生成 Word 报告（模块可开关、可上下移动排序、样式可选；标号自动按 一、二、三… 排序）。')
        # 方式A 页
        if 'report_mods' not in st.session_state:
            st.session_state.report_mods = report_docx.get_default_modules()

        def _rmove(_idx, _step):
            _ms = st.session_state.report_mods
            _j = _idx + _step
            if 0 <= _j < len(_ms):
                _ms[_idx], _ms[_j] = _ms[_j], _ms[_idx]
                for _k, _m in enumerate(_ms):
                    _m['order'] = _k + 1

        st.markdown('**① 模块设置**：用 ▲/▼ 上下移动（拖动排序），启用/关闭、自选样式；标号将按此顺序生成。')
        _hdr = st.columns(6)
        for _i, _t in enumerate(['标号', '模块', '移动', '启用', '样式', '预览']):
            _hdr[_i].markdown(f'**{_t}**')
        _mods = []
        for _i, _m in enumerate(st.session_state.report_mods):
            _c = st.columns(6)
            _c[0].markdown(_m['enabled'] and '一二三四五六七八九十'[_i] or '×')
            _c[1].markdown(_m['name'])
            _c[2].button('▲', key=f'rup_{_m["id"]}', on_click=_rmove, args=(_i, -1),
                         disabled=(_i == 0))
            _c[2].button('▼', key=f'rdn_{_m["id"]}', on_click=_rmove, args=(_i, 1),
                         disabled=(_i == len(st.session_state.report_mods) - 1))
            _en = _c[3].checkbox('启用', value=bool(_m['enabled']), key=f'ren_{_m["id"]}')
            _st = _c[4].selectbox('样式', report_docx.STYLES,
                                  index=(report_docx.STYLES.index(_m['style'])
                                         if _m['style'] in report_docx.STYLES else 0),
                                  key=f'rst_{_m["id"]}', label_visibility='collapsed')
            _c[5].markdown(('一二三四五六七八九十'[_i] + '、' if _m['enabled'] else '')
                           + (_m['name'] if _m['enabled'] else '（关闭）'))
            _mods.append({'id': _m['id'], 'name': _m['name'], 'enabled': _en,
                          'order': _i + 1, 'style': _st})

        st.markdown('**② 字段显示**：勾选 = 该字段不显示。')
        _hide = []
        for _c in ['资产总额', '销售额或营业收入', '净利润', '负债总额', '纳税总额',
                   '净资产收益率(ROE)', '主营业务收入占比', '总资产周转率',
                   '资产负债率', '销售净利率', '净利润率']:
            if st.checkbox(f'不显示：{_c}', value=_c in report_docx.get_default_hide(),
                           key=f'rhid_{_c}'):
                _hide.append(_c)
        for _c in ['销售额同比增长率', '净利润同比增长率',
                   '资产总额同比增长率', '主营业务收入同比增长率']:
            if st.checkbox(f'不显示：{_c}', value=False, key=f'rhid_{_c}'):
                _hide.append(_c)

        st.markdown('**③ 封面信息**')
        _a1, _a2, _a3 = st.columns(3)
        _title = _a1.text_input('报告标题', '年度企业生产经营分析报告')
        _org = _a2.text_input('编制单位', '')
        _date = _a3.text_input('报告日期', pd.Timestamp.now().strftime('%Y年%m月%d日'))
        _rd = st.text_input('地区范围说明', '全部地区')
        _yd = st.text_input('年度说明', f'{int(years_all[0])}~{int(years_all[-1])}')

        _info = {'标题': _title, '地区说明': _rd, '年度说明': _yd,
                 '编制单位': _org, '报告日期': _date}

        if st.button('📄 方式A：用 data/ 全部数据自动生成 Word 报告'):
            _path = report_docx.generate_from_data(
                detail, region_df, info=_info, modules=_mods, hide=_hide,
                out_path=REPORT_DEFAULT_DOCX)
            st.success('已生成：' + _path)
            with open(_path, 'rb') as f:
                st.download_button('⬇ 下载 Word 报告', f.read(),
                                   file_name=os.path.basename(_path))

        if st.button('💾 保存当前模块/字段设置为「报告设置.json」（CLI 与下次默认生效）'):
            with open(REPORT_SETTINGS_JSON, 'w', encoding='utf-8') as f:
                json.dump({'modules': _mods, 'hide': _hide}, f, ensure_ascii=False, indent=2)
            st.success('已保存：' + REPORT_SETTINGS_JSON)

    else:
        st.caption('方式B：先生成「报告填写表单.xlsx」（含 报告设置/字段显示/封面信息/数值 工作表），'
                   '按工作表填写后上传生成报告；顺序、样式、字段显示在表单「报告设置」「字段显示」中设置。')
        # 方式B 页
        if st.button('📄 生成「报告填写表单.xlsx」模板'):
            _fp = report_docx.generate_form()
            st.success('已生成：' + _fp)
            with open(_fp, 'rb') as f:
                st.download_button('⬇ 下载填写表单', f.read(),
                                   file_name=os.path.basename(_fp))

        st.markdown('**上传已填写的表单生成报告**')
        _up = st.file_uploader('上传「报告填写表单.xlsx」', type=['xlsx'], key='rup')
        if _up is not None:
            _tmp = os.path.join(os.path.dirname(os.path.abspath(__file__)), '_upload_form.xlsx')
            with open(_tmp, 'wb') as f:
                f.write(_up.getbuffer())
            try:
                _path2 = report_docx.generate_from_form(_tmp, out_path=REPORT_DEFAULT_DOCX)
                st.success('已生成：' + _path2)
                with open(_path2, 'rb') as f:
                    st.download_button('⬇ 下载 Word 报告', f.read(),
                                       file_name=os.path.basename(_path2))
            except Exception as e:
                st.error('生成失败：' + str(e))