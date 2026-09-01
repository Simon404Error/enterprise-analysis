# -*- coding: utf-8 -*-
"""
Word 报告生成模块（企业生产经营离线分析）
=============================================
用途：按《年度企业年报分析报告》的版式，自动生成 Word 报告。

两种方式：
  方式A：直接从工具计算数据生成报告（缺数据的章节自动跳过）。
  方式B：先生成「报告填写表单.xlsx」，用户填写数值后，据此生成报告。

图表：
  - 默认用 matplotlib 自动生成静态图（完全离线）嵌入 Word；
  - 支持自定义图片：方式A 读取「图表映射.txt」、方式B 读取表单「图片设置」工作表；
    自定义为「是」且给出图片路径时，直接引用该图片，否则自动生成。
  - 报告不含「建议」章节。

运行入口见「生成报告.py」。
"""
import os
import shutil
import tempfile

import numpy as np
import pandas as pd

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'SimSun']
plt.rcParams['axes.unicode_minus'] = False

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

from config import (COL_REGION, COL_YEAR,
                    REPORT_DEFAULT_DOCX, REPORT_FORM_XLSX,
                    REPORT_IMG_DIR, REPORT_IMG_MAP, OUTPUT_XLSX,
                    PUBLISH_ALL, PUBLISH_PART, PUBLISH_NONE, RATIO_COLS, GROWTH_COLS)
from analysis import (region_yearly, add_region_metrics, publish_by_region_year,
                      metrics_for_display)

# ---------------- 图表名称（自定义图片映射用） ----------------
CHART_M1 = '净利率ROE趋势'
CHART_M2 = '各地区盈利对比'
CHART_M3 = '销售额净利润趋势'
CHART_P1 = '公示分布按年'
CHART_NAMES = [CHART_M1, CHART_M2, CHART_M3, CHART_P1]

# 报告章节标题
SEC_OVERALL = '一、生产经营总体情况'
SEC_METRICS = '二、经营指标分析'
SEC_PUBLISH = '三、企业公示分布'

_PCT_FIELDS = ['销售净利率', '净利润率', '净资产收益率(ROE)', '资产负债率',
               '主营业务收入占比', '总资产周转率']


# ============================ 基础工具 ============================
def _load_analysis():
    """加载 data/ 数据 → (明细表, 地区×年度指标表, 按年公示分布表)。"""
    import file_loader
    detail, msgs = file_loader.load_all()
    if detail.empty:
        return detail, pd.DataFrame(), pd.DataFrame()
    region = add_region_metrics(region_yearly(detail))
    _, pub = publish_by_region_year(detail)
    return detail, region, pub


def _fmt(v, digits=2):
    """数值显示：NaN/None -> 'NA'，保留 digits 位小数。"""
    if v is None or pd.isna(v):
        return 'NA'
    return f'{v:,.{digits}f}'


def _overall_table(detail, region):
    """总体情况表：地区×年度，企业数量+关键数值。"""
    cnt = detail.groupby([COL_REGION, COL_YEAR]).size().rename('企业数量').reset_index()
    m = region.copy()
    rows = []
    for _, r in m.sort_values([COL_REGION, COL_YEAR]).iterrows():
        c = cnt[(cnt[COL_REGION] == r[COL_REGION]) & (cnt[COL_YEAR] == r[COL_YEAR])]
        rows.append({
            COL_REGION: r[COL_REGION], COL_YEAR: str(r[COL_YEAR]),
            '企业数量': int(c['企业数量'].iloc[0]) if not c.empty else 0,
            '资产总额': r['资产总额'], '销售额或营业收入': r['销售额或营业收入'],
            '净利润': r['净利润'], '负债总额': r['负债总额'], '纳税总额': r['纳税总额'],
        })
    return pd.DataFrame(rows)


def _metrics_table(region, growth_cols=None):
    """经营指标表：地区×年度，比率类（%）+增长率（%）。"""
    m = metrics_for_display(region)
    cols = [COL_REGION, COL_YEAR] + _PCT_FIELDS
    if growth_cols is None:
        growth_cols = [c for c in GROWTH_COLS if c in m.columns]
    cols += growth_cols
    return m[[c for c in cols if c in m.columns]].copy()


def _read_img_override():
    """读取「图表映射.txt」：每行 图表名称=图片路径。"""
    over = {}
    if os.path.isfile(REPORT_IMG_MAP):
        with open(REPORT_IMG_MAP, encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith('#'):
                    continue
                if '=' in line:
                    k, v = line.split('=', 1)
                    over[k.strip()] = v.strip()
    # 目录兜底：report_img/<图表名称>.png/.jpg
    if os.path.isdir(REPORT_IMG_DIR):
        for name in CHART_NAMES:
            if name in over:
                continue
            for ext in ('.png', '.jpg', '.jpeg'):
                p = os.path.join(REPORT_IMG_DIR, name + ext)
                if os.path.isfile(p):
                    over[name] = p
                    break
    return over


def _resolve_chart(chart_name, gen_func, img_dir, overrides):
    """图表：自定义优先，否则自动生成。返回图片路径或 None。"""
    p = overrides.get(chart_name)
    if p and os.path.isfile(p):
        return p
    return gen_func(img_dir) if gen_func else None


# ============================ 图表生成 ============================
def _plot_yearly_margins(region, img_dir):
    """净利率/ROE 年度趋势折线。"""
    if region.empty:
        return None
    g = region.groupby(COL_YEAR).agg(
        销售额=('销售额或营业收入', 'sum'), 净利润=('净利润', 'sum'),
        权益=('所有者权益合计', 'sum')).reset_index()
    if len(g) < 2 and g['销售额'].iloc[0] == 0:
        return None
    g['净利率'] = g['净利润'] / g['销售额'].replace(0, np.nan) * 100
    g['ROE'] = g['净利润'] / g['权益'].replace(0, np.nan) * 100
    fig, ax = plt.subplots(figsize=(7.2, 3.4))
    ax.plot(g[COL_YEAR], g['净利率'], marker='o', label='净利润率(%)')
    ax.plot(g[COL_YEAR], g['ROE'], marker='s', label='ROE(%)')
    ax.set_xlabel('年度'); ax.set_ylabel('%'); ax.set_title('地区整体 净利润率 / ROE 趋势')
    ax.legend(); ax.grid(True, alpha=0.3)
    fig.tight_layout()
    p = os.path.join(img_dir, CHART_M1 + '.png')
    fig.savefig(p, dpi=150); plt.close(fig)
    return p


def _plot_region_compare(region, img_dir):
    """最新年度各地区 净利率、ROE 对比柱状。"""
    if region.empty:
        return None
    latest = region[COL_YEAR].max()
    bar = region[region[COL_YEAR] == latest].sort_values(COL_REGION)
    if bar.empty:
        return None
    x = bar[COL_REGION].astype(str)
    m = metrics_for_display(bar)
    y1 = m['净利润率'] if '净利润率' in m else None
    y2 = m['净资产收益率(ROE)'] if '净资产收益率(ROE)' in m else None
    if y1 is None or y2 is None:
        return None
    fig, ax = plt.subplots(figsize=(7.2, 3.4))
    import numpy as _np
    idx = _np.arange(len(x))
    w = 0.36
    ax.bar(idx - w/2, y1, w, label='净利润率(%)')
    ax.bar(idx + w/2, y2, w, label='ROE(%)')
    ax.set_xticks(idx); ax.set_xticklabels(x, rotation=30, ha='right')
    ax.set_ylabel('%'); ax.set_title(f'{latest} 年各地区盈利指标对比')
    ax.legend(); ax.grid(True, axis='y', alpha=0.3)
    fig.tight_layout()
    p = os.path.join(img_dir, CHART_M2 + '.png')
    fig.savefig(p, dpi=150); plt.close(fig)
    return p


def _plot_sales_profit(region, img_dir):
    """销售额 / 净利润 年度趋势柱状。"""
    if region.empty:
        return None
    g = region.groupby(COL_YEAR).agg(销售额=('销售额或营业收入', 'sum'),
                                      净利润=('净利润', 'sum')).reset_index()
    fig, ax = plt.subplots(figsize=(7.2, 3.4))
    x = g[COL_YEAR].astype(str)
    idx = np.arange(len(x)); w = 0.36
    ax.bar(idx - w/2, g['销售额'], w, label='销售额')
    ax.bar(idx + w/2, g['净利润'], w, label='净利润')
    ax.set_xticks(idx); ax.set_xticklabels(x)
    ax.set_ylabel('万元'); ax.set_title('销售额与净利润趋势（万元）')
    ax.legend(); ax.grid(True, axis='y', alpha=0.3)
    fig.tight_layout()
    p = os.path.join(img_dir, CHART_M3 + '.png')
    fig.savefig(p, dpi=150); plt.close(fig)
    return p


def _plot_publish(pub, img_dir):
    """公示分布：按 地区×年度 堆叠柱状（限最近若干个组合）。"""
    if pub.empty:
        return None
    d = pub.copy()
    d['label'] = d[COL_REGION].astype(str) + '·' + d[COL_YEAR].astype(str)
    # 组合过多时只取最近年度
    combo = d[COL_YEAR].nunique() * d[COL_REGION].nunique()
    if combo > 15:
        latest_years = sorted(d[COL_YEAR].unique())[-2:]
        d = d[d[COL_YEAR].isin(latest_years)]
    d = d.sort_values([COL_REGION, COL_YEAR])
    fig, ax = plt.subplots(figsize=(7.2, 3.6))
    labels = d['label'].tolist()
    idx = np.arange(len(labels))
    ax.bar(idx, d[PUBLISH_ALL], color='#4C9F70', label=PUBLISH_ALL)
    ax.bar(idx, d[PUBLISH_PART], bottom=d[PUBLISH_ALL], color='#E8A33D', label=PUBLISH_PART)
    ax.bar(idx, d[PUBLISH_NONE],
           bottom=d[PUBLISH_ALL] + d[PUBLISH_PART], color='#C0504D', label=PUBLISH_NONE)
    ax.set_xticks(idx); ax.set_xticklabels(labels, rotation=45, ha='right', fontsize=8)
    ax.set_ylabel('企业数'); ax.set_title('企业公示分布（按地区·年度）')
    ax.legend(); ax.grid(True, axis='y', alpha=0.3)
    fig.tight_layout()
    p = os.path.join(img_dir, CHART_P1 + '.png')
    fig.savefig(p, dpi=150); plt.close(fig)
    return p


def _gen_charts(region, pub, img_dir):
    """生成全部图表，返回 {图表名: 路径}。"""
    return {
        CHART_M1: _plot_yearly_margins(region, img_dir),
        CHART_M2: _plot_region_compare(region, img_dir),
        CHART_M3: _plot_sales_profit(region, img_dir),
        CHART_P1: _plot_publish(pub, img_dir),
    }


# ============================ 叙述文字 ============================
def _overall_narrative(detail, region):
    if region.empty:
        return ''
    years = sorted(region[COL_YEAR].unique())
    n_years = len(years)
    n_regions = region[COL_REGION].nunique()
    n_firms = len(detail)
    texts = [
        f'本报告基于「年度报告记录」数据，共覆盖 {n_regions} 个地区、'
        f'{n_years} 个年度（{years[0]}~{years[-1]}），合计 {n_firms} 条企业年度记录。'
    ]
    # 全年合计
    tot = region[['销售额或营业收入', '净利润', '资产总额', '纳税总额']].sum()
    texts.append(f'全区全年合计：资产总额 {_fmt(tot["资产总额"])} 万元，'
                 f'销售额（营业收入）{_fmt(tot["销售额或营业收入"])} 万元，'
                 f'净利润 {_fmt(tot["净利润"])} 万元，纳税总额 {_fmt(tot["纳税总额"])} 万元。')
    if region['销售额或营业收入'].sum() != 0:
        texts.append(f'全区整体销售净利率为 '
                     f'{_fmt(region["净利润"].sum() / region["销售额或营业收入"].sum() * 100)}%。')
    return ''.join(texts)


def _metrics_narrative(region):
    if region.empty:
        return ''
    latest = region[COL_YEAR].max()
    rows = region[region[COL_YEAR] == latest]
    m = metrics_for_display(rows)
    lines = [f'以下为 {latest} 年度各地区经营指标（比率为 %，NA 表示分母为 0/缺失无法计算）。']
    if '净利润率' in m.columns and m['净利润率'].notna().any():
        best_i = m['净利润率'].idxmax()
        worst_i = m['净利润率'].idxmin()
        lines.append(f'净利润率最高的地区为 {m.loc[best_i, COL_REGION]}'
                     f'（{_fmt(m.loc[best_i, "净利润率"])}%），最低为 {m.loc[worst_i, COL_REGION]}'
                     f'（{_fmt(m.loc[worst_i, "净利润率"])}%）。')
    if '总资产周转率' in m.columns and m['总资产周转率'].notna().any():
        lines.append(f'总资产周转率介于 {_fmt(m["总资产周转率"].min())}%'
                     f'~{_fmt(m["总资产周转率"].max())}%。')
    return ''.join(lines)


def _publish_narrative(pub):
    if pub.empty:
        return ''
    lines = ['公示状况按企业年度记录的 12 项判定（数值为 0 视为空位）。']
    for _, r in pub.sort_values([COL_REGION, COL_YEAR]).iterrows():
        e = int(r.get(PUBLISH_ALL, 0)) + int(r.get(PUBLISH_PART, 0)) + int(r.get(PUBLISH_NONE, 0))
        if e == 0:
            continue
        lines.append(f'{r[COL_REGION]} {r[COL_YEAR]}年共 {e} 家企业：全部公示 '
                     f'{int(r.get(PUBLISH_ALL, 0))} 家、部分公示 {int(r.get(PUBLISH_PART, 0))} 家、'
                     f'全部不公示 {int(r.get(PUBLISH_NONE, 0))} 家。')
    return ''.join(lines)


# ============================ docx 组装 ============================
def _set_cn(run, name='宋体', size=None, bold=None, color=None):
    run.font.name = name
    r = run._element.get_or_add_rPr()
    rF = r.get_or_add_rFonts()
    rF.set(qn('w:eastAsia'), name)
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)


def _para(doc, text, size=12, bold=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
          indent=True, name='宋体', space_after=6):
    p = doc.add_paragraph()
    p.alignment = align
    if indent:
        p.paragraph_format.first_line_indent = Pt(size * 2)
    p.paragraph_format.space_after = Pt(space_after)
    _set_cn(p.add_run(text), name=name, size=size, bold=bold)
    return p


def _heading(doc, text, level=1):
    if level == 1:
        _para(doc, text, size=16, bold=True, indent=False,
              align=WD_ALIGN_PARAGRAPH.LEFT, space_after=10)
    else:
        _para(doc, text, size=13, bold=True, indent=False,
              align=WD_ALIGN_PARAGRAPH.LEFT, space_after=6)


def _add_table(doc, title, headers, rows):
    if title:
        _para(doc, title, size=11, bold=True, indent=False,
              align=WD_ALIGN_PARAGRAPH.LEFT, space_after=2)
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j, h in enumerate(headers):
        c = t.rows[0].cells[j]
        _set_cn(c.paragraphs[0].add_run(str(h)), size=10.5, bold=True)
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            c = t.rows[i + 1].cells[j]
            _set_cn(c.paragraphs[0].add_run(_fmt(v) if isinstance(v, float) else str(v)),
                    size=10.5)
    doc.add_paragraph()


def _add_chart(doc, path):
    if path and os.path.isfile(path):
        doc.add_picture(path, width=Cm(15.5))
        doc.add_paragraph()


def _cover(doc, info):
    for _ in range(6):
        doc.add_paragraph()
    _para(doc, info.get('标题', '年度企业生产经营分析报告'), size=26, bold=True,
          indent=False, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=20)
    sub = info.get('副标题') or ('（' + info.get('地区说明', '全部地区') + ' · '
                                 + info.get('年度说明', '') + '）')
    _para(doc, sub, size=16, bold=True, indent=False,
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)
    for _ in range(3):
        doc.add_paragraph()
    org = info.get('编制单位') or info.get('机构', '')
    date = info.get('报告日期') or info.get('日期', '')
    _para(doc, '编制单位：' + org, size=14, indent=False,
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
    _para(doc, '报告日期：' + date, size=14, indent=False,
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
    doc.add_page_break()


def _build_doc(out_path, info, overall, metrics, growth, pub, charts,
               overall_text='', metrics_text='', publish_text=''):
    doc = Document()
    _cover(doc, info)

    if not overall.empty:
        _heading(doc, SEC_OVERALL, 1)
        if overall_text:
            _para(doc, overall_text)
        _add_table(doc, '表 1  生产经营总体情况（万元）',
                   ['地区', '年度', '企业数量', '资产总额', '销售额或营业收入',
                    '净利润', '负债总额', '纳税总额'],
                   overall.values.tolist())
    if not metrics.empty:
        _heading(doc, SEC_METRICS, 1)
        if metrics_text:
            _para(doc, metrics_text)
        _add_table(doc, '表 2  经营指标分析（%）',
                   ['地区', '年度'] + [c for c in _PCT_FIELDS if c in metrics.columns],
                   metrics[[c for c in [COL_REGION, COL_YEAR] + _PCT_FIELDS
                            if c in metrics.columns]].values.tolist())
        if growth is not None and not growth.empty:
            _add_table(doc, '表 3  增长率（同比，%）',
                       [COL_REGION, COL_YEAR] + list(growth.columns[2:]),
                       growth.values.tolist())
        _add_chart(doc, charts.get(CHART_M1))
        _add_chart(doc, charts.get(CHART_M2))
        _add_chart(doc, charts.get(CHART_M3))
    if not pub.empty:
        _heading(doc, SEC_PUBLISH, 1)
        if publish_text:
            _para(doc, publish_text)
        _add_table(doc, '表 4  企业公示分布（按年）',
                   [COL_REGION, COL_YEAR, PUBLISH_ALL, PUBLISH_PART, PUBLISH_NONE],
                   pub[[COL_REGION, COL_YEAR, PUBLISH_ALL, PUBLISH_PART,
                        PUBLISH_NONE]].values.tolist())
        _add_chart(doc, charts.get(CHART_P1))

    doc.save(out_path)
    return out_path


# ============================ 方式A：自动生成 ============================
_DEFAULT_INFO = '机构默认留空，由表单/调用方提供'


def _table_intro(df, name):
    """表单模式：数据表的简短引导文字。"""
    if df is None or df.empty:
        return ''
    return f'以下为「{name}」数据（金额单位：万元，比率为 %；NA 表示无法计算）。'


def _default_info(detail, region):
    years = sorted(region[COL_YEAR].unique()) if not region.empty else []
    regions = sorted(region[COL_REGION].unique())
    return {
        '标题': '年度企业生产经营分析报告',
        '副标题': '',
        '地区说明': '全部地区' if len(regions) > 1 else (regions[0] if regions else ''),
        '年度说明': f'{years[0]}~{years[-1]}' if len(years) > 1 else
                    (str(years[0]) if years else ''),
        '机构': '',
        '日期': '',
    }


def generate_from_data(detail_df=None, region_df=None, out_path=REPORT_DEFAULT_DOCX, info=None):
    """方式A：从工具计算数据生成 Word 报告。缺数据章节自动跳过。"""
    if detail_df is None:
        detail_df, region_df, _ = _load_analysis()
    if region_df is None and not detail_df.empty:
        region_df = add_region_metrics(region_yearly(detail_df))
    _, pub = publish_by_region_year(detail_df) if not detail_df.empty else (None, pd.DataFrame())

    if info is None:
        info = _default_info(detail_df, region_df)
        info['日期'] = pd.Timestamp.now().strftime('%Y年%m月%d日')

    overrides = _read_img_override()
    img_dir = tempfile.mkdtemp(prefix='rep_')
    try:
        charts = {k: _resolve_chart(k, f, img_dir, overrides) for k, f in [
            (CHART_M1, lambda d: _plot_yearly_margins(region_df, d)),
            (CHART_M2, lambda d: _plot_region_compare(region_df, d)),
            (CHART_M3, lambda d: _plot_sales_profit(region_df, d)),
            (CHART_P1, lambda d: _plot_publish(pub, d)),
        ]}
        overall = _overall_table(detail_df, region_df) if not region_df.empty else pd.DataFrame()
        metrics = _metrics_table(region_df) if not region_df.empty else pd.DataFrame()
        growth_cols = [c for c in GROWTH_COLS if c in metrics.columns]
        growth = metrics[[c for c in [COL_REGION, COL_YEAR] + growth_cols if c in metrics.columns]].copy() \
            if growth_cols else None
        return _build_doc(out_path, info, overall, metrics, growth, pub, charts,
                          overall_text=_overall_narrative(detail_df, region_df),
                          metrics_text=_metrics_narrative(region_df),
                          publish_text=_publish_narrative(pub))
    finally:
        shutil.rmtree(img_dir, ignore_errors=True)


# ============================ 方式B：表单 ============================
_FORM_SHEETS = ['说明', '封面信息', '图片设置', '总体指标', '经营指标', '公示分布']

_COVER_FIELDS = ['标题', '地区说明', '年度说明', '编制单位', '报告日期']


def generate_form(path=REPORT_FORM_XLSX):
    """生成「报告填写表单.xlsx」。"""
    import openpyxl
    wb = openpyxl.Workbook()
    ws0 = wb.active
    ws0.title = '说明'
    ws0.append([])
    for line in [
        '【报告填写表单】使用说明',
        '',
        '1. 「封面信息」：填写报告标题、地区说明、年度说明、编制单位、报告日期。',
        '2. 「图片设置」：默认自动生成图表；如需自定义，将「是否自定义」填为 是，',
        '   并在「图片路径」填入图片文件路径（图表名称见下）。',
        '3. 「总体指标」「经营指标」「公示分布」三个工作表：按表头填写数值；',
        '   留空的行会在报告中跳过。数值单位：金额-万元，比率填写百分数数值（如 15.6 表示 15.6%）。',
        '4. 填完后运行：python 生成报告.py 报告填写表单.xlsx',
        '',
        '图表名称：' + '、'.join(CHART_NAMES),
    ]:
        ws0.append([line])

    ws = wb.create_sheet('封面信息')
    ws.append(['项目', '填写内容'])
    for f in _COVER_FIELDS:
        ws.append([f, ''])

    ws = wb.create_sheet('图片设置')
    ws.append(['图表名称', '是否自定义', '图片路径'])
    for name in CHART_NAMES:
        ws.append([name, '', ''])

    ws = wb.create_sheet('总体指标')
    ws.append(['地区', '年度', '企业数量', '资产总额', '销售额或营业收入',
               '净利润', '负债总额', '纳税总额'])
    ws = wb.create_sheet('经营指标')
    ws.append(['地区', '年度'] + _PCT_FIELDS)
    ws = wb.create_sheet('公示分布')
    ws.append(['地区', '年度', PUBLISH_ALL, PUBLISH_PART, PUBLISH_NONE])

    wb.save(path)
    return path


def _read_form(form_path):
    """读取已填表单 → (info, overall_df, metrics_df, pub_df, 图片覆盖)。"""
    xl = pd.ExcelFile(form_path)
    info = {}
    if '封面信息' in xl.sheet_names:
        df = pd.read_excel(form_path, sheet_name='封面信息')
        for _, r in df.iterrows():
            k, v = r.iloc[0], r.iloc[1]
            if pd.notna(k) and pd.notna(v):
                info[str(k)] = str(v)
    overrides = {}
    if '图片设置' in xl.sheet_names:
        df = pd.read_excel(form_path, sheet_name='图片设置')
        for _, r in df.iterrows():
            name, custom, path = r.iloc[0], str(r.iloc[1]).strip(), str(r.iloc[2]).strip()
            if pd.notna(name) and custom.lower() in ('是', 'yes', '自定义', '1'):
                if path and path.lower() != 'nan':
                    overrides[str(name)] = path
    def _get(sheet):
        if sheet in xl.sheet_names:
            df = pd.read_excel(form_path, sheet_name=sheet)
            return df.dropna(how='all')
        return pd.DataFrame()
    return info, _get('总体指标'), _get('经营指标'), _get('公示分布'), overrides


def generate_from_form(form_path, out_path=None, info=None, overall=None,
                       metrics_df=None, pub_df=None, overrides=None):
    """方式B：读取填写后的表单生成 Word 报告。"""
    if overall is None:
        info, overall, metrics_df, pub_df, overrides = _read_form(form_path)
    out_path = out_path or REPORT_DEFAULT_DOCX
    if not info.get('报告日期'):
        info['报告日期'] = pd.Timestamp.now().strftime('%Y年%m月%d日')
    if not info.get('标题'):
        info['标题'] = '年度企业生产经营分析报告'
    img_dir = tempfile.mkdtemp(prefix='rep_')
    try:
        # 表单模式：仅使用「图片设置」里自定义为「是」且给出路径的图片
        charts = {k: None for k in CHART_NAMES}
        for k in CHART_NAMES:
            p = overrides.get(k)
            if p and os.path.isfile(p):
                charts[k] = p
        return _build_doc(out_path, info, overall, metrics_df, None, pub_df, charts,
                          overall_text=_table_intro(overall, '生产经营总体情况'),
                          metrics_text=_table_intro(metrics_df, '经营指标'),
                          publish_text=_publish_narrative(pub_df))
    finally:
        shutil.rmtree(img_dir, ignore_errors=True)


if __name__ == '__main__':
    print('模块自检：')
    d, r, p = _load_analysis()
    print('数据行数:', len(d), '地区×年度:', len(r), '公示分布行数:', len(p))
    generate_from_data(d, r)
    print('已生成:', REPORT_DEFAULT_DOCX)