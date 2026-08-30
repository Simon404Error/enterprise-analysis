# -*- coding: utf-8 -*-
"""
文件读取模块：递归扫描 data/（含年份子文件夹）下的 Excel（支持 .xlsx 与 .xls），
解析地区名与分卷号，合并同一地区分卷，输出“明细表”。

明细表每行 = 一个企业 × 一个年度（同企业同年度已按数值字段加总）。
年度来源优先级：表头「年度」列 > data/ 下的年份子文件夹名。
本项目分析口径为“地区整体”，因此明细在企业维度之上再提供地区合计用聚合。
"""
import os
import re
import glob

import numpy as np
import pandas as pd

from config import (DATA_DIR, FILE_PATTERN, YEAR_FOLDER_PATTERN,
                    COL_ENTERPRISE, COL_YEAR, COL_REGION, COL_FILE, COL_SRC_YEAR,
                    NUMERIC_COLS, TEXT_COLS)


# 表头列名 → 常见别名（真实表头可能略有差异，用于归一化）
_ALIASES = {
    COL_ENTERPRISE: ['企业名称', '公司名称', '单位名称'],
    COL_YEAR:       ['年度', '年份', '统计年度'],
    '资产总额':       ['资产总额'],
    '所有者权益合计':  ['所有者权益合计', '所有者权益'],
    '销售额或营业收入': ['销售额或营业收入', '销售额', '营业收入', '销售额/营业收入'],
    '利润总额':       ['利润总额'],
    '营业总收入中主营业务收入': ['营业总收入中主营业务收入', '主营业务收入'],
    '净利润':         ['净利润'],
    '负债总额':       ['负债总额'],
    '纳税总额':       ['纳税总额', '纳税额'],
    '资产认缴额':      ['资产认缴额', '认缴额'],
    '资产实缴额':      ['资产实缴额', '实缴额'],
    '是否股权转让':    ['是否股权转让', '股权转让'],
    '是否对外投资':    ['是否对外投资', '对外投资'],
}


def find_data_files(data_dir=None):
    """递归扫描目录，返回 [(路径, 地区名, 分卷号), ...]，按命名规则过滤。支持 .xlsx/.xls。"""
    data_dir = data_dir or DATA_DIR
    if not os.path.isdir(data_dir):
        return []
    files = []
    for path in sorted(glob.glob(os.path.join(data_dir, '**', '*.xlsx'), recursive=True)
                       + glob.glob(os.path.join(data_dir, '**', '*.xls'), recursive=True)):
        name = os.path.basename(path)
        m = FILE_PATTERN.match(name)
        if not m:
            continue
        region, part = (m.group(1) or '').strip(), (m.group(2) or '')
        files.append((path, region, part))
    return files


def folder_year(data_dir, path):
    """从文件相对 data/ 的父目录中提取年份（如 data/2025/x.xlsx -> 2025），找不到返回 NaN。"""
    try:
        rel = os.path.relpath(path, data_dir)
        for p in os.path.dirname(rel).split(os.sep):
            if p and YEAR_FOLDER_PATTERN.match(p):
                return int(p)
    except Exception:
        return np.nan
    return np.nan


def _known_field_names():
    """所有标准字段名 + 全部别名，用于识别表头行。"""
    names = set()
    for std, als in _ALIASES.items():
        names.add(std)
        names.update(a for a in als)
    return names


def _detect_header_row(raw):
    """在头部若干行中找出命中最多的表头行（跳过标题行/空行）。
    真实文件常见格式：第1行标题「年度报告记录」、第2行空、第3行才是表头。"""
    names = _known_field_names()
    best_hits, best_idx = 0, None
    for i in range(min(len(raw), 20)):
        vals = [str(v).strip() for v in raw.iloc[i].tolist()]
        hits = sum(1 for v in vals if v in names)
        if hits > best_hits:
            best_hits, best_idx = hits, i
    return best_idx if best_hits > 0 else None


def read_one_file(path):
    """读取单个 Excel（所有工作表），自动跳过前导标题/空行定位真实表头。
    .xlsx 用 openpyxl、.xls 用 xlrd（由 pandas 自动选择）。"""
    frames = []
    xls = pd.ExcelFile(path)
    for sheet in xls.sheet_names:
        raw = pd.read_excel(xls, sheet_name=sheet, header=None, dtype=object)
        raw = raw.dropna(how='all')
        if raw.empty:
            continue
        hdr = _detect_header_row(raw)
        if hdr is None:
            continue  # 未识别到表头，跳过该工作表
        df = raw.iloc[hdr + 1:].copy()
        df.columns = [str(c).strip() for c in raw.iloc[hdr].tolist()]
        df = df.dropna(how='all')
        if not df.empty:
            frames.append(df)
    if not frames:
        return None
    return pd.concat(frames, ignore_index=True)


def normalize_columns(df):
    """列名去空格并归一到标准字段名；同一标准字段仅保留第一次出现的列。"""
    df.columns = [str(c).strip() for c in df.columns]
    rename, used = {}, set()
    for c in df.columns:
        if not c:
            continue
        std = None
        for s, aliases in _ALIASES.items():
            if c == s or c in aliases:
                std = s
                break
        if std and std not in used:
            rename[c] = std
            used.add(std)
    return df.rename(columns=rename)


def load_all(data_dir=None):
    """加载 data/ 下全部符合规则的 Excel，返回 (明细表, 提示信息列表)。

    明细表每行 = 企业 × 年度（含 地区、年度、数值列、文本列）。
    注意：本项目不强依赖「企业名称」列；地区与年度必然存在。
    """
    data_dir = data_dir or DATA_DIR
    files = find_data_files(data_dir)
    msgs = []
    if not files:
        return pd.DataFrame(), [f'未在 {data_dir} 下找到符合命名规则的 Excel 文件。']

    frames = []
    for path, region, part in files:
        df = read_one_file(path)
        if df is None:
            msgs.append(f'跳过（空文件）：{os.path.basename(path)}')
            continue
        df = normalize_columns(df)
        df[COL_REGION] = region
        df[COL_FILE] = os.path.basename(path)
        df[COL_SRC_YEAR] = folder_year(data_dir, path)
        frames.append(df)

    if not frames:
        return pd.DataFrame(), msgs

    data = pd.concat(frames, ignore_index=True)

    # 年度来源优先级：表头「年度」列 > 年份子文件夹名
    if COL_YEAR in data.columns:
        data[COL_YEAR] = pd.to_numeric(data[COL_YEAR], errors='coerce')
    else:
        msgs.append('表头缺少「年度」列，已采用 data/ 下的年份子文件夹名作为年度。')
        data[COL_YEAR] = data[COL_SRC_YEAR]
    data[COL_YEAR] = pd.to_numeric(data[COL_YEAR], errors='coerce')

    # 数值列数值化 + 文本列兜底
    for c in NUMERIC_COLS:
        if c not in data.columns:
            data[c] = np.nan
        data[c] = pd.to_numeric(data[c], errors='coerce')
    for c in TEXT_COLS:
        if c not in data.columns:
            data[c] = ''
        data[c] = data[c].fillna('').astype(str)

    # 剔除年度为空的行（企业名称不强依赖；若存在也保留）
    if COL_ENTERPRISE in data.columns:
        data[COL_ENTERPRISE] = data[COL_ENTERPRISE].fillna('')
    data = data.dropna(subset=[COL_YEAR])
    if data.empty:
        return data, msgs + ['读取到数据，但年度均为空。']
    data[COL_YEAR] = data[COL_YEAR].astype('Int64')

    # 完全重复行去重（避免重复记录被加总两次）
    before_dup = len(data)
    data = data.drop_duplicates()
    if len(data) < before_dup:
        msgs.append(f'已去重 {before_dup - len(data)} 条完全重复的记录。')

    # 有「企业名称」列时：同一 (地区, 企业, 年度) 多条记录 → 数值加总、文本取首个、来源合并。
    # 无「企业名称」列时：**每行即一家企业**，不合并、保留逐行明细，
    # 以便公示分布按行（=按企业）判定；地区×年度指标由 analysis.region_yearly 另行加总。
    if COL_ENTERPRISE in data.columns:
        agg = {c: (c, lambda s: s.sum(min_count=1)) for c in NUMERIC_COLS}
        for c in TEXT_COLS:
            agg[c] = (c, 'first')
        agg[COL_FILE] = (COL_FILE, lambda s: '+'.join(sorted(set(s))))
        agg[COL_SRC_YEAR] = (COL_SRC_YEAR, 'first')
        data = data.groupby([COL_REGION, COL_ENTERPRISE, COL_YEAR], as_index=False).agg(**agg)
    data = data.sort_values([COL_REGION, COL_YEAR]).reset_index(drop=True)

    msgs.append(f'成功加载 {len(files)} 个文件（含 .xls/.xlsx），合并后共 {len(data)} 条'
                f'（{data[COL_REGION].nunique()} 个地区，年度'
                f' {int(data[COL_YEAR].min())}~{int(data[COL_YEAR].max())}）。')
    return data, msgs


if __name__ == '__main__':
    d, m = load_all()
    for msg in m:
        print(msg)
    print(d.head() if not d.empty else '（空数据）')