# -*- coding: utf-8 -*-
"""
文件读取模块：扫描 data/ 下的 Excel，解析地区名与分卷号，合并同一地区分卷。
返回一张“明细表”：每行 = 一个企业 × 一个年度（同企业同年度已按数值字段加总）。
"""
import os
import re
import glob

import numpy as np
import pandas as pd

from config import (DATA_DIR, FILE_PATTERN, COL_ENTERPRISE, COL_YEAR,
                    COL_REGION, COL_FILE, NUMERIC_COLS)


def find_data_files(data_dir=None):
    """扫描目录，返回 [(文件路径, 地区名, 分卷号), ...]，按命名规则过滤。"""
    data_dir = data_dir or DATA_DIR
    if not os.path.isdir(data_dir):
        return []
    files = []
    for path in sorted(glob.glob(os.path.join(data_dir, '*.xlsx')) + glob.glob(os.path.join(data_dir, '*.xls'))):
        name = os.path.basename(path)
        m = FILE_PATTERN.match(name)
        if not m:
            continue  # 不符合命名规则的文件跳过
        region, part = m.group(1).strip(), (m.group(2) or '')
        files.append((path, region, part))
    return files


def read_one_file(path):
    """读取单个 Excel（所有工作表），返回原始 DataFrame（含文件名）。"""
    frames = []
    xls = pd.ExcelFile(path)
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet, dtype=object)
        if not df.empty:
            frames.append(df)
    if not frames:
        return None
    df = pd.concat(frames, ignore_index=True)
    df[COL_FILE] = os.path.basename(path)
    return df


def load_all(data_dir=None):
    """加载 data/ 下全部符合规则的 Excel，返回 (明细表, 提示信息列表)。

    处理：
      1. 表头固定，直接按第一行读取；列名去空格。
      2. 同地区多个分卷纵向合并。
      3. 数值字段转数值；空值置 NaN。
      4. 完全重复的行去重；同一(地区,企业,年度)按数值字段加总。
      5. 企业名称或年度为空的行剔除。
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
        df.columns = [str(c).strip() for c in df.columns]
        if COL_ENTERPRISE not in df.columns or COL_YEAR not in df.columns:
            msgs.append(f'跳过（缺少表头「{COL_ENTERPRISE}」或「{COL_YEAR}」）：{os.path.basename(path)}')
            continue
        df[COL_REGION] = region
        frames.append(df)

    if not frames:
        return pd.DataFrame(), msgs

    df = pd.concat(frames, ignore_index=True)

    # 数值化
    for c in NUMERIC_COLS:
        if c not in df.columns:
            df[c] = np.nan
        df[c] = pd.to_numeric(df[c], errors='coerce')
    df[COL_YEAR] = pd.to_numeric(df[COL_YEAR], errors='coerce')

    # 剔除企业名称/年度为空的行
    df = df.dropna(subset=[COL_ENTERPRISE, COL_YEAR])
    if df.empty:
        return df, msgs + ['读取到数据，但企业名称/年度均为空。']
    df[COL_YEAR] = df[COL_YEAR].astype('Int64')

    # 完全重复行去重（避免重复记录被加总两次）
    before_dup = len(df)
    df = df.drop_duplicates()
    if len(df) < before_dup:
        msgs.append(f'已去重 {before_dup - len(df)} 条完全重复的记录。')

    # 同一(地区,企业,年度)多条记录：数值字段加总，来源文件合并
    group_cols = [COL_REGION, COL_ENTERPRISE, COL_YEAR]
    agg = {c: (c, lambda s: s.sum(min_count=1)) for c in NUMERIC_COLS}  # min_count=1 保留全空为 NaN
    agg[COL_FILE] = (COL_FILE, lambda s: '+'.join(sorted(set(s))))
    df = df.groupby(group_cols, as_index=False).agg(**agg)
    df = df.sort_values([COL_REGION, COL_ENTERPRISE, COL_YEAR]).reset_index(drop=True)

    msgs.append(f'成功加载 {len(files)} 个文件，合并后共 {len(df)} 条记录'
                f'（{df[COL_REGION].nunique()} 个地区，{df[COL_ENTERPRISE].nunique()} 家企业）。')
    return df, msgs


if __name__ == '__main__':
    # 独立调试：python file_loader.py
    d, m = load_all()
    for msg in m:
        print(msg)
    print(d.head() if not d.empty else '（空数据）')
