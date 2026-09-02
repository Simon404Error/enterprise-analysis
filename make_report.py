# -*- coding: utf-8 -*-
"""
报告生成入口（企业生产经营离线分析）
=====================================
用法：
  python make_report.py -a           方式A：直接从 data/ 计算结果生成 Word 报告
  python make_report.py -f           生成「报告填写表单.xlsx」（方式B 的填写模板）
  python make_report.py 报告填写表单.xlsx   方式B：读取已填表单生成 Word 报告
  python make_report.py              交互菜单

图表：
  - 方式A 默认自动生成（离线）；可通过「图表映射.txt」或 report_img 目录替换为自定义图片。
  - 方式B 在表单「图片设置」工作表指定自定义图片。
报告不含「建议」章节；缺数据的章节自动跳过。
"""
import os
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)

import report_docx
from config import REPORT_DEFAULT_DOCX, REPORT_FORM_XLSX


def _check_env():
    try:
        import docx  # noqa: F401
        import matplotlib  # noqa: F401
    except ImportError as e:
        print(f'缺少依赖：{e}')
        print('请先执行（本包内置 python\\ 下）：')
        print('  python\\python.exe -m pip install python-docx matplotlib')
        return False
    return True


def _way_a():
    print('方式A：从 data/ 数据自动生成报告…')
    path = report_docx.generate_from_data()
    print(f'已生成：{path}')


def _make_form():
    print('生成填写表单…')
    path = report_docx.generate_form()
    print(f'已生成：{path}')
    print('填写后运行：python make_report.py ' + os.path.basename(path))


def _way_b(form_path):
    print(f'读取表单：{form_path}')
    path = report_docx.generate_from_form(form_path)
    print(f'已生成：{path}')


def _menu():
    while True:
        print('=================================')
        print('  报告生成')
        print('  1. 方式A：从 data/ 自动生成报告')
        print('  2. 生成填写表单（方式B 模板）')
        print('  3. 方式B：读取已填表单生成报告')
        print('  0. 退出')
        print('=================================')
        c = input('请选择：').strip()
        if c == '1':
            _way_a(); return
        if c == '2':
            _make_form(); return
        if c == '3':
            p = input('表单路径（如 报告填写表单.xlsx）：').strip()
            _way_b(p); return
        if c == '0':
            return


def main():
    if not _check_env():
        return
    args = sys.argv[1:]
    if not args:
        _menu(); return
    if args[0] in ('-a', '--auto', 'a'):
        _way_a(); return
    if args[0] in ('-f', '--form'):
        _make_form(); return
    if os.path.isfile(args[0]):
        _way_b(args[0]); return
    print('参数无法识别，请用：-a（自动生成）、-f（生成表单）、或直接传入表单 xlsx 路径。')
    _menu()


if __name__ == '__main__':
    main()