# -*- coding: utf-8 -*-
"""从一份「已经带数据透视表」的账单 Excel 里抽出一个干净模板。

仓库里已经自带 `模板/账本模板.xlsx`（假数据，不含任何个人账单）。
只有当你把模板弄坏了、或想换成自己的透视表样式时，才需要重新生成：

    python scripts/make_template.py --src "7、8月账单.xlsx" --out "scripts/模板/账本模板.xlsx"

它会保留原文件的表头样式和透视表结构，清空「合并账单」里的数据行。
"""
from __future__ import annotations

import argparse
import os

import openpyxl


def main() -> int:
    ap = argparse.ArgumentParser(description='从带透视表的账单 xlsx 抽模板')
    ap.add_argument('--src', required=True, help='源 xlsx（含「合并账单」sheet 和一张透视表）')
    ap.add_argument('--out', default=None,
                    help='输出模板路径（默认 <本目录>/模板/账本模板.xlsx）')
    ap.add_argument('--sheet', default='合并账单', help='数据 sheet 名（默认 合并账单）')
    args = ap.parse_args()

    here = os.path.dirname(os.path.abspath(__file__))
    out = args.out or os.path.join(here, '模板', '账本模板.xlsx')
    os.makedirs(os.path.dirname(out), exist_ok=True)

    wb = openpyxl.load_workbook(args.src)
    if args.sheet not in wb.sheetnames:
        raise SystemExit(f'找不到 sheet: {args.sheet}')
    ws = wb[args.sheet]
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)
    ws.auto_filter.ref = 'A1:I1'
    wb.save(out)
    print(f'模板已生成: {out} ({os.path.getsize(out)} bytes)')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
