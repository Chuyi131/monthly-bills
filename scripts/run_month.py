# -*- coding: utf-8 -*-
"""月度一键入口（P1）。

约定：把当月导出的两个文件丢进 `账单/<年月>/`，然后一条命令跑完：

    python 自动化方案/run_month.py --dir 账单/2026-11

它会：
  1. 在目录里自动找微信 xlsx 和支付宝 csv（按文件名特征，不写死文件名）；
  2. 调 bills.py 解析/清洗/分类，生成静态版 xlsx + 处理报告；
  3. 调 apply_pivot.py 套用带真透视表的模板，用 Excel 刷新后另存成品；
  4. 打印摘要（保留行数 / 支出 / 收入 / 自动分类覆盖率 / 待确认行数）。

可选：
    --name 11月账单        自定义成品文件名（默认按目录名推断：2026-11 → 11月账单）
    --wechat/--alipay      直接指定文件（跳过自动查找）
    --no-pivot             不生成带透视表的版本（只用静态汇总）
    --keep-stage           保留中间文件（调试用）
"""
from __future__ import annotations

import argparse
import os
import re
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)

import bills  # noqa: E402


def log(msg: str = '') -> None:
    print(msg, flush=True)


def _norm(path: str) -> str:
    return path.replace('\\', '/')


def find_files(folder: str, wechat=None, alipay=None):
    """在 folder 里找微信 xlsx 与支付宝 csv。返回 (wechat_path, alipay_path)。"""
    files = [os.path.join(folder, f) for f in os.listdir(folder)]
    files = [f for f in files if os.path.isfile(f)]

    def is_output(f):
        name = os.path.basename(f)
        return ('_处理报告' in name) or name.endswith('账单.xlsx') and not any(
            k in name for k in ('微信', 'wechat', 'alipay', '支付宝', '流水'))

    if wechat:
        w = wechat
    else:
        cands = [f for f in files
                 if f.lower().endswith(('.xlsx', '.xls'))
                 and any(k in os.path.basename(f) for k in ('微信', 'wechat'))
                 and not is_output(f)]
        if len(cands) > 1:
            # 优先"账单流水文件"这种标准导出名
            cands.sort(key=lambda f: (0 if '流水' in os.path.basename(f) else 1,
                                      -os.path.getmtime(f)))
        w = cands[0] if cands else None

    if alipay:
        a = alipay
    else:
        cands = [f for f in files
                 if f.lower().endswith(('.csv', '.xls', '.xlsx'))
                 and any(k in os.path.basename(f).lower()
                         for k in ('alipay', '支付宝'))
                 and not is_output(f)]
        if len(cands) > 1:
            cands.sort(key=lambda f: -os.path.getmtime(f))
        a = cands[0] if cands else None
    return w, a


def infer_name(folder: str) -> str:
    """按目录名推断成品文件名：2026-11 -> 11月账单；2026-09-10 -> 9、10月账单。"""
    base = os.path.basename(os.path.abspath(folder).rstrip('/\\'))
    m = re.fullmatch(r'(\d{4})-(\d{1,2})-(\d{1,2})', base)
    if m:
        m1, m2 = int(m.group(2)), int(m.group(3))
        return f'{m1}、{m2}月账单' if m1 != m2 else f'{m1}月账单'
    m = re.fullmatch(r'(\d{4})-(\d{1,2})', base)
    if m:
        return f'{int(m.group(2))}月账单'
    m = re.fullmatch(r'(\d{4})年(\d{1,2})月', base)
    if m:
        return f'{int(m.group(2))}月账单'
    return base + '_账单'


def run_bills(wechat, alipay, stage_xlsx, report_md):
    log(f'[1/3] 解析 + 清洗 + 分类：{os.path.basename(wechat)} + {os.path.basename(alipay)}')
    stats = bills.run(wechat, alipay, stage_xlsx, report_path=report_md)
    return stats


def build_pivot(stage_xlsx, final_xlsx, template=None):
    """调 apply_pivot.py 生成真透视表版本；失败时退回静态版。"""
    script = os.path.join(HERE, 'apply_pivot.py')
    cmd = [sys.executable, script, '--src', stage_xlsx, '--out', final_xlsx]
    if template:
        cmd += ['--template', template]
    log('[2/3] 套用透视表模板 + Excel 刷新')
    r = subprocess.run(cmd, stdout=sys.stdout, stderr=sys.stdout)
    if r.returncode != 0 or not os.path.exists(final_xlsx):
        log('[警告] 透视表版生成失败，输出静态版')
        shutil.copyfile(stage_xlsx, final_xlsx)
        return False
    return True


def main() -> int:
    ap = argparse.ArgumentParser(description='月度一键：账单/<年月>/ -> 成品 xlsx')
    ap.add_argument('--dir', required=True, help='放两个导出文件的目录，如 账单/2026-11')
    ap.add_argument('--name', default=None, help='成品文件名（不含扩展名）')
    ap.add_argument('--wechat', default=None, help='微信 xlsx（默认自动查找）')
    ap.add_argument('--alipay', default=None, help='支付宝 csv（默认自动查找）')
    ap.add_argument('--no-pivot', action='store_true', help='不生成带透视表的版本')
    ap.add_argument('--keep-stage', action='store_true', help='保留中间文件')
    args = ap.parse_args()

    folder = os.path.abspath(args.dir)
    if not os.path.isdir(folder):
        log(f'目录不存在: {folder}')
        return 2

    wechat, alipay = find_files(folder, args.wechat, args.alipay)
    if not wechat or not alipay:
        log('没找到需要的文件：')
        log(f'  微信 xlsx: {wechat or "缺失（文件名需含 微信）"}')
        log(f'  支付宝 csv: {alipay or "缺失（文件名需含 alipay 或 支付宝）"}')
        log(f'目录内容: {[os.path.basename(f) for f in os.listdir(folder)]}')
        return 2

    name = args.name or infer_name(folder)
    final_xlsx = os.path.join(folder, f'{name}.xlsx')
    report_md = os.path.join(folder, f'{name}_处理报告.md')
    stage = os.path.join(folder, f'_{name}_静态.xlsx')

    log('=' * 70)
    log(f'目录: {folder}')
    log(f'微信: {os.path.basename(wechat)}')
    log(f'支付宝: {os.path.basename(alipay)}')
    log(f'成品: {os.path.basename(final_xlsx)}')
    log('=' * 70)

    stats = run_bills(wechat, alipay, stage, report_md)

    if args.no_pivot:
        shutil.move(stage, final_xlsx) if os.path.exists(stage) else None
        pivot_ok = False
    else:
        pivot_ok = build_pivot(stage, final_xlsx)
        if not args.keep_stage and pivot_ok and os.path.exists(stage):
            os.unlink(stage)

    # ---- 摘要 ----
    log('[3/3] 完成')
    log('-' * 70)
    try:
        from openpyxl import load_workbook
        wb = load_workbook(final_xlsx, data_only=True)
        ws = wb['合并账单']
        rows = list(ws.iter_rows(min_row=2, values_only=True))
        exp = sum(float(r[6] or 0) for r in rows if r[4] == '支出')
        inc = sum(float(r[6] or 0) for r in rows if r[4] == '收入')
        need = [r for r in rows if r[4] in ('支出', '收入')]
        done = [r for r in need if r[5]]
        wb.close()
        log(f'  保留 {len(rows)} 行 | 支出 ¥{exp:,.2f} | 收入 ¥{inc:,.2f}')
        log(f'  自动分类 {len(done)}/{len(need)}（{len(done)/max(len(need),1):.1%}）')
        log(f'  透视表: {"已刷新" if pivot_ok else "静态汇总"}')
    except Exception as exc:  # pragma: no cover
        log(f'  （摘要读取失败：{exc}）')
    log(f'  成品: {final_xlsx}')
    log(f'  报告: {report_md}')
    log('-' * 70)
    return 0


if __name__ == '__main__':
    sys.exit(main())
