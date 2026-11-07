"""把 bills.py 的输出（合并账单/待确认/处理说明）套进带真透视表的模板，并刷新透视表。

用法：
    python 自动化方案/apply_pivot.py \
        --src 自动化方案/输出/9、10月账单_自动.xlsx \
        --out 自动化方案/输出/9、10月账单_自动_含透视表.xlsx

流程：
  1. 打开 自动化方案/模板/账本模板.xlsx（内含用户原有样式的真透视表）
  2. 把 src 的「合并账单」数据 + 待确认/处理说明 两个 sheet 复制进模板
  3. 修改透视缓存源区域 → A1:I{末行}
  4. 调用本机 Excel（COM）执行 RefreshAll 并另存为 out，保证透视表数值正确
  5. Excel 不可用时，保留静态 Sheet1 兜底（输出仍可读，只是透视表不刷新）
"""
import argparse
import copy
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

import openpyxl


def copy_style(src_cell, dst_cell):
    dst_cell.font = copy.copy(src_cell.font)
    dst_cell.fill = copy.copy(src_cell.fill)
    dst_cell.border = copy.copy(src_cell.border)
    dst_cell.alignment = copy.copy(src_cell.alignment)
    dst_cell.number_format = src_cell.number_format
    dst_cell.protection = copy.copy(src_cell.protection)


def copy_sheet_values(src_ws, dst_ws):
    for row in src_ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            dst = dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            copy_style(cell, dst)
    for key, dim in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[key].width = dim.width
    for r, dim in src_ws.row_dimensions.items():
        if dim.height:
            dst_ws.row_dimensions[r].height = dim.height
    if src_ws.auto_filter.ref:
        dst_ws.auto_filter.ref = src_ws.auto_filter.ref


def refresh_with_excel(path, out_path):
    """用 Excel COM 刷新透视表并另存。返回 True/False。"""
    ps = r'''
$ErrorActionPreference = "Stop"
$src = "%s"
$dst = "%s"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($src)
    $wb.RefreshAll()
    $wb.SaveAs($dst, 51)
    $wb.Close($false)
    Write-Output "REFRESH_OK"
} catch {
    Write-Output ("REFRESH_FAIL: " + $_.Exception.Message)
} finally {
    $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}
''' % (path, out_path)
    # PowerShell 5.1 需要 UTF-8 BOM 才能正确读取中文路径
    with tempfile.NamedTemporaryFile('w', suffix='.ps1', delete=False, encoding='utf-8-sig') as f:
        f.write(ps)
        tmp = f.name
    log = tmp + '.log'
    try:
        # 注意：沙箱环境不允许创建管道，所以重定向到文件而不是 capture_output
        with open(log, 'w', encoding='utf-8', errors='replace') as f:
            r = subprocess.run(['powershell.exe', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', tmp],
                               stdout=f, stderr=subprocess.STDOUT, timeout=180)
        out = ''
        if os.path.exists(log):
            out = open(log, encoding='utf-8', errors='replace').read()
        ok = 'REFRESH_OK' in out
        print('  COM:', out.strip()[:300])
        return ok
    finally:
        for p in (tmp, log):
            if os.path.exists(p):
                os.unlink(p)


def patch_cache_ref(xlsx_path, last_row):
    """把透视缓存源区域改成 A1:I{last_row}。"""
    tmp = xlsx_path + '.tmp'
    with zipfile.ZipFile(xlsx_path) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'xl/pivotCache/pivotCacheDefinition1.xml':
                s = data.decode('utf-8')
                s = re.sub(r'<worksheetSource ref="[^"]*"', f'<worksheetSource ref="A1:I{last_row}"', s)
                data = s.encode('utf-8')
            zout.writestr(item, data)
    shutil.move(tmp, xlsx_path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--src', required=True, help='bills.py 输出的 xlsx')
    ap.add_argument('--out', required=True, help='最终输出（含真透视表）')
    ap.add_argument('--template', default=None, help='模板路径，默认 自动化方案/模板/账本模板.xlsx')
    args = ap.parse_args()

    here = os.path.dirname(os.path.abspath(__file__))
    tpl = args.template or os.path.join(here, '模板', '账本模板.xlsx')
    if not os.path.exists(tpl):
        print('模板不存在，先运行 make_template.py')
        sys.exit(1)

    print('读取模板:', tpl)
    wb = openpyxl.load_workbook(tpl)
    src_wb = openpyxl.load_workbook(args.src)

    # 1) 合并账单
    src_ws = src_wb['合并账单']
    dst_ws = wb['合并账单']
    if dst_ws.max_row > 1:
        dst_ws.delete_rows(2, dst_ws.max_row - 1)
    n_rows = 0
    for r in range(2, src_ws.max_row + 1):
        for c in range(1, 10):
            s = src_ws.cell(row=r, column=c)
            d = dst_ws.cell(row=r, column=c, value=s.value)
            copy_style(s, d)
        n_rows += 1
    last_row = n_rows + 1
    dst_ws.auto_filter.ref = f'A1:I{last_row}'
    print(f'写入合并账单 {n_rows} 行，末行 {last_row}')

    # 2) 其余 sheet 原样复制
    for name in ('待确认', '处理说明'):
        if name in wb.sheetnames:
            del wb[name]
        src = src_wb[name]
        dst = wb.create_sheet(name)
        copy_sheet_values(src, dst)

    # 3) 先保存，再改缓存源区域
    os.makedirs(os.path.dirname(os.path.abspath(args.out)) or '.', exist_ok=True)
    out_abs = os.path.abspath(args.out)
    stage = out_abs + '.stage.xlsx'
    wb.save(stage)
    patch_cache_ref(stage, last_row)

    # 4) Excel 刷新
    print('调用 Excel 刷新透视表 ...')
    ok = refresh_with_excel(stage, out_abs)
    if ok and os.path.exists(out_abs):
        os.unlink(stage)
        print('完成:', args.out)
    else:
        shutil.move(stage, out_abs)
        print('⚠ Excel 刷新失败，已输出未刷新版本:', out_abs)
        print('  打开后按 Alt+F5（全部刷新）即可，静态统计不受影响。')


if __name__ == '__main__':
    main()
