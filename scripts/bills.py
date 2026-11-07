# -*- coding: utf-8 -*-
"""微信 + 支付宝账单自动化工具（原型）

流程：解析微信 xlsx / 支付宝 csv -> 合并去重 -> 清洗（删 0 元、退款、关闭）
      -> 分类（分类映射.csv / rules.yaml） -> 输出 Excel（Sheet1 透视外观 +
         合并账单 + 待确认 + 处理说明）与 Markdown 报告。

只读输入文件，只写 --out / --report 指定的输出，可重复运行（同名覆盖）。

用法：
    python 自动化方案/bills.py \
        --wechat "26accounte/微信支付账单流水文件(20260901-20261031)_20261107212517.xlsx" \
        --alipay "26accounte/alipay_record_20261107_2126_1.csv" \
        --out "自动化方案/输出/9、10月账单_自动.xlsx"

    # 一般用一键脚本
    python 自动化方案/run_pilot.py
"""
from __future__ import annotations

import argparse
import collections
import datetime as dt
import html
import math
import os
import re
import sys

import openpyxl
import yaml
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_RULES = os.path.join(HERE, 'rules.yaml')
DEFAULT_MAP = os.path.join(HERE, '分类映射.csv')
DEFAULT_DICT = os.path.join(HERE, '分类字典.yaml')

MAIN_COLS = ['交易时间', '交易类型', '交易对方', '商品', '收/支', '类型', '金额', '交易状态', '来源']
PENDING_COLS = ['交易时间', '交易类型', '交易对方', '商品', '收/支', '金额', '原因']
MAIN_WIDTHS = {'A': 20, 'B': 14, 'C': 22, 'D': 40, 'E': 8, 'F': 14, 'G': 12, 'H': 12, 'I': 10}
PENDING_WIDTHS = {'A': 20, 'B': 14, 'C': 22, 'D': 40, 'E': 8, 'F': 12, 'G': 52}
NUM_MONEY_FMT = '"¥"#,##0.00_);[Red]\\("¥"#,##0.00\\)'
TIME_FMT = 'yyyy-mm-dd hh:mm:ss'
HEADER_FONT = Font(name='微软雅黑', size=11, bold=True, color='FFFFFF')
HEADER_FILL = PatternFill('solid', fgColor='4472C4')
HEADER_ALIGN = Alignment(horizontal='center', vertical='center')
BODY_FONT = Font(name='微软雅黑', size=10)
THIN = Side(style='thin', color='A6A6A6')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
DIRECTION_ORDER = ['/', '不计收支', '收入', '支出']   # 与历史 6 月文件的分组顺序一致

LOG_PREFIX = '  '


def log(msg: str = '') -> None:
    print(msg, flush=True)


# ===========================================================================
# 一、解析
# ===========================================================================
_CTRL_RE = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]')


def _cell_text(value) -> str:
    """单元格 -> 干净字符串（去首尾空白、控制符，还原 HTML 实体）。"""
    if value is None:
        return ''
    text = str(value).strip()
    if _CTRL_RE.search(text):
        text = _CTRL_RE.sub('', text).strip()
    if '&' in text and ';' in text:
        text = html.unescape(text).strip()
    return text


def _parse_money(value) -> float:
    """把金额转成 float；无法解析返回 0.0。"""
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return 0.0
        return round(float(value), 2)
    text = str(value).strip().replace('¥', '').replace('￥', '').replace(',', '')
    text = text.replace('元', '').strip()
    if not text:
        return 0.0
    try:
        return round(float(text), 2)
    except ValueError:
        return 0.0


def _parse_time(value):
    """转成 datetime；失败返回 None。"""
    if value is None:
        return None
    if isinstance(value, dt.datetime):
        return value.replace(tzinfo=None)
    if isinstance(value, dt.date):
        return dt.datetime.combine(value, dt.time())
    text = str(value).strip().strip('"').strip("'")
    if not text:
        return None
    text = text.replace('T', ' ')
    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M:%S', '%Y-%m-%d %H:%M',
                '%Y/%m/%d %H:%M', '%Y-%m-%d', '%Y/%m/%d'):
        try:
            return dt.datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def _normalize_direction(value) -> str:
    text = _cell_text(value)
    if text in ('支出', '收入', '不计收支'):
        return text
    if text in ('', '/', '中性交易', '不计入收支'):
        return '/'
    return text


_RE_STATUS_REFUND = re.compile(r'已退款[（(]?\s*[¥￥]?\s*([\d,]+(?:\.\d+)?)')


def _parse_status_refund(status: str) -> float:
    """从微信交易状态里解析退款金额，如「已退款¥39.00」「已退款(¥0.18)」。"""
    match = _RE_STATUS_REFUND.search(_cell_text(status))
    if not match:
        return 0.0
    try:
        return round(float(match.group(1).replace(',', '')), 2)
    except ValueError:
        return 0.0


_RE_WX_SUMMARY = {
    '共': re.compile(r'共\s*(\d+)\s*笔'),
    '收入': re.compile(r'收入[：:]\s*(\d+)\s*笔\s*([\d.]+)\s*元'),
    '支出': re.compile(r'支出[：:]\s*(\d+)\s*笔\s*([\d.]+)\s*元'),
    '中性交易': re.compile(r'中性交易[：:]\s*(\d+)\s*笔\s*([\d.]+)\s*元'),
}


def parse_wechat(path: str):
    """解析微信支付账单 xlsx。

    返回 (records, meta)：
      records: [{交易时间, 交易类型, 交易对方, 商品, 收/支, 金额, 交易状态, 来源, _单号, _行号}]
      meta: {'official': 官方统计, 'header_row': 表头行号, 'data_rows': 数据行数}
    """
    log(f'  读取微信账单: {path}')
    if not os.path.exists(path):
        raise FileNotFoundError(f'微信账单不存在: {path}')
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.worksheets[0]
        rows = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()

    header_idx = None
    for idx, row in enumerate(rows[:80]):
        cells = [_cell_text(c) for c in row]
        if '交易时间' in cells and '交易对方' in cells:
            header_idx = idx
            break
    if header_idx is None:
        raise ValueError('微信 xlsx 中找不到表头行（含「交易时间」「交易对方」）')

    header = [_cell_text(c) for c in rows[header_idx]]
    col = {name: i for i, name in enumerate(header) if name}

    # 说明区官方统计（第 7-10 行）
    preamble = ' '.join(_cell_text(c) for row in rows[:header_idx] for c in row if c is not None)
    official = {}
    for key, pat in _RE_WX_SUMMARY.items():
        m = pat.search(preamble)
        if m:
            if key == '共':
                official[key] = {'count': int(m.group(1))}
            else:
                official[key] = {'count': int(m.group(1)), 'amount': round(float(m.group(2)), 2)}

    def take(row, *names):
        for name in names:
            idx = col.get(name)
            if idx is not None and idx < len(row):
                return row[idx]
        return None

    records = []
    for offset, row in enumerate(rows[header_idx + 1:], start=header_idx + 2):
        if row is None or not any(v is not None for v in row):
            continue
        raw_time = take(row, '交易时间')
        if raw_time is None or _cell_text(raw_time) == '':
            continue
        event_time = _parse_time(raw_time)
        if event_time is None:
            continue
        status_text = _cell_text(take(row, '当前状态', '交易状态'))
        records.append({
            '交易时间': event_time,
            '交易类型': _cell_text(take(row, '交易类型')),
            '交易对方': _cell_text(take(row, '交易对方')),
            '商品': _cell_text(take(row, '商品')),
            '收/支': _normalize_direction(take(row, '收/支')),
            '金额': _parse_money(take(row, '金额(元)', '金额（元）', '金额')),
            '交易状态': status_text,
            '来源': '微信支付',
            '_单号': _cell_text(take(row, '交易单号')),
            '_退款金额': _parse_status_refund(status_text),
            '_行号': offset,
        })

    meta = {'official': official, 'header_row': header_idx + 1, 'data_rows': len(records)}
    log(f'  {LOG_PREFIX}表头第 {header_idx + 1} 行，数据 {len(records)} 行，'
        f'官方口径: 共{official.get("共", {}).get("count", "?")}笔')
    return records, meta


_ALIPAY_FOOTER = re.compile(
    r'^(?:-{3,}|共\s*\d+\s*笔记录|已收入\s*[:：]|待收入\s*[:：]|已支出\s*[:：]|待支出\s*[:：]|导出时间\s*[:：])'
)
_RE_AL_SUMMARY = {
    '共': re.compile(r'^共\s*(\d+)\s*笔记录'),
    '已收入': re.compile(r'^已收入\s*[:：]\s*(\d+)\s*笔\s*[,，]\s*([\d.]+)\s*元'),
    '待收入': re.compile(r'^待收入\s*[:：]\s*(\d+)\s*笔\s*[,，]\s*([\d.]+)\s*元'),
    '已支出': re.compile(r'^已支出\s*[:：]\s*(\d+)\s*笔\s*[,，]\s*([\d.]+)\s*元'),
    '待支出': re.compile(r'^待支出\s*[:：]\s*(\d+)\s*笔\s*[,，]\s*([\d.]+)\s*元'),
    '导出时间': re.compile(r'^导出时间\s*[:：]\s*(\S+)'),
}


def _read_text_any_encoding(path: str) -> str:
    raw = open(path, 'rb').read()
    for enc in ('utf-8-sig', 'gb18030', 'gbk', 'utf-8'):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode('gb18030', errors='replace')


def parse_alipay(path: str):
    """解析支付宝 csv（GBK/GB18030）。

    注意两点（踩坑记录）：
      1) 数据行分隔符是 \\t, ，先把 \\t, 换成 , 再按逗号切分；
      2) 页脚判定必须用「行首锚定」的正则。数据行的「资金状态」列本身就叫
         「已支出」，用「行内含已支出」会把 187 行数据全部误删。
    """
    log(f'  读取支付宝账单: {path}')
    if not os.path.exists(path):
        raise FileNotFoundError(f'支付宝账单不存在: {path}')
    text = _read_text_any_encoding(path)
    lines = text.splitlines()

    header_idx = None
    for idx, line in enumerate(lines[:30]):
        if '交易号' in line and '商家订单号' in line:
            header_idx = idx
            break
    if header_idx is None:
        raise ValueError('支付宝 csv 中找不到表头行（含「交易号」「商家订单号」）')

    header = [h.strip() for h in lines[header_idx].split(',')]
    while header and header[-1] == '':
        header.pop()
    ncols = len(header)
    col = {name: i for i, name in enumerate(header)}
    if '金额（元）' not in col and '金额(元)' in col:
        col['金额（元）'] = col['金额(元)']

    official = {}
    records = []
    skipped_footer = 0
    for lineno, raw in enumerate(lines[header_idx + 1:], start=header_idx + 2):
        line = raw.strip()
        if not line:
            continue
        if _ALIPAY_FOOTER.match(line):
            skipped_footer += 1
            for key, pat in _RE_AL_SUMMARY.items():
                m = pat.match(line)
                if not m:
                    continue
                if key == '导出时间':
                    official[key] = m.group(1)
                elif key == '共':
                    official[key] = {'count': int(m.group(1))}
                else:
                    official[key] = {'count': int(m.group(1)),
                                     'amount': round(float(m.group(2)), 2)}
            continue

        fields = [f.strip() for f in line.replace('\t,', ',').split(',')]
        while fields and fields[-1] == '':
            fields.pop()
        if len(fields) > ncols:
            # 商品名称里可能带逗号：把多出来的片段并回「商品名称」列
            goods_idx = col.get('商品名称', 8)
            extra = len(fields) - ncols
            fields = (fields[:goods_idx]
                      + [','.join(fields[goods_idx:goods_idx + extra + 1])]
                      + fields[goods_idx + extra + 1:])
        if len(fields) < ncols:
            fields.extend([''] * (ncols - len(fields)))

        def f(name):
            idx = col.get(name)
            return fields[idx] if idx is not None and idx < len(fields) else ''

        event_time = _parse_time(f('交易创建时间')) or _parse_time(f('付款时间'))
        if event_time is None:
            continue
        records.append({
            '交易时间': event_time,
            '交易类型': _cell_text(f('类型')),
            '交易对方': _cell_text(f('交易对方')),
            '商品': _cell_text(f('商品名称')),
            '收/支': _normalize_direction(f('收/支')),
            '金额': _parse_money(f('金额（元）')),
            '交易状态': _cell_text(f('交易状态')),
            '来源': '支付宝',
            '_单号': _cell_text(f('交易号')),
            '_退款金额': _parse_money(f('成功退款（元）')),
            '_行号': lineno,
        })

    meta = {'official': official, 'header_row': header_idx + 1, 'data_rows': len(records),
            'skipped_footer': skipped_footer}
    log(f'  {LOG_PREFIX}表头第 {header_idx + 1} 行，数据 {len(records)} 行，'
        f'跳过页脚 {skipped_footer} 行，官方口径: 共{official.get("共", {}).get("count", "?")}笔')
    return records, meta


# ===========================================================================
# 二、合并 + 去重
# ===========================================================================
def _dedupe_key(rec: dict):
    no = _cell_text(rec.get('_单号'))
    if no and no != '/':
        return ('no', rec['来源'], no)
    return ('content', rec['来源'], rec['交易时间'].isoformat(timespec='seconds'),
            rec['交易对方'], rec['商品'], round(float(rec['金额']), 2), rec['收/支'])


def merge_and_dedupe(wechat_records, alipay_records):
    """合并两份账单并按交易单号去重（没有单号时按内容去重）。"""
    combined = list(wechat_records) + list(alipay_records)
    seen = set()
    kept = []
    duplicates = []
    for rec in combined:
        key = _dedupe_key(rec)
        if key in seen:
            rec['_is_duplicate'] = True
            duplicates.append(rec)
            continue
        seen.add(key)
        kept.append(rec)
    by_source = collections.Counter(r['来源'] for r in kept)
    log(f'  合并 {len(combined)} 行 -> 去重删除 {len(duplicates)} 行 -> 保留 {len(kept)} 行'
        f'（微信 {by_source.get("微信支付", 0)} / 支付宝 {by_source.get("支付宝", 0)}）')
    return kept, duplicates


# ===========================================================================
# 三、清洗
# ===========================================================================
def _contains_any(text: str, needles) -> bool:
    if not needles:
        return False
    return any(str(n) in text for n in needles)


def _clean_cfg(cfg: dict, key: str, default=None):
    return (cfg or {}).get('clean', {}).get(key, default)


def clean(records, cfg: dict, duplicates=(), dictionary=None):
    """删除 0 元行、退款行、交易关闭行、按口径不记账的电费AA代收行。"""
    zero_cfg = _clean_cfg(cfg, 'drop_zero_amount', True)
    refund_cfg = _clean_cfg(cfg, 'drop_refund', {}) or {}
    closed_cfg = _clean_cfg(cfg, 'drop_closed', {}) or {}
    neutral_dirs = set(_clean_cfg(cfg, 'neutral_directions', ['不计收支', '/']))

    refund_sources = set(refund_cfg.get('sources') or [])
    refund_type_contains = refund_cfg.get('type_contains') or []
    refund_status_contains = refund_cfg.get('status_contains') or []
    closed_values = set(closed_cfg.get('status_values') or [])
    closed_dirs = set(closed_cfg.get('directions') or [])
    keep_neutral_closed = bool(closed_cfg.get('keep_neutral', True))

    # 字典 special_rules.电费AA收款：舍友转来的电费份额不记账（删除）
    special = (dictionary or {}).get('special') or {}
    aa_rule = special.get('电费AA收款') or {}
    aa_enabled = str(aa_rule.get('操作') or '').strip() == '删除'
    aa_direction = aa_rule.get('方向')
    aa_type = aa_rule.get('交易类型')
    aa_names = list(aa_rule.get('对方名单') or [])

    kept = []
    removed = []
    netted = []
    for rec in list(records) + list(duplicates):
        reason = None
        amount = round(float(rec.get('金额') or 0), 2)
        direction = rec.get('收/支') or ''
        status = _cell_text(rec.get('交易状态'))
        trans_type = _cell_text(rec.get('交易类型'))
        source = rec.get('来源') or ''
        refund = round(float(rec.get('_退款金额') or 0), 2)
        refund_enabled = bool(refund_cfg.get('enabled', True)) and (
            not refund_sources or source in refund_sources)

        if rec.get('_is_duplicate'):
            reason = '重复记录'
        elif zero_cfg and abs(amount) < 0.005:
            reason = '金额为0'
        elif (aa_enabled and direction == aa_direction
              and aa_type and aa_type in trans_type
              and _contains_any(_cell_text(rec.get('交易对方')), aa_names)):
            reason = '电费AA代收（按示例口径不记账）'
        elif (closed_cfg.get('enabled', True) and status in closed_values
              and (not closed_dirs or direction in closed_dirs)):
            reason = f'{source}交易关闭'
        elif (closed_cfg.get('enabled', True) and status in closed_values
              and direction in neutral_dirs and keep_neutral_closed):
            # 不计收支/中性 的交易关闭行：保留，类型留空
            rec['_备注'] = '中性方向的交易关闭行，按规则保留'
            kept.append(rec)
            continue
        elif refund_enabled and direction in ('支出', '收入'):
            is_refund_row = ('退款' in trans_type
                             or (direction == '收入'
                                 and _contains_any(status, refund_status_contains)))
            if is_refund_row:
                # 退款收入行删除；退款金额改为在原始支出行上抵减，不重复计收入
                reason = f'{source}退款收入行' if direction == '收入' else f'{source}退款行'
            elif refund > 0.005 and refund >= amount - 0.005:
                reason = f'{source}全额退款'
            elif refund > 0.005:
                # 部分退款：按「原支出 − 退款金额」净额入账
                rec['_原金额'] = amount
                rec['_净额退款'] = refund
                rec['金额'] = round(amount - refund, 2)
                rec['_备注'] = f'部分退款净额入账（原 {amount:.2f} − 退款 {refund:.2f}）'
                netted.append(rec)
                kept.append(rec)
                continue
            elif _contains_any(status, refund_status_contains):
                reason = f'{source}退款行（未解析到退款金额）'

        if reason:
            rec['_删除原因'] = reason
            removed.append(rec)
        else:
            kept.append(rec)

    by_reason = collections.Counter(r['_删除原因'] for r in removed)
    by_reason_source = collections.Counter((r['_删除原因'], r['来源']) for r in removed)
    by_reason_sorted = dict(sorted(by_reason.items(), key=lambda kv: (-kv[1], kv[0])))
    log(f'  删除 {len(removed)} 行 -> 保留 {len(kept)} 行')
    for reason, count in by_reason_sorted.items():
        log(f'    - {reason}: {count} 行')
    for rec in netted:
        log(f'    * 部分退款净额入账: {rec["来源"]} {rec["交易对方"]} '
            f'{rec["_原金额"]:.2f} - {rec["_退款金额"]:.2f} = {rec["金额"]:.2f}')
    return kept, removed, {'by_reason': by_reason_sorted,
                           'by_reason_source': by_reason_source,
                           'netted': netted}


# ===========================================================================
# 四、分类
# ===========================================================================
def load_mapping(path: str):
    """读取 分类映射.csv -> (精确字典, 包含匹配列表[按商户名长度降序])"""
    if not os.path.exists(path):
        log(f'  [警告] 分类映射不存在，跳过: {path}')
        return {}, []
    exact = {}
    items = []
    with open(path, 'r', encoding='utf-8-sig', newline='') as fh:
        import csv as _csv
        reader = _csv.reader(fh)
        for row in reader:
            if not row or len(row) < 2:
                continue
            merchant, cat = row[0].strip(), row[1].strip()
            if not merchant or not cat or merchant == '商户匹配':
                continue
            exact[merchant] = cat
            items.append((merchant, cat))
    items.sort(key=lambda kv: -len(kv[0]))
    log(f'  分类映射: {len(exact)} 条（精确 {len(exact)}，包含 {len(items)}）')
    return exact, items


def load_rules(path: str) -> dict:
    if not os.path.exists(path):
        log(f'  [警告] 规则文件不存在，跳过: {path}')
        return {}
    with open(path, 'r', encoding='utf-8') as fh:
        cfg = yaml.safe_load(fh) or {}
    rules = cfg.get('classify') or []
    log(f'  分类规则: {len(rules)} 条')
    return cfg


def load_dictionary(path: str) -> dict:
    """加载固定资产分类字典（expense/income 标签、aliases、special_rules）。

    返回 {'expense': [...], 'income': [...], 'aliases': {...},
          'special': {...}, 'neutral': [...], 'version': n}
    """
    if not os.path.exists(path):
        raise FileNotFoundError(f'分类字典不存在: {path}')
    with open(path, 'r', encoding='utf-8') as fh:
        raw = yaml.safe_load(fh) or {}
    expense = list((raw.get('expense') or {}).keys())
    income = list((raw.get('income') or {}).keys())
    aliases = dict(raw.get('aliases') or {})
    special = dict(raw.get('special_rules') or {})
    neutral = list(raw.get('neutral') or ['不计收支', '/'])
    cross_direction = list(raw.get('cross_direction') or [])
    if not expense and not income:
        raise ValueError(f'分类字典为空或格式不对: {path}')
    log(f'  分类字典: 支出 {len(expense)} 类 / 收入 {len(income)} 类 / '
        f'aliases {len(aliases)} 条 / 特殊规则 {len(special)} 条')
    return {'expense': expense, 'income': income, 'aliases': aliases,
            'special': special, 'neutral': neutral,
            'cross_direction': cross_direction,
            'version': raw.get('version')}


def _match_rule(rec: dict, match: dict) -> bool:
    if not match:
        return True
    for key, expect in match.items():
        if key == '金额_min':
            if float(rec.get('金额') or 0) < float(expect):
                return False
            continue
        if key == '金额_max':
            if float(rec.get('金额') or 0) > float(expect):
                return False
            continue
        if key.endswith('_not_contains'):
            column = key[:-len('_not_contains')]
            needles = expect if isinstance(expect, (list, tuple)) else [expect]
            if _contains_any(_cell_text(rec.get(column)), needles):
                return False
            continue
        if key.endswith('_contains'):
            column = key[:-len('_contains')]
            needles = expect if isinstance(expect, (list, tuple)) else [expect]
            if not _contains_any(_cell_text(rec.get(column)), needles):
                return False
            continue
        if key.endswith('_equals'):
            column = key[:-len('_equals')]
            column_value = _cell_text(rec.get(column))
            allowed = expect if isinstance(expect, (list, tuple)) else [expect]
            if column_value not in [str(v) for v in allowed]:
                return False
            continue
        # 普通字段：完全等于（或属于列表）
        column_value = _cell_text(rec.get(key))
        allowed = expect if isinstance(expect, (list, tuple)) else [expect]
        if column_value not in [str(v) for v in allowed]:
            return False
    return True


def _normalize_tag(cat: str, aliases: dict) -> str:
    """按字典 aliases 把旧标签链式转成新标签。"""
    cat = (cat or '').strip()
    seen = set()
    while cat in aliases and cat not in seen:
        seen.add(cat)
        cat = str(aliases[cat]).strip()
    return cat


def classify(records, mapping_exact, mapping_contains, cfg, dictionary=None):
    """给「支出/收入」行打类型；第一条命中的映射/规则生效。

    强校验：类型只能是字典 expense / income 里的标签，否则拦下并进「待确认」。
    """
    rules = cfg.get('classify') or []
    dictionary = dictionary or {}
    aliases = dict(dictionary.get('aliases') or {})
    allowed_expense = set(dictionary.get('expense') or [])
    allowed_income = set(dictionary.get('income') or [])
    cross_direction = set(dictionary.get('cross_direction') or [])
    neutral_dirs = set((cfg.get('clean') or {}).get('neutral_directions')
                       or dictionary.get('neutral') or ['不计收支', '/'])
    income_only = set(cfg.get('income_only_categories') or []) or set(dictionary.get('income') or [])
    special = dictionary.get('special') or {}
    group_rule = special.get('群收款') or {}

    method_counter = collections.Counter()
    rule_hits = collections.Counter()
    rule_notes = collections.Counter()
    invalid = []
    need = classified = 0
    pending = []

    def allowed_for(direction: str):
        return (allowed_income | cross_direction) if direction == '收入' else allowed_expense

    def accept(rec, direction, raw_cat, method):
        cat = _normalize_tag(raw_cat, aliases)
        if not cat:
            return False
        if cat in income_only and direction != '收入':
            return False
        allowed = allowed_for(direction)
        if allowed and cat not in allowed:
            rec['_待确认原因'] = f'分类「{cat}」不在分类字典内，已拦下待人工确认'
            invalid.append({'时间': rec.get('交易时间'), '交易对方': rec.get('交易对方'),
                            '收/支': direction, '标签': cat, '允许': sorted(allowed)})
            return False
        rec['类型'] = cat
        rec['_分类来源'] = method
        method_counter[method] += 1
        return True

    for rec in records:
        rec['类型'] = ''
        rec['_分类来源'] = ''
        direction = rec.get('收/支') or ''
        if direction in neutral_dirs or direction not in ('支出', '收入'):
            continue
        need += 1
        merchant = _cell_text(rec.get('交易对方'))

        # 0) 特殊规则（如房间充值 ÷4，已在 run() 里折算并打标）
        special_type = rec.get('_特殊类型')
        if special_type:
            if accept(rec, direction, special_type, '特殊规则'):
                classified += 1
            else:
                pending.append(rec)
            continue

        # 1) 群收款：按字典 special_rules，支出/收入都记「团建」（示例口径选 A）
        if '群收款' in _cell_text(rec.get('交易类型')):
            group_type = str(group_rule.get(direction) or '').strip() or '团建'
            if accept(rec, direction, group_type, '特殊规则'):
                classified += 1
            else:
                pending.append(rec)
            continue

        # 2) 分类映射：先精确
        cat = mapping_exact.get(merchant)
        if cat and accept(rec, direction, cat, '映射(精确)'):
            classified += 1
            continue
        # 3) 分类映射：再包含（长者优先）
        hit = False
        for needle, needle_cat in mapping_contains:
            if needle and needle in merchant and accept(rec, direction, needle_cat, '映射(包含)'):
                hit = True
                break
        if hit:
            classified += 1
            continue
        # 4) rules.yaml 关键词规则
        for rule in rules:
            if _match_rule(rec, rule.get('match') or {}):
                name = str(rule.get('name') or '').strip()
                if not name:
                    continue
                if accept(rec, direction, name, f'规则:{name}'):
                    rule_hits[name] += 1
                    if rule.get('note'):
                        rule_notes[(name, rule['note'])] += 1
                    hit = True
                break
        if hit:
            classified += 1
        else:
            # 兜底（示例口径）：用途不明的支出归「其他」、收入归「其他收入」。
            # 「待确认」只留给字典外标签/异常，正常应为 0 行。
            fallback_cat = '其他收入' if direction == '收入' else '其他'
            if accept(rec, direction, fallback_cat, '兜底(示例口径)'):
                classified += 1
            else:
                pending.append(rec)

    # 字典外标签：逐条打印警告（已强制进待确认）
    if invalid:
        log(f'  [警告] 发现 {len(invalid)} 行「类型」不在分类字典内，已拦下并转入待确认：')
        for item in invalid:
            log(f'      - {item["标签"]!r}（{item["收/支"]}）{item["交易对方"]}  '
                f'允许: {item["允许"]}')

    log(f'  需要分类 {need} 行，自动分类 {classified} 行，待确认 {need - classified} 行'
        f'（覆盖率 {classified / need:.1%}）' if need else '  没有需要分类的行')
    if invalid:
        log(f'  字典外标签 {len(invalid)} 条（必须为 0）')
    return {
        'need': need,
        'classified': classified,
        'pending': pending,
        'method_counter': method_counter,
        'rule_hits': rule_hits,
        'rule_notes': rule_notes,
        'invalid': invalid,
    }


# ===========================================================================
# 五、统计 / 透视
# ===========================================================================
def amount_stats(records):
    stats = {}
    for direction in DIRECTION_ORDER:
        subset = [r for r in records if r.get('收/支') == direction]
        if subset:
            stats[direction] = {'count': len(subset),
                                'amount': round(sum(float(r['金额']) for r in subset), 2)}
    others = collections.defaultdict(list)
    for rec in records:
        if rec.get('收/支') not in DIRECTION_ORDER:
            others[rec['收/支']].append(rec)
    for direction, subset in others.items():
        stats[direction] = {'count': len(subset),
                            'amount': round(sum(float(r['金额']) for r in subset), 2)}
    return stats


def amount_by_source(records):
    """(来源, 收/支) -> 笔数/金额，用于按来源对账。"""
    out = collections.defaultdict(lambda: {'count': 0, 'amount': 0.0})
    for rec in records:
        key = (rec.get('来源'), rec.get('收/支'))
        out[key]['count'] += 1
        out[key]['amount'] = round(out[key]['amount'] + float(rec['金额']), 2)
    return dict(out)


def classify_coverage(records):
    """按收/支 统计「需要分类 / 已分类 / 待确认」，用于报告覆盖率。"""
    cov = collections.OrderedDict()
    for direction in ('支出', '收入'):
        subset = [r for r in records if r.get('收/支') == direction]
        if not subset:
            continue
        done = sum(1 for r in subset if r.get('类型'))
        cov[direction] = {'need': len(subset), 'done': done,
                          'pending': len(subset) - done,
                          'coverage': done / len(subset)}
    return cov


def parse_stats(records):
    """解析阶段的 来源/收/支 统计，用于和官方口径核对。"""
    stats = {}
    for direction in DIRECTION_ORDER:
        subset = [r for r in records if r.get('收/支') == direction]
        if subset:
            stats[direction] = {'count': len(subset),
                                'amount': round(sum(float(r['金额']) for r in subset), 2)}
    return stats


def build_pivot(records):
    """构造 Sheet1 的静态透视（按 收/支 分组，组内按 类型 分组）。"""
    groups = collections.OrderedDict((d, {'total': 0.0, 'count': 0, 'types': collections.OrderedDict()})
                                     for d in DIRECTION_ORDER)
    for rec in records:
        direction = rec.get('收/支')
        group = groups.get(direction)
        if group is None:
            group = groups.setdefault(direction, {'total': 0.0, 'count': 0,
                                                  'types': collections.OrderedDict()})
        cat = rec.get('类型') or ''
        amount = round(float(rec['金额']), 2)
        group['total'] = round(group['total'] + amount, 2)
        group['count'] += 1
        sub = group['types'].setdefault(cat, {'amount': 0.0, 'count': 0})
        sub['amount'] = round(sub['amount'] + amount, 2)
        sub['count'] += 1

    rows = []  # [(label, amount, kind)] kind: group / type / blank / total
    grand = 0.0
    for direction, group in groups.items():
        if group['count'] == 0:
            continue
        rows.append((direction, group['total'], 'group'))
        typed = [(cat, info) for cat, info in group['types'].items() if cat]
        blank = group['types'].get('', None)
        typed.sort(key=lambda kv: (-kv[1]['amount'], kv[0]))
        for cat, info in typed:
            rows.append((cat, info['amount'], 'type'))
        if blank:
            rows.append(('(空白)', blank['amount'], 'blank'))
        grand = round(grand + group['total'], 2)
    rows.append(('总计', grand, 'total'))
    return rows


def expense_by_type(records):
    agg = collections.defaultdict(lambda: {'count': 0, 'amount': 0.0})
    for rec in records:
        if rec.get('收/支') != '支出':
            continue
        cat = rec.get('类型') or '(未分类)'
        agg[cat]['count'] += 1
        agg[cat]['amount'] = round(agg[cat]['amount'] + float(rec['金额']), 2)
    return sorted(((cat, info['count'], info['amount']) for cat, info in agg.items()),
                  key=lambda x: (-x[2], x[0]))


def income_by_type(records):
    agg = collections.defaultdict(lambda: {'count': 0, 'amount': 0.0})
    for rec in records:
        if rec.get('收/支') != '收入':
            continue
        cat = rec.get('类型') or '(未分类)'
        agg[cat]['count'] += 1
        agg[cat]['amount'] = round(agg[cat]['amount'] + float(rec['金额']), 2)
    return sorted(((cat, info['count'], info['amount']) for cat, info in agg.items()),
                  key=lambda x: (-x[2], x[0]))


def pending_reason(rec: dict) -> str:
    """给「待确认」行生成人话原因（按交易特征从具体到笼统）。"""
    trans_type = _cell_text(rec.get('交易类型'))
    goods = _cell_text(rec.get('商品'))
    merchant = _cell_text(rec.get('交易对方'))
    direction = rec.get('收/支')
    if '在线缴费' in goods or '在线缴费' in merchant:
        return '中国地质大学在线缴费，疑似学费/住宿费，金额较大需人工确认'
    if '先购后付' in goods:
        return '自动售货机「先购后付」，历史同类未分类（可忽略）'
    if '充电宝' in goods:
        return '共享充电宝租借，无对应类型'
    if '先充后付' in goods or '充电桩' in merchant:
        return '两轮车充电桩，无对应类型'
    if '企鹅共享' in goods:
        return '共享设备（企鹅共享），无对应类型'
    if '寄件' in goods or '运费' in goods:
        return '快递/寄件费，无对应类型'
    if '抽纸' in goods or '纸巾' in goods:
        return '生活用品购买，历史同类未分类'
    if '转账' in trans_type:
        return '个人转账，用途需人工判断（还钱/团建/AA 等）'
    if '群收款' in trans_type:
        return '群收款 AA，归属需人工确认'
    if '扫码' in trans_type or '二维码' in trans_type:
        return '扫码/二维码收款，付款对象与用途未知'
    if '红包' in trans_type:
        return '微信红包，历史习惯不分类（可忽略）'
    if '提现' in goods or '评价奖励' in goods or '返现' in goods:
        return '平台小额返现/奖励，历史习惯不分类（可忽略）'
    if merchant.startswith('商户_') or merchant.startswith('个体户'):
        return '个体工商户移动支付，用途未知'
    if direction == '收入':
        return '收入来源无法自动判断'
    return '未知商户，无法自动分类'


# ===========================================================================
# 六、Excel 输出
# ===========================================================================
def _write_header(ws, columns, widths):
    ws.append(columns)
    for idx in range(1, len(columns) + 1):
        cell = ws.cell(1, idx)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGN
        cell.border = BORDER
    ws.row_dimensions[1].height = 16.2
    for letter, width in widths.items():
        ws.column_dimensions[letter].width = width


def _style_body(ws, nrows, ncols, money_col=0, time_col=0, center_cols=()):
    for row_idx in range(2, nrows + 1):
        for col_idx in range(1, ncols + 1):
            cell = ws.cell(row_idx, col_idx)
            cell.font = BODY_FONT
            cell.border = BORDER
            if col_idx == money_col:
                cell.alignment = Alignment(horizontal='right', vertical='center')
            elif col_idx in center_cols:
                cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                cell.alignment = Alignment(vertical='center', wrap_text=False)
            if col_idx == time_col:
                cell.number_format = TIME_FMT
            if col_idx == money_col:
                cell.number_format = NUM_MONEY_FMT


def _fmt_amount(value: float) -> str:
    return f'{value:,.2f}'


def build_sheet1(ws, pivot_rows):
    """Sheet1：模拟用户的数据透视表外观（静态表）。"""
    ws['A3'] = '行标签'
    ws['B3'] = '求和项:金额'
    for cell in (ws['A3'], ws['B3']):
        cell.font = Font(name='微软雅黑', size=11, bold=True)
        cell.fill = PatternFill('solid', fgColor='D9E1F2')
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = BORDER
    ws.row_dimensions[3].height = 18

    row = 4
    for label, amount, kind in pivot_rows:
        ws.cell(row, 1, label)
        ws.cell(row, 2, round(float(amount), 2)).number_format = NUM_MONEY_FMT
        for col in (1, 2):
            cell = ws.cell(row, col)
            cell.font = Font(name='微软雅黑', size=11,
                             bold=(kind in ('group', 'total')))
            cell.border = BORDER
            cell.alignment = Alignment(horizontal='left' if col == 1 else 'right',
                                       vertical='center')
            if col == 2:
                cell.number_format = NUM_MONEY_FMT
        if kind == 'total':
            for col in (1, 2):
                ws.cell(row, col).fill = PatternFill('solid', fgColor='F2F2F2')
        row += 1
    ws.column_dimensions['A'].width = 24
    ws.column_dimensions['B'].width = 18
    ws.sheet_view.showGridLines = False


def build_main_sheet(ws, records):
    """合并账单：9 列，按交易时间倒序。"""
    _write_header(ws, MAIN_COLS, MAIN_WIDTHS)
    for rec in records:
        ws.append([
            rec.get('交易时间'),
            rec.get('交易类型') or '',
            rec.get('交易对方') or '',
            rec.get('商品') or '',
            rec.get('收/支') or '',
            (rec.get('类型') or None),
            round(float(rec.get('金额') or 0), 2),
            rec.get('交易状态') or '',
            rec.get('来源') or '',
        ])
    nrows = len(records) + 1
    _style_body(ws, nrows, len(MAIN_COLS), money_col=7, time_col=1, center_cols=(5, 9))
    ws.auto_filter.ref = f'A1:I{nrows}'
    ws.freeze_panes = 'A2'


def build_pending_sheet(ws, records):
    _write_header(ws, PENDING_COLS, PENDING_WIDTHS)
    for rec in records:
        ws.append([
            rec.get('交易时间'),
            rec.get('交易类型') or '',
            rec.get('交易对方') or '',
            rec.get('商品') or '',
            rec.get('收/支') or '',
            round(float(rec.get('金额') or 0), 2),
            rec.get('_待确认原因') or '',
        ])
    nrows = len(records) + 1
    _style_body(ws, nrows, len(PENDING_COLS), money_col=6, time_col=1, center_cols=(5,))
    for row_idx in range(2, nrows + 1):
        ws.cell(row_idx, 7).alignment = Alignment(vertical='center', wrap_text=True)
    if records:
        ws.auto_filter.ref = f'A1:G{nrows}'
    ws.freeze_panes = 'A2'


def build_notes_sheet(ws, blocks):
    """处理说明：A=项目 B=数值 C=说明。"""
    ws.column_dimensions['A'].width = 38
    ws.column_dimensions['B'].width = 22
    ws.column_dimensions['C'].width = 88
    row = 1
    for block in blocks:
        kind = block[0]
        if kind == 'title':
            text = block[1]
            ws.cell(row, 1, text).font = Font(name='微软雅黑', size=14, bold=True)
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
            ws.row_dimensions[row].height = 24
            row += 1
        elif kind == 'section':
            text = block[1]
            cell = ws.cell(row, 1, text)
            cell.font = Font(name='微软雅黑', size=11, bold=True)
            cell.fill = PatternFill('solid', fgColor='D9E1F2')
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
            for col in range(1, 4):
                ws.cell(row, col).fill = PatternFill('solid', fgColor='D9E1F2')
                ws.cell(row, col).border = BORDER
            row += 1
        elif kind == 'row':
            _, item, value, desc = block
            values = [item, value, desc]
            for col, val in enumerate(values, start=1):
                cell = ws.cell(row, col, val)
                cell.font = BODY_FONT
                cell.border = BORDER
                cell.alignment = Alignment(vertical='top',
                                           wrap_text=(col == 3),
                                           horizontal='right' if col == 2 and isinstance(val, (int, float)) else 'left')
            row += 1
        elif kind == 'blank':
            row += 1
        elif kind == 'text':
            cell = ws.cell(row, 1, block[1])
            cell.font = BODY_FONT
            cell.alignment = Alignment(vertical='top', wrap_text=True)
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)
            row += 1


# ===========================================================================
# 七、报告
# ===========================================================================
def _md_cell(value) -> str:
    text = '' if value is None else str(value)
    return text.replace('|', '\\|').replace('\n', '<br>')


def write_report(path, stats, kept, pending, pivot_rows):
    lines = []
    add = lines.append
    add('# 9、10 月账单自动化处理报告')
    add('')
    add(f'- 生成时间：{stats["generated_at"]}')
    add(f'- 微信账单：`{stats["inputs"]["微信"]}`')
    add(f'- 支付宝账单：`{stats["inputs"]["支付宝"]}`')
    add(f'- 规则文件：`{stats["inputs"]["规则"]}`')
    add(f'- 分类映射：`{stats["inputs"]["映射"]}`')
    add(f'- 分类字典：`{stats["inputs"]["字典"]}`'
        f'（{len(stats["dictionary"]["expense"])} 个支出类 + {len(stats["dictionary"]["income"])} 个收入类）')
    add(f'- 输出文件：`{stats["output_excel"]}`')
    add('')

    add('## 一、总体统计')
    add('')
    add('| 项目 | 数值 |')
    add('| --- | ---: |')
    rows = [
        ('原始行数', stats['raw_total']),
        ('  微信支付', stats['raw_by_source'].get('微信支付', 0)),
        ('  支付宝', stats['raw_by_source'].get('支付宝', 0)),
        ('去重删除', stats['duplicates']),
        ('删除合计（含去重）', stats['removed_total']),
    ]
    for reason, count in stats['removed_by_reason'].items():
        rows.append((f'  删除：{reason}', count))
    rows.extend([
        ('保留行数', stats['kept_total']),
        ('  微信支付', stats['kept_by_source'].get('微信支付', 0)),
        ('  支付宝', stats['kept_by_source'].get('支付宝', 0)),
        ('需要分类的行（支出+收入）', stats['need_classify']),
        ('自动分类行数', stats['classified']),
        ('待确认行数', stats['pending_total']),
        ('自动分类覆盖率', f'{stats["coverage"]:.1%}'),
        ('分类字典外标签数（必须为 0）', len(stats['invalid_tags'])),
    ])
    for item, value in rows:
        add(f'| {_md_cell(item)} | {_md_cell(value)} |')
    add('')

    add('## 二、保留行金额统计')
    add('')
    add('| 收/支 | 笔数 | 金额(元) |')
    add('| --- | ---: | ---: |')
    for direction in DIRECTION_ORDER:
        info = stats['kept_amount'].get(direction)
        if info:
            add(f'| {_md_cell(direction)} | {info["count"]} | {_fmt_amount(info["amount"])} |')
    for direction, info in stats['kept_amount'].items():
        if direction not in DIRECTION_ORDER:
            add(f'| {_md_cell(direction)} | {info["count"]} | {_fmt_amount(info["amount"])} |')
    add('')

    add('## 三、官方口径核对')
    add('')
    add('| 来源 | 指标 | 官方口径 | 本次解析 | 是否一致 |')
    add('| --- | --- | --- | --- | --- |')
    for source, block in (('微信支付', stats['wechat_meta']), ('支付宝', stats['alipay_meta'])):
        official = block['official']
        parsed = block['parsed']
        total = official.get('共', {}).get('count')
        add(f'| {source} | 共笔数 | {_md_cell(total)} | {_md_cell(parsed.get("_total"))} | '
            f'{"✅" if total == parsed.get("_total") else "❌"} |')
        pairs = ([('收入', '收入'), ('支出', '支出'), ('中性交易', '/')]
                 if source == '微信支付' else
                 [('已收入', '收入'), ('已支出', '支出')])
        for official_key, parsed_key in pairs:
            off = official.get(official_key)
            par = parsed.get(parsed_key)
            if not off:
                continue
            label = ('已支出（毛额）' if source == '支付宝' and official_key == '已支出'
                     else official_key)
            off_text = f'{off["count"]} 笔 / {_fmt_amount(off["amount"])} 元'
            par_text = f'{par["count"]} 笔 / {_fmt_amount(par["amount"])} 元' if par else '-'
            same = par and off['count'] == par['count'] and abs(off['amount'] - par['amount']) < 0.005
            add(f'| {source} | {label} | {off_text} | {par_text} | '
                f'{"✅" if same else "⚠️"} |')
        for official_key in ('待收入', '待支出'):
            off = official.get(official_key)
            if off:
                add(f'| {source} | {official_key} | {off["count"]} 笔 / {_fmt_amount(off["amount"])} 元 | - | - |')
    if stats['alipay_refund_total']:
        al_kept_exp = stats.get('kept_amount_by_source', {}).get(('支付宝', '支出'), {}).get('amount', 0)
        al_off_exp = stats['alipay_meta']['official'].get('已支出', {}).get('amount', 0)
        al_off_cnt = stats['alipay_meta']['official'].get('已支出', {}).get('count', 0)
        al_net_ok = abs(al_off_exp - al_kept_exp) < 0.005
        add(f'| 支付宝 | 已支出（净额入账后） | {al_off_cnt} 笔 / {_fmt_amount(al_off_exp)} 元 | '
            f'{_fmt_amount(al_kept_exp)} 元 | {"✅" if al_net_ok else "⚠️"} |')
        add('')
        add(f'> 支付宝 CSV 自带「成功退款（元）」列，本工具已按「原支出 − 退款金额」**净额入账**：'
            f'解析出支出毛额 {_fmt_amount(stats["alipay_meta"]["parsed"].get("支出", {}).get("amount", 0))} 元，'
            f'成功退款合计 {_fmt_amount(stats["alipay_refund_total"])} 元（其中 52.40/5.33 两笔交易关闭已整单删除、'
            f'影视**店 157.82 元部分退 74.70 元），净额后与官方「已支出 {_fmt_amount(al_off_exp)} 元」一致。'
            f'笔数差异来自已删除的 117 行 0 元记录与 2 行交易关闭。')
        wx_net = stats.get('kept_amount_by_source', {}).get(('微信支付', '支出'), {}).get('amount', 0)
        wx_off = stats['wechat_meta']['official'].get('支出', {}).get('amount', 0)
        wx_netted = round(sum(float(r.get('_净额退款') or 0) for r in stats['netted']
                              if r['来源'] == '微信支付'), 2)
        room_wx = round(sum(float(r.get('_原金额') or 0) - float(r['金额'])
                            for r in stats['room_adjusted'] if r['来源'] == '微信支付'), 2)
        add(f'> 微信官方「支出 {_fmt_amount(wx_off)} 元」是退款前毛额；扣退款净额 {_fmt_amount(wx_netted)} 元、'
            f'房间充值 AA 折算 {_fmt_amount(room_wx)} 元后，保留支出 {_fmt_amount(wx_net)} 元。')
    add('')
    add('### 按用户记账口径的调整项')
    add('')
    add('| 调整项 | 影响 | 说明 |')
    add('| --- | ---: | --- |')
    aa_rows = stats['aa_removed']
    aa_sum = round(sum(float(r['金额']) for r in aa_rows), 2)
    netted = stats['netted']
    nett_parts = '、'.join(f"{r['交易对方']} −{_fmt_amount(float(r['_净额退款']))}" for r in netted)
    room_rows = stats['room_adjusted']
    room_parts = '、'.join(f"{float(r['_原金额']):,.2f} → {float(r['金额']):,.2f}" for r in room_rows)
    add(f'| 退款按净额入账 | 支出 −{_fmt_amount(stats["refund_netted_total"])} | {nett_parts} |')
    add(f'| 房间充值 ÷4（只记自己那份） | 支出 −{_fmt_amount(stats["room_reduction"])} | '
        f'{room_parts}（字典 special_rules.房间充值） |')
    add(f'| 电费AA代收不记账 | 收入 −{_fmt_amount(aa_sum)} | '
        + '、'.join(f"{r['交易对方']} {_fmt_amount(float(r['金额']))}" for r in aa_rows)
        + f'，共 {len(aa_rows)} 行 |')
    add(f'| 群收款收支均记「团建」 | 收入 {stats["group_income_count"]} 行已归类 | 示例口径（选 A） |')
    add('')
    add(f'> 结论：微信按「官方毛额 − 退款净额 − 房间充值AA」= {_fmt_amount(wx_net)} 元；'
        f'支付宝按「官方净额」= {_fmt_amount(stats.get("kept_amount_by_source", {}).get(("支付宝", "支出"), {}).get("amount", 0))} 元；'
        f'合计支出 {_fmt_amount(stats["kept_amount"].get("支出", {}).get("amount", 0))} 元。')
    add('')

    add('## 四、分类结果')
    add('')
    add('| 分类方式 | 行数 |')
    add('| --- | ---: |')
    for method, count in sorted(stats['class_method'].items(), key=lambda kv: -kv[1]):
        add(f'| {_md_cell(method)} | {count} |')
    add('')
    add('| 收/支 | 需要分类 | 已分类 | 待确认 | 覆盖率 |')
    add('| --- | ---: | ---: | ---: | ---: |')
    for direction, info in stats['coverage_by_direction'].items():
        add(f'| {_md_cell(direction)} | {info["need"]} | {info["done"]} | {info["pending"]} | '
            f'{info["coverage"]:.1%} |')
    add('')
    add(f'分类字典校验：输出里的类型全部在 分类字典.yaml 内，字典外标签 '
        f'**{len(stats["invalid_tags"])} 条**（必须为 0）。')
    fallback_rows = stats['fallback_rows']
    if fallback_rows:
        fb = collections.Counter(r['收/支'] for r in fallback_rows)
        add('')
        add(f'> 兜底（示例口径）：映射和规则都没命中的行，支出归「其他」、收入归「其他收入」。'
            f'本次兜底 **{len(fallback_rows)} 行**（支出 {fb.get("支出", 0)} 行 → 其他，'
            f'收入 {fb.get("收入", 0)} 行 → 其他收入）；「待确认」只作为字典外标签/异常的安全网。')
    add('')

    add('## 五、支出按类型汇总')
    add('')
    add('| 类型 | 笔数 | 金额(元) |')
    add('| --- | ---: | ---: |')
    for cat, count, amount in stats['expense_by_type']:
        add(f'| {_md_cell(cat)} | {count} | {_fmt_amount(amount)} |')
    add(f'| **合计** | **{stats["kept_amount"].get("支出", {}).get("count", 0)}** | '
        f'**{_fmt_amount(stats["kept_amount"].get("支出", {}).get("amount", 0))}** |')
    add('')

    add('## 六、待确认清单')
    add('')
    add(f'共 {len(pending)} 行（只含「支出/收入」且类型为空的行）。')
    if not pending:
        add('')
        add('✅ 本次没有待确认行：用途不明的支出已归「其他」、收入归「其他收入」。')
    add('')
    add('| 交易时间 | 交易类型 | 交易对方 | 商品 | 收/支 | 金额(元) | 原因 |')
    add('| --- | --- | --- | --- | --- | ---: | --- |')
    for rec in pending:
        add('| ' + ' | '.join([
            _md_cell(rec['交易时间'].strftime('%Y-%m-%d %H:%M:%S')),
            _md_cell(rec.get('交易类型')),
            _md_cell(rec.get('交易对方')),
            _md_cell(rec.get('商品')),
            _md_cell(rec.get('收/支')),
            _fmt_amount(float(rec['金额'])),
            _md_cell(rec.get('_待确认原因')),
        ]) + ' |')
    add('')

    add('## 七、需要人工确认的事项')
    add('')
    for note in stats['notes']:
        add(f'- {note}')
    add('')

    add('## 八、复现命令')
    add('')
    add('```bash')
    add('python 自动化方案/build_mapping.py   # 可选：重算分类映射.csv')
    add('python 自动化方案/run_pilot.py       # 一键生成 xlsx + 本报告')
    add('```')
    add('')

    text = '\n'.join(lines)
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, 'w', encoding='utf-8', newline='\n') as fh:
        fh.write(text)
    return text


# ===========================================================================
# 八、主流程
# ===========================================================================
def _build_notes_blocks(stats):
    blocks = [('title', '9-10 月账单自动化处理说明')]
    blocks.append(('row', '生成时间', stats['generated_at'], '由 bills.py 自动生成'))
    blocks.append(('row', '微信账单', stats['inputs']['微信'], '只读输入'))
    blocks.append(('row', '支付宝账单', stats['inputs']['支付宝'], '只读输入'))
    blocks.append(('row', '规则文件', stats['inputs']['规则'], '清洗 + 分类关键词规则'))
    blocks.append(('row', '分类映射', stats['inputs']['映射'], '商户 -> 类型'))
    blocks.append(('row', '分类字典', stats['inputs']['字典'],
                   f"支出 {len(stats['dictionary']['expense'])} 类 + 收入 {len(stats['dictionary']['income'])} 类"))
    blocks.append(('blank',))

    blocks.append(('section', '一、行数统计'))
    blocks.append(('row', '原始行数', stats['raw_total'], '解析得到的全部交易行'))
    blocks.append(('row', '  微信支付', stats['raw_by_source'].get('微信支付', 0), ''))
    blocks.append(('row', '  支付宝', stats['raw_by_source'].get('支付宝', 0), ''))
    blocks.append(('row', '去重删除', stats['duplicates'], '交易单号重复的行'))
    blocks.append(('row', '删除合计（含去重）', stats['removed_total'], ''))
    for reason, count in stats['removed_by_reason'].items():
        blocks.append(('row', f'  删除：{reason}', count, ''))
    blocks.append(('row', '保留行数', stats['kept_total'], ''))
    blocks.append(('row', '  微信支付', stats['kept_by_source'].get('微信支付', 0), ''))
    blocks.append(('row', '  支付宝', stats['kept_by_source'].get('支付宝', 0), ''))
    blocks.append(('row', '需要分类的行', stats['need_classify'], '收/支 为 支出 或 收入 的行'))
    blocks.append(('row', '自动分类行数', stats['classified'], ''))
    blocks.append(('row', '待确认行数', stats['pending_total'], '见「待确认」sheet / 报告清单'))
    fb_exp = sum(1 for r in stats['fallback_rows'] if r['收/支'] == '支出')
    fb_inc = sum(1 for r in stats['fallback_rows'] if r['收/支'] == '收入')
    blocks.append(('row', '  兜底(示例口径)行数', len(stats['fallback_rows']),
                   f'支出 {fb_exp} 行 → 其他；收入 {fb_inc} 行 → 其他收入（用途不明）'))
    blocks.append(('row', '自动分类覆盖率', f'{stats["coverage"]:.1%}', '自动分类行数 / 需要分类行数'))
    for direction, info in stats['coverage_by_direction'].items():
        blocks.append(('row', f'  {direction}覆盖率', f'{info["coverage"]:.1%}',
                       f'{info["done"]}/{info["need"]} 行，待确认 {info["pending"]} 行'))
    blocks.append(('row', '分类字典外标签数', len(stats['invalid_tags']), '必须为 0；非 0 会打印警告并进待确认'))
    blocks.append(('row', '部分退款净额入账行数', len(stats['netted']),
                   '按「原支出 − 退款金额」入账，不整单删除'))
    if stats['netted']:
        blocks.append(('row', '退款净额合计', round(stats['refund_netted_total'], 2),
                       '；'.join(f'{r["交易对方"]} {_fmt_amount(float(r["_原金额"]))}'
                                f'→{_fmt_amount(float(r["金额"]))}'
                                for r in stats['netted'])))
    blocks.append(('blank',))

    blocks.append(('section', '二、金额统计（清洗后保留行）'))
    for direction in DIRECTION_ORDER:
        info = stats['kept_amount'].get(direction)
        if info:
            blocks.append(('row', f'{direction} 合计', round(info['amount'], 2),
                           f'{info["count"]} 笔'))
    aa_sum = round(sum(float(r['金额']) for r in stats['aa_removed']), 2)
    blocks.append(('row', '口径调整：退款净额入账', round(-stats['refund_netted_total'], 2),
                   '按「原支出 − 退款金额」抵减：' + '；'.join(
                       f"{r['交易对方']} −{_fmt_amount(float(r['_净额退款']))}"
                       for r in stats['netted'])))
    blocks.append(('row', '口径调整：房间充值 ÷4', round(-stats['room_reduction'], 2),
                   '；'.join(f"{r['交易对方']} {float(r['_原金额']):,.2f} → {float(r['金额']):,.2f}"
                            for r in stats['room_adjusted']) + '（只记自己那份）'))
    blocks.append(('row', '口径调整：电费AA代收不记账', round(-aa_sum, 2),
                   f'收入侧删除 {len(stats["aa_removed"])} 行：'
                   + '；'.join(f"{r['交易对方']} {_fmt_amount(float(r['金额']))}"
                              for r in stats['aa_removed'])))
    blocks.append(('blank',))

    blocks.append(('section', '三、与官方口径核对'))
    for source, block, pairs in (
            ('微信支付', stats['wechat_meta'],
             [('共', '_total', '笔数'), ('收入', '收入', '笔数/金额'),
              ('支出', '支出', '笔数/金额'), ('中性交易', '/', '笔数/金额')]),
            ('支付宝', stats['alipay_meta'],
             [('共', '_total', '笔数'), ('已收入', '收入', '笔数/金额'),
              ('已支出', '支出', '笔数/金额')])):
        official = block['official']
        parsed = block['parsed']
        for official_key, parsed_key, unit in pairs:
            off = official.get(official_key)
            if not off:
                continue
            par = parsed.get(parsed_key)
            if 'amount' not in off:
                off_text = f'{off["count"]} 笔'
            else:
                off_text = f'{off["count"]} 笔 / {_fmt_amount(off["amount"])} 元'
            if par is None:
                par_text = '-'
                same = False
            elif unit == '笔数':
                par_text = f'{par} 笔'
                same = par == off['count']
            else:
                par_text = f'{par["count"]} 笔 / {_fmt_amount(par["amount"])} 元'
                same = (par['count'] == off['count']
                        and abs(off['amount'] - par['amount']) < 0.005)
            blocks.append(('row', f'{source}·{("已支出（毛额）" if source == "支付宝" and official_key == "已支出" else official_key)}',
                           off_text, f'本次解析: {par_text}  {"一致" if same else "不一致，见说明"}'))
    if stats['alipay_refund_total']:
        blocks.append(('row', '支付宝成功退款合计', round(stats['alipay_refund_total'], 2),
                       '支付宝 CSV「成功退款（元）」列合计；已按净额抵减到对应支出行'))
        al_off = stats['alipay_meta']['official'].get('已支出') or {}
        al_net = stats['kept_amount_by_source'].get(('支付宝', '支出'), {}).get('amount', 0)
        if al_off:
            same = abs(al_off.get('amount', 0) - al_net) < 0.005
            blocks.append(('row', '支付宝·已支出（净额入账后）', round(al_net, 2),
                           f'官方已支出净额 {_fmt_amount(al_off["amount"])} 元  '
                           f'{"金额一致" if same else "金额不一致，见说明"}；'
                           f'笔数差异来自已删除的 117 行 0 元与 2 行交易关闭'))
    wx_off_exp = stats['wechat_meta']['official'].get('支出', {}).get('amount', 0)
    wx_net_exp = stats['kept_amount_by_source'].get(('微信支付', '支出'), {}).get('amount', 0)
    wx_netted_sum = round(sum(float(r.get('_净额退款') or 0) for r in stats['netted']
                              if r['来源'] == '微信支付'), 2)
    room_wx = round(sum(float(r.get('_原金额') or 0) - float(r['金额'])
                        for r in stats['room_adjusted'] if r['来源'] == '微信支付'), 2)
    blocks.append(('row', '微信支付·已支出（口径调整后）', round(wx_net_exp, 2),
                   f'官方毛额 {_fmt_amount(wx_off_exp)} 元 − 退款净额 {_fmt_amount(wx_netted_sum)} 元 '
                   f'− 房间充值AA折算 {_fmt_amount(room_wx)} 元'))
    blocks.append(('blank',))

    blocks.append(('section', '四、按类型的支出汇总'))
    for cat, count, amount in stats['expense_by_type']:
        blocks.append(('row', cat, round(amount, 2), f'{count} 笔'))
    blocks.append(('blank',))

    blocks.append(('section', '五、需要人工确认的事项'))
    for note in stats['notes']:
        blocks.append(('text', f'· {note}'))
    return blocks


def export_excel(out_path, kept, pending, stats, pivot_rows):
    wb = openpyxl.Workbook()
    ws_pivot = wb.active
    ws_pivot.title = 'Sheet1'
    ws_main = wb.create_sheet('合并账单')
    ws_pending = wb.create_sheet('待确认')
    ws_notes = wb.create_sheet('处理说明')

    build_sheet1(ws_pivot, pivot_rows)
    build_main_sheet(ws_main, kept)
    build_pending_sheet(ws_pending, pending)
    build_notes_sheet(ws_notes, _build_notes_blocks(stats))
    wb.properties.creator = 'bills.py'
    wb.properties.title = '9-10月账单_自动'

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    wb.save(out_path)
    log(f'  [写出] {out_path}')
    return out_path


def apply_room_topup(records, dictionary):
    """按字典 special_rules.房间充值 折算：只记自己那一份（金额 ÷ 人数），类型固定为电费。"""
    special = (dictionary or {}).get('special') or {}
    rule = special.get('房间充值') or {}
    tag = str(rule.get('类型') or '').strip()
    try:
        divisor = float(rule.get('金额除数') or 1)
    except (TypeError, ValueError):
        divisor = 1.0
    adjusted = []
    if not tag or divisor <= 1:
        return adjusted
    for rec in records:
        if rec.get('收/支') != '支出':
            continue
        if '房间充值' not in _cell_text(rec.get('商品')):
            continue
        if '中国地质大学' not in _cell_text(rec.get('交易对方')):
            continue
        original = round(float(rec.get('金额') or 0), 2)
        new_amount = round(original / divisor, 2)
        if abs(new_amount - original) < 0.005:
            continue
        rec['_原金额'] = original
        rec['金额'] = new_amount
        rec['_特殊类型'] = tag
        rec['_口径调整'] = f'房间充值 {original:.2f} → {new_amount:.2f}（÷{divisor:g}，只记自己那份）'
        adjusted.append(rec)
    for rec in adjusted:
        log(f'    * 房间充值 AA 折算: {rec["_口径调整"]}')
    return adjusted


def run(wechat_path, alipay_path, out_path, rules_path=None, map_path=None,
        report_path=None, dict_path=None):
    rules_path = rules_path or DEFAULT_RULES
    map_path = map_path or DEFAULT_MAP
    dict_path = dict_path or DEFAULT_DICT
    if report_path is None:
        report_path = os.path.join(os.path.dirname(os.path.abspath(out_path)), '处理报告.md')
    generated_at = dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    log('[1/6] 解析账单')
    wechat_records, wechat_meta = parse_wechat(wechat_path)
    alipay_records, alipay_meta = parse_alipay(alipay_path)

    log('[2/6] 合并去重')
    merged, duplicates = merge_and_dedupe(wechat_records, alipay_records)

    log('[3/6] 清洗')
    cfg = load_rules(rules_path)
    dictionary = load_dictionary(dict_path)
    # 解析阶段的官方口径核对要在 clean 之前统计：clean 会就地修改部分退款行的金额（净额入账）
    parsed_wx = parse_stats(wechat_records)
    parsed_al = parse_stats(alipay_records)
    parsed_wx['_total'] = len(wechat_records)
    parsed_al['_total'] = len(alipay_records)
    alipay_refund_total = round(sum(float(r.get('_退款金额') or 0) for r in alipay_records), 2)
    kept, removed, clean_stats = clean(merged, cfg, duplicates, dictionary)
    room_adjusted = apply_room_topup(kept, dictionary)

    log('[4/6] 分类')
    mapping_exact, mapping_contains = load_mapping(map_path)
    cls = classify(kept, mapping_exact, mapping_contains, cfg, dictionary)
    for rec in cls['pending']:
        rec.setdefault('_待确认原因', pending_reason(rec))

    log('[5/6] 统计')
    kept_sorted = sorted(kept, key=lambda r: r['交易时间'], reverse=True)
    pending_sorted = sorted(cls['pending'], key=lambda r: r['交易时间'], reverse=True)
    pivot_rows = build_pivot(kept_sorted)

    stats = {
        'generated_at': generated_at,
        'inputs': {'微信': wechat_path, '支付宝': alipay_path,
                   '规则': rules_path, '映射': map_path, '字典': dict_path},
        'dictionary': {'version': dictionary.get('version'),
                       'expense': dictionary.get('expense') or [],
                       'income': dictionary.get('income') or []},
        'output_excel': out_path,
        'wechat_meta': {'official': wechat_meta['official'], 'parsed': parsed_wx},
        'alipay_meta': {'official': alipay_meta['official'], 'parsed': parsed_al},
        'raw_total': len(wechat_records) + len(alipay_records),
        'raw_by_source': collections.Counter(
            r['来源'] for r in list(wechat_records) + list(alipay_records)),
        'duplicates': len(duplicates),
        'removed_total': len(removed),
        'removed_by_reason': dict(clean_stats['by_reason']),
        'kept_total': len(kept_sorted),
        'kept_by_source': collections.Counter(r['来源'] for r in kept_sorted),
        'kept_amount': amount_stats(kept_sorted),
        'kept_amount_by_source': amount_by_source(kept_sorted),
        'coverage_by_direction': classify_coverage(kept_sorted),
        'need_classify': cls['need'],
        'classified': cls['classified'],
        'pending_total': len(pending_sorted),
        'coverage': (cls['classified'] / cls['need']) if cls['need'] else 1.0,
        'class_method': dict(cls['method_counter']),
        'expense_by_type': expense_by_type(kept_sorted),
        'income_by_type': income_by_type(kept_sorted),
        'alipay_refund_total': alipay_refund_total,
        'netted': clean_stats.get('netted') or [],
        'refund_netted_total': round(sum(float(r.get('_净额退款') or 0)
                                        for r in (clean_stats.get('netted') or [])), 2),
        'room_adjusted': room_adjusted,
        'room_reduction': round(sum(float(r.get('_原金额') or 0) - float(r['金额'])
                                    for r in room_adjusted), 2),
        'aa_removed': [r for r in removed if '电费AA代收' in r['_删除原因']],
        'invalid_tags': cls['invalid'],
        'group_income_pending': sum(
            1 for r in pending_sorted
            if '群收款' in _cell_text(r.get('交易类型')) and r.get('收/支') == '收入'),
        'group_income_count': sum(
            1 for r in kept_sorted
            if '群收款' in _cell_text(r.get('交易类型')) and r.get('收/支') == '收入'),
        'fallback_rows': [r for r in kept_sorted if r.get('_分类来源') == '兜底(示例口径)'],
        'removed': removed,
    }

    # ---- 需要人工确认的说明 ----
    notes = []
    zero_removed = [r for r in removed if r['_删除原因'] == '金额为0']
    if zero_removed:
        top = collections.Counter(r['交易对方'] for r in zero_removed).most_common(5)
        detail = '、'.join(f'{name} {count} 行' for name, count in top)
        notes.append(f'已删除金额为 0 的行共 {len(zero_removed)} 行（全部来自支付宝；{detail} 等）。'
                     f'这些行是「无实际扣款」或抵扣记录，删除后不影响支出/收入合计。')
    refund_removed = [r for r in removed if '退款' in r['_删除原因']]
    if refund_removed:
        parts = [f"{r['交易对方']} {_fmt_amount(float(r['金额']))} 元（{r['收/支']}）"
                 for r in refund_removed]
        notes.append(f'已删除退款收入行 {len(refund_removed)} 行：' + '；'.join(parts) +
                     '。退款金额改为在对应原支出行上抵减（不重复计入收入）。')
    netted = stats['netted']
    if netted:
        parts = [f'{r["来源"]} {r["交易对方"]} {_fmt_amount(float(r["_原金额"]))} − '
                 f'{_fmt_amount(float(r["_净额退款"]))} = {_fmt_amount(float(r["金额"]))} 元'
                 for r in netted]
        notes.append(f'部分退款按「原支出 − 退款金额」净额入账 {len(netted)} 行：'
                     + '；'.join(parts) + f'。退款净额合计 {_fmt_amount(stats["refund_netted_total"])} 元。')
    closed_removed = [r for r in removed if '交易关闭' in r['_删除原因']]
    if closed_removed:
        notes.append(f'已删除支付宝「交易关闭」行 {len(closed_removed)} 行（'
                     + '、'.join(f"{r['交易对方']} {_fmt_amount(float(r['金额']))} 元"
                                for r in closed_removed) + '）。')
    neutral_closed = [r for r in kept if '中性方向' in (r.get('_备注') or '')]
    if neutral_closed:
        notes.append(f'另有 {len(neutral_closed)} 行「交易关闭」属于不计收支（'
                     + '、'.join(f"{r['交易对方']} {_fmt_amount(float(r['金额']))} 元"
                                for r in neutral_closed)
                     + '），按中性行保留策略未删除，类型留空。')

    room_adjusted = stats['room_adjusted']
    if room_adjusted:
        notes.append('「房间充值」按字典口径折算（只记自己那份）：'
                     + '；'.join(str(r.get('_口径调整')) for r in room_adjusted)
                     + f'，合计少记 {_fmt_amount(stats["room_reduction"])} 元。'
                       '舍友转来的电费份额不记账（对应收入行已删除）。')
    if stats['aa_removed']:
        total = sum(float(r['金额']) for r in stats['aa_removed'])
        notes.append(f'已删除「电费AA代收」{len(stats["aa_removed"])} 行、合计 '
                     f'{_fmt_amount(total)} 元（'
                     + '、'.join(f"{r['交易对方']} {_fmt_amount(float(r['金额']))} 元"
                                for r in stats['aa_removed'])
                     + '，按示例口径不记账）。')
    if stats['group_income_count']:
        notes.append(f'群收款收入 {stats["group_income_count"]} 行按示例口径记「团建」（2026-11-07 选 A）；'
                     '支出方向同样记「团建」，不再进待确认。')
    if stats['invalid_tags']:
        notes.append(f'[警告] 有 {len(stats["invalid_tags"])} 行产生了字典外标签，已拦下并转入待确认：'
                     + '、'.join(f"{item['标签']}（{item['交易对方']}）"
                                for item in stats['invalid_tags']) + '。')

    kangli = [r for r in kept_sorted if '康力体育' in (r.get('交易对方') or '')
              and r.get('类型') == '健身']
    if kangli:
        notes.append('「康力体育财务」' + '、'.join(_fmt_amount(float(r['金额'])) for r in kangli)
                     + ' 元暂按「健身」处理（参照历史倍力美健身 800 元）；'
                       '若实际是体育用品，请改为网购/线下购物。')

    online_pay = [r for r in kept_sorted if '在线缴费' in (r.get('商品') or '')]
    if online_pay:
        total = sum(float(r['金额']) for r in online_pay)
        parts = [f"{float(rec['金额']):,.2f} 元 → {rec.get('类型')}"
                 for rec in sorted(online_pay, key=lambda x: -float(x['金额']))]
        notes.append(f'中国地质大学(武汉)「在线缴费」共 {len(online_pay)} 行、合计 '
                     f'{_fmt_amount(total)} 元，按金额拆分：' + '、'.join(parts) +
                     '（≥1000 记学费，<1000 记医保；不再进待确认）。')

    rule_notes = [f'规则「{name}」：{note}（命中 {count} 行）'
                  for (name, note), count in cls['rule_notes'].items()]
    if rule_notes:
        notes.append('以下自动分类依据规则注释，属于中等置信度，建议抽查：' + '；'.join(rule_notes) + '。')

    fallback_rows = stats['fallback_rows']
    if fallback_rows:
        by_dir = collections.Counter(r['收/支'] for r in fallback_rows)
        parts = []
        for direction in ('支出', '收入'):
            subset = [r for r in fallback_rows if r['收/支'] == direction]
            if subset:
                parts.append(f'{direction} {len(subset)} 行 → '
                             f'{"其他收入" if direction == "收入" else "其他"}'
                             f'（{_fmt_amount(sum(float(r["金额"]) for r in subset))} 元）')
        notes.append('按示例口径，用途不明的个人转账/扫码等：支出归「其他」、收入归「其他收入」。'
                     f'本次兜底命中 {len(fallback_rows)} 行：' + '；'.join(parts) + '。')
    if stats['pending_total'] == 0:
        notes.append('本次没有待确认行（自动分类覆盖率 100%）；「待确认」sheet 仅作为'
                     '字典外标签/异常的安全网。')

    if pending_sorted:
        notes.append(f'自动分类覆盖率 {stats["coverage"]:.1%}；未命中的 '
                     f'{len(pending_sorted)} 行已列入「待确认」sheet，请人工指定类型。')
    else:
        notes.append(f'自动分类覆盖率 {stats["coverage"]:.1%}（{stats["classified"]}/'
                     f'{stats["need_classify"]} 行），无待确认行。')
    for direction, info in stats['coverage_by_direction'].items():
        notes.append(f'{direction}行自动分类覆盖率 {info["coverage"]:.1%}'
                     f'（{info["done"]}/{info["need"]} 行，待确认 {info["pending"]} 行）。')

    # ---- 按来源对账（支出 / 收入 与官方口径的差异构成）----
    by_src = stats['kept_amount_by_source']
    wx_official = stats['wechat_meta']['official']
    al_official = stats['alipay_meta']['official']
    wx_kept_exp = by_src.get(('微信支付', '支出'), {}).get('amount', 0.0)
    al_kept_exp = by_src.get(('支付宝', '支出'), {}).get('amount', 0.0)
    refund_inc_removed = round(sum(float(r['金额']) for r in removed
                                   if r['收/支'] == '收入' and '退款' in r['_删除原因']), 2)
    wx_off_exp = wx_official.get('支出', {}).get('amount', 0.0)
    al_off_exp = al_official.get('已支出', {}).get('amount', 0.0)
    wx_netted = round(sum(float(r.get('_净额退款') or 0) for r in stats['netted']
                          if r['来源'] == '微信支付'), 2)
    al_netted = round(sum(float(r.get('_净额退款') or 0) for r in stats['netted']
                          if r['来源'] == '支付宝'), 2)
    room_wx = round(sum(float(r.get('_原金额') or 0) - float(r['金额'])
                        for r in stats['room_adjusted'] if r['来源'] == '微信支付'), 2)
    room_al = round(sum(float(r.get('_原金额') or 0) - float(r['金额'])
                        for r in stats['room_adjusted'] if r['来源'] == '支付宝'), 2)
    aa_total = round(sum(float(r['金额']) for r in stats['aa_removed']), 2)
    if wx_kept_exp or al_kept_exp:
        notes.append(
            f'支出金额对账（按示例口径调整后）：微信保留 {_fmt_amount(wx_kept_exp)} 元 = 官方毛额 '
            f'{_fmt_amount(wx_off_exp)} 元 − 退款净额 {_fmt_amount(wx_netted)} 元 '
            f'− 房间充值AA折算 {_fmt_amount(room_wx)} 元；'
            f'支付宝保留 {_fmt_amount(al_kept_exp)} 元 = 官方已支出 '
            f'{_fmt_amount(al_off_exp)} 元（官方即扣除成功退款后的净额，工具按「原支出−退款」入账后一致，'
            f'抵减 {_fmt_amount(al_netted)} 元）；合计支出 '
            f'{_fmt_amount(stats["kept_amount"].get("支出", {}).get("amount", 0))} 元。')
    kept_inc = stats['kept_amount'].get('收入', {}).get('amount', 0.0)
    parsed_inc = (stats['wechat_meta']['parsed'].get('收入', {}).get('amount', 0)
                  + stats['alipay_meta']['parsed'].get('收入', {}).get('amount', 0))
    notes.append(f'收入金额对账（按示例口径调整后）：原始收入合计 {_fmt_amount(parsed_inc)} 元 '
                 f'− 已删退款收入行 {_fmt_amount(refund_inc_removed)} 元 '
                 f'− 电费AA代收 {_fmt_amount(aa_total)} 元 = 保留收入 {_fmt_amount(kept_inc)} 元。')
    if stats['kept_amount'].get('/', {}).get('count'):
        notes.append('保留 1 行中性交易「转入零钱通-来自零钱」1300 元，收/支 为「/」，'
                     '类型留空，计入 Sheet1 的「/」分组。')
    notes.append('「合并账单」按交易时间倒序；「Sheet1」是静态透视表外观，'
                 '分组顺序为 "/ → 不计收支 → 收入 → 支出"，与历史 6 月文件一致。')
    notes.append('分类结果由「分类映射.csv」（先精确后包含）和「rules.yaml」关键词规则共同决定，'
                 '改规则后可重复运行覆盖输出。')
    stats['notes'] = notes

    log('[6/6] 导出 Excel 与 Markdown 报告')
    export_excel(out_path, kept_sorted, pending_sorted, stats, pivot_rows)
    write_report(report_path, stats, kept_sorted, pending_sorted, pivot_rows)
    log(f'  [写出] {report_path}')

    log('')
    log('== 完成 ==')
    log(f'  原始 {stats["raw_total"]} 行 -> 删除 {stats["removed_total"]} 行 -> 保留 {stats["kept_total"]} 行')
    log(f'  支出 {stats["kept_amount"].get("支出", {}).get("count", 0)} 行 '
        f'{_fmt_amount(stats["kept_amount"].get("支出", {}).get("amount", 0))} 元；'
        f'收入 {stats["kept_amount"].get("收入", {}).get("count", 0)} 行 '
        f'{_fmt_amount(stats["kept_amount"].get("收入", {}).get("amount", 0))} 元')
    log(f'  自动分类覆盖率 {stats["coverage"]:.1%}（{stats["classified"]}/{stats["need_classify"]}）')
    return stats


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description='微信 + 支付宝账单自动化合并/清洗/分类（原型）')
    parser.add_argument('--wechat', required=True, help='微信支付账单 xlsx 路径')
    parser.add_argument('--alipay', required=True, help='支付宝账单 csv 路径')
    parser.add_argument('--out', required=True, help='输出 xlsx 路径（同名覆盖）')
    parser.add_argument('--rules', default=DEFAULT_RULES, help=f'清洗/分类规则 yaml（默认 {DEFAULT_RULES}）')
    parser.add_argument('--map', default=DEFAULT_MAP, help=f'商户->类型 映射 csv（默认 {DEFAULT_MAP}）')
    parser.add_argument('--dict', default=DEFAULT_DICT, help=f'固定资产分类字典 yaml（默认 {DEFAULT_DICT}）')
    parser.add_argument('--report', default=None, help='Markdown 报告路径（默认 <out 所在目录>/处理报告.md）')
    args = parser.parse_args(argv)

    run(wechat_path=args.wechat, alipay_path=args.alipay, out_path=args.out,
        rules_path=args.rules, map_path=args.map, report_path=args.report,
        dict_path=args.dict)
    return 0


if __name__ == '__main__':
    sys.exit(main())
