# -*- coding: utf-8 -*-
"""从历史已标注账单中统计「商户 -> 类型」映射，生成 分类映射.csv 初版，并补充 9-10 月新商户。

用法:
    python 自动化方案/build_mapping.py                # 重新生成 自动化方案/分类映射.csv
    python 自动化方案/build_mapping.py --dry-run      # 只打印，不写文件

历史数据来源（只读，绝不修改）:
    3月.xlsx / 26accounte/合并账单_20260401_20260531.xlsx / 6月账单.xlsx / 7、8月账单.xlsx

生成的 CSV 只有两列: 商户匹配,类型
类型必须是 分类字典.yaml 里的 expense / income 标签；历史旧标签先按字典 aliases 统一。
bills.py 的匹配顺序是: 先精确匹配商户，再按「商户匹配」做包含匹配（长者优先）。
因此这里只收录「仅凭交易对方就能确定类型」的商户；像 中国地质大学（商品决定 食堂/电费）、
美团、淘宝闪购（商品决定 外卖/网购）、哈啰出行 这类需要上下文判断的商户一律不进 CSV，
交给 rules.yaml 的关键词规则处理。
"""
from __future__ import annotations

import argparse
import collections
import csv
import io
import os
import sys

import openpyxl
import yaml

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
OUT_CSV = os.path.join(HERE, '分类映射.csv')
DICT_PATH = os.path.join(HERE, '分类字典.yaml')

# ---------------------------------------------------------------------------
# 历史账单（人工标注过的）——只读
# ---------------------------------------------------------------------------
HISTORY = [
    ('3月', '3月.xlsx', 'Sheet1'),
    ('4-5月', '26accounte/合并账单_20260401_20260531.xlsx', '合并账单'),
    ('6月', '6月账单.xlsx', '合并账单'),
    ('7-8月', '7、8月账单.xlsx', '合并账单'),
]

# 这些类型带有个人/场合语义，不能仅凭商户名推广，禁止进入自动映射
# （电费AA 是收入代收，按字典别名会落到「电费」，但方向是收入，绝对不能写进映射）
PERSONAL_CATEGORIES = {
    '电费AA', '团建', '还钱', 'SSL分红', '新加坡姚补', '天协还款', '出书', '团费',
    '买自行车', '出差已报销',
}

# 需要上下文（商品/收支方向/金额）才能确定的商户，禁止凭商户名进入 CSV
EXCLUDE_MERCHANTS = {
    '中国地质大学', '中国地质大学(武汉)', '中国地质大学（武汉）',
    '淘宝闪购', '美团', '哈啰出行',
}

# 白名单式例外：这些商户虽然命中上面的排除逻辑，但仍然可以确定
INCLUDE_OVERRIDE = {
    '妈妈 (美玉(相馆))': '生活费',
    '美玉(相馆)': '生活费',
    '妈妈': '生活费',
    '爸爸 (宝玉照相馆)': '生活费',
    '深圳市腾讯计算机系统有限公司': '娱乐',   # 历史全是绿钻/QQ音乐续费 -> 娱乐
}

# 9-10 月出现、历史没有（或历史样本不足）的新商户，人工补充
# 分类按 分类字典.yaml 收敛：线下超市/便利店/个体户门店 -> 线下购物；
# 线上学习资料 -> 网购；食堂只保留中国地质大学校园卡，其他饭店 -> 餐饮。
CURATED = [
    # 奶茶咖啡
    ('蜜雪冰城', '奶茶咖啡'), ('益禾堂', '奶茶咖啡'), ('茶百道', '奶茶咖啡'),
    ('茉莉奶白', '奶茶咖啡'), ('古茗', '奶茶咖啡'), ('沪上阿姨', '奶茶咖啡'),
    ('霸王茶姬', '奶茶咖啡'),
    # 洗衣机 / 热水 / 饮水机
    ('U净', '洗衣机'), ('慧生活798', '热水'), ('趣智校园', '饮水机'),
    # 学习（打印复印 / 考试报名费）
    ('窝趣云打印', '学习'), ('长江图文', '学习'), ('教育部教育考试院', '学习'),
    # 交通（打车 / 共享单车 / 火车）
    ('哈啰出行', '交通'), ('滴滴出行', '交通'), ('百度平台商家', '交通'),
    ('12306', '交通'), ('中铁网络', '交通'),
    # 外卖
    ('肯德基', '外卖'), ('塔斯汀', '外卖'), ('麦当劳', '外卖'), ('华莱士', '外卖'),
    # 网购 / 电商（线上学习资料也算网购）
    ('拼多多平台商户', '网购'), ('拼多多', '网购'), ('京东', '网购'),
    ('小红书', '网购'), ('浙江天猫技术有限公司', '网购'),
    ('杭州闲鱼信息技术有限公司', '网购'), ('天猫', '网购'),
    ('华研**店', '网购'), ('学丞**店', '网购'),
    # 线下购物（超市 / 便利店 / 零食店 / 个体户门店）
    ('零食很忙', '线下购物'), ('好想来', '线下购物'), ('秭美超市美佳天问店', '线下购物'),
    ('社区便利店', '线下购物'), ('711烟酒便利', '线下购物'), ('711湖北', '线下购物'),
    ('新佳宜', '线下购物'), ('太古可口可乐', '线下购物'), ('可口可乐', '线下购物'),
    ('商户_敦永松', '线下购物'), ('个体户史双英', '线下购物'), ('个体户杜银华', '线下购物'),
    ('华汇壹品', '线下购物'),
    # 水果（水果店单列，不算线下购物）
    ('武汉市洪山区远朋水果店', '水果'), ('新鲜水果地大西区店', '水果'),
    # 餐饮（食堂只留给中国地质大学校园卡）
    ('汤师傅快餐（地大店）', '餐饮'), ('武汉市洪山区朱翔餐饮店（个体工商户）', '餐饮'),
    ('小麦姐早餐', '餐饮'), ('鄧記鍋盔', '餐饮'), ('沙县小吃', '餐饮'),
    ('一米噜日式咖喱蛋包饭', '餐饮'), ('茶港武汉大悦城', '餐饮'),
    # 娱乐（音乐/视频会员、AI 订阅）
    ('哔哩哔哩弹幕网', '娱乐'), ('Valve', '娱乐'), ('小黑盒-清枫科技', '娱乐'),
    ('杭州深度求索', '娱乐'),
    # 医疗 / 健身
    ('湖北省中医院（湖北中医药大学附属医院、湖北省中医药研究院）', '医疗'),
    ('倍力美健身', '健身'), ('康力体育财务', '健身'),
    # 话费
    ('中国电信', '话费'), ('中国移动', '话费'), ('中国联通', '话费'),
]


def load_dictionary(path: str = DICT_PATH) -> dict:
    """加载 分类字典.yaml：aliases + 允许的标签集合。"""
    if not os.path.exists(path):
        raise FileNotFoundError(f'分类字典不存在: {path}')
    with open(path, 'r', encoding='utf-8') as fh:
        raw = yaml.safe_load(fh) or {}
    expense = list((raw.get('expense') or {}).keys())
    income = list((raw.get('income') or {}).keys())
    return {
        'aliases': dict(raw.get('aliases') or {}),
        'expense': expense,
        'income': income,
        'allowed': set(expense) | set(income),
    }


def alias_tag(cat: str, aliases: dict) -> str:
    """按字典 aliases 链式转成新标签。"""
    cat = (cat or '').strip()
    seen = set()
    while cat in aliases and cat not in seen:
        seen.add(cat)
        cat = str(aliases[cat]).strip()
    return cat


def amount_of(rec: dict) -> float:
    for key in ('金额', '金额(元)', '金额（元）'):
        if key in rec:
            try:
                return round(float(rec[key] or 0), 2)
            except (TypeError, ValueError):
                return 0.0
    return 0.0


def get(rec: dict, *names: str) -> str:
    for name in names:
        if name in rec and rec[name] is not None:
            return str(rec[name]).strip()
    return ''


def load_sheet(path: str, sheet: str) -> list[dict]:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    header = [str(c).strip() if c is not None else '' for c in rows[0]]
    out = []
    for row in rows[1:]:
        if not any(v is not None for v in row):
            continue
        out.append(dict(zip(header, row)))
    return out


def collect_history() -> tuple[dict[str, collections.Counter], int]:
    stat: dict[str, collections.Counter] = collections.defaultdict(collections.Counter)
    total = 0
    for label, path, sheet in HISTORY:
        full = os.path.join(ROOT, path)
        if not os.path.exists(full):
            print(f'[警告] 历史文件不存在，跳过: {path}')
            continue
        recs = load_sheet(full, sheet)
        total += len(recs)
        for rec in recs:
            merchant = get(rec, '交易对方')
            cat = get(rec, '类型', '分类')
            if not merchant or not cat:
                continue
            stat[merchant][cat] += 1
        print(f'[历史] {label}: {len(recs)} 行 ({path})')
    return stat, total


def mine(stat: dict[str, collections.Counter], dictionary: dict,
         min_share: float = 0.9, min_count: int = 2) -> tuple[dict[str, str], list]:
    """历史统计 -> 商户映射（高置信度 + 按字典别名收敛，只留字典内标签）。"""
    aliases = dictionary['aliases']
    allowed = dictionary['allowed']
    mined: dict[str, str] = {}
    dropped = []
    for merchant, counter in stat.items():
        n = sum(counter.values())
        raw_cat, top = counter.most_common(1)[0]
        share = top / n
        if merchant in INCLUDE_OVERRIDE:
            cat = alias_tag(INCLUDE_OVERRIDE[merchant], aliases)
            mined[merchant] = cat
            continue
        if merchant in EXCLUDE_MERCHANTS:
            dropped.append((merchant, raw_cat, n, share, '需上下文判断'))
            continue
        if raw_cat in PERSONAL_CATEGORIES:
            why = '收入代收（电费AA，方向不符，不写支出映射）' if raw_cat == '电费AA' else '个人语义类型'
            dropped.append((merchant, raw_cat, n, share, why))
            continue
        if share < min_share:
            dropped.append((merchant, raw_cat, n, share, f'置信度<{min_share}'))
            continue
        if n < min_count:
            dropped.append((merchant, raw_cat, n, share, f'样本<{min_count}'))
            continue
        cat = alias_tag(raw_cat, aliases)
        if cat not in allowed:
            dropped.append((merchant, raw_cat, n, share, f'字典外标签 {cat!r}'))
            continue
        mined[merchant] = cat
    return mined, dropped


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description='生成 分类映射.csv（历史统计 + 9-10 月补充）')
    parser.add_argument('--out', default=OUT_CSV, help='输出 CSV 路径')
    parser.add_argument('--dict', default=DICT_PATH, help='分类字典 yaml 路径')
    parser.add_argument('--dry-run', action='store_true', help='只打印结果，不写文件')
    parser.add_argument('--min-share', type=float, default=0.9)
    parser.add_argument('--min-count', type=int, default=2)
    args = parser.parse_args(argv)

    dictionary = load_dictionary(args.dict)
    print(f"[字典] 支出 {len(dictionary['expense'])} 类 / 收入 {len(dictionary['income'])} 类 / "
          f"aliases {len(dictionary['aliases'])} 条")

    stat, total_rows = collect_history()
    print(f'[历史] 合计 {total_rows} 行, {len(stat)} 个已标注商户')

    mined, dropped = mine(stat, dictionary, args.min_share, args.min_count)
    print(f'[挖掘] 历史高置信度映射 {len(mined)} 条；忽略 {len(dropped)} 条')
    for merchant, cat, n, share, why in sorted(dropped, key=lambda x: -x[2])[:20]:
        print(f'         忽略 {merchant!r} ({cat}, n={n}, share={share:.0%}) <- {why}')

    mapping: dict[str, str] = dict(mined)
    for merchant, cat in CURATED:
        cat = alias_tag(cat, dictionary['aliases'])
        if cat not in dictionary['allowed']:
            print(f'[警告] 补充映射标签不在字典内，跳过: {merchant!r} -> {cat!r}')
            continue
        if merchant in mapping and mapping[merchant] != cat:
            print(f'[补充] 覆盖历史映射: {merchant!r} {mapping[merchant]} -> {cat}')
        mapping[merchant] = cat

    # 去重 + 稳定排序（商户名长度降序，保证包含匹配时长者优先）
    rows = sorted(mapping.items(), key=lambda kv: (-len(kv[0]), kv[0]))
    text = io.StringIO()
    writer = csv.writer(text, lineterminator='\n')
    writer.writerow(['商户匹配', '类型'])
    writer.writerows(rows)
    content = text.getvalue()

    print(f'[结果] 共 {len(rows)} 条映射，标签分布: '
          f'{dict(collections.Counter(c for _, c in rows).most_common())}')
    bad = {c for _, c in rows if c not in dictionary['allowed']}
    print(f"[校验] 字典外标签: {bad if bad else '无'}")
    print('-' * 70)
    print(content.rstrip())
    print('-' * 70)

    if args.dry_run:
        print('[dry-run] 未写文件')
        return 0
    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    with open(args.out, 'w', encoding='utf-8', newline='') as fh:
        fh.write(content)
    print(f'[写出] {args.out} ({len(content.encode("utf-8"))} bytes)')
    return 0


if __name__ == '__main__':
    sys.exit(main())
