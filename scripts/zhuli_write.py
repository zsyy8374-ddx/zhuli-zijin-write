#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
主力资金写入通达信自定义数据序列
  ID=201 → 主力净额（日期, 数值）
  ID=202 → 主买净额（日期, 数值）

用法：
  python3 zhuli_write.py <excel文件路径> [--date YYYYMMDD] [--dry-run]

  --date YYYYMMDD：文件里没有日期列时，统一用这个日期填所有记录；
                   不传 --date 且文件无日期列时，会交互式提示输入日期。

  日期确定优先级：--date 参数 > 文件名中的 8 位日期 > 文件里的「日期」列 > 交互式询问。

支持两种输入（自动识别）：
  1. 通达信导出的「.xls」——实为 Tab 分隔文本（GBK）
  2. 真正的 Excel「.xlsm / .xlsx」——zip 结构

写入结构（参考 ID=101 竞价最大占比）：
  每条 8 字节 = 日期(uint32 小端, 十进制 yyyymmdd) + 数值(float32 小端)
  同日替换、按日期升序；已存在文件合并（不删其他日期），不存在则创建。
  覆盖语义：本次导出里的日期会【覆盖】旧数据 —— 不在本次导出里的股票，其当日旧记录会被清除。

文件命名：{市场标志}_{代码}.dat
  市场标志：6/8 开头→1（沪/科创/板块指数）；0/3 开头→0（深/创业）；9 开头→2（北交所）
"""
import sys
import os
import re
import struct
import zipfile
from collections import defaultdict
from xml.etree import ElementTree as ET

IS_WINDOWS = (os.name == 'nt')

if IS_WINDOWS:
    SIG = r'D:\GP\通达信金融终端(开心果交易版)V2026\T0002\signals'
    DEFAULT_DIR = r'D:\GP\通达信金融终端(开心果交易版)V2026\T0002\export'
    SH_INDEX_DAY = r'D:\GP\通达信金融终端(开心果交易版)V2026\vipdoc\sh\lday\sh000001.day'
else:
    SIG = '/mnt/d/GP/通达信金融终端(开心果交易版)V2026/T0002/signals'
    DEFAULT_DIR = '/mnt/d/GP/通达信金融终端(开心果交易版)V2026/T0002/export'
    SH_INDEX_DAY = '/mnt/d/GP/通达信金融终端(开心果交易版)V2026/vipdoc/sh/lday/sh000001.day'

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

FIELD_CODE = '代码'
FIELD_ZL = '主力净额'
FIELD_ZM = '主买净额'
FIELD_DT = '日期'

# 默认输入文件夹（董哥 2026-09-06 指定）
# 交易日判断用：上证指数日线（有记录=交易日）


def market_flag(code: str) -> int:
    return {'6': 1, '8': 1, '0': 0, '3': 0, '9': 2}.get(code[0], -1)


def _num(x):
    x = (x or '').strip()
    try:
        return float(x)
    except (ValueError, TypeError):
        return None


def extract_date_from_filename(path):
    """从文件名提取 8 位日期（yyyyMMdd 或 yyyy-MM-dd 等），无则返回 None。"""
    base = os.path.basename(path)
    m = re.search(r'(20\d{2})[-_./]?(\d{2})[-_./]?(\d{2})', base)
    if not m:
        return None
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if 1 <= mo <= 12 and 1 <= d <= 31:
        return y * 10000 + mo * 100 + d
    return None


def to_wsl_path(p):
    """Windows 盘符路径 D:/... 或 D:\\... -> /mnt/d/...（仅 WSL 下转换；Windows 下原样返回）。"""
    if IS_WINDOWS:
        return p
    m = re.match(r'^([A-Za-z]):[\\/](.*)$', p)
    if m:
        return '/mnt/' + m.group(1).lower() + '/' + m.group(2).replace('\\', '/')
    return p


def is_trading_day(date_int):
    """判断 date_int(yyyymmdd) 是否交易日：查上证指数日线当天有无记录。"""
    try:
        raw = open(SH_INDEX_DAY, 'rb').read()
    except OSError:
        return True  # 无法判断时放行
    for i in range(0, len(raw) // 32 * 32, 32):
        if struct.unpack('<I', raw[i:i + 4])[0] == date_int:
            return True
    return False


# ---------- 解析 .xlsx / .xlsm ----------
def load_xlsx(path):
    z = zipfile.ZipFile(path)
    ss = []
    try:
        root = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in root.findall(NS + 'si'):
            ss.append(''.join(t.text or '' for t in si.iter(NS + 't')))
    except KeyError:
        pass
    root = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
    rows = []
    for row in root.iter(NS + 'row'):
        cells = {}
        for c in row.findall(NS + 'c'):
            ref = c.get('r') or ''
            m = re.match(r'([A-Z]+)', ref)
            if not m:
                continue
            col = m.group(1)
            t = c.get('t')
            v = c.find(NS + 'v')
            if v is None:
                val = ''
            elif t == 's':
                idx = int(v.text)
                val = ss[idx] if 0 <= idx < len(ss) else ''
            else:
                val = v.text
            cells[col] = val
        rows.append(cells)
    # 第 1 行是表头：列字母 -> 列名
    header = {letter: (name or '').strip() for letter, name in rows[0].items()}
    data = []
    for cells in rows[1:]:
        data.append({name: cells.get(letter, '') for letter, name in header.items()})
    return data


# ---------- 解析 .xls（Tab 分隔文本）----------
def load_tsv(path):
    raw = open(path, 'rb').read()
    # 编码自动识别：UTF-8 优先，失败回退 GBK/GB18030
    text = None
    for enc in ('utf-8-sig', 'utf-8', 'gb18030', 'gbk'):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        text = raw.decode('gbk', errors='replace')
    if '\r\n' in text:
        lines = text.split('\r\n')
    else:
        lines = text.split('\n')
    header = [h.strip() for h in lines[0].split('\t')]
    data = []
    for ln in lines[1:]:
        if not ln.strip():
            continue
        f = ln.split('\t')
        row = {}
        for i, name in enumerate(header):
            row[name] = f[i] if i < len(f) else ''
        data.append(row)
    return data


def load_rows(path):
    with open(path, 'rb') as f:
        head = f.read(4)
    if head[:2] == b'PK':
        return load_xlsx(path)
    return load_tsv(path)


def build(data, default_date=None):
    """data: list[dict{列名:值}] -> (dict201, dict202)  key=市场标志_代码 -> {日期int: 数值float}
    default_date: 文件无日期列时，所有记录统一填这个日期。"""
    d201 = defaultdict(dict)
    d202 = defaultdict(dict)
    for row in data:
        code = (row.get(FIELD_CODE) or '').strip()
        # 清理代码：去掉 =、" 等非数字（Excel 导出 "="920289"" 包裹保前导0）
        code = re.sub(r'\D', '', code)
        if not code:
            continue
        code = code[-6:].zfill(6)  # 取末 6 位并补前导 0
        fl = market_flag(code)
        if fl < 0:
            continue
        if default_date is not None:
            dt = default_date
        else:
            dt_s = (row.get(FIELD_DT) or '').strip()
            if not dt_s.isdigit():
                continue
            dt = int(dt_s)
        key = f'{fl}_{code}'
        v7 = _num(row.get(FIELD_ZL))
        v10 = _num(row.get(FIELD_ZM))
        if v7 is not None:
            d201[key][dt] = v7
        if v10 is not None:
            d202[key][dt] = v10
    return d201, d202


def write_files(data, subdir, dry_run=False):
    d = os.path.join(SIG, subdir)
    os.makedirs(d, exist_ok=True)

    # 本次要写的日期（作为“覆盖/清除”目标）
    clear_dates = set()
    for recs in data.values():
        clear_dates.update(recs.keys())

    if dry_run:
        # 预览模式不读文件、不写盘，只统计
        written = created = appended = 0
        for key in data.keys():
            if os.path.exists(os.path.join(d, key + '.dat')):
                appended += 1
            else:
                created += 1
            written += 1
        return written, created, appended, len(clear_dates), 0, 0

    new_keys = set(data.keys())
    cleared = 0
    removed_empty = 0

    # 第一步：覆盖当天数据 —— 对【不在本次导出里的股票】的文件，删除这些日期记录
    # （在本次导出里的股票，下面写盘时 merged.update 会直接覆盖同日）
    if clear_dates:
        for fname in os.listdir(d):
            if not fname.endswith('.dat'):
                continue
            key = fname[:-4]
            if key in new_keys:
                continue
            path = os.path.join(d, fname)
            with open(path, 'rb') as f:
                raw = f.read()
            recs = {}
            for i in range(0, len(raw) // 8 * 8, 8):
                dd, vv = struct.unpack('<If', raw[i:i + 8])
                recs[dd] = vv
            keep = {dd: vv for dd, vv in recs.items() if dd not in clear_dates}
            if len(keep) != len(recs):
                cleared += 1
                if keep:
                    with open(path, 'wb') as f:
                        for dd in sorted(keep):
                            f.write(struct.pack('<If', dd, keep[dd]))
                else:
                    os.remove(path)
                    removed_empty += 1

    # 第二步：写入新数据（同日替换）
    written = created = appended = 0
    for key, recs in data.items():
        path = os.path.join(d, key + '.dat')
        merged = {}
        if os.path.exists(path):
            with open(path, 'rb') as f:
                raw = f.read()
            for i in range(0, len(raw) // 8 * 8, 8):
                dd, vv = struct.unpack('<If', raw[i:i + 8])
                merged[dd] = vv
            appended += 1
        else:
            created += 1
        merged.update(recs)  # 同日替换，旧日期保留
        with open(path, 'wb') as f:
            for dd in sorted(merged):
                f.write(struct.pack('<If', dd, merged[dd]))
        written += 1
    return written, created, appended, len(clear_dates), cleared, removed_empty


def main():
    args = [a for a in sys.argv[1:]]
    dry_run = '--dry-run' in args
    args = [a for a in args if a != '--dry-run']

    # 解析 --date YYYYMMDD（文件无日期列时用）
    date_arg = None
    rest = []
    i = 0
    while i < len(args):
        a = args[i]
        if a == '--date' and i + 1 < len(args):
            date_arg = args[i + 1]
            i += 2
            continue
        if a.startswith('--date='):
            date_arg = a.split('=', 1)[1]
            i += 1
            continue
        rest.append(a)
        i += 1
    args = rest

    if not args:
        print('用法: python3 zhuli_write.py <excel文件路径> [--date YYYYMMDD] [--dry-run]')
        return 2
    path = args[0]
    path = to_wsl_path(path)
    if not os.path.isabs(path):
        path = os.path.join(DEFAULT_DIR, os.path.basename(path))
    if not os.path.exists(path):
        print(f'文件不存在: {path}')
        return 2

    data = load_rows(path)
    if not data:
        print('文件没有数据行')
        return 2

    # 校验必需列：主力净额 和 主买净额 都要有
    missing = [n for n in (FIELD_ZL, FIELD_ZM) if n not in data[0]]
    if missing:
        print(f'错误：文件缺少列 {chr(34)}{chr(34).join(missing)}{chr(34)}，请换一个含「主力净额」和「主买净额」列的文件')
        return 2

    # 日期优先级：文件名 > 日期列 > --date > 交互式询问
    # 日期优先级：--date 显式指定 > 文件名 > 日期列 > 交互式询问
    default_date = None
    filename_date = extract_date_from_filename(path)
    has_date = FIELD_DT in data[0]
    if date_arg:
        if not date_arg.isdigit() or len(date_arg) != 8:
            print(f'日期格式错误: {date_arg!r}，应为 8 位数字如 20260906')
            return 2
        default_date = int(date_arg)
        print(f'使用 --date 指定日期: {default_date}')
    elif filename_date is not None:
        default_date = filename_date
        print(f'从文件名提取日期: {default_date}')
    elif not has_date:
        default_date = input('Excel 里没有日期列，请输入日期(YYYYMMDD): ').strip()
        if not default_date.isdigit() or len(default_date) != 8:
            print(f'日期格式错误: {default_date!r}，应为 8 位数字如 20260906')
            return 2
        default_date = int(default_date)
        print(f'使用日期: {default_date}')

    # 交易日校验（只针对文件名/--date/交互得到的单一日历日）
    if default_date is not None and not is_trading_day(default_date):
        print(f'⚠️ 警告：{default_date} 不是交易日（无行情数据）')
        new = input('请输入正确日期(YYYYMMDD)，直接回车退出: ').strip()
        if new.isdigit() and len(new) == 8:
            default_date = int(new)
            print(f'改用日期: {default_date}')
        else:
            print('已取消')
            return 2

    d201, d202 = build(data, default_date)

    w201, c201, a201, nd1, cl201, rm201 = write_files(d201, 'signals_user_201', dry_run)
    w202, c202, a202, nd2, cl202, rm202 = write_files(d202, 'signals_user_202', dry_run)

    tag = '[dry-run 预览]' if dry_run else '[已写入]'
    print(f'{tag} 主力净额(ID=201): {w201} 个文件 (新建 {c201}, 覆盖合并 {a201})')
    print(f'{tag} 主买净额(ID=202): {w202} 个文件 (新建 {c202}, 覆盖合并 {a202})')
    if not dry_run:
        print(f'覆盖当天数据：201 清理 {cl201} 个文件(删除空文件 {rm201})，202 清理 {cl202} 个(删除空文件 {rm202})')
        print('写入完成。重启通达信后可在自定义数据中查看 201「主力净额」/ 202「主买净额」。')
    return 0


if __name__ == '__main__':
    sys.exit(main())
