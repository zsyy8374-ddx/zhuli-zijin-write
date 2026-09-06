#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
主力资金写入通达信自定义数据序列
  ID=201 → 主力净额（日期, 数值）
  ID=202 → 主买净额（日期, 数值）

用法：
  python3 zhuli_write.py <excel文件路径> [--dry-run]

支持两种输入（自动识别）：
  1. 通达信导出的「.xls」——实为 Tab 分隔文本（GBK）
  2. 真正的 Excel「.xlsm / .xlsx」——zip 结构

写入结构（参考 ID=101 竞价最大占比）：
  每条 8 字节 = 日期(uint32 小端, 十进制 yyyymmdd) + 数值(float32 小端)
  同日替换、按日期升序；已存在文件合并（不删旧数据），不存在则创建。

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

SIG = '/mnt/d/GP/通达信金融终端(开心果交易版)V2026/T0002/signals'
NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

FIELD_CODE = '代码'
FIELD_ZL = '主力净额'
FIELD_ZM = '主买净额'
FIELD_DT = '日期'


def market_flag(code: str) -> int:
    return {'6': 1, '8': 1, '0': 0, '3': 0, '9': 2}.get(code[0], -1)


def _num(x):
    x = (x or '').strip()
    try:
        return float(x)
    except (ValueError, TypeError):
        return None


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


def build(data):
    """data: list[dict{列名:值}] -> (dict201, dict202)  key=市场标志_代码 -> {日期int: 数值float}"""
    d201 = defaultdict(dict)
    d202 = defaultdict(dict)
    for row in data:
        code = (row.get(FIELD_CODE) or '').strip()
        if not code:
            continue
        code = ('000000' + code)[-6:]
        fl = market_flag(code)
        if fl < 0:
            continue
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
    written = created = appended = 0
    for key, recs in data.items():
        path = os.path.join(d, key + '.dat')
        exists = os.path.exists(path)
        if dry_run:
            # 预览模式不读文件、不写盘，只统计
            if exists:
                appended += 1
            else:
                created += 1
            written += 1
            continue
        merged = {}
        if exists:
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
    return written, created, appended


def main():
    args = [a for a in sys.argv[1:]]
    dry_run = '--dry-run' in args
    args = [a for a in args if a != '--dry-run']
    if not args:
        print('用法: python3 zhuli_write.py <excel文件路径> [--dry-run]')
        return 2
    path = args[0]
    if not os.path.exists(path):
        print(f'文件不存在: {path}')
        return 2

    data = load_rows(path)
    d201, d202 = build(data)

    w201, c201, a201 = write_files(d201, 'signals_user_201', dry_run)
    w202, c202, a202 = write_files(d202, 'signals_user_202', dry_run)

    tag = '[dry-run 预览]' if dry_run else '[已写入]'
    print(f'{tag} 主力净额(ID=201): {w201} 个文件 (新建 {c201}, 追加合并 {a201})')
    print(f'{tag} 主买净额(ID=202): {w202} 个文件 (新建 {c202}, 追加合并 {a202})')
    if not dry_run:
        print('写入完成。重启通达信后可在自定义数据中查看 201「主力净额」/ 202「主买净额」。')
    return 0


if __name__ == '__main__':
    sys.exit(main())
