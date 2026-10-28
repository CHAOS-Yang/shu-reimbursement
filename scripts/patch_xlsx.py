#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""patch_xlsx.py —— 对 xlsx 做「zip 级 XML 手术」，只改必要的节点。

为什么不用 openpyxl 重存？
    重存会重写列分组（min/max）、丢失 <pageSetUpPr> 与自定义打印缩放、
    有时还会塞一个写死的 scale。对报销单这种「有签名图 + 有富文本复选框」的表，
    整表重存风险太高。本脚本只替换目标 XML 片段，其余字节原样保留。

支持：
    --col O=12          设置列宽（可重复）
    --row 15=60         设置行高（可重复）
    --fit-to-page       打开「调整为 1 页宽」，并移除写死的 scale
    --show              只打印当前的列宽/行高/打印设置，不修改
    --sheet 名称         目标 sheet（默认按 --sheet-index 或第一个）
    --out 路径           输出到新文件（默认就地修改，原文件自动留 .bak）

用法示例：
    python patch_xlsx.py 表.xlsx --show
    python patch_xlsx.py 表.xlsx --col O=12 --col L=8.2 --row 15=60 --fit-to-page
"""
import argparse
import html
import os
import re
import shutil
import sys
import zipfile

COL_LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'


def col_to_index(letter):
    """'A' -> 1, 'AA' -> 27"""
    n = 0
    for ch in letter.strip().upper():
        if ch not in COL_LETTERS:
            raise ValueError('非法列标识: %r' % letter)
        n = n * 26 + (COL_LETTERS.index(ch) + 1)
    return n


def index_to_col(n):
    s = ''
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = COL_LETTERS[r] + s
    return s


# ------------------------------------------------------------------ sheet 定位
def sheet_map(z):
    """返回 {sheet 名: 'xl/worksheets/sheetN.xml'}"""
    wb = z.read('xl/workbook.xml').decode('utf-8', 'replace')
    rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8', 'replace')
    rid2target = dict(re.findall(r'<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    # Target 也可能排在 Id 前面，补一次
    for m in re.finditer(r'<Relationship([^>]*)/>', rels):
        attrs = dict(re.findall(r'(\w+)="([^"]*)"', m.group(1)))
        if 'Id' in attrs and 'Target' in attrs:
            rid2target[attrs['Id']] = attrs['Target']

    out = {}
    for m in re.finditer(r'<sheet\b([^>]*)/>', wb):
        attrs = dict(re.findall(r'(\w+:\w+|\w+)="([^"]*)"', m.group(1)))
        name = html.unescape(attrs.get('name', ''))
        rid = attrs.get('r:id') or attrs.get('id')
        tgt = rid2target.get(rid, '')
        if not tgt:
            continue
        tgt = tgt.lstrip('/')
        out[name] = tgt if tgt.startswith('xl/') else 'xl/' + tgt
    return out


# ------------------------------------------------------------------ XML 片段操作
ATTR_RE = re.compile(r'([\w:.-]+)="([^"]*)"')


def parse_attrs(s):
    return [(k, v) for k, v in ATTR_RE.findall(s)]


def attrs_to_str(pairs):
    return ' '.join('%s="%s"' % (k, v) for k, v in pairs)


def set_attr(pairs, key, value):
    out, found = [], False
    for k, v in pairs:
        if k == key:
            out.append((k, str(value)))
            found = True
        else:
            out.append((k, v))
    if not found:
        out.append((key, str(value)))
    return out


def del_attr(pairs, key):
    return [(k, v) for k, v in pairs if k != key]


def fmt_num(v):
    s = ('%.10f' % float(v)).rstrip('0').rstrip('.')
    return s if s else '0'


def patch_cols(xml, changes):
    """changes: {列索引: 宽度}"""
    m = re.search(r'<cols>(.*?)</cols>', xml, re.S)
    existing = []
    if m:
        for cm in re.finditer(r'<col\b([^>]*?)/>', m.group(1)):
            pairs = parse_attrs(cm.group(1))
            d = dict(pairs)
            existing.append((int(d.get('min', 1)), int(d.get('max', 1)), pairs))
    order = sorted(existing)

    for idx, width in sorted(changes.items()):
        hit = None
        for i, (mn, mx, pairs) in enumerate(order):
            if mn <= idx <= mx:
                hit = i
                break
        if hit is None:
            order.append((idx, idx, [('width', fmt_num(width)), ('customWidth', '1')]))
            continue
        mn, mx, pairs = order.pop(hit)
        base = dict(pairs)
        if mn == mx == idx:
            p = set_attr(pairs, 'width', fmt_num(width))
            p = set_attr(p, 'customWidth', '1')
            order.insert(hit, (mn, mx, p))
        else:
            if mn <= idx - 1:
                order.insert(hit, (mn, idx - 1, list(pairs)))
                hit += 1
            p = [('min', str(idx)), ('max', str(idx))]
            p = set_attr(p, 'width', fmt_num(width))
            p = set_attr(p, 'customWidth', '1')
            for k in ('style', 'hidden', 'bestFit', 'outlineLevel', 'collapsed', 'phonetic'):
                if k in base:
                    p = set_attr(p, k, base[k])
            order.insert(hit, (idx, idx, p))
            if idx + 1 <= mx:
                order.insert(hit + 1, (idx + 1, mx, list(pairs)))

    order.sort(key=lambda t: t[0])
    items = []
    for mn, mx, pairs in order:
        pairs = set_attr(pairs, 'min', mn)
        pairs = set_attr(pairs, 'max', mx)
        items.append('<col %s/>' % attrs_to_str(pairs))
    block = '<cols>%s</cols>' % ''.join(items)

    if m:
        return xml[:m.start()] + block + xml[m.end():]
    if '<sheetData' in xml:
        i = xml.index('<sheetData')
        return xml[:i] + block + xml[i:]
    raise RuntimeError('既没有 <cols> 也没有 <sheetData>，无法插入列定义')


def patch_rows(xml, changes):
    """changes: {行号: 行高pt}"""
    for r, h in sorted(changes.items()):
        pat = re.compile(r'<row\b([^>]*?)(/>|>)')
        found = False
        for m in pat.finditer(xml):
            pairs = parse_attrs(m.group(1))
            d = dict(pairs)
            if str(d.get('r')) != str(r):
                continue
            p = set_attr(pairs, 'ht', fmt_num(h))
            p = set_attr(p, 'customHeight', '1')
            xml = xml[:m.start()] + '<row %s%s' % (attrs_to_str(p), m.group(2)) + xml[m.end():]
            found = True
            break
        if not found:
            sys.stderr.write('  ! 第 %d 行在 XML 里没有 <row> 节点（空行），已跳过\n' % r)
    return xml


def patch_fit_to_page(xml):
    # 1) sheetPr / pageSetUpPr
    m = re.search(r'<sheetPr\b([^>]*?)(/>|>(.*?)</sheetPr>)', xml, re.S)
    if m:
        if m.group(2) == '/>':
            repl = '<sheetPr%s><pageSetUpPr fitToPage="1"/></sheetPr>' % m.group(1)
        else:
            inner = m.group(3) or ''
            if '<pageSetUpPr' in inner:
                inner2 = re.sub(r'<pageSetUpPr\b([^>]*?)/>',
                                lambda mm: '<pageSetUpPr %s/>' % attrs_to_str(
                                    set_attr(parse_attrs(mm.group(1)), 'fitToPage', '1')),
                                inner)
            else:
                inner2 = '<pageSetUpPr fitToPage="1"/>' + inner
            repl = '<sheetPr%s>%s</sheetPr>' % (m.group(1), inner2)
        xml = xml[:m.start()] + repl + xml[m.end():]
    else:
        anchor = '<dimension' if '<dimension' in xml else '<sheetData'
        i = xml.index(anchor)
        xml = xml[:i] + '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' + xml[i:]

    # 2) pageSetup：去 scale，设 fitToWidth / fitToHeight
    m = re.search(r'<pageSetup\b([^>]*?)/>', xml, re.S)
    if m:
        pairs = del_attr(parse_attrs(m.group(1)), 'scale')
        pairs = set_attr(pairs, 'fitToWidth', '1')
        pairs = set_attr(pairs, 'fitToHeight', '0')
        xml = xml[:m.start()] + '<pageSetup %s/>' % attrs_to_str(pairs) + xml[m.end():]
    else:
        tag = '<pageSetup fitToWidth="1" fitToHeight="0"/>'
        xml = xml.replace('</worksheet>', tag + '</worksheet>')
    return xml


# ------------------------------------------------------------------ 展示
def show(xml, label):
    print('[%s]' % label)
    m = re.search(r'<cols>(.*?)</cols>', xml, re.S)
    cols = []
    if m:
        for cm in re.finditer(r'<col\b([^>]*?)/>', m.group(1)):
            d = dict(parse_attrs(cm.group(1)))
            mn, mx = int(d.get('min', 1)), int(d.get('max', 1))
            rng = index_to_col(mn) if mn == mx else '%s:%s' % (index_to_col(mn), index_to_col(mx))
            cols.append('  %-8s width=%s%s%s' % (
                rng, d.get('width', '-'),
                ' style=%s' % d['style'] if 'style' in d else '',
                ' HIDDEN' if d.get('hidden') == '1' else ''))
    print('  列宽:' if cols else '  列宽: (无 <cols>，全部默认)')
    for c in cols:
        print(c)

    rows = []
    for rm in re.finditer(r'<row\b([^>]*?)/?>', xml):
        d = dict(parse_attrs(rm.group(1)))
        if 'ht' in d or d.get('hidden') == '1':
            rows.append('  row %-3s ht=%s%s' % (d.get('r'), d.get('ht', '-'),
                                                ' HIDDEN' if d.get('hidden') == '1' else ''))
    print('  行高:' if rows else '  行高: (全部默认)')
    for r in rows:
        print(r)

    ps = re.search(r'<pageSetup\b([^>]*?)/>', xml)
    d = dict(parse_attrs(ps.group(1))) if ps else {}
    has_fit = '<pageSetUpPr' in xml and 'fitToPage="1"' in xml
    print('  打印: fitToPage=%s fitToWidth=%s fitToHeight=%s scale=%s orientation=%s'
          % ('1' if has_fit else '0', d.get('fitToWidth', '-'),
             d.get('fitToHeight', '-'), d.get('scale', '-'), d.get('orientation', '-')))
    print('  图片: %s' % ('有 drawing 部件' if 'drawing r:id' in xml else '此 sheet 无 drawing'))


# ------------------------------------------------------------------ main
def main():
    ap = argparse.ArgumentParser(description='xlsx 的 zip 级 XML 手术（改列宽/行高/打印设置，保留图片与富文本）')
    ap.add_argument('xlsx')
    ap.add_argument('--sheet', help='目标 sheet 名')
    ap.add_argument('--sheet-index', type=int, default=1, help='按顺序取第 N 个 sheet（1 基，--sheet 优先）')
    ap.add_argument('--col', action='append', default=[], metavar='LETTER=WIDTH')
    ap.add_argument('--row', action='append', default=[], metavar='N=HEIGHT')
    ap.add_argument('--fit-to-page', action='store_true', help='启用「调整为 1 页宽」并删除写死的 scale')
    ap.add_argument('--show', action='store_true', help='只打印当前设置')
    ap.add_argument('--out', help='输出路径（默认就地修改并留 .bak）')
    args = ap.parse_args()

    path = args.xlsx
    with zipfile.ZipFile(path) as z:
        smap = sheet_map(z)
        if args.sheet:
            if args.sheet not in smap:
                sys.exit('找不到 sheet %r，可选：%s' % (args.sheet, ' / '.join(smap)))
            target = smap[args.sheet]
        else:
            names = list(smap)
            if not names:
                sys.exit('workbook 里没有 sheet')
            target = smap[names[min(args.sheet_index, len(names)) - 1]]
        xml = z.read(target).decode('utf-8')
        others = [(i, z.read(i.filename)) for i in z.infolist()]

    if args.show or not (args.col or args.row or args.fit_to_page):
        show(xml, target)
        if not (args.col or args.row or args.fit_to_page):
            print('\n（未指定修改项，仅展示。加 --col O=12 / --row 15=60 / --fit-to-page 才会改）')
        return

    col_changes = {}
    for item in args.col:
        k, _, v = item.partition('=')
        col_changes[col_to_index(k)] = float(v)
    row_changes = {}
    for item in args.row:
        k, _, v = item.partition('=')
        row_changes[int(k)] = float(v)

    new_xml = xml
    if col_changes:
        new_xml = patch_cols(new_xml, col_changes)
    if row_changes:
        new_xml = patch_rows(new_xml, row_changes)
    if args.fit_to_page:
        new_xml = patch_fit_to_page(new_xml)

    out = args.out or path
    if out == path:
        bak = path + '.bak'
        if not os.path.exists(bak):
            shutil.copy2(path, bak)
            print('已备份 →', bak)
    tmp = out + '.tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
        for info, data in others:
            if info.filename == target:
                data = new_xml.encode('utf-8')
            zi = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = info.external_attr
            zo.writestr(zi, data)
    os.replace(tmp, out)
    print('已写入 →', out)

    with zipfile.ZipFile(out) as z:
        show(z.read(target).decode('utf-8'), target)

    # 自检：openpyxl 能否读回（能读=结构基本合法）
    # 注意 openpyxl 的 width 单位与 XML 的 width 差 0.71 的 padding，别拿它核对数值
    try:
        import openpyxl
        wb = openpyxl.load_workbook(out)
        print('\n✅ openpyxl 回读成功：%d 个 sheet —— 结构合法，可用 Excel 打开' % len(wb.sheetnames))
    except Exception as e:                                   # noqa: BLE001
        sys.exit('❌ openpyxl 回读失败（XML 可能被破坏）：%s' % e)


if __name__ == '__main__':
    main()
