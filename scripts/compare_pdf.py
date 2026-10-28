# -*- coding: utf-8 -*-
"""
对比与检验：把「我们生成的 PDF」和「Excel 自己导出的 PDF」做几何/底纹/文字/字体四类自动比对。

用法：
    python compare_pdf.py 参考.pdf 生成.pdf 输出目录

产出：
    <输出目录>/对比图.png          上下并排（同区域裁剪 + 局部放大）
    <输出目录>/对比检验报告.md         结构化结论
    <输出目录>/差异热力图.png      两版套准后的像素差异
"""
import os
import re
import shutil
import sys
import subprocess

import pdfplumber
from PIL import Image, ImageChops

DPI = 200
PX_PER_PT = DPI / 72.0


# ------------------------------------------------------------------ 工具
def to_img(pdf, out_prefix):
    subprocess.run(['pdftoppm', '-r', str(DPI), '-png', pdf, out_prefix], check=True)
    for f in sorted(os.listdir(os.path.dirname(out_prefix))):
        if f.startswith(os.path.basename(out_prefix)) and f.endswith('.png'):
            return os.path.join(os.path.dirname(out_prefix), f)
    raise SystemExit('渲染失败')


def cluster(vals, tol=1.2):
    out = []
    for v in sorted(vals):
        if not out or v - out[-1] > tol:
            out.append(v)
    return out


def _gray(val):
    """把 pdfplumber 的 non_stroking_color 归一成 0..1 灰度；失败返回 None。
    注意：单通道灰在 pdfplumber 里可能是 str（'0.851'）也可能是 list/tuple。"""
    try:
        if isinstance(val, str):
            vals = [float(x) for x in val.replace(',', ' ').split()]
        elif isinstance(val, (list, tuple)):
            vals = [float(x) for x in val]
        else:
            vals = [float(val)]
    except (TypeError, ValueError):
        return None
    if len(vals) == 1:
        return vals[0]
    if len(vals) >= 3:
        return sum(vals[:3]) / 3.0
    return None


def segs_of(pdf):
    """返回 竖线x / 横线y / 灰色填充块 / 文字 / 页面尺寸"""
    with pdfplumber.open(pdf) as p:
        pg = p.pages[0]
        V, H, F = [], [], []
        for l in pg.lines:
            if abs(l['x0'] - l['x1']) < 0.3:
                V.append(l['x0'])
            elif abs(l['top'] - l['bottom']) < 0.3:
                H.append(l['top'])
        for r in pg.rects:
            if r['width'] < 2.6 and r['height'] > 2.6:
                V.append(r['x0'])
            elif r['height'] < 2.6 and r['width'] > 2.6:
                H.append(r['top'])
            if r['width'] > 2.6 and r['height'] > 2.6:
                g = _gray(r.get('non_stroking_color'))
                if g is not None and 0.80 < g < 0.92:
                    F.append((r['x0'], r['top'], r['x1'], r['bottom']))
        words = [w['text'] for w in pg.extract_words()]
        return {'V': cluster(V), 'H': cluster(H), 'F': F, 'words': words,
                'page': (pg.width, pg.height)}


def gray_area(F):
    return sum((x1 - x0) * (b - t) for x0, t, x1, b in F)


def bbox_of_fills(F):
    if not F:
        return None
    return (min(f[0] for f in F), min(f[1] for f in F),
            max(f[2] for f in F), max(f[3] for f in F))


def table_bbox(pdf):
    """返回 (几何信息dict, 表格外框(x0,y_top,x1,y_bottom))"""
    d = segs_of(pdf)
    return d, (min(d['V']), min(d['H']), max(d['V']), max(d['H']))


def ink_mask(img, box):
    """区域内非白像素掩膜"""
    sub = img.convert('L').crop(box)
    return sub.point(lambda v: 255 if v < 200 else 0)


def fill_mask(img, box):
    """区域内灰底(#D9D9D9 ± 8)掩膜"""
    sub = img.convert('L').crop(box)
    return sub.point(lambda v: 255 if abs(v - 217) <= 10 else 0)


def iou(m1, m2):
    a = m1.convert('1')
    b = m2.convert('1')
    inter = ImageChops.logical_and(a, b).histogram()
    union = ImageChops.logical_or(a, b).histogram()
    i = inter[255] if len(inter) > 255 else 0
    u = union[255] if len(union) > 255 else 0
    return (i / u) if u else 1.0


# ---------------------------------------------------------------- 版本指纹
SIGMA = 0.917
COL_UNIT = 5.370


def xlsx_fingerprint(xlsx, sheet='上海大学国内差旅费报销单',
                     max_row=18, max_col=17):
    """把 xlsx 的列宽/行高换算成**归一化边界位置**（与打印缩放无关），
    且只保留「实际会画出来」的边线（合并区内部的线 Excel 不画，必须排除）。
    用于判定某份 PDF 到底出自哪一版 xlsx。"""
    import openpyxl
    ws = openpyxl.load_workbook(xlsx)[sheet]
    colw = [0.0] * (max_col + 1)
    for cd in ws.column_dimensions.values():
        for i in range(cd.min or 1, (cd.max or cd.min or 1) + 1):
            if 1 <= i <= max_col:
                colw[i] = 0.0 if cd.hidden else (cd.width or 0.0)
    dh = ws.sheet_format.defaultRowHeight or 14.0
    rowh = [0.0] + [(ws.row_dimensions[r].height or dh)
                    for r in range(1, max_row + 1)]

    merges = [(m.min_row, m.min_col, m.max_row, m.max_col)
              for m in ws.merged_cells.ranges
              if m.max_row <= max_row and m.max_col <= max_col]
    sup_v, sup_h = set(), set()
    for r0, c0, r1, c1 in merges:
        for r in range(r0, r1 + 1):
            for c in range(c0, c1):
                sup_v.add(('V', r, c))
        for r in range(r0, r1):
            for c in range(c0, c1 + 1):
                sup_h.add(('H', r, c))

    cum_w = [0.0] * (max_col + 1)
    for i in range(1, max_col + 1):
        cum_w[i] = cum_w[i - 1] + colw[i]
    tot_w = cum_w[max_col] or 1.0
    body = rowh[3:]                       # 表格外框从第 3 行开始
    tot_h = sum(body) or 1.0
    cum_h = [0.0]
    for h in body:
        cum_h.append(cum_h[-1] + h)

    def bd(r, c):
        b = ws.cell(row=r, column=c).border
        return b

    drawn_v, drawn_h = set(), {0.0, 1.0}
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            b = bd(r, c)
            if b is None:
                continue
            if b.left and b.left.style and ('V', r, c - 1) not in sup_v:
                drawn_v.add(c - 1)
            if b.right and b.right.style and ('V', r, c) not in sup_v:
                drawn_v.add(c)
            if b.top and b.top.style and r >= 4 and ('H', r - 1, c) not in sup_h:
                drawn_h.add(cum_h[r - 3] / tot_h)      # 上边 = 上一行的下边
            if b.bottom and b.bottom.style and r >= 3 and ('H', r, c) not in sup_h:
                drawn_h.add(cum_h[r - 2] / tot_h)      # 下边 = 第 (r-3) 个体的底

    V = sorted({cum_w[c] / tot_w for c in drawn_v} | {0.0, 1.0})
    H = sorted(drawn_h)
    return V, H


def norm_fit(meas, expect):
    """meas: 实测边界坐标；expect: 归一化期望位置（0..1，含首尾）。
    用**最近邻**匹配（对个别线漏检/合并稳健），
    返回 (最大偏差%, 平均偏差%)，单位是「占表宽/表高的百分比」。"""
    lo, hi = min(meas), max(meas)
    span = hi - lo
    if span <= 0:
        return 999.0, 999.0
    nd = [(m - lo) / span for m in meas]
    ds = [min(abs(x - y) for y in nd) * 100 for x in expect]
    return max(ds), sum(ds) / len(ds)


# ------------------------------------------------------------------ 主流程
def run(ref, mine, outdir, title='报销单', xlsx=None, alt=None):
    os.makedirs(outdir, exist_ok=True)
    tmp = os.path.join(outdir, '_tmp')
    os.makedirs(tmp, exist_ok=True)
    ru = to_img(ref, os.path.join(tmp, 'ref'))
    mu = to_img(mine, os.path.join(tmp, 'mine'))
    IU, IM = Image.open(ru), Image.open(mu)

    du, yu = table_bbox(ref)
    dm, ym = table_bbox(mine)

    def edge(d, k, i):
        return d[k][i] if d[k] else 0.0

    pu = (du['page'][0] or 595.276, du['page'][1] or 841.89)
    pm = (dm['page'][0] or 595.276, dm['page'][1] or 841.89)

    rep = []
    rep.append('# 报销单 PDF 检验报告 — %s\n' % title)
    rep.append('| 项目 | Excel 导出（参考） | 本次生成 |\n|---|---|---|\n'
               '| 页面 | %.0f×%.0f pt | %.0f×%.0f pt |\n'
               '| 表格外框 | x %.1f–%.1f / y %.1f–%.1f | x %.1f–%.1f / y %.1f–%.1f |\n'
               '| 表格尺寸 | %.1f × %.1f pt | %.1f × %.1f pt |\n'
               '| 灰色底纹块 | %d 块 / 合计 %.0f pt² | %d 块 / 合计 %.0f pt² |\n'
               '| 横竖线段 | %d 竖 + %d 横 | %d 竖 + %d 横 |\n' % (
                   pu[0], pu[1], pm[0], pm[1],
                   edge(du, 'V', 0), edge(du, 'V', -1), edge(du, 'H', 0), edge(du, 'H', -1),
                   edge(dm, 'V', 0), edge(dm, 'V', -1), edge(dm, 'H', 0), edge(dm, 'H', -1),
                   edge(du, 'V', -1) - edge(du, 'V', 0), edge(du, 'H', -1) - edge(du, 'H', 0),
                   edge(dm, 'V', -1) - edge(dm, 'V', 0), edge(dm, 'H', -1) - edge(dm, 'H', 0),
                   len(du['F']), gray_area(du['F']), len(dm['F']), gray_area(dm['F']),
                   len(du['V']), len(du['H']), len(dm['V']), len(dm['H'])))

    # ---------- 版本指纹：参考 PDF 到底出自哪一版 xlsx ----------
    same_ver = None
    if xlsx and os.path.exists(xlsx):
        cands = [('当前 xlsx（%s）' % os.path.basename(xlsx), xlsx)]
        if alt and os.path.exists(alt):
            cands.append(('另一版 xlsx（%s）' % os.path.basename(alt), alt))
        rep.append('\n## 0. 版本指纹（尺度无关）：参考 PDF 出自哪一版 xlsx？\n\n'
                   '把两份 xlsx 的列宽/行高换算成**归一化边界位置**，再与参考 PDF 实测的'
                   '竖线/横线归一化位置比较。偏差越小说明越可能是同一版本。\n\n'
                   '| 候选 xlsx | 列边界最大偏差 | 行边界最大偏差 | 判定 |\n|---|---|---|---|\n')
        best = None
        for name, path in cands:
            ev, eh = xlsx_fingerprint(path)
            mv, _ = norm_fit(du['V'], ev)
            mh, _ = norm_fit(du['H'], eh)
            verdict = '**吻合**' if (mv < 1.2 and mh < 1.8) else '不吻合'
            rep.append('| %s | %.2f%% | %.2f%% | %s |\n' % (name, mv, mh, verdict))
            if name.startswith('当前'):
                same_ver = (mv < 1.2 and mh < 1.8)
            if best is None or (mv + mh) < best[0]:
                best = (mv + mh, name, verdict)
        rep.append('\n> 参考 PDF 与本轮生成所依据的 xlsx 若不同版，'
                   '则两者的行列位置必然对不上，这属于**版本差异**而非渲染错误。\n'
                   '> 最接近的候选：**%s**（%s）。\n' % (best[1], best[2]))

    # 列边界对照
    rep.append('\n## 1. 列/行几何对照\n\n| # | 参考竖线 x | 本次竖线 x | 差 |\n|---|---|---|---|\n')
    for i in range(max(len(du['V']), len(dm['V']))):
        a = du['V'][i] if i < len(du['V']) else None
        b = dm['V'][i] if i < len(dm['V']) else None
        d = '' if (a is None or b is None) else '%+.1f' % (b - a)
        rep.append('| %d | %s | %s | %s |\n' % (i + 1, '-' if a is None else '%.1f' % a,
                                                '-' if b is None else '%.1f' % b, d))
    rep.append('\n| # | 参考横线 y | 本次横线 y | 差 |\n|---|---|---|---|\n')
    for i in range(max(len(du['H']), len(dm['H']))):
        a = du['H'][i] if i < len(du['H']) else None
        b = dm['H'][i] if i < len(dm['H']) else None
        d = '' if (a is None or b is None) else '%+.1f' % (b - a)
        rep.append('| %d | %s | %s | %s |\n' % (i + 1, '-' if a is None else '%.1f' % a,
                                                '-' if b is None else '%.1f' % b, d))

    # 文字完整性
    KEY = ['6879.87', '1029.87', '1749.87', '3030', '1830', '1200', '2100', '320', '400',
           '住宿费', '伙食补助费', '市内交通费', '大写']
    if xlsx and os.path.exists(xlsx):
        try:
            import openpyxl
            _ws = openpyxl.load_workbook(xlsx).worksheets[-2]
            nums = []
            for row in _ws.iter_rows():
                for cc in row:
                    if isinstance(cc.value, (int, float)) and cc.value:
                        s = ('%.2f' % cc.value).rstrip('0').rstrip('.')
                        if s not in nums:
                            nums.append(s)
            KEY = nums + [k for k in KEY if not k.replace('.', '').isdigit()]
        except Exception:
            pass
    rep.append('\n## 2. 关键数字/文字是否完整\n\n| 片段 | 参考 PDF | 本次生成 |\n|---|---|---|\n')
    for k in KEY:
        rep.append('| `%s` | %s | %s |\n' % (k, '有' if any(k in w for w in du['words']) else '**无**',
                                             '有' if any(k in w for w in dm['words']) else '**无**'))

    # 丢字形检测：参考里出现「1029」+「87」分离 或 「1030」
    bad_u = any(w == '1030' for w in du['words']) or any(w.startswith('1029') and w.endswith('87') and w != '1029.87' for w in du['words'])
    rep.append('\n**丢字形检测**：参考 PDF %s；本次生成 %s\n' % (
        '存在「1029.87 → 1029 87 / 1030」的丢小数点现象 ⚠️' if bad_u else '正常',
        '正常 ✅' if not any(w == '1030' for w in dm['words']) else '存在异常 ⚠️'))

    # 像素级套准
    def norm(img, bb):
        x0, t0, x1, b1 = bb
        return img.crop((int(x0 * PX_PER_PT), int(t0 * PX_PER_PT),
                         int(x1 * PX_PER_PT), int(b1 * PX_PER_PT)))

    bu, bm = norm(IU, yu), norm(IM, ym)
    hole = (min(bu.width, bm.width), min(bu.height, bm.height))
    rep_img_ink = iou(ink_mask(bu, (0, 0) + hole), ink_mask(bm.resize(bu.size), (0, 0) + hole))
    rep_img_fill = iou(fill_mask(bu, (0, 0) + hole), fill_mask(bm.resize(bu.size), (0, 0) + hole))
    diff = ImageChops.difference(bu.convert('L'), bm.resize(bu.size).convert('L'))
    diff.save(os.path.join(outdir, '差异热力图.png'))
    caveat = ''
    if same_ver is False:
        caveat = ('\n> ⚠️ **上面两个 IoU 数值在本次比对中不具备判定意义**：版本指纹已证明两份 PDF\n'
                  '> 出自不同版本的 xlsx（行高/列宽不同），按外框拉伸对齐后像素位置本就对不上。\n'
                  '> 判定「表格对应方式对不对」请以对源表的《逐格检验报告》为准；\n'
                  '> 本报告的作用是**给出差异来源的定量归因**。\n')
    elif same_ver:
        caveat = '\n> 两份 PDF 出自同一版本 xlsx，IoU 可直接作为版式一致性的量度。\n'
    rep.append('\n## 3. 像素级套准（已按表格外框对齐后）\n\n'
               '- 墨迹（文字+线）重合率 IoU：**%.1f%%**\n'
               '- 灰色底纹重合率 IoU：**%.1f%%**\n%s' % (
                   rep_img_ink * 100, rep_img_fill * 100, caveat))

    # 对比图
    def stack(box, scale, path, label_h=26):
        a, b = IU.crop(box), IM.crop(box)
        w = max(a.width, b.width)
        c = Image.new('RGB', (w, a.height + b.height + label_h * 2 + 16), 'white')
        c.paste(a, (0, label_h))
        c.paste(b, (0, label_h + a.height + label_h + 16))
        c = c.resize((int(c.width * scale), int(c.height * scale)), Image.LANCZOS)
        c.save(path)
        return path

    top = min(yu[1], ym[1])
    stack((int(14 * PX_PER_PT), int((top - 6) * PX_PER_PT),
           int(582 * PX_PER_PT), int((top + 300) * PX_PER_PT)), 0.86,
          os.path.join(outdir, '对比图_上部.png'))
    stack((int(14 * PX_PER_PT), int((top + 290) * PX_PER_PT),
           int(582 * PX_PER_PT), int((top + 470) * PX_PER_PT)), 0.86,
          os.path.join(outdir, '对比图_下部.png'))
    rep.append('\n> 对比图：上半为你导出的 PDF，下半为本次生成（同区域裁剪）。\n')

    with open(os.path.join(outdir, '对比检验报告.md'), 'w') as f:
        f.write(''.join(rep))
    print(''.join(rep))
    shutil.rmtree(tmp, ignore_errors=True)
    return os.path.join(outdir, '对比检验报告.md')


if __name__ == '__main__':
    a = sys.argv
    run(a[1], a[2], a[3],
        a[4] if len(a) > 4 else '报销单',
        a[5] if len(a) > 5 else None,
        a[6] if len(a) > 6 else None)
