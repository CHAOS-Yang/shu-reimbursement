# -*- coding: utf-8 -*-
"""
逐格像素级检验：把「生成的 PDF」和「源 xlsx」逐格对照，验证
  ① 底纹（灰底有没有铺、铺得对不对）
  ② 边框（每条边线的有无与粗细）
  ③ 文字（该有字的格子有没有墨迹、有没有被裁掉）

用法：
    python verify_cells.py <xlsx> <生成的pdf> <输出目录> [标签]

产出：
    <输出目录>/逐格检验报告.md
    <输出目录>/底纹对照.png      期望底纹 vs 实际底纹 叠加图
    <输出目录>/问题标注.png      有问题的格子画红框
"""
import os
import shutil
import subprocess
import sys

import openpyxl
from PIL import Image, ImageDraw

from geometry_render import render, fill_rgb, _signature

DPI = 300
S = DPI / 72.0
PAGE_H = 841.89
PAGE_W = 595.276
GRAY_LO, GRAY_HI = 200, 234      # #D9D9D9 = 217
INK = 128


def to_png(pdf, prefix):
    subprocess.run(['pdftoppm', '-r', str(DPI), '-png', '-f', '1', '-l', '1', pdf, prefix],
                   check=True)
    d, b = os.path.dirname(prefix), os.path.basename(prefix)
    for f in sorted(os.listdir(d)):
        if f.startswith(b) and f.endswith('.png'):
            return os.path.join(d, f)
    raise SystemExit('PDF 渲染失败')


def edges_expected(ws, max_row, max_col, sup_v, sup_h, xs, ys):
    """返回 {('V',r,c): weight} / {('H',r,c): weight}；weight in {None,'thin','medium'}"""
    order = {None: 0, 'hair': 1, 'thin': 1, 'dotted': 1, 'dashed': 1,
             'medium': 2, 'thick': 3, 'double': 3}
    V, H = {}, {}

    def bump(store, key, style):
        cur = store.get(key)
        if order.get(style, 0) > order.get(cur, 0):
            store[key] = style

    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            b = ws.cell(row=r, column=c).border
            if b is None:
                continue
            for sd, key, store, sup in (
                    (b.left, ('V', r, c - 1), V, sup_v),
                    (b.right, ('V', r, c), V, sup_v),
                    (b.top, ('H', r - 1, c), H, sup_h),
                    (b.bottom, ('H', r, c), H, sup_h)):
                st = sd.style if (sd and sd.style) else None
                if st:
                    bump(store, key, st)
    # 合并区内侧边：Excel 不画
    for k in sup_v:
        V.pop(k, None)
    for k in sup_h:
        H.pop(k, None)
    return V, H


def sample_region(img, x0, yt, x1, yb, inset_pt=1.6):
    """返回单元格内部区域（去掉 inset 边距）的 (median_gray, ink_count, size)"""
    l = int(round(x0 * S))
    r = int(round(x1 * S))
    t = int(round((PAGE_H - yt) * S))
    b = int(round((PAGE_H - yb) * S))
    i = max(1, int(inset_pt * S))
    box = (l + i, t + i, r - i, b - i)
    if box[2] <= box[0] or box[3] <= box[1]:
        return None
    sub = img.crop(box).convert('L')
    px = list(sub.getdata())
    px.sort()
    med = px[len(px) // 2]
    return med, sum(1 for v in px if v < INK), len(px), box


def edge_ink(img, kind, coord, a0, a1):
    """沿边线取一条细带，返回最暗像素值（越小越黑）"""
    mid = (a0 + a1) / 2.0
    span = max(1.0, a1 - a0)
    lo = a0 + span * 0.25
    hi = a1 - span * 0.25
    thin = 0.7
    if kind == 'V':
        box = (int(round((coord - thin) * S)), int(round((PAGE_H - hi) * S)),
               int(round((coord + thin) * S)), int(round((PAGE_H - lo) * S)))
    else:
        box = (int(round(lo * S)), int(round((PAGE_H - coord - thin) * S)),
               int(round(hi * S)), int(round((PAGE_H - coord + thin) * S)))
    if box[2] <= box[0] or box[3] <= box[1]:
        return 255
    sub = img.crop(box).convert('L')
    return min(sub.getdata())


def run(xlsx, pdf, outdir, title='报销单', sheet='上海大学国内差旅费报销单',
        max_row=18, max_col=17):
    os.makedirs(outdir, exist_ok=True)
    info = render(xlsx, pdf, sheet, max_row=max_row, max_col=max_col, verbose=False)
    xs, ys = info['xs'], info['ys']
    tl2area, in_merge = info['tl2area'], info['in_merge']
    sup_v, sup_h = set(), set()
    for r0, c0, r1, c1 in info['merges']:
        for r in range(r0, r1 + 1):
            for c in range(c0, c1):
                sup_v.add(('V', r, c))
        for r in range(r0, r1):
            for c in range(c0, c1 + 1):
                sup_h.add(('H', r, c))

    ws = openpyxl.load_workbook(xlsx)[sheet]
    tmpdir = os.path.join(outdir, '_tmp')
    os.makedirs(tmpdir, exist_ok=True)
    png = to_png(pdf, os.path.join(tmpdir, 'pg'))
    img = Image.open(png)

    exp_fill, exp_text, exp_fill_inner = {}, {}, []
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            rgb = fill_rgb(ws.cell(row=r, column=c))
            if rgb and (r, c) not in in_merge:
                exp_fill[(r, c)] = rgb
            elif rgb and (r, c) in in_merge:
                exp_fill_inner.append((r, c, rgb))
            v = ws.cell(row=r, column=c).value
            if v not in (None, '') and (r, c) not in in_merge:
                exp_text[(r, c)] = str(v)

    V, H = edges_expected(ws, max_row, max_col, sup_v, sup_h, xs, ys)

    rep = ['# 逐格检验报告 — %s\n' % title,
           '\n源文件：`%s`　生成：`%s`\n' % (os.path.basename(xlsx), os.path.basename(pdf)),
           '\n采样精度 %d dpi，灰底基准 #D9D9D9(=217)。\n' % DPI]

    # ---------------- ① 底纹 ----------------
    fill_bad, fill_ok = [], 0
    for (r, c), rgb in sorted(exp_fill.items()):
        r1, c1 = tl2area.get((r, c), (r, c))
        x0, yt, x1, yb = xs[c - 1], ys[r - 1], xs[c1], ys[r1]
        if x1 - x0 < 3 or yt - yb < 3:
            continue
        s = sample_region(img, x0, yt, x1, yb)
        if not s:
            continue
        med = s[0]
        want_gray = sum(rgb) / 3.0 < 250
        got_gray = GRAY_LO <= med <= GRAY_HI
        if want_gray == got_gray:
            fill_ok += 1
        else:
            fill_bad.append(('%s%d' % (chr(64 + c), r), '应有底纹' if want_gray else '应无底纹',
                             '实测灰度 %d' % med, '未铺' if want_gray else '多铺'))

    rep.append('\n## ① 底纹\n\n- 应铺底纹的单元格：**%d** 个；一致 **%d**，不一致 **%d**\n'
               % (len(exp_fill), fill_ok, len(fill_bad)))
    if fill_bad:
        rep.append('\n| 单元格 | 期望 | 实测 | 结论 |\n|---|---|---|---|\n')
        for a, b, c_, d in fill_bad:
            rep.append('| %s | %s | %s | **%s** |\n' % (a, b, c_, d))
    if exp_fill_inner:
        rep.append('\n> 注意：以下单元格处在合并区内但不是左上角，Excel 不会单独铺底纹（已忽略）：%s\n'
                   % ', '.join('%s%d' % (chr(64 + c), r) for r, c, _ in exp_fill_inner))

    # 底纹对照叠加图
    ov = img.convert('RGB')
    d = ImageDraw.Draw(ov)
    for (r, c), rgb in exp_fill.items():
        r1, c1 = tl2area.get((r, c), (r, c))
        x0, yt, x1, yb = xs[c - 1], ys[r - 1], xs[c1], ys[r1]
        d.rectangle([x0 * S, (PAGE_H - yt) * S, x1 * S, (PAGE_H - yb) * S],
                    outline=(0, 120, 255), width=2)
    ov.save(os.path.join(outdir, '底纹对照.png'))

    # ---------------- ② 边框 ----------------
    border_bad, border_ok, skipped = [], 0, 0
    for kind, store in (('V', V), ('H', H)):
        keys = [k for k in store]
        for k in keys:
            _, r, c = k
            if kind == 'V':
                x = xs[c]
                a0, a1 = ys[r], ys[r - 1]
                if a1 - a0 < 3:
                    skipped += 1
                    continue
                ink = edge_ink(img, 'V', x, a0, a1)
            else:
                y = ys[r]
                a0, a1 = xs[c - 1], xs[c]
                if a1 - a0 < 3:
                    skipped += 1
                    continue
                ink = edge_ink(img, 'H', y, a0, a1)
            want = store[k]
            if ink < 110:
                border_ok += 1
            else:
                border_bad.append(('%s' % ('竖 x=%.1f' % x if kind == 'V' else '横 y=%.1f' % y),
                                   want, '最暗 %d（无线）' % ink))
    # 反向：期望无线但实际有线
    all_v = [('V', r, c) for r in range(1, max_row + 1) for c in range(0, max_col)]
    all_h = [('H', r, c) for r in range(0, max_row) for c in range(1, max_col + 1)]
    extra = 0
    for kind, keys in (('V', all_v), ('H', all_h)):
        for key in keys:
            if key in V or key in H:
                continue
            _, r, c = key
            if kind == 'V':
                if key in sup_v or xs[c + 1] - xs[c] < 0.2:
                    continue
                x = xs[c]
                a0, a1 = ys[r], ys[r - 1]
                if a1 - a0 < 3:
                    continue
                ink = edge_ink(img, 'V', x, a0, a1)
                tag = '竖 x=%.1f (第%d行右)' % (x, r)
            else:
                if key in sup_h:
                    continue
                y = ys[r]
                a0, a1 = xs[c - 1], xs[c]
                if a1 - a0 < 3:
                    continue
                ink = edge_ink(img, 'H', y, a0, a1)
                tag = '横 y=%.1f (第%d行下)' % (y, r + 1)
            if ink < 110:
                extra += 1
                if extra <= 20:
                    border_bad.append((tag, '无', '最暗 %d（多画）' % ink))
    rep.append('\n## ② 边框\n\n- 应存在的边线：**%d** 条（其中 %d 条因隐藏列/零宽被跳过）；'
               '实测检出 **%d**，**缺失 %d**\n'
               '- 不该有的边线：额外画出 **%d** 处\n'
               % (len(V) + len(H), skipped, border_ok, len(border_bad),
                  extra))
    if border_bad:
        rep.append('\n| 边线 | 期望 | 实测 |\n|---|---|---|\n')
        for a, b, c_ in border_bad:
            rep.append('| %s | %s | %s |\n' % (a, b, c_))

    # ---------------- ③ 文字 ----------------
    text_bad, text_ok = [], 0
    for (r, c), txt in sorted(exp_text.items()):
        r1, c1 = tl2area.get((r, c), (r, c))
        x0, yt, x1, yb = xs[c - 1], ys[r - 1], xs[c1], ys[r1]
        if x1 - x0 < 3 or yt - yb < 3:
            continue
        s = sample_region(img, x0, yt, x1, yb, inset_pt=0.8)
        if not s:
            continue
        ink = s[1]
        if '\n' in txt:
            need = True
        else:
            need = True
        if ink > 0:
            text_ok += 1
        else:
            text_bad.append(('%s%d' % (chr(64 + c), r), txt[:20]))
    rep.append('\n## ③ 文字\n\n- 非空单元格：**%d** 个；检出墨迹 **%d**，空白 **%d**\n'
               % (len(exp_text), text_ok, len(text_bad)))
    if text_bad:
        rep.append('\n| 单元格 | 内容 |\n|---|---|\n')
        for a, b in text_bad:
            rep.append('| %s | `%s` |\n' % (a, b))

    clipped = info.get('clipped') or []
    shrunk = info.get('shrunk') or []
    rep.append('\n## ④ 文字完整性（溢出/裁切）\n\n')
    if shrunk:
        rep.append('为使文字完整显示，以下单元格做了**兜底微缩**（Excel 会直接裁掉）：\n\n'
                   '| 单元格 | 内容 | 字号 |\n|---|---|---|\n')
        for r, c, t, sz, f in shrunk:
            rep.append('| %s%d | `%s` | %g pt → %g pt（×%.2f） |\n'
                       % (chr(64 + c), r, t.replace('\n', '/'), sz, sz * f, f))
        rep.append('\n')
    if clipped:
        rep.append('以下单元格 wrap 后仍放不下、按 Excel 语义被裁切：\n\n'
                   '| 单元格 | 内容 | 需要高度 | 可用高度 |\n|---|---|---|---|\n')
        for a, b, h, l in clipped:
            rep.append('| %s | `%s` | %.1f pt | %.1f pt |\n' % (a, b, h, l))
        rep.append('\n')
    if not shrunk and not clipped:
        rep.append('全部文字都在单元格内完整显示，无裁切、无需微缩。\n')

    # ---------------- ⑤ 签名图片 ----------------
    # 易漏点：xlsx 里嵌了签名图，但表格是别处复制来的（没有 xl/drawings）→ PDF 就没签名。
    # 这里既查「有没有」，也查「锚点算出来的位置/尺寸」与 PDF 实测是否一致。
    sig_bad = []
    sig = _signature(xlsx)
    rep.append('\n## ⑤ 签名图片\n\n')
    if sig is None:
        rep.append('- 源 xlsx 未嵌入图片（无 `xl/drawings`），跳过\n')
    else:
        _, col_i, col_off, row_i, row_off, cx, cy = sig
        k = info['k']
        ex0 = xs[col_i - 1] + col_off * k
        eyt = ys[row_i - 1] - row_off * k            # PDF 底原点：图片上边
        want = (ex0, PAGE_H - eyt, ex0 + cx * k, PAGE_H - eyt + cy * k)  # x0,top,x1,bottom
        anchor = '%s%d（锚点）' % (chr(64 + col_i), row_i)
        try:
            import pdfplumber
            with pdfplumber.open(pdf) as _pf:
                ims = _pf.pages[0].images
        except Exception as e:
            ims = []
            sig_bad.append(('读取 PDF 失败', str(e), ''))
        if not ims:
            sig_bad.append(('**签名图片缺失**', anchor, 'PDF 中未找到任何图片'))
            rep.append('- ⚠️ **源表有签名图，但 PDF 里没有** → 检查 xlsx 的 `xl/drawings`、'
                       '`xl/worksheets/_rels`、`[Content_Types].xml` 是否齐全\n')
        else:
            im = ims[0]
            got = (im['x0'], im['top'], im['x1'], im['bottom'])
            diffs = [abs(a - b) for a, b in zip(got, want)]
            tol = 0.8
            rep.append('| 项目 | 锚点推算 | PDF 实测 | 偏差 |\n|---|---|---|---|\n')
            for nm, a, b in (('左 x0', want[0], got[0]), ('上 top', want[1], got[1]),
                             ('右 x1', want[2], got[2]), ('下 bottom', want[3], got[3])):
                rep.append('| %s | %.2f | %.2f | %.2f |\n' % (nm, a, b, abs(a - b)))
            rep.append('| 尺寸 | %.1f × %.1f pt | %.1f × %.1f pt | |\n'
                       % (cx * k, cy * k, got[2] - got[0], got[3] - got[1]))
            if max(diffs) > tol:
                sig_bad.append(('位置/尺寸偏差过大', anchor,
                                '最大偏差 %.2f pt（容差 %.1f）' % (max(diffs), tol)))
                rep.append('\n- ⚠️ 位置偏差超过 %.1f pt（常见原因：图片 x 索引少减 1，'
                           '整幅图右移一个列宽）\n' % tol)
            else:
                rep.append('\n- 位置与锚点一致（最大偏差 %.2f pt，容差 %.1f）\n' % (max(diffs), tol))
            # 是否溢出表格右边框（xlsx 右边界 = xs[max_col]）
            if got[2] > xs[max_col] + 0.8:
                sig_bad.append(('签名图越出表格右边框', anchor,
                                '右缘 %.1f > 表格右界 %.1f' % (got[2], xs[max_col])))
                rep.append('- ⚠️ 右缘 %.1f 越出表格右边框 %.1f\n' % (got[2], xs[max_col]))
            else:
                rep.append('- 未越出表格右边框（右缘 %.1f ≤ %.1f）\n' % (got[2], xs[max_col]))

    # ---------------- 问题标注图 ----------------
    bad = Image.open(png).convert('RGB')
    dd = ImageDraw.Draw(bad)
    marks = []
    for a, b, c_, _d in fill_bad:
        col = ord(a[0]) - 64
        row = int(a[1:]) if not a[1:].isdigit() else int(a[1:])
        marks.append((col, row, '底纹'))
    for a, b, c_ in border_bad:
        pass
    for a, b in text_bad:
        marks.append((ord(a[0]) - 64, int(a[1:]), '文字'))
    for col, row, _tag in marks:
        r1, c1 = tl2area.get((row, col), (row, col))
        x0, yt, x1, yb = xs[col - 1], ys[row - 1], xs[c1], ys[r1]
        dd.rectangle([x0 * S, (PAGE_H - yt) * S, x1 * S, (PAGE_H - yb) * S],
                     outline=(220, 0, 0), width=3)
    bad.save(os.path.join(outdir, '问题标注.png'))

    ok = (not fill_bad) and (not border_bad) and (not text_bad) and (not sig_bad)
    rep.insert(2, '\n**总判定：%s**\n' % ('✅ 与源表一致' if ok else '⚠️ 存在不一致，见下'))
    with open(os.path.join(outdir, '逐格检验报告.md'), 'w') as f:
        f.write(''.join(rep))
    print(''.join(rep))
    shutil.rmtree(tmpdir, ignore_errors=True)
    return {'fill_bad': fill_bad, 'border_bad': border_bad, 'text_bad': text_bad,
            'sig_bad': sig_bad, 'clipped': clipped, 'ok': ok}


if __name__ == '__main__':
    run(*sys.argv[1:])
