# -*- coding: utf-8 -*-
"""
几何保真渲染器 v3 —— 把「标准 Excel 报销单」按真实几何 + 真实样式重建为 PDF。

v1 的问题：只画 GRID 网格、完全不读底纹，与 Excel 自带导出差异巨大。
v2 的问题：文字超出单元格时会**溢出画到边框外**（Excel 是裁掉的）。

v3 做法：
  1. 边框：逐格读 border 的 top/bottom/left/right 与样式，合并区内侧边不画；
     外框/表头分隔用 medium(1.80pt)、普通内线用 thin(0.96pt)。
  2. 底纹：读 fill(patternType + fgColor 的 theme/tint)，按 Excel tint 算法还原 RGB
     （本模板 theme0=FFFFFF, tint=-0.14994 -> #D9D9D9），合并区按左上角整块铺。
  3. 几何：列宽 = 5.370pt/unit（实测）、行高 ×0.917、字号 ×0.949，再统一缩放到页宽。
  4. 文字：逐块用 canvas 裁剪到单元格矩形内绘制（和 Excel 一样，超出部分被裁），
     先做「自动微缩」（最小 0.88×）尽量让内容放得下，避免有意义的文字被裁掉。
  5. 字体：宋体 -> Songti.ttc(subfontIndex 6/1/0)，纯 ASCII 单元格用 Times；
     TTF 全量嵌入（uni=yes），**不会再出现「1029.87 -> 1029 87 / 1030」丢字形**。
  6. 签名图片：读 drawing1.xml 的 twoCellAnchor 定位。
"""
import os
import re
import colorsys
import zipfile

import openpyxl
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas as rl_canvas

# ---------------------------------------------------------------- 字体
SONGTI = '/System/Library/Fonts/Supplemental/Songti.ttc'
_FONTS = {}


def register_fonts():
    if _FONTS:
        return _FONTS
    _FONTS.update(song='STSong-Light', song_b='STSong-Light', hei='STSong-Light')
    if os.path.exists(SONGTI):
        for key, idx in (('song', 6), ('song_b', 1), ('hei', 0)):
            name = 'Songti-%d' % idx
            try:
                pdfmetrics.registerFont(TTFont(name, SONGTI, subfontIndex=idx))
                # Paragraph 需要字体族映射，否则 ps2tt() 会报错
                pdfmetrics.registerFontFamily(name, normal=name, bold=name,
                                              italic=name, boldItalic=name)
                _FONTS[key] = name
            except Exception:
                pass
    return _FONTS


# ---------------------------------------------------------------- 颜色
def excel_tint(rgb, tint):
    r, g, b = rgb
    h, l, s = colorsys.rgb_to_hls(r / 255.0, g / 255.0, b / 255.0)
    l = l * (1.0 + tint) if tint < 0 else l + (1.0 - l) * tint
    r, g, b = colorsys.hls_to_rgb(h, max(0.0, min(1.0, l)), s)
    return (int(round(r * 255)), int(round(g * 255)), int(round(b * 255)))


THEME_LT1 = (0xFF, 0xFF, 0xFF)


def fill_rgb(cell):
    f = cell.fill
    if not f or not f.patternType or f.patternType == 'none':
        return None
    fc = f.fgColor
    if fc is None:
        return None
    if fc.type == 'rgb' and isinstance(fc.rgb, str):
        v = fc.rgb
        if len(v) == 8:
            v = v[2:]
        return tuple(int(v[i:i + 2], 16) for i in (0, 2, 4))
    if fc.type == 'theme' and fc.theme in (0, 1):
        return excel_tint(THEME_LT1 if fc.theme == 0 else (0, 0, 0), fc.tint or 0.0)
    return None


# ---------------------------------------------------------------- 数值
def eval_value(v, ws):
    if isinstance(v, str) and v.startswith('='):
        expr = v[1:].strip()
        if re.fullmatch(r'[A-Z]{1,3}\d+(?:\s*\+\s*[A-Z]{1,3}\d+)*', expr):
            tot = 0.0
            for p in re.split(r'\+', expr):
                x = ws[p.strip()].value
                if isinstance(x, (int, float)):
                    tot += x
            return tot
        m = re.fullmatch(r'SUM\(([A-Z]{1,3}\d+):([A-Z]{1,3}\d+)\)', expr)
        if m:
            tot = 0.0
            for row in ws[m.group(1):m.group(2)]:
                for cc in row:
                    if isinstance(cc.value, (int, float)):
                        tot += cc.value
            return tot
        return None
    return v


def fmt_value(v):
    if v is None:
        return ''
    if isinstance(v, float):
        if abs(v - round(v)) < 1e-9:
            return str(int(round(v)))
        return ('%.2f' % v).rstrip('0').rstrip('.')
    return str(v)


# ---------------------------------------------------------------- 常量
SIGMA = 0.917        # 行高/字号缩放（实测自 Excel 导出）
COL_UNIT = 5.370     # 1 个 Excel 列宽单位 = 5.370pt
FSCALE = 0.949       # 字号比例（11pt -> 10.44pt）
LEADING = 1.23       # 行距倍数（实测 12.84/10.44）
THIN = 0.96
MEDIUM = 1.80
PADX = 1.0           # 单元格左右内边距
PADY = 0.4
# Excel 的行为是「wrap 后按行换行、超出单元格直接裁掉」，不缩放字号。
# 但裁掉的是有意义的内容，对报销单这种要给人看的单据不合适，
# 故保留一个**克制的兜底微缩**：只有当 wrap 后确实放不下时才逐档缩到放得下（最低 0.88）。
# 实际命中情况会在渲染日志与检验报告里标出来。
SHRINKS = (1.0, 0.97, 0.94, 0.91, 0.88)
MARGIN_X = 18.5
TOP_Y = 54.0
PAGE = (595.276, 841.89)
BIG = 1e6


def render(xlsx, out_pdf, sheet_name='上海大学国内差旅费报销单',
           max_row=18, max_col=17, verbose=True):
    F = register_fonts()
    ws = openpyxl.load_workbook(xlsx)[sheet_name]

    # ---------- 几何 ----------
    colw = [0.0] * (max_col + 1)
    for cd in ws.column_dimensions.values():
        if cd.width is None:
            continue
        for i in range(cd.min or 1, (cd.max or cd.min or 1) + 1):
            if 1 <= i <= max_col:
                colw[i] = 0.0 if cd.hidden else COL_UNIT * cd.width
    dh = ws.sheet_format.defaultRowHeight or 14.0
    rowh = [0.0] + [(ws.row_dimensions[r].height or dh) for r in range(1, max_row + 1)]

    k = min((PAGE[0] - 2 * MARGIN_X) / sum(colw[1:]), 1.0)
    colw = [w * k for w in colw]
    rowh = [h * SIGMA * k for h in rowh]
    fs_scale = FSCALE * k

    # ---------- 合并区 ----------
    merges = [(m.min_row, m.min_col, m.max_row, m.max_col)
              for m in ws.merged_cells.ranges
              if m.max_row <= max_row and m.max_col <= max_col]
    tl2area = {(r0, c0): (r1, c1) for r0, c0, r1, c1 in merges}
    in_merge = set()
    for r0, c0, r1, c1 in merges:
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                if (r, c) != (r0, c0):
                    in_merge.add((r, c))
    sup_v, sup_h = set(), set()
    for r0, c0, r1, c1 in merges:
        for r in range(r0, r1 + 1):
            for c in range(c0, c1):
                sup_v.add(('V', r, c))
        for r in range(r0, r1):
            for c in range(c0, c1 + 1):
                sup_h.add(('H', r, c))

    # ---------- 网格坐标（ys 为 PDF 底原点：ys[r-1] = 第 r 行上边，ys[r] = 下边） ----------
    x_left = (PAGE[0] - sum(colw[1:])) / 2.0
    y_top = PAGE[1] - TOP_Y
    xs = [x_left]
    for i in range(1, max_col + 1):
        xs.append(xs[-1] + colw[i])
    ys = [y_top]
    for r in range(1, max_row + 1):
        ys.append(ys[-1] - rowh[r])

    def cellrect(r0, c0, r1, c1):
        x0, x1 = xs[c0 - 1], xs[c1]
        yt, yb = ys[r0 - 1], ys[r1]          # yt = 上边(大), yb = 下边(小)
        return x0, yt, x1, yb, (x1 - x0), (yt - yb)

    # ---------- 文字块 ----------
    styles = []
    scache = {}

    def pstyle(size, font, align):
        key = (round(size, 3), font, align)
        if key not in scache:
            scache[key] = ParagraphStyle('s%r' % (key,), fontName=font, fontSize=size,
                                         leading=size * LEADING, alignment=align,
                                         spaceBefore=0, spaceAfter=0)
        return scache[key]

    def para_text(s):
        """转成 Paragraph 可用的标记文本。

        两个坑：
        1) reportlab 会折叠连续空白（Sheet 里“合 计 金 额”“万   仟”靠空格撑开，
           折叠后整行会挤在一起）；
        2) 空格宽度：Excel 的 SimSun 空格 = 0.5em，而 macOS Songti 的空格只有 0.25em，
           单用一个普通空格会偏窄一半。
        故：先做 XML 转义，再把每个 ASCII 空格换成两个 &nbsp;（0.25em × 2 = 0.5em）。
        """
        t = s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        t = t.replace(' ', '&nbsp;&nbsp;')
        return t.replace('\n', '<br/>')

    blocks = []          # (r0,c0,r1,c1, Paragraph, align, wrap, valign)
    shrunk = []
    clipped = []         # 按 Excel 语义被裁掉的单元格
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            if (r, c) in in_merge:
                continue
            cell = ws.cell(row=r, column=c)
            txt = fmt_value(eval_value(cell.value, ws))
            if txt == '':
                continue
            r1, c1 = tl2area.get((r, c), (r, c))
            x0, yt, x1, yb, cw, ch = cellrect(r, c, r1, c1)
            bold = bool(cell.font.bold)
            wrap = bool(cell.alignment.wrap_text) or ('\n' in txt)
            ascii_only = all(ord(ch) < 128 for ch in txt)
            if ascii_only:
                font = 'Times-Bold' if bold else 'Times-Roman'
            elif cell.font.name in ('黑体', 'SimHei'):
                font = F['hei']
            else:
                font = F['song_b'] if bold else F['song']
            al = {'left': 0, 'center': 1, 'right': 2}.get(cell.alignment.horizontal, 1)
            va = 'top' if cell.alignment.vertical == 'top' else 'mid'
            base = float(cell.font.size or 11.0) * fs_scale
            avail = max(cw - 2 * PADX * k, 1.0)
            limit = ch - 2 * PADY * k
            para, used = None, 1.0
            for f in SHRINKS:
                size = base * f
                p = Paragraph(para_text(txt), pstyle(size, font, al))
                pw, ph = p.wrap(avail if wrap else BIG, BIG)
                para, used = p, f
                if ph <= limit:
                    break
            blocks.append((r, c, r1, c1, para, al, wrap, va))
            if used < 1.0:
                shrunk.append((r, c, txt[:18], float(cell.font.size or 11.0), used))
            if ph > limit + 0.5:
                cc = chr(64 + c) if c <= 26 else '?'
                clipped.append(('%s%d' % (cc, r), txt.replace('\n', '/')[:22],
                                round(ph, 1), round(limit, 1)))
            if al == 0:
                styles.append(('ALIGN', (c - 1, r - 1), (c - 1, r - 1), 'LEFT'))
            elif al == 2:
                styles.append(('ALIGN', (c - 1, r - 1), (c - 1, r - 1), 'RIGHT'))
            if cell.alignment.vertical == 'top':
                styles.append(('VALIGN', (c - 1, r - 1), (c - 1, r - 1), 'TOP'))

    # ---------- 底纹 ----------
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            if (r, c) in in_merge:
                continue
            rgb = fill_rgb(ws.cell(row=r, column=c))
            if not rgb:
                continue
            r1, c1 = tl2area.get((r, c), (r, c))
            styles.append(('BACKGROUND', (c - 1, r - 1), (c1 - 1, r1 - 1),
                           colors.Color(*[v / 255.0 for v in rgb])))
    for r0, c0, r1, c1 in merges:
        if (r0, c0) != (r1, c1):
            styles.append(('SPAN', (c0 - 1, r0 - 1), (c1 - 1, r1 - 1)))

    # ---------- 空表（只出底纹） ----------
    grid = [['' for _ in range(max_col)] for _ in range(max_row)]
    tbl = Table(grid, colWidths=colw[1:], rowHeights=rowh[1:], repeatRows=0)
    tbl.setStyle(TableStyle([('LEFTPADDING', (0, 0), (-1, -1), 0),
                             ('RIGHTPADDING', (0, 0), (-1, -1), 0),
                             ('TOPPADDING', (0, 0), (-1, -1), 0),
                             ('BOTTOMPADDING', (0, 0), (-1, -1), 0)] + styles))

    cv = rl_canvas.Canvas(out_pdf, pagesize=PAGE)
    tbl.wrapOn(cv, sum(colw[1:]), sum(rowh[1:]))
    tbl.drawOn(cv, x_left, y_top - sum(rowh[1:]))

    # ---------- 边框线段 ----------
    segs = []
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            b = ws.cell(row=r, column=c).border
            if b is None:
                continue
            for name, edge, a, b1, b2, kind in (
                    ('left', ('V', r, c - 1), xs[c - 1], ys[r - 1], ys[r], 'v'),
                    ('right', ('V', r, c), xs[c], ys[r - 1], ys[r], 'v'),
                    ('top', ('H', r - 1, c), ys[r - 1], xs[c - 1], xs[c], 'h'),
                    ('bottom', ('H', r, c), ys[r], xs[c - 1], xs[c], 'h')):
                sd = getattr(b, name)
                if not sd or not sd.style:
                    continue
                if kind == 'v' and edge in sup_v:
                    continue
                if kind == 'h' and edge in sup_h:
                    continue
                segs.append((kind, a, b1, b2, sd.style))
    order = {'thin': 0, 'hair': 0, 'dotted': 0, 'dashed': 0,
             'medium': 1, 'thick': 2, 'double': 2}
    for kind, a, b1, b2, st in sorted(segs, key=lambda s: order.get(s[4], 0)):
        cv.setStrokeColorRGB(0, 0, 0)
        cv.setLineWidth((MEDIUM if st in ('medium', 'thick') else THIN) * k)
        if kind == 'v':
            cv.line(a, b1, a, b2)
        else:
            cv.line(b1, a, b2, a)

    # ---------- 文字（裁剪到单元格） ----------
    for r, c, r1, c1, para, al, wrap, va in blocks:
        x0, yt, x1, yb, cw, ch = cellrect(r, c, r1, c1)
        pw, ph = para.wrap(max(cw - 2 * PADX * k, 1.0) if wrap else BIG, BIG)
        if al == 0:
            px = x0 + PADX * k
        elif al == 2:
            px = x1 - PADX * k - pw
        else:
            px = x0 + (cw - pw) / 2.0
        if va == 'top' or ph > ch:
            py = yt - PADY * k - ph
        else:
            py = yb + (ch - ph) / 2.0
        cv.saveState()
        path = cv.beginPath()
        path.rect(x0, yb, cw, ch)
        cv.clipPath(path, stroke=0, fill=0)
        para.drawOn(cv, px, py)
        cv.restoreState()

    # ---------- 签名图片 ----------
    sig = _signature(xlsx)
    if sig:
        blob, col_i, col_off, row_i, row_off, cx, cy = sig
        tmp = '/tmp/_sig_render.jpeg'
        with open(tmp, 'wb') as fh:
            fh.write(blob)
        # xs[j] = 第 (j+1) 列的左边，故 col_i 列（1-based）的左边界是 xs[col_i-1]
        px = xs[col_i - 1] + col_off * k
        py = ys[row_i - 1] - row_off * k - cy * k
        cv.drawImage(tmp, px, py, width=cx * k, height=cy * k, mask=None)
    cv.showPage()
    cv.save()
    if verbose:
        print('PDF -> %s | 表宽 %.1fpt k=%.4f 文字块 %d 微缩 %d 裁切 %d' %
              (os.path.basename(out_pdf), sum(colw[1:]), k, len(blocks), len(shrunk),
               len(clipped)))
        for r, c, t, sz, f in shrunk:
            cc = chr(64 + c) if c <= 26 else '?'
            print('   · 微缩 %s%d  %s  %gpt -> %gpt (×%.2f)' %
                  (cc, r, t, sz, sz * f, f))
    return {'width': sum(colw[1:]), 'k': k, 'shrunk': shrunk, 'clipped': clipped,
            'xs': xs, 'ys': ys, 'colw': colw, 'rowh': rowh,
            'merges': merges, 'in_merge': in_merge, 'tl2area': tl2area}


def _signature(xlsx):
    with zipfile.ZipFile(xlsx) as z:
        names = z.namelist()
        if 'xl/drawings/drawing1.xml' not in names:
            return None
        xml = z.read('xl/drawings/drawing1.xml').decode('utf-8')
        img = None
        for n in sorted(names):
            if n.startswith('xl/media/'):
                img = z.read(n)
                break
        if img is None:
            return None
        m = re.search(r'<xdr:from><xdr:col>(\d+)</xdr:col><xdr:colOff>(-?\d+)</xdr:colOff>'
                      r'<xdr:row>(\d+)</xdr:row><xdr:rowOff>(-?\d+)</xdr:rowOff>', xml)
        if not m:
            return None
        m2 = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"/>', xml)
        return (img, int(m.group(1)) + 1, int(m.group(2)) / 12700.0,
                int(m.group(3)) + 1, int(m.group(4)) / 12700.0,
                int(m2.group(1)) / 12700.0, int(m2.group(2)) / 12700.0)


# ---------------------------------------------------------------- CLI
def main():
    import argparse
    ap = argparse.ArgumentParser(
        description='把 xlsx 的某个 sheet 按真实几何 + 真实样式重建为 PDF（不依赖 LibreOffice）。')
    ap.add_argument('xlsx', help='源 xlsx 路径')
    ap.add_argument('pdf', help='输出 PDF 路径')
    ap.add_argument('--sheet', default='上海大学国内差旅费报销单',
                    help='要渲染的 sheet 名（按名字取，默认报销单）')
    args = ap.parse_args()
    info = render(args.xlsx, args.pdf, sheet_name=args.sheet)
    print('written:', args.pdf)
    if isinstance(info, dict):
        print('table width %.1fpt  scale k=%.3f  shrunk=%s  clipped=%s'
              % (info.get('width', 0), info.get('k', 0),
                 info.get('shrunk'), info.get('clipped')))


if __name__ == '__main__':
    main()
