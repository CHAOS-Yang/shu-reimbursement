# 07 · 工具：xlsx 几何保真渲染与逐格核验

三个脚本 + 一个补丁工具，都在 `scripts/`。目标：**没有 LibreOffice 也能导出「像 Excel 打印出来」的 PDF，并且能逐格证明它对。**

## 0. 环境

```bash
bash scripts/setup_env.sh
# 或在已有 venv 里：
pip install reportlab openpyxl pdfplumber pillow pymupdf
```

- macOS 自带 `pdftoppm`（poppler）与 `pdffonts`。
- 字体：`/System/Library/Fonts/Supplemental/Songti.ttc`。Linux 换 Noto CJK 并改脚本 `SONGTI` 常量。

## 1. `geometry_render.py` —— 按真实几何 + 真实样式重建

```bash
python scripts/geometry_render.py 差旅_202610成都.xlsx 报销单.pdf \
       --sheet 上海大学国内差旅费报销单
```

它读 xlsx 的列宽、行高、底纹、边框、合并区、字体、公式、签名图，还原成 PDF。

### 关键常数（换模板要重新标定）

| 常数 | 值 | 含义 |
|------|-----|------|
| `COL_UNIT` | **5.370** | 每个列宽单位 = 5.370 pt |
| `SIGMA` | **0.917** | PDF 行高 / xlsx 行高 |
| `FSCALE` | **0.949** | PDF 字号 / xlsx 字号 |
| `LEADING` | **1.23** | 行距系数 |
| 整体缩放 k | `min((页宽 - 2×边距) / 表宽, 1.0)` | 模拟 Excel「调整为 1 页宽」，行高与字号都要 ×k |

### 标定方法（换模板必做）

1. `pdftotext -layout` 或 `pdfplumber` 量出参考 PDF 的外框与每条竖线 x、横线 y；
2. 用 xlsx 列宽单位算累积值 → `COL_UNIT = 表宽 / 总列宽单位`；
3. `SIGMA = 实测相邻横线间距 / xlsx 行高`；
4. **交叉验证**：把 xlsx 边界归一化后与 PDF 实测边界做最近邻比对（见 `compare_pdf.py` 的 `xlsx_fingerprint` / `norm_fit`），**最大偏差 < 0.5%** 才算常数可靠。

### 九个坑（都踩过）

1. **合并区内侧边抑制要用三元组键**：边标识是 `('V', r, c)` / `('H', r, c)`，`sup_v`/`sup_h` 里若只存 `(r, c)` 则永不匹配 → 情况说明区会冒出多余竖线。
2. **坐标原点**：reportlab 是**底原点**。`ys[r-1]` = 第 r 行**上边**（数值大），`ys[r]` = 下边。别按屏幕坐标映射。
3. **不要用缩略图肉眼判断底纹铺没铺**——缩到 0.3 倍后 `#D9D9D9` 极易看漏，必须采样像素。
4. **`pdfplumber` 的单通道灰可能是字符串**：`rect['non_stroking_color'] == '0.851'`（不是 list）。过滤器写 `isinstance(c, (list, tuple))` 会全部漏掉。
5. **`pdfplumber` 的文本层不受裁剪影响**：被 `clipPath` 裁掉的字**仍能**被 `extract_words()` 读到 → 判断「有没有被裁」必须看像素。
6. **`Paragraph` 需要 `registerFontFamily()`**，否则报 `ps2tt` 错。自动微缩（`SHRINKS`）要克制，最低 0.88，并在报告里标出。
7. **每个 ASCII 空格要换成两个 `&nbsp;`**。两个原因叠加：
   - reportlab 按 XML 语义**折叠**连续空白 → 「合 计 金 额」的空格被吞；
   - 度量差异：Excel 的 SimSun 空格 = **0.5em**，macOS Songti 的空格只有 **0.25em**。
   实测校验：参考 PDF 相邻字符推进 = 10.44（汉字）+ 5.27（空格），比值 0.505；本机 `pdfmetrics.stringWidth(' ', 'Songti-6', 11) = 2.75` → 必须 ×2。
8. **签名/图片的 x 索引用 `xs[col_i - 1]`**。`xs[j]` 是「第 (j+1) 列的左边」，而 `<xdr:from><xdr:col>` 是 **0 基**列号 → `col_i = col + 1`（1 基），左边 = `xs[col_i - 1]`。写成 `xs[col_i]` 会让整幅图**右移一个列宽**，压出右边框。同理 `ys[row_i - 1]`。
9. **改了提交包里的 xlsx 要同步到源目录**：渲染器从你指定的 xlsx 读，不同步就会出现「改了锚点但重渲染没变化」的假象。

## 2. `verify_cells.py` —— 逐格像素级硬检验

```bash
python scripts/verify_cells.py 表.xlsx 报销单.pdf PDF核验/ 标签
```

300 dpi 渲染 → 按渲染器自己的几何取每格内部 → 中位灰度判底纹、最暗像素判边线、暗像素数判文字。

- 期望值**全部从 xlsx 现读**（fill / border / merges / value），所以这是「对源表」的硬检验，不是自说自话。
- 产出：`逐格检验报告.md`、`底纹对照.png`（蓝框 = 应铺底纹处）、`问题标注.png`（红框 = 问题格）。
- **判据**：底纹一致、边框缺失 0 / 多余 0、非空单元格全部有墨迹、无裁切（或裁切已列出）、**签名图**存在且位置尺寸与锚点推算一致（容差 0.8pt）、未越出表格右边框。

## 3. `compare_pdf.py` —— 与「别人导出的 PDF」比对 + 版本指纹

```bash
python scripts/compare_pdf.py 参考.pdf 我生成的.pdf 对比/ 标签 当前.xlsx 另一版.xlsx
```

- 统计页面 / 外框 / 底纹块 / 线段数，逐条列竖线横线，检索关键数字，算套准后 IoU；
- 产出 `对比图_上部/下部.png`、`差异热力图.png`；
- **★版本指纹**：几何对不上时**先别改渲染器**。用 `xlsx_fingerprint` + `norm_fit` 判定参考 PDF 出自哪一版 xlsx。
  实测：参考 PDF 与「备份的旧版」偏差 0.11%/0.21%（吻合），与当前版偏差 2.33%/3.83%（不吻合） → 说明是**用户改了列宽行高**，不是渲染错。

## 4. `patch_xlsx.py` —— zip 级 XML 补丁

```bash
# 调列宽 / 行高 / 打印设置，保留图片与富文本
python scripts/patch_xlsx.py 表.xlsx --col O=12 --col L=8.2 --row 15=60 --fit-to-page
```

只改 XML、不重存，所以**列分组、hidden 属性、签名图、富文本 run 全部原样保留**。
这也是修改带签名/带复选框的表时**唯一安全**的路径。

## 5. 给从模板复制来的 xlsx 补签名图

**症状**：渲染出的 PDF 签字栏空白，用户说「你忘记放图了」。

**先查**：`zipfile.namelist()` 里有没有 `xl/media/*` 与 `xl/drawings/drawing1.xml`。
空白模板复制出来的表**本来就没有**这些部件 → 不是渲染器丢的。

补齐要 5 处，缺一处 Excel 就报「修复」对话框：

| 部件 | 内容 |
|------|------|
| `xl/media/image1.jpeg` | 从有签名的那份**原样拷字节** |
| `xl/drawings/drawing1.xml` | 原样拷（`twoCellAnchor editAs="oneCell"` + `from`/`to` + `a:ext`） |
| `xl/drawings/_rels/drawing1.xml.rels` | `rId1` → `../media/image1.jpeg` |
| `xl/worksheets/_rels/sheetN.xml.rels` | **新建**：`rId1` → `../drawings/drawing1.xml` |
| `[Content_Types].xml` | 加 `<Default Extension="jpeg" .../>` + drawing 的 `<Override>` |

另外在 worksheet 的 `</worksheet>` 前加 `<drawing r:id="rId1"/>`；根元素若没声明 `xmlns:r` 要补上，否则 XML 非法。

**锚点**：`editAs="oneCell"` → 位置跟 `from` 走、尺寸取 `a:ext`。所以行高被改大后，签名会**掉到签字行下方**，需要把锚点行往前挪（例：`row 15 → 14`（0 基）、`rowOff 25400 → 155700`），并把 `to` 锚点调到自洽。
