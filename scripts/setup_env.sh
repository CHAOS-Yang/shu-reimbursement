#!/usr/bin/env bash
# 建立本技能所需的隔离 Python 环境（不污染系统环境）
#
#   bash scripts/setup_env.sh              # 默认 ./.venv
#   VENV_DIR=~/envs/shu bash scripts/setup_env.sh
#
set -euo pipefail

PY="${PYTHON:-python3}"
VENV="${VENV_DIR:-.venv}"

echo "==> Python: $($PY -V 2>&1)"
echo "==> venv  : $VENV"
"$PY" -m venv "$VENV"

"$VENV/bin/python" -m pip install --quiet --upgrade pip
"$VENV/bin/python" -m pip install --quiet \
    reportlab openpyxl pdfplumber pillow pymupdf

"$VENV/bin/python" - <<'PY'
import importlib
mods = ['reportlab', 'openpyxl', 'pdfplumber', 'PIL', 'fitz']
for m in mods:
    mod = importlib.import_module(m)
    print('  %-12s %s' % (m, getattr(mod, '__version__', getattr(mod, 'Version', 'ok'))))
PY

echo
echo "==> 字体检查"
for f in \
  "/System/Library/Fonts/Supplemental/Songti.ttc" \
  "/System/Library/Fonts/Supplemental/Arial Unicode.ttf"; do
  if [ -f "$f" ]; then echo "  ✅ $f"; else echo "  ⚠️  缺少 $f（脚本里改字体常量）"; fi
done

echo
echo "==> 外部命令"
for c in pdftotext pdftoppm pdffonts; do
  if command -v "$c" >/dev/null 2>&1; then echo "  ✅ $c"; else echo "  ⚠️  $c 未安装（macOS: brew install poppler）"; fi
done

echo
echo "完成。用法："
echo "  $VENV/bin/python scripts/geometry_render.py 表.xlsx 出.pdf --sheet 上海大学国内差旅费报销单"
echo "  $VENV/bin/python scripts/verify_cells.py   表.xlsx 出.pdf PDF核验/"
echo "  $VENV/bin/python scripts/patch_xlsx.py     表.xlsx --show"
