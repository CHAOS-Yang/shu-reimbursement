# CHANGELOG

> 项目起点：[**@Prism-Lab-SHU / SHU-reimbursement**](https://github.com/Prism-Lab-SHU/SHU-reimbursement)（MIT）。
> 本仓库在它的流程骨架之上，补齐 2026 新规、表格坐标与工具链。详见 [ATTRIBUTION.md](ATTRIBUTION.md)。

格式参考 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/)，版本号遵循 [语义化版本](https://semver.org/lang/zh-CN/)。

## [2.0.0] — 2026-10-27

**重构版。合并前身 `shu-reimbursement` v1.4.3 与配套工具 `xlsx-to-pdf` v1.3.0，分层、脱敏、补章。**
逐项继承对照见 [ATTRIBUTION.md](ATTRIBUTION.md)。

### 新增

- `SKILL.md` §0 **三条前置判断**：出差日 vs 2026-07-30、经费卡首字母、差旅 vs 日常报销
- `SKILL.md` §4 **十条铁律**：把原先散落的"必看""血泪教训"收敛成可执行禁令
- `references/08-出国境报销.md`：PIM 与智能报销双线、7 项凭证、柜照 5 个工作日、Airbnb 口径、合住分摊、注册费费率件、未成行口径
- `references/09-日常报销-科目口径.md`：办公费 vs 材料费判据、400–1000 元 10 类低值物资、**包干制负面清单**、会费科目、退单整改路径
- `references/11-实操坑清单.md` §10：**PDF 原凭证叠层加中文标注**（reportlab overlay + `show_pdf_page` + `saveIncr`），含"不要用 insert_font，否则 23 MB"的教训
- `scripts/patch_xlsx.py`：zip 级 XML 手术工具，支持列宽（能拆分 `min/max` 分组、保留 `hidden`）、行高、`fitToPage`、`--show` 自检
- `scripts/setup_env.sh`：一键建立隔离环境并检查字体与 poppler 命令
- `templates/README.md`：不内置校方模板，给出防盗链下载姿势
- `examples/01-报销单填写示意.md`：虚构单据演示坐标与小计
- `ATTRIBUTION.md`：继承链、逐项对照、脱敏对照、上游致谢

### 变更

- `geometry_render.py` 入口由硬编码 `BASE` / `TRIPS` 改为 `argparse`（`xlsx pdf --sheet NAME`），渲染逻辑未动
- 报销规则速查 → 拆为 `references/00`–`03`；填表坐标 → `references/04`–`05`；排版导出 → `references/06`–`07`
- 会议/培训补助的两分支改为对照表 + 算例，明确"看钱由谁出，不看有没有统一安排"
- 时效性事项改为状态表，标注已过期项

### 移除

- **不再维护单文件形态**（v1.4.3 的 SKILL.md 全文已按 §2.1 映射拆入 references）
- 删除所有个人与内部信息：姓名、工号、真实经费卡号、银行卡尾号、批件号/团组流水号/报销单号、本机绝对路径

### 修复

- 修正 v1.4.3 内部引用错位：`references/05`（报销单）/`06`（审批单）与路由表编号不一致
- 修正「住宿限额速查」指向错误章节的问题

---

## [1.4.3] — 2026-09-24（前身最终版）

> 单文件 `SKILL.md`（487 行）。以下为合并前的最后状态。

### 新增

- 上大内〔2026〕127/128 号全套规则（2026-07-30 施行）
- 判经费卡类型 → 选适用办法的前置判断
- 住宿费按省市分档限额表（非科研 3 档 / 科研 4 档）
- 会议/培训补助「食宿自理」分支
- 市内交通费「包干 与 凭票据实 二选一」
- 报销单排版规范：列宽 `O=12 / L=8.2 / N=1.0`、大写行前导空格 3、情况说明区行高公式
- zip 级 XML 手术细则（列分组、数字实体、`html.unescape`、保留签名图）
- 审批单 2026 版坐标（`A19/C19`、`A15`、新增 `A14`）与 Wingdings 2 复选框打勾手法
- 会费/学会会员费不得进差旅单的论证
- 错版表格排查、写表前关 Excel、`~$` 锁文件处置

---

## [1.3.0] — 2026-09-26（配套工具 `xlsx-to-pdf` 最终版）

- `geometry_render.py` 几何保真重建（`COL_UNIT` / `SIGMA` / `FSCALE` / `LEADING` 实测常数）
- `verify_cells.py` 逐格像素核验（底纹 / 边框 / 文字 / 裁切 / 签名图）
- `compare_pdf.py` 与参考 PDF 比对 + **版本指纹**（区分"渲染错了"与"用户改了列宽"）
- 补签名图 5 处部件与锚点调整
- 打印设置 `fitToPage`，去掉写死的 `scale`

---

## 更早

`shu-reimbursement` v1.0.0–v1.4.2、`xlsx-to-pdf` v1.0.0–v1.2.x 为内部迭代版本，未单独发布。
其结论已全部并入 v2.0.0。
