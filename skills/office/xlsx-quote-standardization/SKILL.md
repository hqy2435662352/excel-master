---
name: xlsx-quote-standardization
description: 将空调报价 Excel 文件中的客户 sheet 迁移为 24 列统一标准化格式，含智能列映射（自动从表头识别列含义）、7 个派生公式生成、样式应用和完整验证流水线。当用户需要标准化报价数据、统一不同客户的报价表格式、跨客户对比报价/利润率，或处理含多数据块（不同日期报价快照）的空调报价 Excel 文件时使用此技能。适用于报价汇总、报价标准化、报价迁移、报价格式统一等场景。
compatibility: requires openpyxl; Python ≥ 3.8
tags: [xlsx, quote, standardization, migration, openpyxl]
---

# 报价汇总表标准化

将空调报价汇总 Excel 文件中的客户 sheet 迁移为 24 列标准化格式，使下游能进行跨客户横向比较和财务分析自动化。

本技能自包含——所有参考文档和代码模板均在 `references/` 目录下。

---

## 快速启动

1. 复制 `references/migrate_template.py` 到工作目录
2. 修改顶部 CONFIG 区的 4 项配置（**模板里是占位符 `r"<源文件路径>"` / `r"<输出文件路径>"`，不是真值，必须全文替换**）：
   - `SRC` — 源文件路径
   - `OUT` — 输出文件路径
   - `SKIP_SHEETS` — 跳过的 sheet（如 `{"333", "FRESH转口"}`）
   - `SHEET_OUTPUT_NAMES` — 源 sheet 名 → 输出 sheet 名映射
3. 运行迁移脚本
4. 运行验证：`python scripts/verify_semantic.py <源.xlsx> <输出.xlsx>`

---

## 架构

三层模型：`L1 定价哲学（333 sheet）→ L2 区域母模板（埃及机型报价模板）→ L3 客户实例`

一个客户 sheet 可含多个独立数据块（不同日期/型号组合的报价快照），所有块合并到同一个输出 sheet。同一标题下 C 列为空的行分割**子块**——W（总盈亏）每个子块独立计算。

输出文件每客户一个 sheet，块间 3 行空行分隔，子块间 1 行空行分隔。

---

## 一、24 列 Schema

| # | 列 | 字段 | 类型 | 说明 |
|---|-----|------|------|------|
| 1-6 | A-F | 类别/产品类别/订单明细/工厂型号/配置描述/压缩机 | 文本 | 物料号（C 列）格式：`^Z[0-9A-Z]{8,}$` |
| 7 | G | 数量 | 数字(>0) | — |
| 8-9 | H-I | 历史报价/客户目标价 | 数字 | — |
| 10 | J | 报价 | 数字 | **蓝字黄底**（用户输入列） |
| 11-14 | K-N | 财务费用/OA信保/返点/其他费用 | 数字 | 源缺值 → 填 0 |
| 15 | O | 净价 | **公式** | `=IFERROR(J-K-L-M-N,0)` |
| 16 | P | 原型机成本 | 数字 | 源缺值 → 橙底待裁决 |
| 17 | Q | 铜管成本(3m) | 数字 | 源缺值 → 填 0 |
| 18 | R | 结算价 | **公式** | `=IFERROR(P+Q,0)` |
| 19 | S | 净收入 | **公式** | `=IFERROR(O*G,0)` |
| 20 | T | 毛利 | **公式** | `=IFERROR((O-R)*G,0)` |
| 21 | U | 单机型损益率 | **公式** | `=IFERROR(IF(S=0,0,T/S),0)` |
| 22 | V | 系列盈亏 | **比率公式** | `=IFERROR(SUM(T系列)/SUM(S系列),0)` |
| 23 | W | 总盈亏 | **比率公式** | `=IFERROR(SUM(T子块)/SUM(S子块),0)` |
| 24 | X | 铜管规格 | 文本/数字 | — |

### 公式语义

7 个公式列 (O/R/S/T/U/V/W) 总是写入 + IFERROR 防护，源数据缺值时返回 0。

**S/T 公式基于净价 O，不是报价 J**。这是最常见的出错点——K/L/M/N 列（财务/OA/返点/其他）从报价中扣除后才是真正的落地价。如果用报价 J 计算净收入和毛利，核心利润指标会失实 10%-30%。

| 公式列 | 正确写法 | 错误写法及后果 |
|--------|---------|---------------|
| S 净收入 | `O*G` | ❌ `J*G` — 净收入虚高 |
| T 毛利 | `(O-R)*G` | ❌ `(J-R)*G` — 毛利完全失真 |

---

## 二、样式规范

| 元素 | 字体 | 填充色 |
|------|------|--------|
| 标题行 | Arial 12pt Bold #000000 | #D9E1F2 浅蓝 |
| 表头行 | Arial 10pt Bold #FFFFFF | #4472C4 标准蓝 |
| 普通数据 | Arial 10pt #000000 | 无 |
| 报价列 J | Arial 10pt #0000FF（蓝） | #FFF2CC 浅黄 |
| 公式列 O/R/S/T/U | Arial 10pt Italic #008000（绿） | #F2F2F2 浅灰 |
| 聚合列 V/W | Arial 10pt Bold #C00000（红） | #E2EFDA 浅绿 |
| P 列缺值 | Arial 10pt #000000 | #FCE4D6 浅橙 |

**数字格式**：
- H-R, X → `\$#,##0.00_);[Red]\(\$#,##0.00\)`（USD，负数红字括号）
- S → `\$#,##0.00_);[Red]\(\$#,##0.00\)`
- T → `\$#,##0.00;\-\$#,##0.00`（USD，负数减号前缀）
- U-W → `0.00%`

**合并规则**：
- 标题行 A:X 整行合并；每个数据块一个标题
- A 列：同系列前向填充后合并
- V 列：每系列首行写公式后合并整个系列
- W 列：每子块首数据行写公式后合并整个子块

完整样式常量和列宽定义见 `references/migrate_template.py` 的 STYLES 区。

---

## 三、核心机制：智能列映射

### 为什么需要智能列映射

源文件中不同客户 sheet 的列布局完全不同——报价列可能在 Col 8-12 的任意位置。更隐蔽的问题是：**同一 sheet 的不同数据块**，表头列含义也可能不同。

例如某客户的两个报价块：
- 块 1 表头：`报价 | 财务费用 | OA信保 | 返点 | 净价`
- 块 2 表头：`报价 | 净价 | OA信保 | 返点`（没有"财务费用"列，Col 12 直接是"净价"）

如果用同一个列映射套所有块，块 2 会把"净价"值误当"财务费用"写入 K 列——数据静默错位，只有对账脚本能发现。

### 解决方案：每块独立从表头自动识别

`smart_col_map_from_header(ws, header_row)` 从源表头逐字识别列含义，**主循环每块调用一次**：

- **报价列**（4 级 fallback）：精确匹配 `"报价"` → `"PRICE"`（英文表头）→ `"价格"`（iclima 等用词）→ 包含"报价"但排除历史报价/目标价/向ELB终端报价等
- **K/L/M/N/P/Q**：按表头字面精确匹配（"财务费用"/"OA信保"/"返点"/"其他费用"/"原型机成本"/"铜管成本"）
- **A-F, G**：固定 1:1 映射（前 7 列在所有 sheet 中一致）

```python
for blk in blocks:
    col_map, _ = smart_col_map_from_header(ws_src, blk["hr"])
    write_data_rows(ws_out, ..., col_map)
```

函数实现在 `references/migrate_template.py`。`detect_blocks` 返回的每个 block 含 `hr` 字段（表头行号），自动处理表头在 R2/R3 等不同位置的情况。

### 边界情况

- 表头使用非标准词（如 `"AUX 报价teriak"`）时可能误匹配 → 先抽样表头确认字面
- 同一 sheet 不同块表头完全重排时需扩展排除词列表
- **`"价格"` vs `"报价"` 表头异构**：部分客户 sheet（如 iclima）块 1-4 表头使用 `"报价"` 而块 5-9 使用 `"价格"`。smart_col_map 的第 3 级 fallback 自动处理，**前提是模板的 fallback 链已包含 `"价格"`**。迁移前检查目标文件的表头用词，若遗漏则需更新 include 链。对应 `migrate_template.py` 约第 95-98 行。
- **`"降价报价"` 不应在排除列表中**：埃及询盘块 2 表头同时含 `"向ELB终端报价"`（Col9）和 `"降价报价"`（Col10），后者是实际的报价列。排除列表需精准——只排除 `"向ELB终端报价"` 保留 `"降价报价"`，否则 J 列会漏配。对应 `migrate_template.py` 第 102 行 exclude 元组。

---

## 四、迁移流程

### 4.1 分析源数据

对每个客户 sheet：
- 用 `references/detect_blocks.py` 的 `detect_blocks(ws)` 自动探测所有数据块、子块、总计行
- 确认每块的标题行、表头行、数据范围
- 抽样检查表头字面，确认列映射能正确识别

### 4.2 配置并运行

复制 `references/migrate_template.py`，只修改 CONFIG 区的 4 项配置。不需要手写 col_map 或行范围——模板的主循环自动完成：

1. `detect_blocks` — 探测每 sheet 的所有块
2. `smart_col_map_from_header` — 每块独立识别列映射
3. `write_data_rows` — 写数据行 + 公式 + 样式
4. `apply_vw_with_subblocks` — 写 V/W 聚合公式 + 合并单元格
5. `fix_column_widths` — 设置列宽

### 4.3 最小测试先行

文件含多个客户时，先跑 2-3 个代表性 sheet（含直映射 + 偏移映射类型），确认无误后再全量。

### 4.4 后处理

模板自动设置列宽。如需修复行高（避免 `#######`），用 `officecli view <文件> issues --json` 获取 `suggest.rowHeight` 后设置。

### 4.5 验证

```bash
# 结构合法性检查
officecli view <输出文件> issues --json

# 语义逐行对账 — 推荐，按列含义比对
python scripts/verify_semantic.py <源.xlsx> <输出.xlsx>

# 辅助验证
python scripts/verify_totals.py <源.xlsx> <输出.xlsx>     # J 列块总和对比
python scripts/verify_columns.py <源.xlsx> <输出.xlsx>    # 逐列总和对比
```

---

## 五、常见陷阱

1. **公式 S/T 基于净价 O 而非报价 J**：最高危错误——K/L/M/N 从报价扣除后才得净价，用 J 算利润会让指标失真 10%-30%。写完公式后抽样至少 5 个 sheet 的 S/T 值与源对比。

2. **write_data_rows 必须用 data_only=True 读缓存值**：`wb_src_f`（data_only=False）是为探测保留的（看公式字面值），**写入数据时必须用 `wb_src_v`（data_only=True）**。因为源 P/Q 列常含公式（如 `=4.66*3`、`=102%*153.17`、`=J3-L3-M3-N3-O3`），`data_only=False` 读到公式字符串，`safe_num()` 无法解析 → 输出全 0。症状：对账脚本报告 Q/P 列全差异，源有值但输出全 0。

3. **一个 col_map 不能套用所有块**：同一 sheet 不同块表头可能不同。始终用 `smart_col_map_from_header` 每块独立识别，禁止复用。

4. **每块都必须有自己的标题行**：漏写标题 = 数据在但丢失业务语境。写完后检查每个数据块都有 A:X 合并标题行。

5. **标题行和表头位置不固定**：部分 sheet 标题在 R2（R1 为空），表头在 R3。`detect_blocks` 返回 `hr` 字段自动处理。

6. **col_map 反查顺序陷阱**：`write_data_rows` 用 `t→s` 反查，遍历 dict 时第一个匹配的 key 胜出。禁止对同一目标列映射两个源列（如 `{10:10, 11:10}`）。

7. **源数据末尾的"数量小计行"**：A/C/D 空 + G 有数字 + J 空 → `detect_blocks` 自动滤除，不写入输出。

8. **总计行不写入**：源文件 Col 18-24 含"总"字或 `=SUM(...)` 公式的行属于总计行，跳过不写（公式常有 bug、块内 V/W 公式已自动聚合）。

9. **sheet 名称可能含尾随空格**：部分源文件 sheet 名如 `"ELB "`（尾随空格），`officecli outline` 会将其标准化显示为 `ELB`（无空格），导致配置字段写 `"ELB"` 找不到 sheet。用 `openpyxl` 的 `wb.sheetnames` + `repr()` 获取精确名称。

10. **J 列总和正确 ≠ 数据完整**：漏配 K/L/M/N 列时 J 不变，但 O 列公式退化为 `J-0-0-0-0 = J`（净价偏 1.5%）。验证必须用 `verify_semantic.py`（逐行比对）而非仅 `verify_totals.py`（仅 J 列总和）。

11. **文件锁**：Excel 打开源文件时换文件名保存。

12. **对账脚本列号错位**：报告差异时先检查对账脚本中 `src.cell(r, N)` 的列号 N 是否正确。给出差异样本前 5 行 + 源/输出实际值以便判断。

13. **V/W 列宽 12pt 太窄**：V/W 列写 `=IFERROR(SUM(Tn:Tm)/SUM(Sn:Sm),0)` 公式后，`officecli view issues --json` 会报 `text overflow (merged Vn:Vm): 3 lines at 10.0pt need 36pt`。列宽 12 容纳不了 30+ 字符公式。修复：迁移后批量加宽 V/W 列到 16pt。
    ```python
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    wb = load_workbook(OUT)
    for s in wb.sheetnames:
        ws = wb[s]
        if ws.max_column >= 22: ws.column_dimensions[get_column_letter(22)].width = 16  # V
        if ws.max_column >= 23: ws.column_dimensions[get_column_letter(23)].width = 16  # W
    wb.save(OUT)
    ```
    注意：此调整**不会**改变公式计算结果（公式只依赖 T/S 列值），仅消除静态分析告警。

14. **LibreOffice `--outdir` 在 MSYS 终端路径翻倍**：在 git-bash 跑 `soffice --outdir /d/hermes/xlsx/_recalc ...` 时，outdir 会被自动加前缀 `D:\\` 变成 `D:\\d\\hermes\\xlsx\\_recalc\\...`。**修复**：用 Windows 风格路径 `--outdir "D:\\hermes\\xlsx\\_recalc"`，并先 `rm -rf` 防止残留污染；重算完成后用 `cp` 或 Python `shutil.copy` 把重算结果拷回原位（openpyxl 写入时公式未缓存，data_only=True 读不到值）。

15. **`"价格"` 表头未被识别为报价列**：iclima 等 sheet 的某些块表头用 `"价格"` 而非 `"报价"`。旧版 smart_col_map 只有 3 级 fallback（报价 → PRICE → 含报价），匹配不到 `"价格"` → 该块 J 列全为 None。**修复**：在 `"PRICE"` 之后新增第 3 级 exact match `"价格"`。凡是 fallback 链遗漏新词，优先增加一级精确匹配，避免在 contains-报价 模糊搜索中误伤。

16. **排除列表误杀有效报价列**：`"降价报价"` 曾被列入 exclude 元组，但它是有效报价列——埃及询盘块 2 的 Col10。被排除后 fallback 链找到更前的 `"向ELB终端报价"`（Col9），输出错误列值。**经验**：排除列表只排除明确非报价列（历史、目标价、涨跌幅等），`"降价"` 前缀通常有效，不应排除。新增排除词前在源文件中抽样验证。

更多详细陷阱见 `references/MIGRATION_RULES.md` §常见坑。

---

## 六、协作约定

本技能的工作方式遵循以下约定，确保人机协作高效：

- **先测试后全量**：每完成配置修改后生成测试文件 + 简要报告，停下等确认，再全量跑
- **按语义核价**：对账时按列含义（报价/原型机/铜管/结算价/净收入/毛利）比对，而非写死列号
- **收到中断信号立即中止**：用户说"等等"/"停"时立即停止所有运行中进程，生成状态报告
- **分阶段执行**：用户指定"先做前 N 张 sheet"时严格遵守，完成一批后停下汇报

---

## 参考资源

| 文件 | 内容 | 何时查阅 |
|------|------|---------|
| `references/报价汇总-标准化-MEN基线.md` | 项目基线 —— WHY + 架构设计 | 首次了解项目背景 |
| `references/MIGRATION_RULES.md` | 规则手册 —— 详细规则和原理 | 遇到复杂映射或边缘情况 |
| `references/migrate_template.py` | 可执行迁移模板 —— 复制后改 CONFIG 即可运行 | 每次执行迁移时 |
| `references/detect_blocks.py` | 块自动探测器 —— 被模板引用 | 调试探测逻辑时 |
| `scripts/verify_semantic.py` | 语义化逐行对账脚本 | 每次迁移后必运行 |
| `scripts/verify_totals.py` | J 列块总和验证 | 辅助验证 |
| `scripts/verify_columns.py` | 逐列总和验证 | 辅助验证 |
| `scripts/verify_colmap.py` | 列映射冲突检测 | 迁移前检查 col_map |
