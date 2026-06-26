# 报价汇总迁移规则手册

本文档记录了标准化迁移的底层设计原理和完整规则。配合 SKILL.md 使用——SKILL.md 覆盖操作层面，本文档提供原理级细节。

---

## 1. 架构总览

### 1.1 三层模板-实例架构

```
L1 哲学层（333 sheet）→ L2 区域母模板（埃及机型报价模板）→ L3 客户实例（IBC/白鲸/...）
```

### 1.2 单 sheet 多数据块

**铁律**: 一个客户 sheet 可含多个独立数据块，所有块合并到同一个输出 sheet。

```
源 sheet:        标题1 → 表头1 → 数据1... → 空行 → 标题2 → 表头2 → 数据2...
迁移后输出 sheet:  标题1(merged A:X) → 表头1 → 数据1...
                    空行
                    标题2(merged A:X) → 表头2 → 数据2...
```

### 1.3 子块概念

同一标题下，C列为空的行分割子块。每个子块的 W（总盈亏）独立计算。

---

## 2. 24 列 Schema 定义

| # | 列 | 字段名 | 类型 | 必填 | 公式 / 规则 |
|---|-----|--------|------|------|------------|
| 1 | A | 类别/SERIES | 字符串 | ✅ | 原始字符串保留（不映射枚举） |
| 2 | B | 产品类别 | 字符串 | ✅ | — |
| 3 | C | 订单明细 | 字符串 | ✅ | 正则 `^Z[0-9A-Z]{8,}$` |
| 4 | D | 工厂型号 | 字符串 | ✅ | — |
| 5 | E | 配置描述 | 字符串 | ✅ | — |
| 6 | F | 压缩机 | 字符串 | ✅ | — |
| 7 | G | 数量 | 数字 | ✅ | > 0；源 None → 迁移 None |
| 8 | H | 历史报价 | 数字 | ❌ | ≥ 0 |
| 9 | I | 客户目标价 | 数字 | ❌ | ≥ 0 |
| 10 | J | 报价 | 数字 | ✅ | 蓝色输入；源 None → 公式返回 0 |
| 11 | K | 财务费用 | 数字 | ✅ | ≥ 0；源 None → 0 |
| 12 | L | OA信保 | 数字 | ✅ | ≥ 0；源 None → 0 |
| 13 | M | 返点 | 数字 | ✅ | ≥ 0；源 None → 0 |
| 14 | N | 其他费用 | 数字 | ✅ | ≥ 0；源 None → 0 |
| 15 | O | **净价** | **公式** | ✅ | **`=IFERROR(J-K-L-M-N,0)`** |
| 16 | P | 原型机成本 | 数字 | ✅ | > 0；源文本/None → 迁移 None |
| 17 | Q | 铜管成本(3m) | 数字 | ✅ | ≥ 0；源 None → 0 |
| 18 | R | **结算价** | **公式** | ✅ | **`=IFERROR(P+Q,0)`** |
| 19 | S | **净收入** | **公式** | ✅ | **`=IFERROR(O*G,0)`** ← 基于**净价 O**, 不是报价 J |
| 20 | T | **毛利** | **公式** | ✅ | **`=IFERROR((O-R)*G,0)`** ← 基于**净价 O**, 不是报价 J |
| 21 | U | **单机型损益率** | **公式** | ✅ | **`=IFERROR(IF(S=0,0,T/S),0)`** |
| 22 | V | **系列盈亏** | **比率公式** | ❌ | **`=IFERROR(SUM(T系列)/SUM(S系列),0)`** |
| 23 | W | **总盈亏** | **比率公式** | ❌ | **`=IFERROR(SUM(T子块)/SUM(S子块),0)`** |
| 24 | X | 铜管规格 | 数字/文本 | ❌ | 源 None → 0；文本原样保留 |

---

## 3. 7 个派生公式（铁律：P0 总是写入 + IFERROR 防护）

| 列 | 公式 | 写入条件 |
|----|------|---------|
| O | `=IFERROR(J{r}-K{r}-L{r}-M{r}-N{r},0)` | 总是写入 |
| R | `=IFERROR(P{r}+Q{r},0)` | 总是写入 |
| S | `=IFERROR(O{r}*G{r},0)` | 总是写入 |
| T | `=IFERROR((O{r}-R{r})*G{r},0)` | 总是写入 |
| U | `=IFERROR(IF(S{r}=0,0,T{r}/S{r}),0)` | 总是写入 |
| V | `=IFERROR(SUM(T{first}:T{last})/SUM(S{first}:S{last}),0)` | 每个系列首行 |
| W | `=IFERROR(SUM(T{ds}:T{de})/SUM(S{ds}:S{de}),0)` | 每个子块首数据行 |

**P0 修复背景**: 源数据缺值时公式返回 0，等源数据补录后公式自动重算。

**V/W 锚点铁律**:
- V 锚点：每系列首行
- W 锚点：每子块首数据行（ds），不是固定 W3
- V 范围：严格限制在系列内
- W 范围：严格限制在子块内

---

## 4. 样式规则

### 4.1 字体规范

```python
FONT_TITLE    = Font(name='Arial', size=12, bold=True, color='000000')   # 标题行
FONT_HEADER   = Font(name='Arial', size=10, bold=True, color='FFFFFF')   # 表头
FONT_BODY     = Font(name='Arial', size=10)                              # 普通数据
FONT_FORMULA  = Font(name='Arial', size=10, italic=True, color='008000') # 公式（绿）
FONT_INPUT    = Font(name='Arial', size=10, color='0000FF')              # 用户输入（蓝）
FONT_AGGREGATE= Font(name='Arial', size=10, bold=True, color='C00000')   # 聚合公式（红）
```

### 4.2 填充色

```python
FILL_TITLE     = PatternFill('solid', fgColor='D9E1F2')  # 浅蓝（标题）
FILL_HEADER    = PatternFill('solid', fgColor='4472C4')  # 标准蓝（表头）
FILL_INPUT     = PatternFill('solid', fgColor='FFF2CC')  # 浅黄（J 报价输入）
FILL_FORMULA   = PatternFill('solid', fgColor='F2F2F2')  # 浅灰（公式列）
FILL_AGGREGATE = PatternFill('solid', fgColor='E2EFDA')  # 浅绿（V/W 聚合）
FILL_PENDING   = PatternFill('solid', fgColor='FCE4D6')  # 浅橙（待裁决）
```

### 4.3 数字格式

| 列范围 | 格式 | 说明 |
|--------|------|------|
| 1-7 | General | 文本列 + 数量 |
| 8-19 | `\$#,##0.00_);[Red]\(\$#,##0.00\)` | USD + 红字负数 |
| 20 | `\$#,##0.00;\-\$#,##0.00` | USD + 减号负数（毛利） |
| 21-23 | `0.00%` | 百分比 2 位小数 |
| 24 | `\$#,##0.00_);[Red]\(\$#,##0.00\)` | USD（铜管规格如为数字） |

### 4.4 颜色编码语义

| 颜色 | 用途 | 实现 |
|------|------|------|
| 蓝字 #0000FF | 用户硬编码输入 | J 报价列 (FONT_INPUT + FILL_INPUT) |
| 黑字 #000000 | 普通数据 | A-F, G-N, P-Q, X |
| 绿字 #008000 | 派生公式 | O/R/S/T/U (FONT_FORMULA) |
| 红字 #C00000 | 聚合公式 | V/W (FONT_AGGREGATE) |
| 黄底 #FFF2CC | 重要输入 | J 报价列 |
| 灰底 #F2F2F2 | 公式区 | O/R/S/T/U |
| 浅绿底 #E2EFDA | 聚合区 | V/W |
| 橙底 #FCE4D6 | 待裁决 | P 列缺值 |

---

## 5. 合并单元格规则

| 范围 | 锚点 | 规则 |
|------|------|------|
| 标题行 A1:X1 | A1 | 整行合并 |
| 块标题 A{ds}:X{ds} | 块第一行 | 合并承载该块标题 |
| A 列系列 | 系列首行 | 前向填充合并 |
| V 列系列 | 系列首行 | V 公式写入首行后合并整个系列 |
| W 列子块 | 子块首数据行 | W 公式写入锚点后合并到子块末 |

---

## 6. 常见坑

1. **IBC 报价列**: 源 Col K=PRICE → 输出 J，非 Col J=客户目标价
2. **子块分割**: 同一标题下 C 列为空的行分割子块，W 各自独立
3. **MSYS 路径**: 不用 terminal 跑 Python，用 execute_code
4. **文件锁定**: Excel 打开时无法覆盖保存
5. **公式列宽**: 不能按公式文本算列宽，要按显示值（数字）
6. **CJK 宽字符**: 中文/日文约占 2 倍宽度
7. **P 列缺值**: 标橙底 FCE4D6，不填 0
8. **永远不要假设源列和标准列 1:1 对应**: 每个新客户必须先分析源列含义再建立 COL_MAP
9. **S/T 公式必须基于净价 O，不是报价 J**：用 `=J*G` 和 `=(J-R)*G` 是严重语义错误——K/L/M/N 从报价扣除后才得净价，直接用 J 计算会让净收入/毛利失真 10%-30%。正确写法是 `=O*G` 和 `=(O-R)*G`。
10. **块间列含义可能不同**：即使同一个客户 sheet，不同报价块的表头可能完全不同（如块 1: 客户目标价/报价/财务/净价/...；块 2: 报价/净价/返点/...）。不能用一个 col_map 套所有块，必须用 `smart_col_map_from_header` 每块独立识别。
11. **块间 col_map 误复用的隐蔽性**：如果某 sheet 块 2 的 col_map 错位（如把"净价"列误当"财务费用"），输出对账时只有块 2 数据行差异，块 1 完美。症状是部分行对账失败，根因是块间 col_map 误复用。预防：每块 `smart_col_map_from_header()` 自动识别。
12. **始终用智能 col_map，不要硬编码**：不要为每张表手写 col_map 字典。始终用 `smart_col_map_from_header(ws, header_row)` 从源表头逐字自动识别列含义。详见 §7。
13. **智能 col_map 的边界**：如果表头使用非标准词（如"AUX 报价teriak"），`find_col` 退化为模糊匹配，可能误匹配。预防：先抽样表头确认字面。
14. **对账脚本列号可能错位**：报告"大量数据行差异"时，先检查对账脚本中 `src.cell(src_r, N)` 的列号 N 是否对得上源实际列定义。报告差异时给出前 5 行差异样本 + 源/输出实际值以便判断。
15. **块内表头可能在 R3+**：部分 sheet（如贸易商、埃及询盘）R1 空、R2 标题、R3 表头。`detect_blocks` 返回 `hr` 字段自动处理，`smart_col_map_from_header` 传入即可。

---

## 7. 智能 col_map 规则

### 7.1 问题背景

在同一 sheet（如 NICE AIR）中，不同报价块的表头列定义可能完全不同：

| 块 | 表头行 | Col 11 | Col 12 | Col 13 | Col 14 | Col 15 |
|----|--------|--------|--------|--------|--------|--------|
| 块 1 (R2) | `报价 2025.12 铜价92000 汇7.00` | 报价 | 财务费用 | OA信保 | 返点 | 净价 |
| 块 2 (R25) | `报价 2026.6 铜价105000 汇率6.7` | 报价 | **净价** | OA信保 | 返点 | (空) |

如果用同一个列映射套所有块，块 2 会把源 Col 12 "净价"误当"财务费用"写入 K 列——数据静默错位。

### 7.2 smart_col_map_from_header 实现

每块独立从表头行逐字识别列含义：

```python
def smart_col_map_from_header(ws, header_row):
    """从源表头自动识别 col_map. 目标布局固定 24 列."""
    cells = {c: ws.cell(header_row, c).value for c in range(1, 26)}

    def find_col(predicate, exclude_words=()):
        for c in range(8, 25):
            v = cells.get(c)
            if isinstance(v, str) and predicate(v) and not any(w in v for w in exclude_words):
                return c
        return None

    # 报价列 (3 级 fallback)
    quote_col = None
    # 优先级 1: "报价" 独立成词
    for c in range(8, 25):
        if cells.get(c) == "报价":
            quote_col = c; break
    # 优先级 2: "PRICE" (IBC/白鲸英文)
    if not quote_col:
        for c in range(8, 25):
            if cells.get(c) == "PRICE":
                quote_col = c; break
    # 优先级 3: 含"报价"但非排除词
    if not quote_col:
        exclude = ("历史报价", "客户目标价", "客户最近目标价", "目标价", "降价报价", "原报价",
                   "上次报价", "涨跌幅", "涨幅", "4/13报价", "最近一次报价",
                   "上一次报价", "上一次报价（92000/7.0）", "科特迪瓦成交价报价", "TRK成交价（含管）")
        for c in range(8, 25):
            v = cells.get(c)
            if isinstance(v, str) and "报价" in v and v not in exclude:
                quote_col = c; break

    # 净价/原型机/铜管/结算价 (按字面匹配)
    n_price_col = find_col(lambda v: v == "净价")
    proto_col = find_col(lambda v: v.startswith("原型机成本"))
    copper_col = find_col(lambda v: v.startswith("铜管成本"))
    settle_col = find_col(lambda v: v == "结算价")

    # K/L/M/N (财务费用/OA信保/返点/其他费用, 按字面精确匹配)
    k_col = find_col(lambda v: v == "财务费用")
    l_col = find_col(lambda v: v == "OA信保")
    m_col = find_col(lambda v: v == "返点")
    n_col = find_col(lambda v: v == "其他费用")
    if not n_col:
        n_col = find_col(lambda v: v == "其他成本")

    # S (净收入/总金额) 和 T (毛利)
    s_col = find_col(lambda v: v in ("净收入", "总金额"))
    t_col = find_col(lambda v: v == "毛利")

    # 铜管规格/连接管线 (Col 24/25)
    extra_col = None
    for c in (24, 25):
        v = cells.get(c)
        if v and ("铜管" in str(v) or "连接" in str(v)):
            extra_col = c; break

    col_map = {1:1, 2:2, 3:3, 4:4, 5:5, 6:6, 7:7}
    if quote_col: col_map[quote_col] = 10
    if k_col: col_map[k_col] = 11
    if l_col: col_map[l_col] = 12
    if m_col: col_map[m_col] = 13
    if n_col: col_map[n_col] = 14
    if proto_col: col_map[proto_col] = 16
    if copper_col: col_map[copper_col] = 17
    if settle_col: col_map[settle_col] = 18
    if extra_col: col_map[extra_col] = 24

    sheet_meta = {"s_col": s_col, "t_col": t_col}
    return col_map, sheet_meta
```

### 7.3 主循环调用（每块独立识别）

```python
for i, blk in enumerate(blocks):
    # 智能生成 col_map (每块独立, 从该块 header_row 自动识别)
    if "hr" in blk and blk["hr"]:
        col_map, sheet_meta = smart_col_map_from_header(ws_src_f, blk["hr"])
    else:
        col_map, sheet_meta = smart_col_map_from_header(ws_src_f, 2)  # 退化
    # ... 用 col_map 写数据
```

### 7.4 已知边界

- 表头用了非标准词（如 "AUX 报价teriak"）时 `find_col` 退化为"含'报价'非排除词"匹配，**可能误匹配**。预防：先 `get` 抽样表头确认字面。
- 同一 sheet 不同块表头完全重排时（少见但可能），需要扩展 `find_col` 的 exclude 列表。
- 块内表头可能在 R3+（如贸易商/埃及询盘），`detect_blocks` 的 `hr` 字段已正确处理。
