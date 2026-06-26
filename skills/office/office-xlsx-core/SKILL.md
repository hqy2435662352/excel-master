---
name: office-xlsx-core
description: 任何 .xlsx 文件操作前务必加载此 skill——告诉你在冻结窗格、条件格式、图表、迷你图、数据验证、批处理、CSV导入等场景下最优工具是什么、为什么、怎么最快。不加载的代价：用 openpyxl 写 30 行做一个一行命令的事。需要具体命令签名时查 officecli help xlsx。
version: 4.1.0
tags: [excel, xlsx, officecli, decision-framework, tier-classification]
---

# office-xlsx-core — xlsx 操作决策框架

**核心原则：本 skill 不维护命令字典。** OfficeCLI 是活的真相源——当需要了解命令签名、属性名、参数格式时，直接查 CLI：

```bash
officecli help xlsx                  # 列出所有 xlsx 元素
officecli help xlsx <element>        # 某元素的完整属性签名 + 示例
officecli help all --jsonl           # 全量导出 JSON（一次性全面了解）
```

本 skill 的价值在 CLI 做不到的事：**面对一个 xlsx 任务时，告诉你最优解是什么、为什么、怎么最快。**

**随附文件**：
- `references/officecli-xlsx-capabilities.md` — OfficeCLI xlsx 能力全景图
- `references/token-costs.md` — 报价汇总文件实测 Token 消耗数据
- `references/CAPABILITY_GAPS.md` — **已知能力缺口**（v1 边界，遇到时触发需求讨论）
- `scripts/officecli-safe` — Wrapper 脚本，硬拦截裸 `query "cell"` + 500KB 截断兜底

---

## 一、环境配置（每次新终端必设）

```bash
export MSYS_NO_PATHCONV=1
```

Git Bash 会把 `/body`、`/Sheet1/A1` 等 OfficeCLI 元素路径当成文件系统路径转换。`MSYS_NO_PATHCONV=1` 全局禁用此行为。

---

## 二、强制工作流

### 2.1 标准操作序列

```
备份 → 观察 → 操作 → 验证 → 报告
 cp      view    set/add  validate  JSON
```

**备份和验证不可跳过。** 如果文件被 Excel 占用无法 cp，换文件名保存（`_v2.xlsx`）。

### 2.2 修改前确认（经济方式）

| ✅ 允许 | 🚫 禁止 |
|---------|---------|
| `view stats`（0.03K token） | `get /Sheet --depth 1 --json`（全量 dump） |
| `get /Sheet/G2 --json`（抽样 2-3 个） | 全列逐行 `get` |
| `query "cell[type=Error]"`（精确筛选） | `query "cell"`（裸奔 → wrapper 拦截） |

### 2.3 标准输出 JSON（强制）

**每个任务结束时必须输出此 JSON。这不是可选的——不输出等于任务未完成。** 放在回复的最后一段，用代码块包裹：

```json
{
  "status": "SUCCESS | PARTIAL | FAILED | REJECTED",
  "output_file": "path/to/result.xlsx",
  "backup_file": "path/to/result.backup.xlsx",
  "issues_found": 0,
  "summary": "简短描述做了什么"
}
```

### 2.4 不确定属性名 → 查 help

宁可查 help 多花 0.5 秒，不要猜错属性名导致命令失败 + 重试。

---

## 三、渐进式阅读协议（三级）

`dump` 输出的 JSON 不直接全量读取，按三级阶梯升级。

### Step 1：看骨架（默认，零上下文消耗）
```python
import json, os
path = "spec.json"
size_kb = os.path.getsize(path) / 1024
print(f"文件大小: {size_kb:.0f} KB")
with open(path) as f:
    data = json.load(f)
if isinstance(data, list):
    print(f"条目数: {len(data)}")
    if data: print(f"首条 keys: {list(data[0].keys())}")
else:
    print(f"顶层 keys: {list(data.keys())}")
```

### Step 2：照局部（需要更多信息时）
```python
def peek_context(data, err_index, window=5):
    start = max(0, err_index - window)
    end = min(len(data), err_index + window + 1)
    for i in range(start, end):
        marker = ">>>" if i == err_index else "   "
        print(f"{marker} [{i}] {data[i]}")
```

### Step 3：全量读取（硬触发条件）
只在以下条件**同时满足**时允许：
1. 同一任务 Python 脚本连续报错 ≥ 3 次
2. 局部切片无法定位根因
3. 调用前输出 `## ESCALATION TO FULL READ` 自证（含已尝试方式、失败原因、预估大小）
4. 全量读取后优先用 `batch --input` 回放，避免再次 dump→read 循环

---

## 四、Tier 1：日常必用（13 项，每次都要能用）

每条包含：**命令模式 + 决策要点**。不写完整命令签名（查 `officecli help xlsx <element>`）。

### T1.1 看结构
```bash
officecli view <file> outline          # Sheet 列表 + 表名
officecli view <file> stats            # 行数/列数/公式数/错误数
officecli view <file> issues --json    # 问题检测（溢出/断链/公式错误）
```
**决策**：永远是第一步。`outline` → `stats` → `issues`，渐进。不要一上来就 `get`。

### T1.2 精准读取
```bash
officecli get <file> /Sheet1/A1 --json           # 单单元格
officecli get <file> /Sheet1 --depth 0 --json    # Sheet 元信息
officecli query <file> "cell[type=Error]" --json # 精确筛选
```
**决策**：抽样 2-3 个关键单元格确认格式/值，不要逐列遍历。`query` 必须带选择器——裸 `cell` 被 wrapper 硬拦截。

### T1.3 改值
```bash
officecli set <file> /Sheet1/A1 --prop value="2026" --json
officecli set <file> /Sheet1/A1 --prop formula="SUM(A1:A10)" --json
```
**决策**：改值最简单，但改前确认当前值（`get` 抽样），避免 no-op。

### T1.4 批操作
```bash
officecli dump <file> /Sheet1 -o spec.json       # 导出
# → Python 处理 spec.json                        # 本地计算
officecli batch <file> --input spec.json --json  # 回写
```
**决策**：数据量 > 20 行时必用 dump→batch 模式。数据不进入 Agent 上下文——零 Token 成本。

### T1.5 格式
```bash
officecli set <file> /Sheet1/A1:A100 --prop bold=true --prop fill=FFFF00
officecli set <file> /Sheet1/A1:A100 --prop font.color=red --prop font.size=12
officecli set <file> /Sheet1/A1:A100 --prop alignment.horizontal=center
```
**决策**：范围操作一次性设置，不要逐单元格。改格式前先 `get` 抽样——如果已经是要设的值，跳过。

### T1.6 冻结窗格
```bash
officecli set <file> /Sheet1 --prop freeze=A2    # 冻结首行
officecli set <file> /Sheet1 --prop freeze=B3    # 冻结首行+首列
```
**决策**：一行命令。不要用 openpyxl 写 5 行做同一件事。

### T1.7 自动筛选
```bash
officecli add <file> /Sheet1 --type autofilter --prop ref=A1:Z100
```
**决策**：套在表头范围上。如果已经有 table，table 自带筛选，不需要重复加。

### T1.8 条件格式
```bash
officecli add <file> /Sheet1 --type conditionalformatting \
  --prop ref=G2:G100 --prop rule="cellValue>5000" \
  --prop fill=FF4444 --prop font.color=FFFFFF
```
**决策**：三种核心规则类型——cellValue / formula / colorScale。简单规则（如「大于某值红底」）直接用 CLI，别用 openpyxl。

### T1.9 套用表格
```bash
officecli add <file> /Sheet1 --type table \
  --prop ref=A1:Z100 --prop headerRow=true \
  --prop style="TableStyleMedium2"
```
**决策**：table 自动带筛选 + 斑马纹 + 结构化引用。如果只是要斑马纹不要筛选，用条件格式 `formula=MOD(ROW(),2)=0`。

### T1.10 验证
```bash
officecli validate <file>                         # OpenXML schema 校验
officecli view <file> issues --json               # 问题清单
```
**决策**：任何写入操作后必跑。`validate` 是 schema 级，`view issues` 是语义级（行高溢出、公式求值错误）。

### T1.11 CSV 导入
```bash
officecli import <file> /Sheet1 data.csv          # CSV → xlsx
```
**决策**：一条命令替代 Python csv + openpyxl 10 行。

### T1.12 打印设置
```bash
officecli set <file> /Sheet1 --prop printArea=A1:Z100
officecli set <file> /Sheet1 --prop printTitleRows="1:2"
```
**决策**：`printTitleRows` = 每页重复打印的表头行。`printArea` = 打印范围。

### T1.13 数据验证
```bash
officecli add <file> /Sheet1 --type validation \
  --prop ref=D2:D100 --prop type=list \
  --prop formula1='"是,否"'
```
**决策**：三种常用 type——list（下拉）/ whole（整数范围）/ decimal。注意 list 的 formula1 需要双层引号。

---

## 五、Tier 2：需要决策框架（6 项，不需要独立 skill）

### T2.1 图表
```bash
officecli add <file> /Sheet1 --type chart \
  --prop chartType=bar --prop dataRange="Sheet1!A1:B10" \
  --prop categories="Sheet1!A2:A10" --prop title="Sales"
```
**决策树**：
- 简单柱/折/饼图 → `officecli add chart`，1 条命令搞定
- 双轴/组合图/复杂配色 → 考虑 openpyxl（chart 元素组合复杂时 CLI 参数太多）
- 「只是想看数据趋势」→ 优先考虑 `sparkline`（T2.2），更轻量

常见 chartType：bar / line / pie / scatter / area

### T2.2 迷你图
```bash
officecli add <file> /Sheet1 --type sparkline \
  --prop ref=H2 --prop dataRange="Sheet1!B2:G2" \
  --prop type=line
```
**决策**：单元格内嵌的趋势线。适合「每行一个趋势」的场景。比 chart 轻量得多。

### T2.3 命名区域
```bash
officecli add <file> --type namedrange \
  --prop name="SalesData" --prop ref="Sheet1!A1:B100"
```
**决策**：给区域起名，后续公式用 `=SUM(SalesData)` 代替 `=SUM(Sheet1!A1:B100)`。

### T2.4 排序
```bash
officecli sort <file> /Sheet1/A1:B100 --by A --order asc
```
**决策**：单列排序直接用 CLI。多列排序（先按 A 升序再按 B 降序）可能需要 openpyxl。

### T2.5 工作表保护
```bash
officecli set <file> /Sheet1 --prop protect=true --prop password="1234"
```
**决策**：保护后默认所有单元格锁定。如果只想保护部分区域，先 `set <range> --prop locked=false` 解锁不需要保护的单元格，再 protect。

### T2.6 标签颜色
```bash
officecli set <file> /Sheet1 --prop tabColor=FF0000
```
**决策**：零成本美化。常用颜色编码：红=注意/待审核，绿=通过，黄=草稿。

---

## 六、最优工具选择矩阵

| 场景 | 最优工具 | 原因 |
|------|---------|------|
| 单元素操作（Tier 1 全部） | `officecli add/set` | 1 条命令，无 Python 开销 |
| 批量格式修改 | `officecli set <range>` | 比 openpyxl 逐行循环快 |
| 批量数据写入 | `officecli batch --input spec.json` | dump→改→回写，零上下文 |
| 复杂逻辑（块探测/列映射/条件判断） | openpyxl (Python) | 需要 Python 判断逻辑 |
| CLI 不原生支持 | `raw-set` / `add-part` | OpenXML 万能逃生舱 |
| 数据导入 | `officecli import` | 一条命令 |
| 模板填充 | `officecli merge` | 模板 + JSON 一键输出 |
| 创建空白 | `officecli create` | 比 openpyxl Workbook() 更轻 |
| 元素移动/重排 | `officecli move / swap` | 比删+插简单 |
| 单步声明式操作（条件格式/数据验证/命名区域）且 Token 敏感 | openpyxl 一条 FormulaRule / DataValidation | officecli 的完整 JSON 往返开销可能超过操作本身的价值。如果只需要一条规则且 token 预算紧张，openpyxl 更省 |

**反模式**
- 用 openpyxl 写 30 行做一个 `officecli set --prop freeze=A2` 的事
- 逐个 `officecli set` 循环，而不是 `set <range>` 批量
- 明明有 `officecli import` 却 Python csv + openpyxl 逐行写

---

## 七、只读操作优先级（从省 Token 到费 Token）

| 需求 | 首选（省） | 禁止（费） |
|------|-----------|-----------|
| 看结构 | `view outline` | `get / --json` |
| 看行数 | `view stats` | `get /Sheet --depth 1 --json` |
| 扫描问题 | `view issues --json` | 逐单元格 `query` |
| 确认单值 | `get /Sheet/A1 --json`（抽样 2-3 个） | 全列 loop |
| 精确筛选 | `query "cell[type=Error]"` | `query "cell"` 裸奔 |
| 大量数据 | `dump /Sheet -o spec.json`（进文件） | `query` / `get --json`（进上下文） |

---

## 八、常用查询选择器

| 选择器 | 含义 |
|--------|------|
| `cell:has(formula)` | 公式单元格 |
| `cell:empty` | 空单元格 |
| `row[Col>5000]` | 某列值大于阈值 |
| `row[Col~=text]` | 某列包含文本 |
| `cell[type=Number or type=Date]` | 数字/日期类型 |

---

## 九、常见坑

### Git Bash 路径转换（Windows）
**症状**：`/body` → `C:/Program Files/Git/body`
**修复**：`export MSYS_NO_PATHCONV=1`

### Shell 路径引号
**症状**：`/slide[1]` → `no matches found`
**修复**：单引号包裹：`'/slide[1]'`

### 全量 dump 大文件
**症状**：dump 输出 JSON 超上下文
**修复**：按 Sheet 分块 — `officecli dump file.xlsx /Sheet1 -o spec.json`

### 冗余写入
**症状**：20 条 `set bold=true` 但原始文件已是 bold（no-op 白烧 Token）
**修复**：格式 set 前先抽样 `get`，已是目标值则跳过

### 公式错误不一定是 bug
部分公式（VLOOKUP 查源数据、模板预留公式）是有意为之。先标记，不盲目覆盖。

### 文件被 Excel 占用（PermissionError）
**症状**：openpyxl 保存时 `PermissionError`
**修复**：换文件名保存（`_v2.xlsx`），或先关闭 Excel。

### officecli 路径必须用 Windows 风格
**症状**：`officecli validate "/d/hermes/xlsx/file.xlsx"` → `File not found: D:\d\...`
**修复**：始终用 `D:\hermes\xlsx\file.xlsx`。`/d/...` 在 git-bash 能 `ls` 但 officecli 不认。

---

## 🔒 硬拦截（wrapper 已保护）

以下操作被 `officecli-safe` wrapper 自动拒绝，Agent 无需自行判断：
- 裸 `query "cell"`（无选择器）→ 直接拒绝
- 输出 > 500KB → 截断 + 提示用 `dump -o` 代替

---

## ⚠️ 任务结束检查清单

在回复用户之前，确认以下三项全部完成：

- [ ] 是否输出了标准 JSON 报告？（§2.3 — 强制）
- [ ] 是否跑了 `validate` + `view issues`？（§2.1 — 不可跳过）
- [ ] 备份文件是否存在？（§2.1 — 不可跳过）

这三项缺一不可。缺任何一项，任务不算完成。
