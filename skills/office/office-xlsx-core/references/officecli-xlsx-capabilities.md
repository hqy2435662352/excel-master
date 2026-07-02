# OfficeCLI xlsx 能力全景图

> 基于 `officecli help xlsx` 输出（v1.0.116），35 个元素 + sheet 级 20+ 属性的完整清单。
> **不含命令签名**——需要时跑 `officecli help xlsx <element>` 获取实时 schema。

## 基础结构（已淬火）

| 元素 | 说明 | office-master 状态 |
|------|------|-------------------|
| workbook | 工作簿级别属性 | ✓ 使用中 |
| sheet | 工作表（含 freeze/tabColor/hidden/protect/printArea 等 20+ 属性） | ✓ 基础使用，高级属性未触碰 |
| row | 行 | ✓ |
| column | 列（含列宽等） | ✓ |
| cell | 单元格（含 comment/hyperlink/run 子元素） | ✓ 深度使用 |
| range | 范围操作（含 sort 子元素） | ✓ 写入使用 |

## 数据操作

| 元素 | 说明 | 状态 |
|------|------|------|
| table | 结构化表格 | △ 命令存在，未实战 |
| detectedtable | OfficeCLI 自动检测的表 | ✗ |
| pivottable | 数据透视表 | △ 命令存在，未实战 |
| slicer | 透视表切片器（UI 交互） | ✗ |

## 格式与条件

| 元素 | 说明 | 状态 |
|------|------|------|
| conditionalformatting | 条件格式容器 | ✗ |
| cellis / topn / aboveaverage / duplicatevalues / uniquevalues / containstext / dateoccurring / databar / colorscale / iconset / formulacf / cfextended | 13 种条件格式规则 | ✗ 全部未使用 |

## 可视化

| 元素 | 说明 | 状态 |
|------|------|------|
| chart (+ chart-axis + chart-series) | 图表（含轴和系列配置） | △ 命令存在，未实战 |
| sparkline | 迷你图（line/column/winloss） | ✗ |
| picture | 图片（含裁剪/旋转/翻转） | ✗ |
| shape | 形状/绘图 | ✗ |
| ole | OLE 对象嵌入 | ✗ |

## 数据质量

| 元素 | 说明 | 状态 |
|------|------|------|
| validation | 数据验证（下拉列表/数字范围等） | ✗ |
| namedrange | 命名区域（含 scope/comment） | ✗ |
| autofilter | 自动筛选（含多列 criteria） | ✗ |

## 页面布局

| 元素 | 说明 | 状态 |
|------|------|------|
| pagebreak (+ rowbreak + colbreak) | 分页符 | ✗ |

## 逃生舱

| 命令 | 说明 | 状态 |
|------|------|------|
| raw / raw-set / add-part | OpenXML 原始 XML 操作 | ✗ 万能钥匙，从未使用 |

## 批量/文件级命令（OfficeCLI 子命令，非 xlsx 元素）

| 命令 | 说明 | 状态 |
|------|------|------|
| import | CSV/TSV 导入 | ✗ |
| create | 创建空白文档 | ✗ |
| merge | 模板 + JSON 合并 | ✗ |
| move / swap | 元素移动/交换 | ✗ |
| open / close / save | 驻留进程（批量加速） | ✗ |

## 覆盖率统计

- 已淬火（实战验证）：6 / 35 元素 = **17%**
- 命令存在但零实战：4 / 35 = **11%**
- 完全空白：25 / 35 = **71%**
