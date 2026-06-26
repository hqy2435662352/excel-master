# Token 消耗实测数据

> 测试文件：报价汇总.-更新.xlsx（433KB，21 Sheet，17,291 单元格，2,359 公式，93 错误单元格）

## query 命令输出大小实测

| 命令 | 匹配数 | 输出大小 | 占上下文 |
|------|--------|---------|---------|
| `query "cell"` | 17,291 | **10.4 MB** | 🚫 致命 |
| `query "cell:has(formula)"` | 2,359 | 1.8 MB | 🚫 重度 |
| `query "cell[value>0]"` | ~7,000 | 2.4 MB | 🚫 重度 |
| `query "cell[type=Error]"` | 93 | **72 KB** | ✅ 安全 |
| `query "merge"` | 若干 | 107 KB | ✅ 安全 |

## 关键发现

1. **433KB xlsx → 10.4MB JSON**：膨胀比 25:1。不是 officecli 的问题，是单元格级结构化 JSON 的固有特性。
2. **筛选条件决定一切**：`cell[type=Error]` 只返回 93 条（72KB），完全可用。不带筛选的 `query "cell"` 返回全表（10.4MB），一次清空额度。
3. **即使用筛选，匹配数仍是关键**：`cell:has(formula)` 匹配 2,359 条 → 1.8MB，仍需截断。
4. **`view issues --json` 比 `query` 聚合度更高**：175 个 issues 只有几十 KB，是公式/错误扫描的首选方式。

## 轻量替代方案（省 Token）

| 需求 | 命令 | 典型消耗 |
|------|------|---------|
| 文件结构 | `view stats` | ~300 bytes |
| Sheet 大纲 | `view outline` | ~500 bytes |
| 问题扫描 | `view issues --json` | ~2-5 KB |
| 单值确认 | `get /Sheet/A1 --json` | ~300 bytes |
| 大量数据 | `dump /Sheet -o spec.json` | 进文件，不进上下文 |
