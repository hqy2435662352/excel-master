# excel-master

专门处理 Excel (.xlsx) 文件的 Hermes Agent profile。基于 [OfficeCLI](https://github.com/iOfficeAI/OfficeCLI) + openpyxl 双引擎，Tier 1/2 决策框架覆盖 19 种日常 xlsx 操作。

## Agent 部署指南

以下指令可直接交给 Hermes Agent 自动执行。每步末尾附验证命令——Agent 应确认 `[OK]` 后继续。

### 前提条件

- Windows 10+，已安装 [Hermes Agent](https://hermes-agent.nousresearch.com)
- 网络可访问 GitHub

---

### 步骤 1：安装 OfficeCLI

```powershell
# PowerShell（管理员）
irm https://d.officecli.ai/install.ps1 | iex
```

```bash
# 验证
officecli --version
# [OK] 输出类似 "officecli v1.0.x"
```

### 步骤 2：创建 profile 骨架

```bash
hermes profile create excel-master
```

```bash
# 验证
ls "$HOME/AppData/Local/hermes/profiles/excel-master/SOUL.md"
# [OK] 文件存在
```

### 步骤 3：部署本仓库文件

```bash
PROFILE_DIR="$HOME/AppData/Local/hermes/profiles/excel-master"
cp SOUL.md "$PROFILE_DIR/"
cp -r skills/office "$PROFILE_DIR/skills/"
cp .gitignore "$PROFILE_DIR/"
```

```bash
# 验证
ls "$PROFILE_DIR/skills/office/office-xlsx-core/SKILL.md"
# [OK] 文件存在
ls "$PROFILE_DIR/skills/office/office-xlsx-core/scripts/officecli-safe"
# [OK] 文件存在
```

### 步骤 4：配置 API key

```bash
cp config.yaml.template "$PROFILE_DIR/config.yaml"
```

然后编辑 `config.yaml`，将 `YOUR_DEEPSEEK_API_KEY` 替换为实际 key。

可选：修改 `model.default` 为其他模型（如 `deepseek-v4-pro` 开思考模式）。

```bash
# 验证
grep -c "YOUR_DEEPSEEK_API_KEY" "$PROFILE_DIR/config.yaml"
# [OK] 输出 0（说明已替换）
```

### 步骤 5：验证 profile 可用

```bash
hermes profile use excel-master
hermes profile show
```

```bash
# 验证：启动 excel-master 并测试
# 发送 "hello, 确认你已加载 office-xlsx-core skill"
# [OK] Agent 回复中确认已加载
```

---

## 依赖

| 依赖 | 版本 | 用途 |
|------|------|------|
| [OfficeCLI](https://github.com/iOfficeAI/OfficeCLI) | ≥ v1.0.116 | xlsx CLI 操作引擎 |
| [Hermes Agent](https://hermes-agent.nousresearch.com) | 最新 | Agent 运行时 |
| openpyxl | 随 OfficeCLI 环境 | 复杂数据处理回退方案 |

---

## 架构

```
excel-master/
├── SOUL.md                      # 人格层：xlsx 独占 + 需求洁癖 + 反模式嗅觉 + 冷幽默
├── config.yaml.template         # 配置模板（复制后填 key）
├── .gitignore                   # 白名单策略，隔离运行时文件
└── skills/office/
    ├── office-xlsx-core/        # 核心：Tier 1(13项) + Tier 2(6项) 决策框架
    │   ├── SKILL.md
    │   ├── scripts/officecli-safe    # Token 防线 wrapper
    │   └── references/
    │       ├── CAPABILITY_GAPS.md    # 已知能力缺口
    │       ├── officecli-xlsx-capabilities.md
    │       └── token-costs.md
    └── xlsx-quote-standardization/   # 示例：空调报价标准化（业务专用 skill）
```

---

## 能力边界

| 能做（v1） | 不能做（已知缺口，触发需求讨论） |
|-----------|-------------------------------|
| 查看结构/问题/统计 | 数据透视表 + 切片器 |
| 改值/改格式/改公式 | 公式审计/依赖分析 |
| 冻结窗格/筛选/条件格式/表格 | raw XML 操作 |
| 图表/迷你图/命名区域 | 多工作簿合并/外部引用 |
| 数据验证/CSV 导入/打印设置 | openpyxl 深度（块探测/列映射） |
| 批操作（dump → Python → batch） | 超大文件（>200MB）未实测 |

详见 `skills/office/office-xlsx-core/references/CAPABILITY_GAPS.md`
