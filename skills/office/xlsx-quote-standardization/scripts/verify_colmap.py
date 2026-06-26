"""
COL_MAP 冲突检测器 — 在执行迁移前扫描所有 col_map，找出：
1. 同一目标列写了多个源列（dict 顺序陷阱）
2. 关键目标列 (10=J, 16=P, 17=Q, 24=X) 缺失映射

使用：
    python verify_colmap.py <migrate_script.py>

输出：每条冲突打印源文件:行号 + 重复的目标列号 + 候选源列。
退出码 1 = 有冲突（必须修复），0 = OK。
"""
import ast
import sys
from pathlib import Path

KEY_TARGETS = {
    1: "类别",
    2: "产品类别",
    3: "订单明细",
    4: "工厂型号",
    5: "配置描述",
    6: "压缩机",
    7: "数量",
    10: "报价 (J)",
    16: "原型机成本 (P)",
    17: "铜管成本 (Q)",
    24: "铜管规格 (X)",
}


def find_dicts(tree):
    """找出所有 dict literal 并返回 (lineno, dict_items_list)"""
    out = []
    for node in ast.walk(tree):
        if isinstance(node, ast.Dict):
            items = []
            for k, v in zip(node.keys, node.values):
                if isinstance(k, ast.Constant) and isinstance(v, ast.Constant):
                    items.append((k.value, v.value))
            if items:
                out.append((node.lineno, items))
    return out


def main(path):
    src = Path(path).read_text(encoding="utf-8")
    tree = ast.parse(src)
    dicts = find_dicts(tree)

    errors = 0
    seen_dups = set()
    missing_keys = {}

    for lineno, items in dicts:
        # Skip dicts that look like HEADERS_24 or COL_FMT (string-keyed)
        if all(isinstance(k, str) for k, _ in items):
            continue
        # 找目标列冲突
        target_to_sources = {}
        for s, t in items:
            target_to_sources.setdefault(t, []).append(s)

        for t, sources in target_to_sources.items():
            if len(sources) > 1 and t not in seen_dups:
                seen_dups.add(t)
                # 只关心数字键
                if isinstance(t, int):
                    err = f"  L{lineno}: 目标列 {t} ← 源列 {sources}（多个）→ 第 1 个胜出，会覆盖其余"
                    if t in KEY_TARGETS:
                        err += f"   ⚠ 这是关键列 [{KEY_TARGETS[t]}]"
                    print(err)
                    errors += 1

        # 关键列缺失检查（仅对含较多键的 dict，认为是 col_map）
        if len(items) >= 5:
            present_targets = {t for _, t in items if isinstance(t, int)}
            for k, name in KEY_TARGETS.items():
                if k not in present_targets:
                    key = (lineno, k)
                    if key not in missing_keys:
                        print(f"  L{lineno}: 关键列 {k}({name}) 未映射（可能是有意跳过 P 列缺值？）")
                        missing_keys[key] = True

    print()
    if errors:
        print(f"✗ {errors} 个冲突需修复")
        return 1
    print("✓ 所有 COL_MAP 无冲突")
    return 0


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python verify_colmap.py <migrate_script.py>")
        sys.exit(2)
    sys.exit(main(sys.argv[1]))
