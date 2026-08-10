# CLI 用法速查 — 方便脱离 GUI 调试

合并逻辑入口脚本：`merge_stock_files.py`，所有合并规则均在此处实现。

## 一键演示（最快验证）

```bash
.build/venv/bin/python3 merge_stock_files.py --demo
```

会在 `/tmp/stock_demo/` 生成：
- `src/subA.xlsx` （含 `色A/色C` 两个全空数据列 + 末尾 `辅助列/换算率/备注`）
- `src/subB.xlsx` （列顺序与 subA 不同，含 `辅助列/辅助2`）
- `idx.xlsx` （分类表，组 A → [subA, subB]）

然后立即执行一次合并，并把日志打到 stdout。

## 自定义路径

```bash
.build/venv/bin/python3 merge_stock_files.py \
  --src <源子表目录> \
  --index <分类表.xlsx> \
  --out <输出目录> \
  --unsafe-output  # 若 out 在 src 树下需要它
```

不传任何参数 → 读 `./config.json`。

## 只看分类表结构（不实际合并）

```bash
.build/venv/bin/python3 merge_stock_files.py --demo --dry-run
```

## 演示数据自定义目录

```bash
.build/venv/bin/python3 merge_stock_files.py --demo --demo-dir ~/Desktop/stock_demo
```

默认演示数据落在 `~/stock_demo/`，合并产物在 `~/stock_demo/out/<时间戳>/<组名>/`。

## 常用调试片段

```python
# 直接调用核心逻辑（无需 GUI/CLI）
import sys; sys.path.insert(0, '/Users/kaiyuan/work/stock')
import merge_stock_files as core

result = core.process({
    "src_dir": str(Path.home() / "stock_demo/src"),
    "index_file": str(Path.home() / "stock_demo/idx.xlsx"),
    "output_dir": str(Path.home() / "stock_demo/out"),
    "allow_unsafe_output": True,
})
print("\n".join(result.log))
print("actual_output_dir =", result.actual_output_dir)
# 实际产物: actual_output_dir / <组名> / <组名>.xlsx
```

## 列数规则（2026-08）

1. **预剔除**（每张子表）：按表头名命中 `辅助列 / 辅助2 / 换算率 / 备注` 一律删除，与数据无关
2. **自动取宽**：合并目标列数 = 组内所有子表（预剔除后）的最大真实列数
3. **空列压缩**（整组合并后）：数据行全空列剔除（表头空不影响），同时更新 ColumnMap 索引
