"""
按 分类.xlsx 把 stock/2026-7-17/ 下的子表拼接成 A 列命名的大表，并按组归档。

输出结构（每组一个文件夹）：
  <OUTPUT_DIR>/<A列名>/
    <A列名>.xlsx       ← 该组所有子表纵向拼接（保留各自表头，按辅助列过滤）
    <子表名>.xlsx      ← 分类表 C 列起的源子表
    ...

规则：
- 每个子表的最后两行通常为 (空行, 合计行)；合计行的特征是「型号 == '合计'」
  兜底：当某张子表没有此字样时，取末尾倒数第二个非空数据行
- 辅助列（最后一列）= 显示 → 保留；不显示 → 丢弃；合计行保留
- 写表时统一丢掉辅助列（输出列数 = num_cols - 1）
- 行内容：补齐 / 截断到统一列数 NUM_COLS（默认 41）；完全空行跳过

样式（在合并文件上自动应用）：
- 表头行：深蓝底 #4472C4 + 白字 + 粗体 + 居中
- 数据行：G 列起按列固定底色（36 色黄金角分布，每列一眼区分其归属档位 1级/2级/D1..D22/A/A1..A12）
- 空数据列不涂色；同一列内空单元格也不涂色，让其它已涂色的格子视觉一致
- 合计行：浅黄 #FFE699 + 粗体（覆盖整行）
- 全表灰色细边框
- 冻结第 1 行
- 列宽按内容自动估算，上限 50

本模块提供：
  load_config()         -> dict
  process(cfg, progress_cb) -> ProcessResult   (供 GUI / CLI 共用)
  main()                                          (CLI 入口)
"""
from __future__ import annotations

import json
import os
import shutil
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


# ------------------- 路径解析 -------------------
def exe_dir() -> Path:
    """脚本运行/被打包后所在目录：开发时是脚本所在目录，打包后是 exe 同级目录。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent.resolve()


# 路径全部留空：打包后的二进制必须依赖同级目录的 config.json 才能运行
DEFAULT_CONFIG = {
    "src_dir": "",
    "index_file": "",
    "output_dir": "",
    "num_cols": 41,
}


def load_config() -> dict:
    """优先读 <exe_dir>/config.json，未提供字段用默认值。"""
    cfg_path = exe_dir() / "config.json"
    cfg = dict(DEFAULT_CONFIG)
    if cfg_path.exists():
        try:
            user_cfg = json.loads(cfg_path.read_text(encoding="utf-8"))
            if isinstance(user_cfg, dict):
                cfg.update({k: v for k, v in user_cfg.items() if k in DEFAULT_CONFIG})
        except Exception as e:
            print(f"[警告] config.json 解析失败，已忽略: {e}")
    else:
        example = exe_dir() / "config.example.json"
        if not example.exists():
            example.write_text(
                json.dumps(DEFAULT_CONFIG, indent=2, ensure_ascii=False),
                encoding="utf-8",
            )
    return cfg


def save_config(cfg: dict, path: Path | None = None) -> Path:
    """把 cfg 写到 <exe_dir>/config.json（GUI 持久化用）。"""
    p = path or (exe_dir() / "config.json")
    p.write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8")
    return p


# ------------------- 数据结构 -------------------
class UnsafeOutputDirError(Exception):
    """output_dir 指向危险路径(如 ~/Desktop)时抛出,需要 cfg['allow_unsafe_output']=True 才放行。"""


def _resolve_real(p: Path) -> Path:
    """解析 symlink、相对路径,得到真实绝对路径。用于和家目录/根目录比较。"""
    return p.expanduser().resolve()


_DANGEROUS_OUTPUT_DIRS: tuple[tuple[Path, str], ...] = (
    (Path.home(), "用户主目录 (~)"),
    (Path.home() / "Desktop", "桌面 (~)"),
    (Path.home() / "Documents", "文档 (~)"),
    (Path.home() / "Downloads", "下载 (~)"),
    (Path("/"), "根目录 (/)"),
)


def _is_dangerous_output_dir(p: Path) -> str | None:
    """返回命中的危险路径说明;否则 None。仅在严格相等(resolve 之后)命中。"""
    real = _resolve_real(p)
    for dangerous, label in _DANGEROUS_OUTPUT_DIRS:
        try:
            if real == dangerous.resolve():
                return label
        except Exception:
            continue
    return None


def _safe_prepare_output_dir(output_dir: Path, allow_unsafe: bool = False) -> None:
    """清理并准备 output_dir。

    安全策略(防 rmtree 误删):
    1. resolve 后,若等于 ~/Desktop ~/Documents ~/Downloads ~ / 之一 → 抛 UnsafeOutputDirError
       (除非显式传入 allow_unsafe=True;此选项仅供 CLI 高级用户)
    2. 若目录已存在,先统计文件数;不阻断,但通过日志让调用方知情
    3. 上述都通过才执行 rmtree + mkdir
    """
    if not allow_unsafe:
        label = _is_dangerous_output_dir(output_dir)
        if label is not None:
            raise UnsafeOutputDirError(
                f"output_dir 指向危险位置: {label}\n"
                f"  → 为了避免误删整个 {label},脚本拒绝使用此路径作为输出。\n"
                f"  请新建一个子目录(例如 ~/work/stock/_output)再设置 output_dir。"
            )

    if output_dir.exists():
        shutil.rmtree(output_dir)
    output_dir.mkdir(parents=True)


@dataclass
class ProcessResult:
    groups_total: int = 0
    groups_merged: int = 0
    rows_total: int = 0                # 过滤后写出的行数（不含"不显示"被丢的行）
    rows_filtered_total: int = 0       # 过滤掉多少行（"不显示"）
    files_copied: int = 0
    missing_files: list[str] = field(default_factory=list)
    log: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return self.groups_merged > 0


# ------------------- 样式常量 (professional client-facing preset) -------------------
# 设计目标: 报表可直接邮件给客户 → 干净、专业、易读、有视觉层次
#
# 视觉系统:
#   - 表头:深海军蓝底 + 白色粗体, 高度 32px, 冻结首行
#   - 数据行:相邻 section 用浅/深交替(斑马纹), 隔 5 列轻微色相切换增强可读性
#   - 数字右对齐(便于上下比较);文本列(型号/类型)居中
#   - 合计行:温暖橙黄底 + 粗体 + 上方双线,视觉跳出
#   - 全表细线边框, 颜色柔和不刺眼
#   - 客户名首列(A):浅灰底,视觉锚定 "这是谁的库存"
#   - 行高 22px, 列宽自适应 + padding 2
#
HEAD_FILL = PatternFill("solid", fgColor="1F3864")                # 深海军蓝
HEAD_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
HEAD_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)

TOTAL_FILL = PatternFill("solid", fgColor="E65100")       # 深橙(最醒目)
TOTAL_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
TOTAL_ALIGN = Alignment(horizontal="center", vertical="center")

# 数据行
DATA_FONT = Font(name="Calibri", size=11, color="1F1F1F")
DATA_FONT_NUM = Font(name="Calibri", size=11, color="1F1F1F")
DATA_FONT_BOLD = Font(name="Calibri", bold=True, size=11, color="1F1F1F")
DATA_ALIGN_TEXT = Alignment(horizontal="center", vertical="center", wrap_text=True)
DATA_ALIGN_NUM = Alignment(horizontal="right", vertical="center")
DATA_ALIGN_FIRST = Alignment(horizontal="left", vertical="center", indent=1)   # 客户名首列

# Section 斑马纹 (偶数/奇数交替)
SECTION_FILL_EVEN = PatternFill("solid", fgColor="F0F0F0")   # 浅灰(偶数 section)
SECTION_FILL_ODD = PatternFill("solid", fgColor="FFFFFF")     # 纯白(奇数 section)
SECTION_FILL_NONE = PatternFill(fill_type=None)

# 客户名首列底色
CLIENT_FILL = PatternFill("solid", fgColor="EAEFF5")                # 浅蓝灰

# 边框
BORDER_SIDE = Side(style="thin", color="BDBDBF")
BORDER_THICK = Side(style="medium", color="1F3864")
BORDER = Border(left=BORDER_SIDE, right=BORDER_SIDE, top=BORDER_SIDE, bottom=BORDER_SIDE)
TOTAL_BORDER = Border(left=BORDER_THICK, right=BORDER_THICK, top=BORDER_THICK, bottom=BORDER_THICK)

# ------------------- 数据列底色调色 -------------------
# 4 色强对比循环: 蓝 / 绿 / 黄 / 橙 (每 4 列换色)
# 饱和度 0.45 (中等), 亮度 0.92 (淡,数字清晰可读)
def _build_column_fills(n: int = 24) -> list[PatternFill]:
    fills: list[PatternFill] = []
    import colorsys
    palette = [
        (0.60, 0.92, 0.45),   # 蓝色  (D1-D4)
        (0.32, 0.92, 0.45),   # 绿色  (D5-D8)
        (0.14, 0.92, 0.45),   # 黄色  (D9-D12)
        (0.06, 0.92, 0.45),   # 橙色  (D13-D16)
        (0.60, 0.87, 0.45),   # 蓝色2 (D17-D20)
        (0.32, 0.87, 0.45),   # 绿色2 (D21-D22 + A组)
        (0.14, 0.87, 0.45),   # 黄色2
        (0.06, 0.87, 0.45),   # 橙色2
    ]
    for i in range(n):
        h, l, s = palette[i % len(palette)]
        r, g, b = colorsys.hls_to_rgb(h, l, s)
        hexc = "{:02X}{:02X}{:02X}".format(int(r * 255), int(g * 255), int(b * 255))
        fills.append(PatternFill("solid", fgColor=hexc))
    return fills


COLUMN_FILLS = _build_column_fills(24)
COLUMN_FONT = Font(name="Calibri", size=11, color="1F1F1F")


def _is_data_column_name(header_name: str) -> bool:
    """判断表头名是否属于"应上底色的数据列"。
    数据列的语义:每行有非零/非空数字(D1..D22 / A / A1..A12 / 1级 / 2级)

    注意:**换算率、备注**这些"辅助信息列"**不上底色**——它们要么是单个数字、要么是文字评论,
    视觉上和数据矩阵混在一起反而混乱(而且我们最终也不会输出这些列,所以更不需要识别)。
    """
    if header_name is None:
        return False
    n = _normalize_header(header_name)
    if not n:
        return False
    # 1 级 / 2 级
    if n in ("1级", "2级", "一级", "二级", "等级1", "等级2"):
        return True
    # D1..D22
    if len(n) >= 2 and n[0] == "d" and n[1:].isdigit():
        return True
    # A / A1..A12  (排除 "辅助列" 等别名,因为 _normalize 已去空格,纯 "a" 才命中)
    if n == "a" or (n.startswith("a") and n[1:].isdigit()):
        return True
    return False


def _compute_used_data_cols(
    ws, sections: list[dict], header_rows: set[int], cm: ColumnMap | None
) -> set[int]:
    """返回「数据列」中实际有非空数据的 Excel 列号(1-based)集合。

    改为按**表头名字**判定谁是数据列 (vs 之前按">=7 + helper_idx" 位置):
    - 凡表头匹配 _is_data_column_name:1级 / 2级 / D1..D22 / A / A1..A12 / 换算率 / 备注 ...
    - 凡表头匹配 cm.helper_idx / cm.nosort_idx / 型号 / 类型 / 客户名:不上色
    - 这种按名字匹配方式天然兼容源表列位置的随意挪动

    空列不参与上色,严格满足「空行空列不处理」需求。
    """
    # 找第一个 header 行,解析哪些 ws 列号是"数据列"
    if not header_rows:
        return set()
    first_header = min(header_rows)
    data_cols_ws: set[int] = set()
    # 排除集合 (ws 列号 1-based)
    excluded: set[int] = set()
    if cm is not None:
        # ws 列号 = cm.col_idx + 1 - (cm.col_idx 之前有几个被丢的列)
        # 简化:遍历第一行,逐个判断名字,排除被识别的列名
        pass
    for c in range(1, ws.max_column + 1):
        name = ws.cell(first_header, c).value
        if _is_data_column_name(name):
            data_cols_ws.add(c)
    # 排除特殊列:模型 / 类型 — 它们在 ws 中已重映射到不同列号
    # 但在我们的实现里,我们**已经在 _parse_column_map** 阶段标记了 model / type 列
    # helper / nosort 列已从 ws 中删除,所以 ws 中已不存在
    # 因此 data_cols_ws 已经天然排除了 helper/nosort(已被 _strip 删除)
    #     也天然排除了客户名(它的表头是 None / 空字符串, _is_data_column_name 返回 False)
    # 只需要再排除: cm.model_idx 和 cm.type_idx 对应的 ws 列号
    for orig_idx in (cm.model_idx if cm else None, cm.type_idx if cm else None):
        if orig_idx is None or orig_idx < 0:
            continue
        # 翻译 orig_idx → ws col:假设每张子表的 strip 不变(在 _apply_styles 调用前都是同结构),
        # 用第一个 section 的数据行做样本是不行的。
        # 我们改用遍历第一行(表头行),看哪个 ws 列号对应 model/type 的名字
        # 但更简单:遍历 ws 的所有列,匹配名字
        for c in range(1, ws.max_column + 1):
            n = ws.cell(first_header, c).value
            if n is None:
                continue
            target = ("型号" if orig_idx == cm.model_idx else "类型") if cm else ""
            if target and _normalize_header(n) == _normalize_header(target):
                excluded.add(c)
                break
    data_cols_ws -= excluded

    # 过滤"实际有数据"的列
    used: set[int] = set()
    for sec in sections:
        for r in range(sec["data_first_row"], sec["data_last_row"] + 1):
            if r in header_rows:
                continue
            if _is_total_row_in_ws(ws, r, cm):
                continue
            for c in data_cols_ws:
                v = ws.cell(r, c).value
                if v is None:
                    continue
                if isinstance(v, str) and not v.strip():
                    continue
                used.add(c)
    return used


def _is_total_row_in_ws(ws, row_idx_1based: int, cm: ColumnMap | None) -> bool:
    """判断已写到 ws 的某行是否为合计行。用 column_map 决定「型号」所在 Excel 列。"""
    model_col_1based: int | None = None
    if cm is not None and cm.model_idx is not None:
        # 模型:helper_idx 通常是原表的最后一列(原 N 列),输出 ws 只保留到第 helper_idx 列。
        # 但「型号 / 类型」是写到 helper_idx 之前的位置,所以 col = cm.model_idx + 1
        model_col_1based = cm.model_idx + 1
    else:
        model_col_1based = 2  # 兜底
    v = ws.cell(row_idx_1based, model_col_1based).value
    return bool(v) and isinstance(v, str) and v.strip() == "合计"


# ------------------- 数据读取 & 过滤 -------------------
def _is_blank_row(row: list) -> bool:
    return all(v is None or (isinstance(v, str) and not v.strip()) for v in row)


def _read_sub_file_rows(path: Path, num_cols: int) -> tuple[list[list], ColumnMap]:
    """读取子表全部行 + 解析表头列映射。

    返回 (rows, column_map):
    - rows: 已补齐 / 截断到 num_cols,跳过完全空行的二维数组
    - column_map: 「型号」「类型」「辅助列」的位置(基于表头名字识别,源表列顺序可变化)
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    raw: list[list] = []
    header: list | None = None
    for row in ws.iter_rows(values_only=True):
        row_list = list(row)
        if len(row_list) < num_cols:
            row_list.extend([None] * (num_cols - len(row_list)))
        else:
            row_list = row_list[:num_cols]
        if _is_blank_row(row_list):
            continue
        if header is None:
            header = row_list
        raw.append(row_list)
    cm = _parse_column_map(header) if header else ColumnMap()
    return raw, cm


# ------------------- 表头列名识别 (容错:源表列位置可变化) -------------------
# 源表格中可能存在这些列(其它列保持原顺序输出):
#   - 型号    :用于识别「合计行」,主排序键 2
#   - 类型    :主排序键 1
#   - 辅助列  :最后一列(可能命名「显示」「辅助」「是否显示」「隐藏」等),决定行是否被丢弃
# 无论你把这些列挪到 A/B/C/...,只要表头名字匹配,脚本都能工作。
COL_NAME_MODEL = "型号"
COL_NAME_TYPE = "类型"
# 辅助列的表头可能有多种写法(逐项匹配)
COL_NAME_HELPER_ALIASES = ("显示", "辅助", "是否显示", "显示列", "显示控制", "隐藏")
# 「辅助列 / 辅助」是过滤控制(显示/不显示); 「辅助2」是排序控制(不排序 / 空)
COL_NAME_HELPER_FILTER_ALIASES = ("显示", "是否显示", "显示列", "显示控制", "隐藏", "辅助列", "辅助")
COL_NAME_HELPER_SORT_ALIASES = ("辅助2", "不排序", "不排序列", "排序控制", "辅助 2", "aux2")
DEFAULT_HELPER_COL_IDX = -1   # 默认仍是最后一列(向后兼容旧表)
DEFAULT_NOSORT_COL_IDX = -1   # 默认不存在该列


@dataclass
class ColumnMap:
    """从一张子表的表头行解析出的列映射。所有索引都是 0-based 数组下标,
    适用于 _read_sub_file_rows 返回的 row 列表。

    关键字段:
    - model_idx:  「型号」所在列(用于合计行识别 + 排序次键)
    - type_idx:   「类型」所在列(用于排序主键)
    - helper_idx: 过滤控制列(显示 / 不显示; 不显示的行被丢)
    - nosort_idx: 排序控制列(不排序 / 空; 标了"不排序"的行不参与排序,保留原顺序)
    - max_used_col_idx: 表头最大下标(用于数据列染色范围判定)
    """
    model_idx: int | None = None   # 「型号」所在列
    type_idx: int | None = None    # 「类型」所在列
    helper_idx: int = DEFAULT_HELPER_COL_IDX   # 过滤控制列(默认最后一列)
    nosort_idx: int | None = None  # 排序控制列(默认不存在 → 全量排序)
    max_used_col_idx: int = 0


def _normalize_header(v) -> str:
    """表头名归一化:去空白、统一小写、去掉所有空格与全角空格,用于模糊匹配。"""
    if v is None:
        return ""
    s = str(v).strip().casefold()
    return s.replace(" ", "").replace("\u3000", "")


def _parse_column_map(header: list) -> ColumnMap:
    """从表头行构建 ColumnMap。

    字段识别规则:
    - 「型号」「类型」:精确匹配(casefold)
    - helper(filter): 「显示 / 辅助 / 是否显示 / 显示列 / 显示控制 / 隐藏 / 辅助列」任一别名
                      — 取**最后**一个匹配列,通常对应源表那个"最后显示控制列"
    - nosort:        「辅助2 / 不排序列 / 排序控制 / 辅助 2 / aux2」任一别名

    找不到对应列时:
    - 型号/类型 → None(兜底退化)
    - helper → -1(原约定:辅助视为末列,自动被丢)
    - nosort → None(不启用"不排序"功能,全部数据行参与排序)
    """
    cm = ColumnMap()
    for i, raw in enumerate(header):
        name = _normalize_header(raw)
        if name == "型号" and cm.model_idx is None:
            cm.model_idx = i
        elif name == "类型" and cm.type_idx is None:
            cm.type_idx = i
        else:
            # 别名匹配:按优先级顺序遍历,后写的覆盖前面的(谁靠后谁赢,符合用户把"辅助列"挪到中间的场景)
            filter_match = any(_normalize_header(a) == name for a in COL_NAME_HELPER_FILTER_ALIASES)
            nosort_match = any(_normalize_header(a) == name for a in COL_NAME_HELPER_SORT_ALIASES)
            if filter_match:
                cm.helper_idx = i
            if nosort_match:
                cm.nosort_idx = i
    cm.max_used_col_idx = len(header) - 1 if header else 0
    return cm


def _is_total_row(row: list, cm: ColumnMap) -> bool:
    """判定某行是否为合计行:型号列 == '合计'。
    型号列可能已被挪到任意位置(通过 column_map 解析);找不到则按 False。
    """
    idx = cm.model_idx if cm is not None else None
    if idx is None or idx >= len(row):
        return False
    v = row[idx]
    return isinstance(v, str) and v.strip() == "合计"


def _fallback_total_idx(rows: list[list], cm: ColumnMap) -> int | None:
    """兜底:没有任何型号=='合计'的行时,取末尾倒数第二个非空数据行作为合计。"""
    if len(rows) < 2:
        return None
    return len(rows) - 2


def _filter_rows(
    rows: list[list], cm: ColumnMap
) -> tuple[list[list], list[list], list[list], int, int]:
    """返回 (header_row, data_rows, total_rows, drop_count, total_count)。

    - header_row: 第一行(子表的列名)单独保留
    - total_rows: 第一个型号=='合计'的行;兜底用最后第二个非空数据行
    - data_rows: 其余数据行
    - 辅助列(ColumnMap.helper_idx)== '不显示' 且不是合计行 → 丢弃
    - 辅助列未匹配时回退到最后一列
    """
    total_idx = None
    for i, r in enumerate(rows):
        if i == 0:
            continue
        if _is_total_row(r, cm):
            total_idx = i
            break
    if total_idx is None:
        total_idx = _fallback_total_idx(rows, cm)

    helper_idx_raw = cm.helper_idx if cm is not None else -1
    drop = 0
    header: list[list] | None = None
    totals: list[list] = []
    data: list[list] = []
    for i, r in enumerate(rows):
        if i == 0:
            header = r
            continue
        if i == total_idx:
            totals.append(r)
            continue
        helper = None
        if helper_idx_raw is not None and len(r) > 0:
            resolved = helper_idx_raw if helper_idx_raw >= 0 else len(r) + helper_idx_raw
            if 0 <= resolved < len(r):
                helper = r[resolved]
        if isinstance(helper, str) and helper.strip() == "不显示":
            drop += 1
            continue
        data.append(r)
    return header or [], data, totals, drop, len(rows)


# ------------------- 数据排序 -------------------
def _sort_key_for_data_row(row: list, cm: ColumnMap):
    """生成数据行的排序键:(类型, 型号)。

    - 优先用 cm.type_idx / cm.model_idx
    - 表头中找不到对应列时,退化到旧位置 row[2] / row[1]
    - None 排到末尾
    - 字符串 .casefold() 排序
    """
    def normalize(v):
        if v is None:
            return None
        if isinstance(v, str):
            return v.strip().casefold()
        return str(v).casefold()

    if cm is not None:
        type_idx = (cm.type_idx if cm.type_idx is not None else 2)
        model_idx = (cm.model_idx if cm.model_idx is not None else 1)
    else:
        type_idx, model_idx = 2, 1
    a = normalize(row[type_idx]) if type_idx < len(row) else None
    b = normalize(row[model_idx]) if model_idx < len(row) else None
    return (
        (1, "") if a is None else (0, a),
        (1, "") if b is None else (0, b),
    )


def _is_nosort_row(row: list, cm: ColumnMap) -> bool:
    """判断数据行是否被 nosort 列(辅助2 / 不排序列)标记为「不排序」。

    仅当 cm.nosort_idx 存在(>= 0)且行长度足够时才检查;否则返回 False(全部参与排序)。
    """
    if cm is None or cm.nosort_idx is None or cm.nosort_idx < 0:
        return False
    if cm.nosort_idx >= len(row):
        return False
    v = row[cm.nosort_idx]
    if isinstance(v, str):
        s = v.strip()
        if not s:
            return False
        return s in ("不排序", "否", "N", "n", "No", "NO", "no", "0", "false", "False")
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v == 0
    return False


def _sort_data_rows(rows: list[list], cm: ColumnMap) -> list[list]:
    """对数据行按 类型 → 型号 升序排序,保留「不排序」标记行的原顺序。

    排序策略:
    1. 把数据行拆成 pinned(标了不排序,按原顺序)+ sortable(未标,正常排序)
    2. sortable 按 (类型, 型号) 排序后,接在 pinned 之后
       (pinned 在前:用户标"不排序"通常想让这些行优先展示,例如特殊型号置顶)
    3. 在稳定排序下,sortable 内同类型同型号仍保持原顺序

    没 cm.nosort_idx 时退化为纯排序(向后兼容)。
    """
    if cm is None or cm.nosort_idx is None:
        return sorted(rows, key=lambda r: _sort_key_for_data_row(r, cm))
    pinned, sortable = [], []
    for r in rows:
        if _is_nosort_row(r, cm):
            pinned.append(r)
        else:
            sortable.append(r)
    sortable_sorted = sorted(sortable, key=lambda r: _sort_key_for_data_row(r, cm))
    return pinned + sortable_sorted


# ------------------- 报表样式应用 -------------------
def _visual_width(s: str) -> float:
    """估算单元格内容在 Excel 列宽单位下的视觉宽度。

    - CJK(中日韩)统一按 2 个单位计(Excel 中英字约 1 字符宽 = 7 px,中文约 14 px)
    - 数字/英文/常见 ASCII 按 1 个单位计
    - 全角符号(·、—、… 等)按 2 计
    """
    if not s:
        return 0.0
    w = 0.0
    for ch in s:
        code = ord(ch)
        # CJK 基本区 + 扩展 A-F + 标点 + 全角符号
        if (
            0x1100 <= code <= 0x115F          # Hangul Jamo
            or 0x2E80 <= code <= 0x303E       # CJK 标点 / 部首
            or 0x3041 <= code <= 0x33FF       # 平假名 / 片假名 / CJK 符号
            or 0x3400 <= code <= 0x4DBF       # CJK 扩展 A
            or 0x4E00 <= code <= 0x9FFF       # CJK 基本
            or 0xA000 <= code <= 0xA4CF       # 彝文
            or 0xAC00 <= code <= 0xD7A3       # 韩文音节
            or 0xF900 <= code <= 0xFAFF       # CJK 兼容
            or 0xFE30 <= code <= 0xFE4F       # CJK 兼容形式
            or 0xFF00 <= code <= 0xFF60       # 全角 ASCII / 全角标点
            or 0xFFE0 <= code <= 0xFFE6       # 全角符号
        ):
            w += 2.0
        else:
            w += 1.0
    return w


def _autosize_columns(ws, min_w: float = 8.0, max_w: float = 22.0) -> None:
    """按列内容自适应宽度(用每列最长 cell 的真实视觉宽度 + padding)。

    设计原则:
    - **必须**用 max(不是 P95):少数长型号(如 `IN12P004GL60120`)若被截,客户读不出来
    - 默认 padding = 2,空出 1 字符边距
    - 上限 max_w = 22(防止极长 token 拉爆整张表;超长用截断或换行兜底)
    - 下限 min_w = 8(数字列 "1234567" 也至少要装下)
    - CJK 字符按 2 个单位计(用 _visual_width)
    """
    from openpyxl.utils import get_column_letter
    for col_idx, col_cells in enumerate(ws.columns, 1):
        max_len = 0.0
        for cell in col_cells:
            if cell.value is None:
                continue
            s = str(cell.value)
            max_len = max(max_len, _visual_width(s))
        # padding = 2 (左边距 1 + 右边距 1)
        target = max_len + 2
        # 软上限:若 target > max_w, 仍给到 max_w;超过部分会被显示为溢出
        # (但绝大多数情况 max_w=22 足够)
        width = max(min_w, min(max_w, target))
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def _is_numeric_header(name: str) -> bool:
    """判断表头是否代表"数字列"(用于决定右对齐 + 数字格式)。
    1级 / 2级 / D1..D22 / A / A1..A12 都是数字列。
    """
    if name is None:
        return False
    n = _normalize_header(name)
    if not n:
        return False
    if n in ("1级", "2级", "一级", "二级", "等级1", "等级2"):
        return True
    if len(n) >= 2 and n[0] == "d" and n[1:].isdigit():
        return True
    if n == "a" or (n.startswith("a") and n[1:].isdigit()):
        return True
    return False


def _apply_styles(ws, sections: list[dict], cm: ColumnMap | None = None) -> None:
    """应用专业报表样式 (client-facing preset)。

    视觉系统:
    1. 表头行:深海军蓝底 + 白色粗体 + 居中 + 行高 32 + wrap_text
    2. 数据行:section 斑马纹 (奇浅 / 偶白)
       - 文本列(型号/类型):居中 + 允许换行 (避免超长型号挤出)
       - 数字列(1级/2级/D..A):右对齐 + 千分位格式
       - 客户名首列(A):浅蓝灰底 + 左对齐 + 缩进 1
    3. 合计行:橙黄底 + 粗体 + 顶边加粗 + 全行右对齐(数字)/居中(文本)
    4. 数据列底色循环 (蓝族 12 + 紫族 12),帮助识别列归属
    5. 全表细线边框 (D9D9D9),不刺眼
    6. 冻结首行 + 冻结首列 (A 列客户名)

    cm (ColumnMap) 用于判定「合计行」(型号列位置由 cm 决定)。
    """
    n_cols = ws.max_column
    n_rows = ws.max_row
    header_rows = {sec["header_row"] for sec in sections}

    # 把每个数据列标上序号 + 数据列属性 (用于背景色 + 数字格式)
    # 用 _is_data_column_name 找数据列
    data_col_ordered: list[tuple[int, str]] = []   # (ws 列号 1-based, header 名)
    for c in range(1, n_cols + 1):
        name = ws.cell(1, c).value
        if _is_data_column_name(name):
            data_col_ordered.append((c, str(name) if name else ""))

    # ===== 1) 全表边框 =====
    for row in ws.iter_rows(min_row=1, max_row=n_rows, max_col=n_cols):
        for cell in row:
            cell.border = BORDER

    # ===== 2) 表头行 =====
    for sec in sections:
        hr = sec["header_row"]
        for cell in ws[hr]:
            cell.fill = HEAD_FILL
            cell.font = HEAD_FONT
            cell.alignment = HEAD_ALIGN
        ws.row_dimensions[hr].height = 36

    # ===== 3) 数据行染色 + 对齐 =====
    # 对每个 section 内部按行号判断"奇偶行"做斑马纹 (section 内重置计数,跨 section 重新开始)
    # section 整体偶数/奇数 (在分组中的位置) 也切换底色,但为了客户视觉清爽,
    # 这里采用"全表统一白底 + 数据列底色循环",不去做斑马纹(避免和数据列底色冲突)
    for sec_idx, sec in enumerate(sections):
        for r in range(sec["data_first_row"], sec["data_last_row"] + 1):
            is_total = _is_total_row_in_ws(ws, r, cm)
            for cell in ws[r]:
                col = cell.column
                cell.border = BORDER if not is_total else TOTAL_BORDER
                if is_total:
                    # 合计行
                    cell.fill = TOTAL_FILL
                    cell.font = TOTAL_FONT
                    if col == 1:
                        cell.alignment = DATA_ALIGN_FIRST
                    else:
                        # 数字列右对齐,文本列居中
                        name = ws.cell(sec["header_row"], col).value if sec["header_row"] <= ws.max_row else None
                        if _is_numeric_header(name):
                            cell.alignment = DATA_ALIGN_NUM
                            cell.number_format = '#,##0;-#,##0;""'
                        else:
                            cell.alignment = TOTAL_ALIGN
                    continue

                # 数据行
                # section 斑马纹: 偶数 section 浅灰底, 奇数 section 白底
                sec_zebra = SECTION_FILL_EVEN if sec_idx % 2 == 0 else SECTION_FILL_ODD
                if col == 1:
                    # 客户名首列: 蓝色浅底 + 粗体
                    cell.fill = CLIENT_FILL
                    cell.font = DATA_FONT_BOLD
                    cell.alignment = DATA_ALIGN_FIRST
                else:
                    # 检查是否在数据列内
                    data_idx = next((i for i, (wc, _) in enumerate(data_col_ordered) if wc == col), None)
                    if data_idx is not None and cell.value is not None and not (isinstance(cell.value, str) and not cell.value.strip()):
                        # 数据列有值:上列底色 + 数字格式
                        fill = COLUMN_FILLS[data_idx % len(COLUMN_FILLS)]
                        cell.fill = fill
                        cell.font = COLUMN_FONT
                        cell.alignment = DATA_ALIGN_NUM
                        if isinstance(cell.value, (int, float)):
                            cell.number_format = '#,##0;-#,##0;""'
                        elif isinstance(cell.value, str):
                            try:
                                v = float(cell.value)
                                cell.value = v
                                cell.number_format = '#,##0;-#,##0;""'
                                cell.alignment = DATA_ALIGN_NUM
                            except (ValueError, TypeError):
                                cell.alignment = DATA_ALIGN_TEXT
                    else:
                        # 非数据列 / 空格: section 斑马纹底色
                        cell.fill = sec_zebra
                        cell.font = DATA_FONT
                        cell.alignment = DATA_ALIGN_TEXT

    # ===== 4) 行高 =====
    for r in range(2, n_rows + 1):
        if r in header_rows:
            continue
        ws.row_dimensions[r].height = 22

    # ===== 5) 冻结首行 + 客户列 =====
    ws.freeze_panes = "B2"

    # ===== 6) 自适应列宽 =====
    _autosize_columns(ws)


# ------------------- 合并主函数 -------------------
def _build_merged_file(
    group: dict,
    available: set[str],
    src_dir: Path,
    output_dir: Path,
    num_cols: int,
    log: list[str],
) -> tuple[Path, int, int] | None:
    """返回 (path, rows_kept, rows_dropped). 仍然无数据时返回 None."""
    out_name = group["a_col"]
    if not out_name:
        return None

    group_dir = output_dir / out_name
    group_dir.mkdir(parents=True, exist_ok=True)

    merged_path = group_dir / f"{out_name}.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = out_name[:31]

    # 默认 Workbook 已含 1 行；删除它并把指针归零，避免 header_row 不等于 1
    ws.delete_rows(1)
    current_row = 0

    output_col_count = num_cols - 1  # 丢掉最后一列(即辅助列)
    sections: list[dict] = []
    rows_kept_total = 0
    rows_dropped_total = 0
    # 全局 cm:以第一张有效子表为准,后续写出 / 排序 / 染色均按它进行。
    # (分类表的源表通常用同一模板,但即使不同也会被合并到同一 ws,只要头部一致即可)
    global_cm: ColumnMap | None = None

    for sub in group["sub_files"]:
        candidate = src_dir / f"{sub}.xlsx"
        if sub not in available or not candidate.exists():
            log.append(f"  [跳过-缺文件] {sub}")
            continue
        raw_rows, cm = _read_sub_file_rows(candidate, num_cols)
        if global_cm is None:
            global_cm = cm
        header_row, data_rows, total_rows, drop, total = _filter_rows(raw_rows, cm)
        # 数据行排序:类型 → 型号 升序(表头与合计行不参与排序)
        data_rows_sorted = _sort_data_rows(data_rows, cm) if data_rows else []
        kept_rows = ([header_row] if header_row else []) + data_rows_sorted + total_rows
        log.append(
            f"  [过滤] {sub}: 保留 {len(kept_rows)} / 全部 {total}  (丢弃 {drop})"
        )

        if not kept_rows:
            continue

        # 决定"丢哪些列":
        # 1) 辅助列 helper_idx (默认最后一列)
        # 2) 排序列 nosort_idx (若存在)
        # 3) 表头名为「换算率 / 备注」(用户明确要求删除,不展示)
        # 4) 兜底:cm 啥都没识别时丢最后一列(向后兼容老数据)
        # 5) 保留源表所有其它列,包括全 None 的空列(用户明确要求保留做视觉对齐)
        drop_cols: set[int] = set()
        h_idx = cm.helper_idx
        if h_idx is not None:
            resolved = h_idx if h_idx >= 0 else (num_cols - 1 if num_cols > 0 else -1)
            if resolved >= 0:
                drop_cols.add(resolved)
        n_idx = cm.nosort_idx
        if n_idx is not None and 0 <= n_idx < num_cols:
            drop_cols.add(n_idx)
        # 丢掉名为"换算率"或"备注"的列(遍历表头找)
        hidden_aliases_norm = {_normalize_header(a) for a in ("换算率", "换算", "备注")}
        # header_row 是来自 raw_rows[0], 索引 0-based
        if cm.helper_idx is not None and header_row:
            for i, name in enumerate(header_row):
                if _normalize_header(name) in hidden_aliases_norm:
                    if 0 <= i < num_cols:
                        drop_cols.add(i)
        # 兜底:若什么都没识别,且 num_cols > 0,丢最后一列(老数据)
        if not drop_cols and num_cols > 0:
            drop_cols.add(num_cols - 1)
        # 大的下标先丢(从右往左删,不会导致下标错位)
        drop_cols_sorted = sorted(drop_cols, reverse=True)

        def _strip(row: list) -> list:
            out = row
            for c in drop_cols_sorted:
                if 0 <= c < len(out):
                    out = out[:c] + out[c + 1:]
            return out

        # 写表:每张子表"第一行(表头)+ 后续(数据+合计)",丢 cm 指定的辅助列
        current_row += 1
        section_header_row = current_row
        ws.append(_strip(kept_rows[0]))                   # 表头
        for r in kept_rows[1:]:
            current_row += 1
            ws.append(_strip(r))                          # 数据 + 合计行
        data_first_row = section_header_row + 1
        data_last_row = current_row

        rows_kept_total += len(kept_rows)
        rows_dropped_total += drop
        sections.append({
            "header_row": section_header_row,
            "data_first_row": data_first_row,
            "data_last_row": data_last_row,
        })

    if rows_kept_total == 0:
        wb.close()
        log.append(f"[空组跳过] {out_name} (无任何子表匹配或全部被过滤)")
        shutil.rmtree(group_dir)
        return None

    # 列数补齐到统一 output_col_count（防止短行被错认）
    if ws.max_column < output_col_count:
        for r in range(1, ws.max_row + 1):
            row = ws[r]
            for _ in range(output_col_count - ws.max_column):
                # openpyxl 在追加时已自动扩到 max；这里是兜底防御
                pass

    _apply_styles(ws, sections, global_cm)
    wb.save(merged_path)
    return merged_path, rows_kept_total, rows_dropped_total


def _copy_source_files(
    group: dict,
    available: set[str],
    src_dir: Path,
    output_dir: Path,
) -> int:
    out_name = group["a_col"]
    group_dir = output_dir / out_name
    copied = 0
    for sub in group["sub_files"]:
        src = src_dir / f"{sub}.xlsx"
        dst = group_dir / f"{sub}.xlsx"
        if sub not in available or not src.exists():
            continue
        shutil.copy2(src, dst)
        copied += 1
    return copied


def load_classification(index_path: Path) -> list[dict]:
    """读取分类表，返回每行配置: {a_col, b_col, sub_files}。"""
    wb = openpyxl.load_workbook(index_path, data_only=True)
    ws = wb.active
    rows: list[dict] = []
    for row in ws.iter_rows(values_only=True):
        a_val, b_val = row[0], row[1]
        if a_val is None and b_val is None:
            continue
        sub_files = [str(v).strip() for v in row[2:] if v is not None and str(v).strip()]
        rows.append(
            {
                "a_col": str(a_val).strip() if a_val is not None else "",
                "b_col": str(b_val).strip() if b_val is not None else "",
                "sub_files": sub_files,
            }
        )
    return rows


# ------------------- 进程入口 -------------------
ProgressCb = Callable[[int, int, str], None]


def process(cfg: dict, progress_cb: ProgressCb | None = None) -> ProcessResult:
    """主入口。cfg 需含 src_dir/index_file/output_dir/num_cols。
    progress_cb(current_idx, total, message) 在每个组开始时调用一次。"""
    result = ProcessResult()

    src_dir = Path(cfg["src_dir"])
    index_file = Path(cfg["index_file"])
    output_dir = Path(cfg["output_dir"])
    num_cols = int(cfg.get("num_cols", 41))

    if not src_dir or src_dir == Path("."):
        raise ValueError("src_dir 未配置")
    if not index_file or index_file == Path("."):
        raise ValueError("index_file 未配置")
    if not output_dir or output_dir == Path("."):
        raise ValueError("output_dir 未配置")
    if not index_file.exists():
        raise FileNotFoundError(f"索引文件不存在: {index_file}")
    if not src_dir.exists():
        raise FileNotFoundError(f"源目录不存在: {src_dir}")

    _safe_prepare_output_dir(output_dir, allow_unsafe=bool(cfg.get("allow_unsafe_output", False)))

    available = {f[:-5] for f in os.listdir(src_dir) if f.lower().endswith(".xlsx")}
    groups = load_classification(index_file)
    result.groups_total = len(groups)
    if not groups:
        return result

    for idx, g in enumerate(groups, start=1):
        out_name = g["a_col"]
        log = result.log
        log.append(f"=== {out_name} ===")
        if progress_cb:
            progress_cb(idx - 1, len(groups), f"正在合并: {out_name}")
        merged = _build_merged_file(g, available, src_dir, output_dir, num_cols, log)
        if merged is None:
            continue
        merged_path, kept, dropped = merged
        result.rows_total += kept
        result.rows_filtered_total += dropped
        copied = _copy_source_files(g, available, src_dir, output_dir)
        result.files_copied += copied
        result.missing_files.extend(
            sub for sub in g["sub_files"]
            if sub not in available or not (src_dir / f"{sub}.xlsx").exists()
        )
        log.append(f"  合并文件: {merged_path.name}  (有效 {kept} 行，过滤 {dropped} 行)")
        log.append(f"  源子表: 已复制 {copied} / {len(g['sub_files'])} 个")
        result.groups_merged += 1
        if progress_cb:
            progress_cb(idx, len(groups), f"完成: {out_name}")

    return result


# ------------------- CLI 入口 -------------------
def main() -> int:
    cfg = load_config()
    print(f"SRC_DIR    = {cfg['src_dir']}")
    print(f"INDEX_FILE = {cfg['index_file']}")
    print(f"OUTPUT_DIR = {cfg['output_dir']}")
    print(f"NUM_COLS   = {cfg.get('num_cols', 41)}")
    print()

    try:
        result = process(cfg, progress_cb=lambda c, t, m: None)
    except (ValueError, FileNotFoundError) as e:
        print(f"[错误] {e}")
        return 1

    print("\n".join(result.log))
    print()
    print(f"[完成] 合并 {result.groups_merged} / {result.groups_total} 组，"
          f"有效 {result.rows_total} 行，过滤 {result.rows_filtered_total} 行，"
          f"复制 {result.files_copied} 个源文件。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
