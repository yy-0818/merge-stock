# Stock Merge — 打包指南

仓库不再维护 `build.sh` / `build.bat` 之类的脚本。本文件是统一的打包说明。

## 产物

| 平台 | 产物 | 大小（参考） |
|---|---|---|
| macOS（Apple Silicon / Intel） | `dist/merge_stock.app` | ~540 MB（PySide6 占用） |
| Windows 10/11 | `dist/merge_stock/merge_stock.exe` | ~200 MB |

跨平台不能本地交叉编译——必须**在目标平台上执行打包**。例如在 mac 上跑命令得到 `.app`，在 Windows 上跑命令得到 `.exe`。

## 前置条件

- Python 3.10+（macOS 自带 3.9 也可，PySide6 6.10 要求 3.9+）
- 能访问 PyPI

## 通用步骤

```bash
# 1. 准备虚拟环境
python3 -m venv .build/venv

# Windows:
#   py -3 -m venv .build/venv

# 2. 装依赖（PyInstaller + 运行时）
. .build/venv/bin/activate          # macOS / Linux
# .build\venv\Scripts\activate      # Windows (cmd / PowerShell)

pip install --upgrade pip
pip install openpyxl PySide6 pyinstaller
```

## 打包

### macOS

```bash
# 在项目根目录下
pyinstaller --noconfirm merge_stock.spec
```

得到 `dist/merge_stock.app`。**双击**即可启动；首次启动后 GUI 内填三个路径，或编辑 `merge_stock.app/Contents/Resources/config.json`。

### Windows

```cmd
pyinstaller --noconfirm merge_stock.spec
```

得到 `dist/merge_stock/merge_stock.exe` 与一堆 DLL。

## 验证

```bash
open dist/merge_stock.app                # macOS 启动
# start dist\merge_stock\merge_stock.exe  # Windows 启动
```

启动后：

1. 在 GUI 中填 `src_dir` / `index_file` / `output_dir`
2. 点 **保存路径**
3. 点 **开始合并**
4. 日志区应看到 `[过滤] <子表>: 保留 X / 全部 Y (丢弃 Z)`

## 发送产物

把整个 `dist/merge_stock.app`（macOS）或 `dist/merge_stock/`（Windows）压缩后给同事；他们解压双击即用。**不要只发单文件**——PyInstaller 是一组目录。

## 重新打包时机

- 改了 `gui.py` 或 `merge_stock_files.py`
- 升了依赖版本

不需要做的事：

- 改源码后清缓存——PyInstaller 只追踪 `gui.py` 入口，spec 已是最新版
- 删除 `.build/build/`（PyInstaller 会自己管理）

## 故障排查

| 症状 | 排查 |
|---|---|
| 双击 .app 没反应 | 看 `Console.app` 里的 crash log；可能签名问题，本机首次启动需右键"打开"绕过 Gatekeeper |
| `ModuleNotFoundError: openpyxl` | 漏装依赖，重跑 `pip install openpyxl` |
| Windows 提示缺 MSVCR/VCRUNTIME140.dll | 装 [Microsoft Visual C++ Redistributable](https://aka.ms/vs/17/release/vc_redist.x64.exe) |
| .exe 闪退 | 用 `cmd` 直接运行 `merge_stock.exe`，看 stderr 输出 |

## 目录结构

```
stock/
├── gui.py
├── merge_stock_files.py
├── merge_stock.spec           ← 单一 spec，两平台通用
├── config.example.json
├── 分类.xlsx
├── BUILD.md                   ← 本文件
└── .build/
    ├── venv/
    └── dist/
        ├── merge_stock/       ← COLLECT 中间产物 (Windows exe 在这里)
        └── merge_stock.app/   ← macOS 启动产物
```