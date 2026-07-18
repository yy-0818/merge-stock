# Windows 打包说明

本项目用 PySide6 + PyInstaller 打包。Mac 上 `./build.sh` 已经验证通过。Windows 端在条件满足时同样可以直接执行 `build.bat`。

---

## 1. 前置条件

| 项 | 要求 |
|---|---|
| 操作系统 | Windows 10 / 11（64 位） |
| Python | **3.10 或更高**（PySide6 6.10+ 的硬性要求） |
| 命令 | `python` 命令在 PATH 中（建议 Python 官方安装版，安装时勾选 *Add to PATH*） |
| 磁盘 | 至少 **2 GB 可用空间**（PySide6 + 缓存） |

⚠️ 如果你的 Python 是 3.9 或更低，`pip install PySide6` 会失败（找不到匹配版本）。需要先装 3.10+。

验证方式：

```cmd
python --version
python -m pip --version
```

## 2. 一键打包

在 `stock/` 目录下：

```cmd
build.bat
```

脚本会自动：
1. 建 `.build\venv\` venv
2. 装 `openpyxl`、`PySide6`、`pyinstaller`
3. 调用 PyInstaller 生成 `dist\merge_stock\merge_stock.exe`
4. 把 `config.example.json` 拷成 `config.json` 模板
5. 暂停，让你看最终目录结构

**时间**：第一次约 2-5 分钟（含 PySide6 安装 + 收集 Qt6 插件）。后续增量构建约 30 秒。

## 3. 产物结构

打包完成后的目录：

```
dist\
  merge_stock\
    merge_stock.exe        ← 双击启动 GUI
    _internal\             ← PySide6 + Qt6 运行时（不要删）
    config.json            ← 默认空模板
```

**目录大小**：约 200-250 MB（PySide6 6 占用大头）。正常情况下这是必须的，没法显著缩小。

## 4. 分发给同事

**必须把整个 `dist\merge_stock\` 目录打包成 zip 发送**（仅发 `merge_stock.exe` 会因为找不到 `_internal\` 报错）：

```cmd
cd dist
powershell Compress-Archive -Path merge_stock -DestinationPath merge_stock-windows.zip
```

把 `merge_stock-windows.zip` 发给同事。同事侧：

1. 解压到任意目录（路径**不要有中文/空格**也 OK，但建议放纯英文路径避免意外）
2. 双击 `merge_stock.exe`
3. 第一次启动：编辑同目录的 `config.json` 填三个路径，**或** 直接在 GUI 里点 "选择…" 按钮让文件管理器选
4. 点 "开始合并"，日志区实时显示进度
5. 完成后会自动弹出输出文件夹（取消勾选 "完成后打开输出目录" 可关闭此行为）

## 5. 常见问题

### Q1: `pip install PySide6` 失败 / 太慢
- 换清华源：
  ```cmd
  python -m pip install --index-url https://pypi.tuna.tsinghua.edu.cn/simple openpyxl PySide6 pyinstaller
  ```
- 或临时只装 pip 源：
  ```cmd
  set PIP_INDEX_URL=https://pypi.tuna.tsinghua.edu.cn/simple
  build.bat
  ```

### Q2: `merge_stock.exe` 双击后没反应 / 闪退
- 打开 PowerShell/cmd，切换到该目录，手动执行：
  ```cmd
  cd merge_stock
  .\merge_stock.exe
  ```
- 这时能看到 PyInstaller 的 stderr，能定位问题。

### Q3: 双击后弹窗 "缺少 MSVCP140.dll / VCRUNTIME140.dll"
- 装 [Microsoft Visual C++ 2015-2022 Redistributable (x64)](https://aka.ms/vs/17/release/vc_redist.x64.exe)
- 同事机器几乎肯定有这个组件；如果提示缺失再装。

### Q4: Windows 防火墙弹窗
- PyInstaller 打包的 exe 不联网，但微软的 SmartScreen 会"未识别应用"拦截，需要：
  1. 点 "更多信息" → "仍要运行"
  2. 或给 exe 签名（需要代码签名证书）

## 6. 自检：自动化模式

GUI 里内置一个环境变量驱动的自动测试钩子，方便打包后无人值守验证：

```cmd
set STOCK_MERGE_AUTO_RUN=1
set QT_QPA_PLATFORM=offscreen
merge_stock.exe
```

`STOCK_MERGE_AUTO_RUN=1` 让 GUI 启动后 300ms 自动点 "开始合并"，完成后 500ms 自动退出。
`QT_QPA_PLATFORM=offscreen` 让它在没有显示器的环境也能跑（CI、Docker）。

正常用户使用**不需要**设置这两个变量。

## 7. 文件总览

```
stock\
├── gui.py                 ← GUI 入口（被 PyInstaller 打包）
├── merge_stock_files.py   ← 纯逻辑函数（被 gui.py 引用）
├── merge_stock.spec       ← PyInstaller spec（macOS / Windows 共用）
├── BUILD.md               ← 跨平台打包指南（替代旧 build.sh / build.bat）
├── win-build.md           ← 本文件（Windows 用户/部署补充）
├── config.example.json    ← 默认空模板（打包时拷到产物目录）
├── 分类.xlsx              ← 索引文件（用户准备）
└── <你的源数据目录>\
```

## 8. 输出规则变化（v1.1 起）

合并时启用新版规则：

- **辅助列过滤**：源表最后一列名为「辅助列」，取值为「不显示」的行会被丢弃；「显示」行为「合计」保留。合计行识别 = 型号列 == `合计`；如某表无此字样则取末尾倒数第二非空行作为兜底。
- **报表美化**：表头行深蓝底 + 白字粗体；数据相邻子表带状底色；合计行浅黄 + 粗体；全表灰细边框；冻结首行；列宽按内容自适应（上限 50）。
- **输出列数**：辅助列被丢掉，合并表为 40 列（原来 41 列）。
