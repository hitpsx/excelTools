# 工程概要（供 AI / 新协作者快速上手）

## 这是什么

Excel 工具箱：Tkinter GUI + openpyxl 的 Excel 合并/拆分桌面工具，PyInstaller 打包成单文件 exe 分发给不懂技术的办公用户。全中文界面，提示语刻意使用大白话（无技术黑话），出错弹窗而非日志。

## 架构

- `main.py` —— GUI 层。`ExcelToolApp` 类持有全部界面状态；合并/拆分各一个 Tab（`_ScrollableFrame` 可滚动容器）；长任务放 `threading.Thread`，通过 `root.after(0, ...)` 回主线程更新进度条/状态栏/弹窗。
- `core/merger.py` —— 合并逻辑。4 种模式由 `run_merge(MergeOptions, files, progress_cb)` 分发；`scan_header_differences` 供 GUI 在合并前扫描表头差异。
- `core/splitter.py` —— 拆分逻辑。3 种模式由 `run_split(SplitOptions, input_file, progress_cb)` 分发。
- GUI 层与 core 层通过 `MergeOptions` / `SplitOptions` dataclass + `progress_cb(current, total, msg)` 回调解耦；core 层不得 import tkinter。

## 关键约定

- **只支持 .xlsx**；空行（整行 None/空白）在所有模式下一律跳过。
- 合并模式串：`files_to_sheets` / `files_to_rows` / `files_to_union_rows` / `sheets_to_rows`；拆分模式串：`by_sheet` / `by_rows` / `by_column`。新增模式需同时改 core 分发器和 GUI 单选项。
- 表头语义：`include_header=True` = 仅首文件/首 Sheet 保留表头，其余跳过；`False` = 全部不写表头。
- 按列值拆分时，None/空串/NaN 统一归到 `blank` 分组；输出文件名经 `_safe_filename` 过滤 Windows 非法字符。
- Sheet 名超 31 字符截断，重名自动加 `_N` 后缀。
- 注释/提示语用中文，代码标识符用英文；UI 文案保持大白话风格。

## 构建与测试

- `python test_core.py` —— 纯断言自测，覆盖全部模式及边界（列号超限、空值分组、表头差异），改 core 后必跑。
- `python generate_testdata.py` —— 重新生成 `testdata/`。
- `python build.py` —— 打包（自动清理 dist/build/spec）。

## ⚠️ 硬约束：Windows 7 兼容

- **必须用 Python 3.8.x 构建和打包**（3.9+ 打出的 exe 无法在 Win7 运行）。`requirements.txt` 的版本范围都是按 3.8 兼容选的，升级依赖前先确认还支持 3.8。
- GUI 里 `shcore.SetProcessDpiAwareness` 的 try/except 不能删（Win7 无 shcore.dll）。
- 避免使用 Python 3.9+ 语法（如 `list[str]` 内置泛型——项目用 `from __future__ import annotations` + `typing.List`）。
