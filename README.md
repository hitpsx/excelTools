# Excel 工具箱 (ExcelTools)

一个面向普通办公用户的 Excel 整理小工具：把零散的 Excel 表合并成一份，或把一份大表拆成多份。GUI 基于 Tkinter，提示语全部使用大白话，出错直接弹窗。

## 功能

### 合并（4 种模式）

- **多文件 → 多 Sheet**：多个 Excel 合成一个工作簿，每个文件单独放一个工作表
- **多文件 → 单 Sheet 堆叠**：所有内容追加到同一张工作表（可选是否只保留首个表头）
- **多文件 → 按表头并集合并**：按所有文件表头去重后的并集统一列，缺失的列留空
- **单文件多 Sheet → 单 Sheet**：把一个工作簿里的多张工作表内容堆到一起

合并前会自动扫描表头差异，发现不一致时弹窗友好提示（不阻塞）。

### 拆分（3 种模式）

- **按工作表拆**：每张工作表单独存成一个文件
- **按行数拆**：如每 1000 行切成一个文件（可选保留表头）
- **按列值拆**：按某列的不同值分组，空值统一归到 `xxx_blank.xlsx`

仅支持 `.xlsx` 格式（不支持老版 `.xls`）。

## 环境要求

- Python **3.8.x**（⚠️ 必须保持 3.8，打包产物才能兼容 Windows 7；3.9+ 打出的 exe 无法在 Win7 运行）
- 依赖见 `requirements.txt`：`openpyxl`、`pyinstaller`

## 开发与测试

```bash
pip install -r requirements.txt

# 生成 testdata/ 下的 10 个测试文件
python generate_testdata.py

# 跑核心逻辑自测（不依赖 GUI,覆盖全部合并/拆分模式）
python test_core.py

# 启动 GUI
python main.py
```

## 打包

```bash
python build.py
```

产物为 `dist/ExcelTools.exe`（PyInstaller 单文件、无控制台窗口、UPX 压缩）。`build.py` 会先清理 `dist/`、`build/` 和旧的 `.spec` 文件。

## 目录结构

```
├── main.py               # GUI 入口:主题样式、两个 Tab(合并/拆分)、事件与线程调度
├── core/
│   ├── merger.py         # 合并核心逻辑 + 表头差异扫描,run_merge 统一入口
│   └── splitter.py       # 拆分核心逻辑,run_split 统一入口
├── test_core.py          # 核心逻辑自测脚本(纯断言,不依赖 pytest)
├── generate_testdata.py  # 生成 testdata/ 下的演示/测试 Excel 文件
├── build.py              # PyInstaller 打包脚本
└── testdata/             # 覆盖常见场景的测试 Excel 文件
```

## 兼容性

- 支持 Windows 7 SP1 及以上（前提：用 Python 3.8 打包）
- Win7 上界面字体会从 "Microsoft YaHei UI" 回退到默认字体，Tab 上的 emoji 可能显示为方框，仅影响观感
