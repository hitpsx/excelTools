"""Excel 工具箱 GUI 入口 (人性化提示语 + 防呆版)。

设计目标:
- 提示语零技术黑话,大白话
- 无日志面板,出错直接弹窗
- 拆分参数按所选模式联动显示(防呆:只让填当前模式需要的设置)
- 合并前扫描表头差异,友好提示但不阻塞
- 卡片化、Banner、状态栏
"""
from __future__ import annotations

import os
import threading
import tkinter as tk
import traceback
from tkinter import filedialog, messagebox, Tk, StringVar, BooleanVar, Toplevel
from tkinter import ttk
from typing import List

from core.merger import MergeOptions, run_merge, scan_header_differences, HeaderDiff
from core.splitter import SplitOptions, run_split


# ================= 主题 =================
APP_TITLE = "Excel 工具箱"
APP_SUBTITLE = "帮你把零散的 Excel 表整理成一份,也能把一份大表拆成多份"
THEME_NAME = "clam"

COLOR_BG = "#F1F5F9"
COLOR_PANEL = "#FFFFFF"
COLOR_PRIMARY = "#2563EB"
COLOR_PRIMARY_DARK = "#1D4ED8"
COLOR_PRIMARY_LIGHT = "#DBEAFE"
COLOR_TEXT = "#0F172A"
COLOR_TEXT_LIGHT = "#64748B"
COLOR_BORDER = "#E2E8F0"
COLOR_HEADER = "#EFF6FF"
COLOR_BANNER = "#1E3A8A"
COLOR_BANNER_TXT = "#FFFFFF"
COLOR_STRIPE = "#F8FAFC"

FONT_FAMILY = "Microsoft YaHei UI"


class _ScrollableFrame(ttk.Frame):
    """可垂直滚动的容器,用于 Tab 内容。

    内容超出可视区域时显示滚动条,同时支持鼠标滚轮。
    进入区域时绑定滚轮,离开时解绑,避免与子组件冲突。
    """

    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)

        self._canvas = tk.Canvas(self, highlightthickness=0, bg=COLOR_BG)
        self._vscroll = ttk.Scrollbar(self, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=self._vscroll.set)

        self.inner = ttk.Frame(self._canvas, style="TFrame")
        self._window_id = self._canvas.create_window(
            (0, 0), window=self.inner, anchor="nw"
        )

        self._canvas.pack(side="left", fill="both", expand=True)
        self._vscroll.pack(side="right", fill="y")

        # 内容大小变化时更新滚动区域
        self.inner.bind(
            "<Configure>",
            lambda e: self._canvas.configure(scrollregion=self._canvas.bbox("all")),
        )
        # 容器宽度变化时让内容宽度跟随
        self._canvas.bind(
            "<Configure>",
            lambda e: self._canvas.itemconfigure(self._window_id, width=e.width),
        )

        # 滚轮:进入 Canvas 区域时绑定全局滚轮,离开时解绑
        self._canvas.bind("<Enter>", self._bind_wheel)
        self._canvas.bind("<Leave>", self._unbind_wheel)

    def _bind_wheel(self, _event) -> None:
        self._canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _unbind_wheel(self, _event) -> None:
        self._canvas.unbind_all("<MouseWheel>")

    def _on_mousewheel(self, event) -> None:
        # 只处理当前可见 Tab
        if self.winfo_ismapped():
            self._canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
FONT_TITLE = (FONT_FAMILY, 20, "bold")
FONT_SUBTITLE = (FONT_FAMILY, 10)
FONT_LABEL = (FONT_FAMILY, 10)
FONT_BOLD = (FONT_FAMILY, 10, "bold")
FONT_BTN = (FONT_FAMILY, 10, "bold")
FONT_BTN_BIG = (FONT_FAMILY, 11, "bold")
FONT_TAB = (FONT_FAMILY, 11, "bold")
FONT_SECTION = (FONT_FAMILY, 10, "bold")


class ExcelToolApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("800x660")
        self.root.minsize(740, 580)
        self.root.configure(bg=COLOR_BG)

        # 合并页变量
        self.merge_mode = StringVar(value="files_to_sheets")
        self.merge_output = StringVar()
        self.merge_include_header = BooleanVar(value=True)
        self.merge_files: List[str] = []

        # 拆分页变量
        self.split_mode = StringVar(value="by_sheet")
        self.split_input = StringVar()
        self.split_output_dir = StringVar()
        self.split_rows = StringVar(value="1000")
        self.split_col = StringVar(value="1")
        self.split_keep_header = BooleanVar(value=True)

        # 拆分模式联动(防呆):模式变化时切换参数区
        self.split_mode.trace_add("write", self._on_split_mode_change)

        self._configure_style()
        self._build_ui()

    # ---------- 样式 ----------
    def _configure_style(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use(THEME_NAME)
        except Exception:
            style.theme_use("default")

        style.configure(".", font=FONT_LABEL, foreground=COLOR_TEXT, background=COLOR_BG)
        style.configure("TFrame", background=COLOR_BG)
        style.configure("Card.TFrame", background=COLOR_PANEL)
        style.configure("TLabel", background=COLOR_BG, foreground=COLOR_TEXT)
        style.configure("Card.TLabel", background=COLOR_PANEL, foreground=COLOR_TEXT)
        style.configure("Hint.TLabel", background=COLOR_PANEL,
                        foreground=COLOR_TEXT_LIGHT, font=(FONT_FAMILY, 9))

        # 卡片化 LabelFrame
        style.configure(
            "Card.TLabelframe",
            background=COLOR_PANEL, foreground=COLOR_TEXT,
            bordercolor=COLOR_BORDER, lightcolor=COLOR_BORDER, darkcolor=COLOR_BORDER,
            relief="solid", borderwidth=1,
        )
        style.configure(
            "Card.TLabelframe.Label",
            background=COLOR_PANEL, foreground=COLOR_PRIMARY_DARK,
            font=FONT_SECTION, padding=(2, 0),
        )

        # 主按钮 (实心蓝)
        style.configure(
            "Primary.TButton",
            font=FONT_BTN_BIG, foreground="#FFFFFF", background=COLOR_PRIMARY,
            bordercolor=COLOR_PRIMARY_DARK, borderwidth=0, focuscolor="none",
            padding=(24, 12),
        )
        style.map(
            "Primary.TButton",
            background=[("active", COLOR_PRIMARY_DARK), ("disabled", "#94A3B8")],
            foreground=[("disabled", "#F1F5F9")],
        )

        # 次级按钮 (白底蓝边)
        style.configure(
            "Secondary.TButton",
            font=FONT_BTN, foreground=COLOR_PRIMARY, background=COLOR_PANEL,
            bordercolor=COLOR_PRIMARY, borderwidth=1, focuscolor="none",
            padding=(14, 7),
        )
        style.map(
            "Secondary.TButton",
            background=[("active", COLOR_PRIMARY_LIGHT)],
            foreground=[("active", COLOR_PRIMARY_DARK)],
        )

        # 单选/复选
        style.configure("Card.TRadiobutton",
                        background=COLOR_PANEL, foreground=COLOR_TEXT,
                        font=FONT_LABEL, padding=(4, 4))
        style.configure("Card.TCheckbutton",
                        background=COLOR_PANEL, foreground=COLOR_TEXT,
                        font=FONT_LABEL, padding=(4, 4))

        # Notebook
        style.configure("TNotebook", background=COLOR_BG, borderwidth=0,
                        tabmargins=(0, 0, 0, 0))
        style.configure(
            "TNotebook.Tab",
            font=FONT_TAB, padding=(40, 10),
            background=COLOR_BG, foreground=COLOR_TEXT_LIGHT, borderwidth=0,
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", COLOR_PANEL)],
            foreground=[("selected", COLOR_PRIMARY), ("active", COLOR_PRIMARY_DARK)],
        )

        # Entry
        style.configure(
            "Card.TEntry",
            fieldbackground="#FFFFFF", foreground=COLOR_TEXT,
            bordercolor=COLOR_BORDER, lightcolor=COLOR_BORDER, darkcolor=COLOR_BORDER,
            borderwidth=1, padding=8,
        )

        # Progressbar
        style.configure(
            "Custom.Horizontal.TProgressbar",
            troughcolor=COLOR_PRIMARY_LIGHT, background=COLOR_PRIMARY,
            bordercolor=COLOR_PRIMARY_LIGHT,
            lightcolor=COLOR_PRIMARY, darkcolor=COLOR_PRIMARY,
            thickness=12,
        )

        # Treeview
        style.configure(
            "Files.Treeview",
            background="#FFFFFF", fieldbackground="#FFFFFF", foreground=COLOR_TEXT,
            rowheight=28, bordercolor=COLOR_BORDER, borderwidth=1,
            font=(FONT_FAMILY, 9),
        )
        style.configure(
            "Files.Treeview.Heading",
            background=COLOR_HEADER, foreground=COLOR_PRIMARY_DARK,
            font=FONT_BOLD, relief="flat", borderwidth=0, padding=(8, 6),
        )
        style.map(
            "Files.Treeview",
            background=[("selected", COLOR_PRIMARY_LIGHT)],
            foreground=[("selected", COLOR_PRIMARY_DARK)],
        )

    # ---------- 顶层 UI ----------
    def _build_ui(self) -> None:
        self._build_banner()

        nb_frame = ttk.Frame(self.root)
        nb_frame.pack(fill="both", expand=True, padx=16, pady=(6, 4))

        nb = ttk.Notebook(nb_frame)
        nb.pack(fill="both", expand=True)
        self.nb = nb

        # 可滚动的 Tab 容器
        merge_tab = _ScrollableFrame(nb)
        split_tab = _ScrollableFrame(nb)
        nb.add(merge_tab, text="  📥  合 并  ")
        nb.add(split_tab, text="  📤  拆 分  ")

        self._build_merge_tab(merge_tab.inner)
        self._build_split_tab(split_tab.inner)
        self._build_status_bar()
        self._set_status("准备就绪,请选择上方功能开始使用")

        # 初始化拆分参数区的显示
        self._on_split_mode_change()

    def _build_banner(self) -> None:
        banner = tk.Frame(self.root, bg=COLOR_BANNER, height=88)
        banner.pack(fill="x", side="top")
        banner.pack_propagate(False)

        inner = tk.Frame(banner, bg=COLOR_BANNER)
        inner.pack(fill="both", expand=True, padx=28, pady=14)

        title = tk.Label(
            inner, text="Excel 工具箱",
            bg=COLOR_BANNER, fg=COLOR_BANNER_TXT, font=FONT_TITLE,
        )
        title.pack(anchor="w")

        sub = tk.Label(
            inner, text=APP_SUBTITLE,
            bg=COLOR_BANNER, fg="#BFDBFE", font=FONT_SUBTITLE,
        )
        sub.pack(anchor="w", pady=(2, 0))

        version = tk.Label(
            inner, text="v1.0", bg=COLOR_BANNER, fg="#93C5FD",
            font=(FONT_FAMILY, 9),
        )
        version.pack(side="right", anchor="e", pady=(6, 0))

    def _build_status_bar(self) -> None:
        bar = tk.Frame(self.root, bg="#E2E8F0", height=28)
        bar.pack(side="bottom", fill="x")
        bar.pack_propagate(False)
        self.status_label = tk.Label(
            bar, text="", bg="#E2E8F0", fg=COLOR_TEXT_LIGHT,
            font=(FONT_FAMILY, 9), padx=14,
        )
        self.status_label.pack(side="left", fill="y")

    def _set_status(self, msg: str) -> None:
        self.status_label.config(text=msg)

    # ---------- 合并 Tab ----------
    def _build_merge_tab(self, parent) -> None:
        # 1. 想怎么合并
        mode_lf = ttk.LabelFrame(parent, text="① 想怎么合并?", style="Card.TLabelframe", padding=12)
        mode_lf.pack(fill="x", pady=(4, 6))
        ttk.Radiobutton(
            mode_lf, text="好几个 Excel 表 → 合成 1 个,每个表单独放一个工作表",
            variable=self.merge_mode, value="files_to_sheets", style="Card.TRadiobutton",
        ).grid(row=0, column=0, sticky="w", padx=6, pady=3)
        ttk.Radiobutton(
            mode_lf, text="好几个 Excel 表 → 合成 1 个,所有内容堆在同一张工作表里",
            variable=self.merge_mode, value="files_to_rows", style="Card.TRadiobutton",
        ).grid(row=1, column=0, sticky="w", padx=6, pady=3)
        ttk.Radiobutton(
            mode_lf, text="好几个 Excel 表 → 合成 1 个,按表头去重后统一列",
            variable=self.merge_mode, value="files_to_union_rows", style="Card.TRadiobutton",
        ).grid(row=2, column=0, sticky="w", padx=6, pady=3)
        ttk.Radiobutton(
            mode_lf, text="1 个 Excel 表里有多张工作表 → 把内容都堆到第一张里",
            variable=self.merge_mode, value="sheets_to_rows", style="Card.TRadiobutton",
        ).grid(row=3, column=0, sticky="w", padx=6, pady=3)
        ttk.Checkbutton(
            mode_lf, text="重复出现的表头只保留第一次(适用于堆在一起的情况)",
            variable=self.merge_include_header, style="Card.TCheckbutton",
        ).grid(row=4, column=0, sticky="w", padx=6, pady=3)

        # 2. 选文件
        file_lf = ttk.LabelFrame(parent, text="② 选要合并的文件", style="Card.TLabelframe", padding=12)
        file_lf.pack(fill="x", pady=(0, 6))
        ttk.Label(file_lf, text="可以按住 Ctrl 多选;也可以直接选整个文件夹",
                  style="Hint.TLabel").pack(anchor="w", pady=(0, 4))

        tree_frame = ttk.Frame(file_lf, style="Card.TFrame")
        tree_frame.pack(fill="x", pady=(0, 8))

        self.merge_tree = ttk.Treeview(
            tree_frame, columns=("idx", "path"), show="headings",
            height=5, style="Files.Treeview", selectmode="extended",
        )
        self.merge_tree.heading("idx", text="序号")
        self.merge_tree.heading("path", text="文件位置")
        self.merge_tree.column("idx", width=56, anchor="center", stretch=False)
        self.merge_tree.column("path", width=560, anchor="w")
        self.merge_tree.tag_configure("stripe", background=COLOR_STRIPE)
        self.merge_tree.pack(side="left", fill="both", expand=True)

        sb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.merge_tree.yview)
        self.merge_tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")

        btn_row = ttk.Frame(file_lf, style="Card.TFrame")
        btn_row.pack(fill="x")
        ttk.Button(btn_row, text="+ 挑文件", style="Secondary.TButton",
                   command=self._on_merge_add).pack(side="left", padx=2)
        ttk.Button(btn_row, text="+ 整文件夹", style="Secondary.TButton",
                   command=self._on_merge_add_dir).pack(side="left", padx=2)
        ttk.Button(btn_row, text="移除选中", style="Secondary.TButton",
                   command=self._on_merge_remove).pack(side="left", padx=2)
        ttk.Button(btn_row, text="清空", style="Secondary.TButton",
                   command=self._on_merge_clear).pack(side="left", padx=2)

        # 3. 输出位置
        out_lf = ttk.LabelFrame(parent, text="③ 合并后的文件放哪?", style="Card.TLabelframe", padding=12)
        out_lf.pack(fill="x", pady=(0, 6))
        ttk.Entry(out_lf, textvariable=self.merge_output, style="Card.TEntry").pack(
            side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(out_lf, text="选个地方", style="Secondary.TButton",
                   command=self._on_merge_choose_output).pack(side="left")

        # 4. 进度 + 大按钮
        run_lf = ttk.LabelFrame(parent, text="④ 开干!", style="Card.TLabelframe", padding=12)
        run_lf.pack(fill="x", pady=(0, 6))
        self.merge_progress = ttk.Progressbar(
            run_lf, mode="determinate", style="Custom.Horizontal.TProgressbar"
        )
        self.merge_progress.pack(fill="x", pady=(0, 10))
        ttk.Button(run_lf, text="  开  始  合  并  ", style="Primary.TButton",
                   command=self._on_merge_start).pack(anchor="e")

    # ---------- 拆分 Tab ----------
    def _build_split_tab(self, parent) -> None:
        # 1. 想怎么拆
        mode_lf = ttk.LabelFrame(parent, text="① 想怎么拆?", style="Card.TLabelframe", padding=12)
        mode_lf.pack(fill="x", pady=(4, 6))
        ttk.Radiobutton(
            mode_lf, text="表里有多张工作表 → 每张工作表单独存成 1 个文件",
            variable=self.split_mode, value="by_sheet", style="Card.TRadiobutton",
        ).grid(row=0, column=0, sticky="w", padx=6, pady=3)
        ttk.Radiobutton(
            mode_lf, text="按行数切 → 比如每 1000 行切成一个文件",
            variable=self.split_mode, value="by_rows", style="Card.TRadiobutton",
        ).grid(row=1, column=0, sticky="w", padx=6, pady=3)
        ttk.Radiobutton(
            mode_lf, text="按某一列的值切 → 比如把 \"部门=A\" 的都放一个文件",
            variable=self.split_mode, value="by_column", style="Card.TRadiobutton",
        ).grid(row=2, column=0, sticky="w", padx=6, pady=3)

        # 2. 选文件
        in_lf = ttk.LabelFrame(parent, text="② 选要拆的那个 Excel", style="Card.TLabelframe", padding=12)
        in_lf.pack(fill="x", pady=(0, 6))
        ttk.Entry(in_lf, textvariable=self.split_input, style="Card.TEntry").pack(
            side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(in_lf, text="挑文件", style="Secondary.TButton",
                   command=self._on_split_choose_input).pack(side="left")

        # 3. 参数 (防呆:只显示当前模式需要的设置)
        self.param_lf = ttk.LabelFrame(parent, text="③ 拆分时的小设置", style="Card.TLabelframe", padding=12)
        self.param_lf.pack(fill="x", pady=(0, 6))

        # 3a. 按行数切的设置
        self.rows_frame = ttk.Frame(self.param_lf, style="Card.TFrame")
        ttk.Label(self.rows_frame, text="每个文件最多多少行:",
                  style="Card.TLabel").pack(side="left", padx=6)
        ttk.Entry(self.rows_frame, textvariable=self.split_rows,
                  style="Card.TEntry", width=12).pack(side="left", padx=4)
        ttk.Label(self.rows_frame, text="比如填 1000,就是每 1000 行存成一个文件",
                  style="Hint.TLabel").pack(side="left", padx=8)

        # 3b. 按列值切的设置
        self.col_frame = ttk.Frame(self.param_lf, style="Card.TFrame")
        ttk.Label(self.col_frame, text="按第几列的值分组:",
                  style="Card.TLabel").pack(side="left", padx=6)
        ttk.Entry(self.col_frame, textvariable=self.split_col,
                  style="Card.TEntry", width=12).pack(side="left", padx=4)
        ttk.Label(self.col_frame, text="填 1 表示第 1 列,填 2 表示第 2 列",
                  style="Hint.TLabel").pack(side="left", padx=8)

        # 3c. 按工作表拆的提示
        self.sheet_frame = ttk.Frame(self.param_lf, style="Card.TFrame")
        ttk.Label(self.sheet_frame,
                  text="这种拆法不用额外设置,每张工作表都会单独存成一个文件",
                  style="Hint.TLabel").pack(side="left", padx=6)

        # 3d. 通用:保留表头
        self.header_frame = ttk.Frame(self.param_lf, style="Card.TFrame")
        self.header_frame.pack(fill="x", pady=(6, 0))
        ttk.Checkbutton(self.header_frame, text="拆出来的每个文件都保留表头",
                        variable=self.split_keep_header, style="Card.TCheckbutton"
                        ).pack(side="left", padx=6)

        # 4. 输出目录
        out_lf = ttk.LabelFrame(parent, text="④ 拆出来的文件放哪?", style="Card.TLabelframe", padding=12)
        out_lf.pack(fill="x", pady=(0, 6))
        ttk.Entry(out_lf, textvariable=self.split_output_dir, style="Card.TEntry").pack(
            side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(out_lf, text="选个文件夹", style="Secondary.TButton",
                   command=self._on_split_choose_output).pack(side="left")

        # 5. 进度 + 大按钮
        run_lf = ttk.LabelFrame(parent, text="⑤ 开干!", style="Card.TLabelframe", padding=12)
        run_lf.pack(fill="x", pady=(0, 6))
        self.split_progress = ttk.Progressbar(
            run_lf, mode="determinate", style="Custom.Horizontal.TProgressbar"
        )
        self.split_progress.pack(fill="x", pady=(0, 10))
        ttk.Button(run_lf, text="  开  始  拆  分  ", style="Primary.TButton",
                   command=self._on_split_start).pack(anchor="e")

    def _on_split_mode_change(self, *args) -> None:
        """防呆:根据所选模式,只显示对应的设置项。"""
        if not hasattr(self, "rows_frame"):
            return
        for f in (self.rows_frame, self.col_frame, self.sheet_frame):
            f.pack_forget()
        mode = self.split_mode.get()
        if mode == "by_rows":
            self.rows_frame.pack(fill="x", pady=(0, 2))
        elif mode == "by_column":
            self.col_frame.pack(fill="x", pady=(0, 2))
        else:  # by_sheet
            self.sheet_frame.pack(fill="x", pady=(0, 2))

    # ---------- 通用工具 ----------
    def _set_progress(self, bar, current: int, total: int) -> None:
        if total <= 0:
            bar["value"] = 0
            bar["maximum"] = 100
            return
        bar["maximum"] = total
        bar["value"] = current

    def _refresh_merge_tree(self) -> None:
        for iid in self.merge_tree.get_children():
            self.merge_tree.delete(iid)
        for i, p in enumerate(self.merge_files, start=1):
            tag = ("stripe",) if i % 2 == 0 else ()
            self.merge_tree.insert("", "end", values=(i, p), tags=tag)

    @staticmethod
    def _format_header_diffs(diffs: List[HeaderDiff]) -> str:
        """把表头差异整理成大白话文本。"""
        lines: List[str] = []
        for d in diffs:
            if not d.missing:
                continue
            missing_text = "、".join(d.missing)
            lines.append(f"· {d.file_name}: 缺少列 「{missing_text}」")
        return "\n".join(lines) if lines else ""

    # ---------- 合并事件 ----------
    def _on_merge_add(self) -> None:
        paths = filedialog.askopenfilenames(
            title="挑要合并的 Excel",
            filetypes=[("Excel 表", "*.xlsx"), ("所有文件", "*.*")],
        )
        added = 0
        for p in paths:
            if p not in self.merge_files:
                self.merge_files.append(p)
                added += 1
        if added:
            self._refresh_merge_tree()
            self._set_status(f"已经加了 {len(self.merge_files)} 个文件")

    def _on_merge_add_dir(self) -> None:
        d = filedialog.askdirectory(title="选整个文件夹(里面所有 Excel 都会被加入)")
        if not d:
            return
        added = 0
        for name in sorted(os.listdir(d)):
            if name.lower().endswith(".xlsx") and not name.startswith("~$"):
                full = os.path.join(d, name)
                if full not in self.merge_files:
                    self.merge_files.append(full)
                    added += 1
        if added:
            self._refresh_merge_tree()
            self._set_status(f"已从文件夹加入 {added} 个,共 {len(self.merge_files)} 个")
        else:
            messagebox.showinfo("提示", "这个文件夹里没找到 Excel 表 ( .xlsx )")

    def _on_merge_remove(self) -> None:
        sel = self.merge_tree.selection()
        if not sel:
            return
        indexes = sorted(
            [self.merge_tree.index(iid) for iid in sel], reverse=True
        )
        for idx in indexes:
            del self.merge_files[idx]
        self._refresh_merge_tree()
        self._set_status(f"已移除,还剩 {len(self.merge_files)} 个")

    def _on_merge_clear(self) -> None:
        if not self.merge_files:
            return
        self.merge_files.clear()
        self._refresh_merge_tree()
        self._set_status("文件列表已清空")

    def _on_merge_choose_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="合并后要保存到哪?",
            defaultextension=".xlsx",
            filetypes=[("Excel 表", "*.xlsx")],
        )
        if path:
            self.merge_output.set(path)

    def _on_merge_start(self) -> None:
        if not self.merge_output.get().strip():
            messagebox.showwarning("少一步", "还没选合并后保存到哪,点第 ③ 步选一下")
            return
        mode = self.merge_mode.get()
        if mode == "sheets_to_rows":
            if len(self.merge_files) != 1:
                messagebox.showwarning("少一步",
                                       "你选的是「把 1 个 Excel 里的多张工作表合并」,\n"
                                       "所以第 ② 步只能挑 1 个 Excel 文件哦")
                return
        elif not self.merge_files:
            messagebox.showwarning("少一步", "还没挑要合并的 Excel,点第 ② 步挑一下")
            return

        # 对堆在一起的合并模式,先做表头差异扫描
        if mode in ("files_to_rows", "files_to_union_rows"):
            union_headers, diffs = scan_header_differences(self.merge_files)
            diff_text = self._format_header_diffs(diffs)
            if diff_text:
                if mode == "files_to_union_rows":
                    msg = (
                        "检测到以下文件的表头和大家的并集不一样:\n\n"
                        f"{diff_text}\n\n"
                        "我最后会按所有表头去重后的并集来合并,缺失的列会空着。\n"
                        "你还要继续吗?"
                    )
                else:
                    msg = (
                        "检测到以下文件的表头和第一个文件不一样:\n\n"
                        f"{diff_text}\n\n"
                        "继续合并的话,列对不上的地方可能会串列或空着。\n"
                        "建议选上面的「按表头去重后统一列」模式,你还要继续吗?"
                    )
                if not messagebox.askyesno("表头好像不太一样", msg):
                    self._set_status("已取消合并")
                    return

        self.merge_progress["value"] = 0
        self._set_status("合并中,请稍等…")

        files_snapshot = list(self.merge_files)
        opts = MergeOptions(
            mode=mode,
            output_path=self.merge_output.get().strip(),
            include_header=self.merge_include_header.get(),
        )

        t = threading.Thread(
            target=self._run_merge_task, args=(opts, files_snapshot), daemon=True
        )
        t.start()

    def _run_merge_task(self, opts: MergeOptions, files: List[str]) -> None:
        def cb(cur, total, msg):
            self.root.after(0, self._set_progress, self.merge_progress, cur, total)
            self.root.after(0, self._set_status, msg)

        try:
            n = run_merge(opts, files, progress_cb=cb)
            self.root.after(0, self._set_status, f"合并完成: {opts.output_path}")
            self.root.after(0, lambda: messagebox.showinfo("搞定", f"合并完成,共生成 {n} 个结果"))
        except Exception as e:
            traceback.print_exc()
            self.root.after(0, self._set_status, "合并失败")
            self.root.after(0, lambda: messagebox.showerror(
                "合并没成功",
                f"出错原因: {e}\n\n常见原因:\n"
                "· 文件正被 Excel 打开,先关掉再试\n"
                "· 选了 .xls 老格式,目前只支持 .xlsx\n"
                "· 文件已损坏或受密码保护"))

    # ---------- 拆分事件 ----------
    def _on_split_choose_input(self) -> None:
        path = filedialog.askopenfilename(
            title="挑要拆分的 Excel",
            filetypes=[("Excel 表", "*.xlsx")],
        )
        if path:
            self.split_input.set(path)
            self._set_status(f"已选择: {os.path.basename(path)}")

    def _on_split_choose_output(self) -> None:
        d = filedialog.askdirectory(title="拆出来的文件放哪个文件夹?")
        if d:
            self.split_output_dir.set(d)
            self._set_status(f"输出文件夹: {d}")

    def _on_split_start(self) -> None:
        if not self.split_input.get().strip() or not os.path.isfile(self.split_input.get()):
            messagebox.showwarning("少一步", "还没挑要拆的 Excel,点第 ② 步挑一下")
            return
        if not self.split_output_dir.get().strip():
            messagebox.showwarning("少一步", "还没选拆出来的文件放哪,点第 ④ 步选一下")
            return

        # 只在对应模式下校验对应参数(防呆:别的模式的脏数据不影响)
        mode = self.split_mode.get()
        rows = 1000
        col = 1
        if mode == "by_rows":
            try:
                rows = int(self.split_rows.get())
                if rows <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showwarning("数字不太对",
                                       "「每个文件最多多少行」请填一个正整数,比如 1000")
                return
        elif mode == "by_column":
            try:
                col = int(self.split_col.get())
                if col <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showwarning("数字不太对",
                                       "「按第几列的值分组」请填一个正整数,1 表示第 1 列")
                return

        self.split_progress["value"] = 0
        self._set_status("拆分中,请稍等…")

        opts = SplitOptions(
            mode=mode,
            output_dir=self.split_output_dir.get().strip(),
            rows_per_file=rows,
            column_index=col,
            keep_header=self.split_keep_header.get(),
        )
        in_file = self.split_input.get().strip()

        t = threading.Thread(
            target=self._run_split_task, args=(opts, in_file), daemon=True
        )
        t.start()

    def _run_split_task(self, opts: SplitOptions, in_file: str) -> None:
        def cb(cur, total, msg):
            self.root.after(0, self._set_progress, self.split_progress, cur, total)
            self.root.after(0, self._set_status, msg)

        try:
            n = run_split(opts, in_file, progress_cb=cb)
            self.root.after(0, self._set_status, f"拆分完成: {opts.output_dir}")
            self.root.after(0, lambda: messagebox.showinfo("搞定", f"拆分完成,共 {n} 个文件"))
        except Exception as e:
            traceback.print_exc()
            self.root.after(0, self._set_status, "拆分失败")
            self.root.after(0, lambda: messagebox.showerror(
                "拆分没成功",
                f"出错原因: {e}\n\n常见原因:\n"
                "· 文件正被 Excel 打开,先关掉再试\n"
                "· 选了 .xls 老格式,目前只支持 .xlsx\n"
                "· 填的列号超过了表格实际的列数"))


# ================= 入口 =================
def main() -> None:
    root = tk.Tk()
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    ExcelToolApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
