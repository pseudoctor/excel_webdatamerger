"""
excel_datamerger GUI界面 v1.0
功能：
- 列名映射配置管理
- 数据质量报告
- 智能去重
- 模糊匹配选项
- 列名映射预览
"""
import json
import os
import threading
import tkinter as tk
import webbrowser
from datetime import datetime
from tkinter import filedialog, messagebox, scrolledtext, ttk

import pandas as pd

from .config_manager import ConfigManager
from .io_utils import read_file, save_to_excel
from .logger import setup_logger
from .merger import ExcelMergerCore

logger = setup_logger("ExcelMergerGUI")

class ExcelMergerGUI:
    """excel_datamerger v1.0"""

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("excel_datamerger v1.0")
        self.root.geometry("1000x800")
        self.root.minsize(950, 750)

        # 适配 macOS 深色模式 - 增强对比度
        self.root.configure(bg="#1a1a1a")
        self.root.option_add("*Foreground", "#FFFFFF")
        self.root.option_add("*Background", "#1a1a1a")
        self.root.option_add("*Button.Background", "#404040")
        self.root.option_add("*Button.Foreground", "#FFFFFF")
        self.root.option_add("*TLabel.Foreground", "#FFFFFF")
        self.root.option_add("*TCheckbutton.Foreground", "#FFFFFF")
        # Checkbutton 增强对比度
        self.root.option_add("*Checkbutton.selectColor", "#404040")
        self.root.option_add("*Checkbutton.activeBackground", "#1a1a1a")
        self.root.option_add("*Checkbutton.activeForeground", "#FFFFFF")

        # 配置管理器
        self.config_manager = ConfigManager()

        # 文件列表
        self.file_paths = []
        self.progress_var = tk.DoubleVar()
        self.status_text = tk.StringVar(value="就绪")

        # 选项
        self.remove_duplicates = tk.BooleanVar(value=False)
        self.normalize_columns = tk.BooleanVar(value=True)
        self.enable_fuzzy_match = tk.BooleanVar(value=False)  # 新增：模糊匹配
        self.smart_dedup = tk.BooleanVar(value=False)  # 新增：智能去重
        self.dedup_keys = tk.StringVar(value="")  # 新增：去重关键字段

        self._build_ui()

    # ======================================================
    # 构建界面
    # ======================================================
    def _build_ui(self):
        # 文件区
        file_frame = tk.LabelFrame(self.root, text="📂 已上传文件", font=("Helvetica", 11, "bold"))
        file_frame.pack(fill=tk.BOTH, padx=10, pady=10, expand=False)

        self.listbox = tk.Listbox(file_frame, height=6, width=100, font=("Consolas", 10), bg="#3c3f41", fg="white")
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.listbox.bind("<<ListboxSelect>>", self.update_preview)

        scrollbar = tk.Scrollbar(file_frame, command=self.listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.config(yscrollcommand=scrollbar.set)

        # 按钮区
        btn_frame = tk.Frame(self.root, bg="#1a1a1a")
        btn_frame.pack(fill=tk.X, padx=10, pady=5)
        # macOS优化：使用highlightbackground确保按钮可见
        tk.Button(btn_frame, text="添加文件", command=self.add_files,
                 bg="#707070", fg="#000000", font=("Helvetica", 11, "bold"),
                 relief=tk.RAISED, bd=3, cursor="hand2",
                 highlightbackground="#707070", highlightcolor="#FFFFFF",
                 activebackground="#909090", activeforeground="#000000").pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="删除选中", command=self.remove_selected,
                 bg="#707070", fg="#000000", font=("Helvetica", 11, "bold"),
                 relief=tk.RAISED, bd=3, cursor="hand2",
                 highlightbackground="#707070", highlightcolor="#FFFFFF",
                 activebackground="#909090", activeforeground="#000000").pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="清空列表", command=self.clear_all,
                 bg="#707070", fg="#000000", font=("Helvetica", 11, "bold"),
                 relief=tk.RAISED, bd=3, cursor="hand2",
                 highlightbackground="#707070", highlightcolor="#FFFFFF",
                 activebackground="#909090", activeforeground="#000000").pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="⚙️ 列名映射配置", command=self.open_config_window,
                 bg="#4CAF50", fg="#000000", font=("Helvetica", 11, "bold"),
                 relief=tk.RAISED, bd=3, cursor="hand2",
                 highlightbackground="#4CAF50", highlightcolor="#FFFFFF",
                 activebackground="#66BB6A", activeforeground="#000000").pack(side=tk.RIGHT, padx=5)

        # 选项区（增强版）
        opt_frame = tk.LabelFrame(self.root, text="🧩 功能选项", font=("Helvetica", 11, "bold"))
        opt_frame.pack(fill=tk.X, padx=10, pady=5)

        # 第一行选项
        row1 = tk.Frame(opt_frame, bg="#1a1a1a")
        row1.pack(fill=tk.X, padx=5, pady=2)
        tk.Checkbutton(row1, text="统一列名（使用映射规则）",
                      variable=self.normalize_columns,
                      bg="#1a1a1a", fg="#FFFFFF", selectcolor="#404040",
                      activebackground="#1a1a1a", activeforeground="#FFFFFF").pack(side=tk.LEFT, padx=10)
        tk.Checkbutton(row1, text="启用模糊匹配",
                      variable=self.enable_fuzzy_match,
                      bg="#1a1a1a", fg="#FFFFFF", selectcolor="#404040",
                      activebackground="#1a1a1a", activeforeground="#FFFFFF").pack(side=tk.LEFT, padx=10)

        # 第二行选项
        row2 = tk.Frame(opt_frame, bg="#1a1a1a")
        row2.pack(fill=tk.X, padx=5, pady=2)
        tk.Checkbutton(row2, text="删除重复行",
                      variable=self.remove_duplicates,
                      bg="#1a1a1a", fg="#FFFFFF", selectcolor="#404040",
                      activebackground="#1a1a1a", activeforeground="#FFFFFF").pack(side=tk.LEFT, padx=10)
        tk.Checkbutton(row2, text="智能去重（基于关键字段）",
                      variable=self.smart_dedup,
                      bg="#1a1a1a", fg="#FFFFFF", selectcolor="#404040",
                      activebackground="#1a1a1a", activeforeground="#FFFFFF").pack(side=tk.LEFT, padx=10)

        # 第三行：去重关键字段输入
        row3 = tk.Frame(opt_frame, bg="#1a1a1a")
        row3.pack(fill=tk.X, padx=5, pady=2)
        tk.Label(row3, text="去重关键字段（逗号分隔）:",
                fg="#FFFFFF", bg="#1a1a1a", font=("Helvetica", 10)).pack(side=tk.LEFT, padx=10)
        tk.Entry(row3, textvariable=self.dedup_keys, width=40,
                bg="#404040", fg="#FFFFFF", insertbackground="#FFFFFF",
                font=("Helvetica", 10)).pack(side=tk.LEFT, padx=5)

        # 文件预览区
        preview_frame = tk.LabelFrame(self.root, text="👁 文件预览（前5行）", font=("Helvetica", 11, "bold"))
        preview_frame.pack(fill=tk.BOTH, padx=10, pady=5, expand=True)
        self.preview_text = tk.Text(preview_frame, height=8, wrap="none", font=("Consolas", 9),
                                    bg="#1e1e1e", fg="white")
        self.preview_text.pack(fill=tk.BOTH, expand=True)

        # 进度条区
        prog_frame = tk.Frame(self.root, bg="#1a1a1a")
        prog_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Progressbar(prog_frame, variable=self.progress_var, maximum=100).pack(fill=tk.X, padx=5)
        tk.Label(prog_frame, textvariable=self.status_text, anchor="w",
                fg="#FFFFFF", bg="#1a1a1a", font=("Helvetica", 10)).pack(fill=tk.X, padx=5)

        # 日志显示区
        log_frame = tk.LabelFrame(self.root, text="📜 实时日志", font=("Helvetica", 11, "bold"))
        log_frame.pack(fill=tk.BOTH, padx=10, pady=5, expand=True)
        self.log_text = tk.Text(log_frame, height=8, wrap="word", font=("Consolas", 9),
                                bg="#1e1e1e", fg="#c5c5c5")
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # 启动按钮 - macOS优化
        tk.Button(self.root, text="🚀 开始合并", font=("Helvetica", 16, "bold"),
                  bg="#42A5F5", fg="#000000", command=self.run_in_thread,
                  relief=tk.RAISED, bd=4, height=2, cursor="hand2",
                  highlightbackground="#42A5F5", highlightcolor="#FFFFFF",
                  activebackground="#64B5F6", activeforeground="#000000").pack(fill=tk.X, padx=10, pady=10)

    # ======================================================
    # 文件操作
    # ======================================================
    def add_files(self):
        files = filedialog.askopenfilenames(
            title="选择要合并的文件",
            filetypes=[
                ("支持格式", "*.xlsx *.xls *.csv *.txt"),
                ("Excel 文件", "*.xlsx *.xls"),
                ("CSV 文件", "*.csv"),
                ("文本文件", "*.txt")
            ]
        )
        for f in files:
            if f not in self.file_paths:
                self.file_paths.append(f)
                self.listbox.insert(tk.END, os.path.basename(f))
        self.status_text.set(f"已添加 {len(files)} 个文件")

    def remove_selected(self):
        for i in reversed(self.listbox.curselection()):
            self.listbox.delete(i)
            self.file_paths.pop(i)
        self.status_text.set("已删除选中文件")

    def clear_all(self):
        self.file_paths.clear()
        self.listbox.delete(0, tk.END)
        self.status_text.set("文件列表已清空")

    # ======================================================
    # 文件预览
    # ======================================================
    def update_preview(self, event):
        sel = self.listbox.curselection()
        if not sel:
            return
        path = self.file_paths[sel[0]]
        try:
            sheets = read_file(path)
            df = next(iter(sheets.values()))
            preview = df.head(5).to_string(index=False)
        except Exception as e:
            preview = f"⚠️ 预览失败：{e}"
        self.preview_text.delete("1.0", tk.END)
        self.preview_text.insert(tk.END, preview)

    # ======================================================
    # 后台线程启动
    # ======================================================
    def run_in_thread(self):
        thread = threading.Thread(target=self.start_merge_safe, daemon=True)
        thread.start()

    def start_merge_safe(self):
        try:
            self.start_merge()
        except Exception as e:
            import traceback
            msg = traceback.format_exc()
            self.log(f"❌ 发生错误: {e}\n{msg}")
            messagebox.showerror("错误", f"{e}")

    # ======================================================
    # 核心合并逻辑（增强版）
    # ======================================================
    def start_merge(self):
        if not self.file_paths:
            messagebox.showwarning("提示", "请先选择要合并的文件！")
            return

        output = filedialog.asksaveasfilename(
            title="保存合并结果",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")]
        )
        if not output:
            return

        # 使用配置管理器创建合并核心
        merger = ExcelMergerCore(self.config_manager)
        all_dfs = []
        total_mapping_report = {}  # 收集所有文件的列名映射报告

        # 第一阶段：读取文件
        for i, f in enumerate(self.file_paths):
            try:
                self.status_text.set(f"读取文件: {os.path.basename(f)} ({i+1}/{len(self.file_paths)})")
                self.progress_var.set((i+1) / len(self.file_paths) * 40)
                self.root.update_idletasks()

                sheets = read_file(f)
                for name, df in sheets.items():
                    if df.empty:
                        self.log(f"⚠️ 跳过空表: {os.path.basename(f)} - {name}")
                        continue

                    # 列名归一化
                    if self.normalize_columns.get():
                        original_cols = list(df.columns)
                        df = merger.normalize_columns(
                            df,
                            enable_fuzzy=self.enable_fuzzy_match.get()
                        )
                        # 检查是否有重复列名被处理
                        new_cols = list(df.columns)
                        if any('_' in col and col.rsplit('_', 1)[-1].isdigit() for col in new_cols):
                            self.log(f"⚠️ 检测到重复列名，已自动添加后缀: {os.path.basename(f)} - {name}")

                        # 收集映射报告
                        mapping = merger.get_mapping_report()
                        if mapping:
                            total_mapping_report[f"{os.path.basename(f)}-{name}"] = mapping

                    # 添加来源标识（去掉文件扩展名）
                    filename_without_ext = os.path.splitext(os.path.basename(f))[0]
                    df.insert(0, "来源文件", filename_without_ext)
                    df.insert(1, "工作表", name)
                    all_dfs.append(df)

                    # 记录统计信息
                    stats = merger.get_summary_stats(df)
                    self.log(f"✅ {os.path.basename(f)} - {name} | {stats}")

            except Exception as e:
                self.log(f"⚠️ 文件跳过: {os.path.basename(f)} ({e})")
                continue

        if not all_dfs:
            messagebox.showinfo("提示", "没有可合并的数据。")
            return

        # 显示列名映射报告
        if total_mapping_report:
            self._show_mapping_report(total_mapping_report)

        # 第二阶段：合并数据
        self.status_text.set("正在合并数据...")
        self.progress_var.set(50)
        self.root.update_idletasks()

        merged = pd.concat(all_dfs, join="outer", ignore_index=True, sort=False)
        self.log(f"📊 合并完成 | 总计 {len(merged)} 行 × {len(merged.columns)} 列")

        # 第三阶段：去重处理
        original_count = len(merged)

        if self.smart_dedup.get() and self.dedup_keys.get().strip():
            # 智能去重（基于关键字段）
            key_cols = [k.strip() for k in self.dedup_keys.get().split(",")]
            self.status_text.set(f"智能去重中（关键字段: {key_cols}）...")
            self.progress_var.set(70)
            merged = merger.deduplicate_smart(merged, key_columns=key_cols)
            removed = original_count - len(merged)
            if removed > 0:
                self.log(f"🧹 智能去重: 删除 {removed} 行重复数据")
        elif self.remove_duplicates.get():
            # 全行去重
            self.status_text.set("删除重复行...")
            self.progress_var.set(70)
            merged = merger.deduplicate_smart(merged)
            removed = original_count - len(merged)
            if removed > 0:
                self.log(f"🧹 全行去重: 删除 {removed} 行重复数据")

        # 第四阶段：数据质量报告
        self.status_text.set("生成数据质量报告...")
        self.progress_var.set(85)
        quality_report = merger.validate_data(merged)
        self._show_quality_report(quality_report)

        # 第五阶段：保存文件
        self.status_text.set("正在保存结果...")
        self.progress_var.set(90)
        save_to_excel(merged, output)

        self.progress_var.set(100)
        self.status_text.set("合并完成 ✅")
        self.log(f"💾 合并完成，文件已保存至: {output}")

        # 自动打开输出目录（已禁用）
        # folder = os.path.dirname(output) or os.getcwd()
        # if os.path.exists(folder):
        #     webbrowser.open(folder)

        messagebox.showinfo("成功", f"合并完成！\n最终数据: {len(merged)} 行\n文件位置:\n{output}")

    # ======================================================
    # 日志输出
    # ======================================================
    def log(self, msg):
        logger.info(msg)
        self.log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {msg}\n")
        self.log_text.see(tk.END)

    # ======================================================
    # 新增功能：列名映射配置窗口
    # ======================================================
    def open_config_window(self):
        """打开列名映射配置窗口"""
        config_win = tk.Toplevel(self.root)
        config_win.title("列名映射配置管理")
        config_win.geometry("700x600")
        config_win.configure(bg="#1a1a1a")

        # 说明文本
        info_frame = tk.Frame(config_win, bg="#1a1a1a")
        info_frame.pack(fill=tk.X, padx=10, pady=10)
        tk.Label(info_frame, text="配置列名映射规则，格式：标准列名 → 别名列表",
                fg="#FFFFFF", bg="#1a1a1a", font=("Helvetica", 11, "bold")).pack(anchor="w")

        # 配置编辑区
        edit_frame = tk.Frame(config_win, bg="#1a1a1a")
        edit_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 使用Text widget显示配置
        text_widget = scrolledtext.ScrolledText(
            edit_frame,
            height=20,
            font=("Consolas", 10),
            bg="#1e1e1e",
            fg="white"
        )
        text_widget.pack(fill=tk.BOTH, expand=True)

        # 加载当前配置
        mappings = self.config_manager.get_mappings()
        config_text = json.dumps(mappings, ensure_ascii=False, indent=2)
        text_widget.insert("1.0", config_text)

        # 按钮区
        btn_frame = tk.Frame(config_win, bg="#1a1a1a")
        btn_frame.pack(fill=tk.X, padx=10, pady=10)

        def save_config():
            try:
                new_config = json.loads(text_widget.get("1.0", tk.END))
                self.config_manager.save_mappings(new_config)
                messagebox.showinfo("成功", "配置已保存！")
                config_win.destroy()
            except json.JSONDecodeError as e:
                messagebox.showerror("错误", f"JSON格式错误：{e}")

        def reset_config():
            if messagebox.askyesno("确认", "确定要重置为默认配置吗？"):
                self.config_manager.reset_to_default()
                self.config_manager.save_mappings()
                text_widget.delete("1.0", tk.END)
                config_text = json.dumps(
                    self.config_manager.get_mappings(),
                    ensure_ascii=False,
                    indent=2
                )
                text_widget.insert("1.0", config_text)
                messagebox.showinfo("成功", "已重置为默认配置！")

        tk.Button(btn_frame, text="保存配置", command=save_config,
                 bg="#4CAF50", fg="#000000", font=("Helvetica", 11, "bold"),
                 relief=tk.RAISED, bd=3, cursor="hand2",
                 highlightbackground="#4CAF50", activebackground="#66BB6A",
                 activeforeground="#000000").pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="重置为默认", command=reset_config,
                 bg="#FF9800", fg="#000000", font=("Helvetica", 11, "bold"),
                 relief=tk.RAISED, bd=3, cursor="hand2",
                 highlightbackground="#FF9800", activebackground="#FFB74D",
                 activeforeground="#000000").pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="取消", command=config_win.destroy,
                 bg="#f44336", fg="#FFFFFF", font=("Helvetica", 11, "bold"),
                 relief=tk.RAISED, bd=3, cursor="hand2",
                 highlightbackground="#f44336", activebackground="#EF5350",
                 activeforeground="#FFFFFF").pack(side=tk.RIGHT, padx=5)

    # ======================================================
    # 新增功能：显示列名映射报告
    # ======================================================
    def _show_mapping_report(self, total_report):
        """显示列名映射报告"""
        self.log("=" * 50)
        self.log("📋 列名映射报告")
        self.log("=" * 50)

        for file_sheet, mappings in total_report.items():
            self.log(f"\n文件: {file_sheet}")
            for orig, (mapped, match_type) in mappings.items():
                if orig != mapped:  # 只显示被映射的列
                    self.log(f"  • {orig} → {mapped} [{match_type}]")

        self.log("=" * 50)

    # ======================================================
    # 新增功能：显示数据质量报告
    # ======================================================
    def _show_quality_report(self, report):
        """显示数据质量报告"""
        self.log("=" * 50)
        self.log("📊 数据质量报告")
        self.log("=" * 50)
        self.log(f"总行数: {report['总行数']}")
        self.log(f"总列数: {report['总列数']}")
        self.log(f"重复行数: {report['重复行数']}")

        # 显示空值率高的列
        self.log("\n空值情况（仅显示空值率>0的列）:")
        null_stats = report["空值统计"]
        has_null = False
        for col, stats in null_stats.items():
            if stats["数量"] > 0:
                has_null = True
                self.log(f"  • {col}: {stats['数量']} 行 ({stats['百分比']}%)")

        if not has_null:
            self.log("  ✅ 无空值")

        self.log("=" * 50)

    # ======================================================
    def run(self):
        self.root.mainloop()
