"""
主窗口 —— 整合所有 UI 面板。
"""
import os
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime

import ttkbootstrap as tb

from engine.excel_scanner import scan_xlsx_files, filter_by_name
from engine.excel_reader import read_excel
from engine.validator import validate_table, ValidationError
from engine.merger import merge_tables
from engine.lua_generator import write_lua_file


class MainWindow:
    def __init__(self):
        self.root = tb.Window(themename="cosmo")
        self.root.title("Lua 配置转换器")
        self.root.geometry("900x720")
        self.root.minsize(700, 550)

        # ---- 状态变量 ----
        self.input_dir = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.search_text = tk.StringVar()
        self.all_files: list[str] = []
        self.filtered_files: list[str] = []

        # ---- 搜索框变更跟踪 ----
        self.search_timer_id: str | None = None

        # ---- 分隔条初始比例（文件列表占比，0.0~1.0） ----
        self.pane_ratio = 0.6

        self._build_ui()
        self._load_last_paths()
        # 窗口渲染后再设置分隔条初始位置
        self.root.after(150, self._init_sash)

    def _build_ui(self):
        # ===== 顶部：路径配置（两行） =====
        top_frame = ttk.Frame(self.root, padding=(10, 10, 10, 5))
        top_frame.pack(fill=tk.X)

        # 第一行：导入路径
        ttk.Label(top_frame, text="导入路径:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        ttk.Entry(top_frame, textvariable=self.input_dir, state="readonly").grid(
            row=0, column=1, sticky=tk.EW, padx=(0, 5))
        ttk.Button(top_frame, text="选择导入文件夹", command=self._select_input_dir).grid(
            row=0, column=2)

        # 第二行：导出路径
        ttk.Label(top_frame, text="导出路径:").grid(row=1, column=0, sticky=tk.W, padx=(0, 5), pady=(6, 0))
        ttk.Entry(top_frame, textvariable=self.output_dir, state="readonly").grid(
            row=1, column=1, sticky=tk.EW, padx=(0, 5), pady=(6, 0))
        ttk.Button(top_frame, text="选择导出文件夹", command=self._select_output_dir).grid(
            row=1, column=2, pady=(6, 0))

        top_frame.columnconfigure(1, weight=1)

        # ===== 中部：垂直分割（文件列表 + 日志面板） =====
        self.mid_pane = ttk.PanedWindow(self.root, orient=tk.VERTICAL)
        self.mid_pane.pack(fill=tk.BOTH, expand=True, padx=10, pady=(5, 0))

        # -- 上半：搜索框 + 文件列表 --
        upper_frame = ttk.Frame(self.mid_pane, padding=(0, 0, 0, 0))
        self.mid_pane.add(upper_frame, weight=4)

        search_frame = ttk.Frame(upper_frame)
        search_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(search_frame, text="搜索文件:").pack(side=tk.LEFT, padx=(0, 5))
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_text)
        self.search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(search_frame, text="清除", command=self._clear_search).pack(side=tk.LEFT)

        self.search_text.trace_add("write", self._on_search_changed)

        list_frame = ttk.Frame(upper_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)

        columns = ("filename", "path")
        self.tree = ttk.Treeview(
            list_frame, columns=columns, show="tree headings",
            selectmode="extended")
        self.tree.heading("#0", text="")
        self.tree.heading("filename", text="文件名")
        self.tree.heading("path", text="路径")
        self.tree.column("#0", width=0, stretch=False)
        self.tree.column("filename", width=200, minwidth=100)
        self.tree.column("path", width=500, minwidth=200)

        tree_scroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_scroll.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # 右键菜单
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="导出", command=self._export_selected)
        self.context_menu.add_command(label="打开", command=self._open_selected)
        self.tree.bind("<Button-3>", self._on_right_click)
        self.tree.bind("<Double-1>", lambda e: self._open_selected())

        # -- 下半：日志面板（可拖拽分隔条调整高度） --
        log_frame = ttk.Frame(self.mid_pane, padding=(0, 5, 0, 0))
        self.mid_pane.add(log_frame, weight=1)

        log_header = ttk.Frame(log_frame)
        log_header.pack(fill=tk.X)
        ttk.Label(log_header, text="输出日志", font=("", 9, "bold")).pack(side=tk.LEFT)
        ttk.Button(log_header, text="清空日志", command=self._clear_log,
                   style="secondary.TButton", width=8).pack(side=tk.RIGHT)

        log_text_frame = ttk.Frame(log_frame)
        log_text_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(
            log_text_frame, wrap=tk.WORD, state=tk.DISABLED,
            font=("Consolas", 9), bg="#1e1e1e", fg="#d4d4d4",
            insertbackground="#d4d4d4", relief=tk.FLAT,
            borderwidth=0, padx=8, pady=6)
        log_text_scroll = ttk.Scrollbar(log_text_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_text_scroll.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_text_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # 日志标签颜色（仅标签着色，正文保持默认前景色）
        self.log_text.tag_configure("time", foreground="#888888")
        self.log_text.tag_configure("tag_info", foreground="#6f9fcf")
        self.log_text.tag_configure("tag_success", foreground="#6fdf6f")
        self.log_text.tag_configure("tag_error", foreground="#ff6b6b")
        self.log_text.tag_configure("tag_warn", foreground="#ffcc66")

        # ===== 底部：状态栏 + 操作按钮 =====
        bottom_frame = ttk.Frame(self.root, padding=(10, 5, 10, 10))
        bottom_frame.pack(fill=tk.X)

        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(bottom_frame, textvariable=self.status_var, anchor=tk.W).pack(
            side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Button(bottom_frame, text="全部导出", command=self._export_all,
                   style="success.TButton").pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(bottom_frame, text="刷新列表", command=self._refresh_file_list,
                   style="info.TButton").pack(side=tk.RIGHT, padx=(5, 0))

        # 初始日志消息
        self._log("就绪，等待操作...", "info")

    # ---- 分隔条 ----

    def _init_sash(self):
        """设置文件列表与日志面板的初始分隔条位置。
        pane_ratio 控制文件列表占中部区域的比例（默认 0.78 = 78%）。
        """
        try:
            total = self.mid_pane.winfo_height()
            if total > 100:
                sash_y = int(total * self.pane_ratio)
                self.mid_pane.sashpos(0, sash_y)
        except Exception:
            pass  # 窗口尚未就绪则跳过

    def _log(self, message: str, tag: str = "info"):
        """向日志面板追加一条消息。格式: [时间] [类型] 正文
        仅类型标签着色，正文使用默认文本色。
        """
        tag_labels = {"info": "INFO", "success": "OK", "error": "ERR", "warn": "WARN"}
        label = tag_labels.get(tag, tag.upper())

        self.log_text.configure(state=tk.NORMAL)
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] ", "time")
        self.log_text.insert(tk.END, f"[{label}] ", f"tag_{tag}")
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _clear_log(self):
        """清空日志面板。"""
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state=tk.DISABLED)
        self._log("日志已清空", "info")

    # ---- 事件处理 ----

    def _select_input_dir(self):
        path = filedialog.askdirectory(title="选择 Excel 导入文件夹")
        if path:
            self.input_dir.set(path)
            self._save_last_paths()
            self._refresh_file_list()
            self._log(f"导入路径: {path}", "info")

    def _select_output_dir(self):
        path = filedialog.askdirectory(title="选择 Lua 导出文件夹")
        if path:
            self.output_dir.set(path)
            self._save_last_paths()
            self._log(f"导出路径: {path}", "info")

    def _on_search_changed(self, *args):
        if self.search_timer_id:
            self.root.after_cancel(self.search_timer_id)
        self.search_timer_id = self.root.after(300, self._refresh_file_list)

    def _clear_search(self):
        self.search_text.set("")

    def _refresh_file_list(self):
        input_dir = self.input_dir.get()
        if not input_dir or not os.path.isdir(input_dir):
            self.all_files = []
            self.filtered_files = []
            self._update_treeview()
            self.status_var.set("请先选择导入文件夹")
            return

        try:
            self.all_files = scan_xlsx_files(input_dir)
        except Exception as e:
            self._log(f"扫描文件夹失败: {e}", "error")
            self.all_files = []

        self.filtered_files = filter_by_name(self.all_files, self.search_text.get())
        self._update_treeview()
        self.status_var.set(f"共 {len(self.filtered_files)} 个文件"
                            f"（总计 {len(self.all_files)} 个）")

    def _update_treeview(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for fpath in self.filtered_files:
            fname = os.path.basename(fpath)
            try:
                rel_path = os.path.relpath(fpath, self.input_dir.get())
            except ValueError:
                rel_path = fpath
            self.tree.insert("", tk.END, values=(fname, rel_path), iid=fpath)

    def _on_right_click(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            if item not in self.tree.selection():
                self.tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)

    def _get_selected_files(self) -> list[str]:
        return list(self.tree.selection())

    def _export_selected(self):
        selected = self._get_selected_files()
        if not selected:
            self._log("请先选择要导出的文件", "warn")
            return
        self._do_export(selected)

    def _export_all(self):
        if not self.filtered_files:
            self._log("没有可导出的文件", "warn")
            return
        self._do_export(list(self.filtered_files))

    def _open_selected(self):
        selected = self._get_selected_files()
        if not selected:
            return
        for fpath in selected:
            try:
                os.startfile(fpath)
                self._log(f"打开文件: {os.path.basename(fpath)}", "info")
            except Exception as e:
                self._log(f"无法打开文件: {os.path.basename(fpath)} - {e}", "error")

    def _do_export(self, files: list[str]):
        output_dir = self.output_dir.get()
        if not output_dir:
            self._log("请先选择导出文件夹", "warn")
            return

        self._log(f"开始导出 {len(files)} 个文件...", "info")
        self.status_var.set("正在导出...")

        def export_thread():
            try:
                result = self._run_export(files, output_dir)
                self.root.after(0, lambda res=result: self._on_export_done(res))
            except Exception as e:
                err_msg = str(e)
                self.root.after(0, lambda msg=err_msg: self._on_export_error(msg))

        threading.Thread(target=export_thread, daemon=True).start()

    def _run_export(self, files: list[str], output_dir: str) -> str:
        all_tables = []
        errors = []

        for fpath in files:
            try:
                tables = read_excel(fpath)
                all_tables.extend(tables)
                # 在线程中通过 after 输出日志
                self.root.after(0, lambda f=fpath, n=len(tables):
                    self._log(f"读取: {os.path.basename(f)} ({n} 张表)", "info"))
            except Exception as e:
                err_detail = f"读取失败 [{os.path.basename(fpath)}]:\n  {e}"
                errors.append(err_detail)
                self.root.after(0, lambda msg=err_detail: self._log(msg, "error"))

        if errors:
            raise Exception("导出过程中发生错误:\n\n" + "\n\n".join(errors))

        if not all_tables:
            return "没有找到有效的配置表（无 A1 表名定义）"

        merged_map = merge_tables(all_tables)
        self.root.after(0, lambda: self._log(f"合并为 {len(merged_map)} 张表", "info"))

        exported = []
        for name, merged in merged_map.items():
            try:
                validate_table(merged)
            except ValidationError as e:
                raise Exception(f"校验失败 [{name}]:\n{e.message}")

            file_path = write_lua_file(merged, output_dir)
            exported.append(f"  {name}.lua  ({len(merged.rows)} 行数据)")
            self.root.after(0, lambda n=name, r=len(merged.rows):
                self._log(f"  导出: {n}.lua ({r} 行)", "success"))

        return "\n".join(exported)

    def _on_export_done(self, message: str):
        self.status_var.set("导出完成")
        self._log("导出完成:", "success")
        for line in message.split("\n"):
            self._log(f"  {line}", "success")

    def _on_export_error(self, error_msg: str):
        self.status_var.set("导出失败")
        for line in error_msg.split("\n"):
            self._log(line, "error")

    # ---- 路径持久化 ----

    def _save_last_paths(self):
        """保存最后使用的路径到配置文件。"""
        import json
        config = {
            "input_dir": self.input_dir.get(),
            "output_dir": self.output_dir.get(),
        }
        config_path = self._get_config_path()
        try:
            os.makedirs(os.path.dirname(config_path), exist_ok=True)
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False)
        except Exception:
            pass  # 静默失败，不影响主流程

    def _load_last_paths(self):
        """加载上次使用的路径。"""
        import json
        config_path = self._get_config_path()
        try:
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                if config.get("input_dir"):
                    self.input_dir.set(config["input_dir"])
                    self._refresh_file_list()
                if config.get("output_dir"):
                    self.output_dir.set(config["output_dir"])
        except Exception:
            pass

    def _get_config_path(self) -> str:
        """获取配置文件路径。"""
        appdata = os.environ.get("APPDATA", os.path.expanduser("~"))
        return os.path.join(appdata, "LuaConfigConverter", "config.json")

    def run(self):
        self.root.mainloop()
