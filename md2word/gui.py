"""Tkinter GUI：选择 Markdown/输出文件，并在界面内直接编辑样式配置后转换。"""

from __future__ import annotations

import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, ttk

from .config import DEFAULT_CONFIG, deep_merge


def run_gui() -> None:
    root = tk.Tk()
    App(root)
    root.mainloop()


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("md2word - Markdown 转 Word")
        root.geometry("720x580")

        self.md_var = tk.StringVar()
        self.out_var = tk.StringVar()
        self.log_queue: queue.Queue[str] = queue.Queue()

        pad = {"padx": 10, "pady": 4}

        # 文件选择区
        file_frame = ttk.LabelFrame(root, text="文件", padding=6)
        file_frame.pack(fill=tk.X, **pad)
        self._row(file_frame, 0, "Markdown 文件:", self.md_var, self._pick_md)
        self._row(file_frame, 1, "输出文件:", self.out_var, self._pick_out)

        # 配置区
        cfg_frame = ttk.LabelFrame(root, text="样式配置", padding=6)
        cfg_frame.pack(fill=tk.BOTH, expand=True, **pad)
        self._build_config_editor(cfg_frame)

        self.convert_btn = ttk.Button(root, text="开始转换", command=self._convert)
        self.convert_btn.pack(pady=6)

        self.log = tk.Text(root, height=10, state=tk.DISABLED, wrap=tk.WORD)
        self.log.pack(fill=tk.BOTH, expand=False, padx=10, pady=(0, 8))

        self.root.after(100, self._poll_log)

    # ---------- 文件选择 ----------

    def _row(self, parent, row, label, var, command):
        ttk.Label(parent, text=label, width=14).grid(row=row, column=0, sticky=tk.W, pady=3)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky=tk.EW, padx=4)
        ttk.Button(parent, text="浏览...", command=command).grid(row=row, column=2)
        parent.columnconfigure(1, weight=1)

    def _pick_md(self):
        path = filedialog.askopenfilename(
            title="选择 Markdown 文件",
            filetypes=[("Markdown", "*.md *.markdown"), ("所有文件", "*.*")])
        if path:
            self.md_var.set(path)
            if not self.out_var.get():
                self.out_var.set(os.path.splitext(path)[0] + ".docx")

    def _pick_out(self):
        path = filedialog.asksaveasfilename(
            title="选择输出位置", defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx")])
        if path:
            self.out_var.set(path)

    # ---------- 配置编辑器 ----------

    def _build_config_editor(self, parent):
        self.notebook = ttk.Notebook(parent)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.page_vars = self._build_page_tab(self.notebook)
        self.body_vars = self._build_body_tab(self.notebook)
        self.heading_vars = self._build_headings_tab(self.notebook)
        self.code_vars, self.table_vars = self._build_code_table_tab(self.notebook)

    def _add_tab(self, notebook, title):
        frame = ttk.Frame(notebook, padding=8)
        notebook.add(frame, text=title)
        return frame

    def _entry(self, parent, row, label, var, width=12):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky=tk.W, pady=3)
        ttk.Entry(parent, textvariable=var, width=width).grid(row=row, column=1, sticky=tk.W, padx=4)

    def _combo(self, parent, row, label, var, values, width=12):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky=tk.W, pady=3)
        cb = ttk.Combobox(parent, textvariable=var, values=values, width=width, state="readonly")
        cb.grid(row=row, column=1, sticky=tk.W, padx=4)

    def _check(self, parent, row, label, var):
        ttk.Checkbutton(parent, text=label, variable=var).grid(row=row, column=0, columnspan=2, sticky=tk.W, pady=3)

    def _build_page_tab(self, notebook):
        frame = self._add_tab(notebook, "页面")
        defaults = DEFAULT_CONFIG["page"]
        vars_ = {
            "size": tk.StringVar(value=defaults["size"]),
            "width_cm": tk.StringVar(value=str(defaults["width_cm"])),
            "height_cm": tk.StringVar(value=str(defaults["height_cm"])),
            "margin_top_cm": tk.StringVar(value=str(defaults["margin_top_cm"])),
            "margin_bottom_cm": tk.StringVar(value=str(defaults["margin_bottom_cm"])),
            "margin_left_cm": tk.StringVar(value=str(defaults["margin_left_cm"])),
            "margin_right_cm": tk.StringVar(value=str(defaults["margin_right_cm"])),
        }
        self._combo(frame, 0, "页面尺寸", vars_["size"], ["A4", "A5", "Letter", "custom"])
        self._entry(frame, 1, "宽度 (cm)", vars_["width_cm"])
        self._entry(frame, 2, "高度 (cm)", vars_["height_cm"])
        self._entry(frame, 3, "上边距 (cm)", vars_["margin_top_cm"])
        self._entry(frame, 4, "下边距 (cm)", vars_["margin_bottom_cm"])
        self._entry(frame, 5, "左边距 (cm)", vars_["margin_left_cm"])
        self._entry(frame, 6, "右边距 (cm)", vars_["margin_right_cm"])
        return vars_

    def _build_body_tab(self, notebook):
        frame = self._add_tab(notebook, "正文")
        defaults = DEFAULT_CONFIG["body"]
        vars_ = {
            "font_en": tk.StringVar(value=defaults["font_en"]),
            "font_zh": tk.StringVar(value=defaults["font_zh"]),
            "font_size": tk.StringVar(value=str(defaults["font_size"])),
            "line_spacing": tk.StringVar(value=str(defaults["line_spacing"])),
            "first_line_indent": tk.BooleanVar(value=defaults["first_line_indent"]),
            "color": tk.StringVar(value=defaults["color"]),
        }
        self._entry(frame, 0, "英文字体", vars_["font_en"], width=24)
        self._entry(frame, 1, "中文字体", vars_["font_zh"], width=24)
        self._entry(frame, 2, "字号 (pt)", vars_["font_size"])
        self._entry(frame, 3, "行距 (倍)", vars_["line_spacing"])
        self._entry(frame, 4, "颜色 (RRGGBB)", vars_["color"])
        self._check(frame, 5, "段落首行缩进两字符", vars_["first_line_indent"])
        return vars_

    def _build_headings_tab(self, notebook):
        frame = self._add_tab(notebook, "标题")
        defaults = DEFAULT_CONFIG["headings"]

        ttk.Label(frame, text="级别").grid(row=0, column=0, padx=4)
        ttk.Label(frame, text="中文字体").grid(row=0, column=1, padx=4)
        ttk.Label(frame, text="字号").grid(row=0, column=2, padx=4)
        ttk.Label(frame, text="加粗").grid(row=0, column=3, padx=4)
        ttk.Label(frame, text="对齐").grid(row=0, column=4, padx=4)
        ttk.Label(frame, text="段前间距").grid(row=0, column=5, padx=4)
        ttk.Label(frame, text="段后间距").grid(row=0, column=6, padx=4)

        vars_ = {}
        for i, level in enumerate(("h1", "h2", "h3"), start=1):
            d = defaults[level]
            v = {
                "font_zh": tk.StringVar(value=d["font_zh"]),
                "font_size": tk.StringVar(value=str(d["font_size"])),
                "bold": tk.BooleanVar(value=d.get("bold", True)),
                "align": tk.StringVar(value=d.get("align", "left")),
                "space_before": tk.StringVar(value=str(d.get("space_before", 0))),
                "space_after": tk.StringVar(value=str(d.get("space_after", 0))),
            }
            vars_[level] = v
            ttk.Label(frame, text=level.upper()).grid(row=i, column=0, pady=2)
            ttk.Entry(frame, textvariable=v["font_zh"], width=12).grid(row=i, column=1, padx=4)
            ttk.Entry(frame, textvariable=v["font_size"], width=6).grid(row=i, column=2, padx=4)
            ttk.Checkbutton(frame, variable=v["bold"]).grid(row=i, column=3)
            ttk.Combobox(frame, textvariable=v["align"], values=["left", "center", "right", "justify"],
                         width=8, state="readonly").grid(row=i, column=4, padx=4)
            ttk.Entry(frame, textvariable=v["space_before"], width=6).grid(row=i, column=5, padx=4)
            ttk.Entry(frame, textvariable=v["space_after"], width=6).grid(row=i, column=6, padx=4)
        return vars_

    def _build_code_table_tab(self, notebook):
        frame = self._add_tab(notebook, "代码 / 表格")
        code_defaults = DEFAULT_CONFIG["code"]
        table_defaults = DEFAULT_CONFIG["table"]

        code = ttk.LabelFrame(frame, text="代码块", padding=6)
        code.pack(fill=tk.X, pady=4)
        code_vars = {
            "font_en": tk.StringVar(value=code_defaults["font_en"]),
            "font_zh": tk.StringVar(value=code_defaults["font_zh"]),
            "font_size": tk.StringVar(value=str(code_defaults["font_size"])),
            "line_spacing": tk.StringVar(value=str(code_defaults["line_spacing"])),
            "color": tk.StringVar(value=code_defaults["color"]),
        }
        self._entry(code, 0, "英文字体", code_vars["font_en"], width=20)
        self._entry(code, 1, "中文字体", code_vars["font_zh"], width=20)
        self._entry(code, 2, "字号 (pt)", code_vars["font_size"])
        self._entry(code, 3, "行距 (倍)", code_vars["line_spacing"])
        self._entry(code, 4, "颜色 (RRGGBB)", code_vars["color"])

        table = ttk.LabelFrame(frame, text="表格", padding=6)
        table.pack(fill=tk.X, pady=4)
        table_vars = {
            "style": tk.StringVar(value=table_defaults["style"]),
            "header_align": tk.StringVar(value=table_defaults["header_align"]),
            "header_bold": tk.BooleanVar(value=table_defaults["header_bold"]),
        }
        self._combo(table, 0, "表格样式", table_vars["style"], ["grid", "three_line"])
        self._combo(table, 1, "表头对齐", table_vars["header_align"], ["left", "center", "right"])
        self._check(table, 2, "表头加粗", table_vars["header_bold"])

        return code_vars, table_vars

    # ---------- 日志 ----------

    def _log(self, msg: str):
        self.log.config(state=tk.NORMAL)
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)
        self.log.config(state=tk.DISABLED)

    def _poll_log(self):
        try:
            while True:
                self._log(self.log_queue.get_nowait())
        except queue.Empty:
            pass
        self.root.after(100, self._poll_log)

    # ---------- 转换 ----------

    def _get_numeric(self, var, name):
        value = var.get().strip()
        if not value:
            raise ValueError(f"{name} 不能为空")
        try:
            return int(value) if "." not in value else float(value)
        except ValueError as exc:
            raise ValueError(f"{name} 必须是数字") from exc

    def _collect_config(self) -> dict:
        cfg: dict = {"page": {}, "body": {}, "headings": {}, "code": {}, "table": {}}

        cfg["page"]["size"] = self.page_vars["size"].get()
        cfg["page"]["width_cm"] = self._get_numeric(self.page_vars["width_cm"], "页面宽度")
        cfg["page"]["height_cm"] = self._get_numeric(self.page_vars["height_cm"], "页面高度")
        cfg["page"]["margin_top_cm"] = self._get_numeric(self.page_vars["margin_top_cm"], "上边距")
        cfg["page"]["margin_bottom_cm"] = self._get_numeric(self.page_vars["margin_bottom_cm"], "下边距")
        cfg["page"]["margin_left_cm"] = self._get_numeric(self.page_vars["margin_left_cm"], "左边距")
        cfg["page"]["margin_right_cm"] = self._get_numeric(self.page_vars["margin_right_cm"], "右边距")

        cfg["body"]["font_en"] = self.body_vars["font_en"].get().strip() or DEFAULT_CONFIG["body"]["font_en"]
        cfg["body"]["font_zh"] = self.body_vars["font_zh"].get().strip() or DEFAULT_CONFIG["body"]["font_zh"]
        cfg["body"]["font_size"] = self._get_numeric(self.body_vars["font_size"], "正文字号")
        cfg["body"]["line_spacing"] = self._get_numeric(self.body_vars["line_spacing"], "正文行距")
        cfg["body"]["first_line_indent"] = bool(self.body_vars["first_line_indent"].get())
        cfg["body"]["color"] = self.body_vars["color"].get().strip() or DEFAULT_CONFIG["body"]["color"]

        cfg["code"]["font_en"] = self.code_vars["font_en"].get().strip() or DEFAULT_CONFIG["code"]["font_en"]
        cfg["code"]["font_zh"] = self.code_vars["font_zh"].get().strip() or DEFAULT_CONFIG["code"]["font_zh"]
        cfg["code"]["font_size"] = self._get_numeric(self.code_vars["font_size"], "代码字号")
        cfg["code"]["line_spacing"] = self._get_numeric(self.code_vars["line_spacing"], "代码行距")
        cfg["code"]["color"] = self.code_vars["color"].get().strip() or DEFAULT_CONFIG["code"]["color"]

        cfg["table"]["style"] = self.table_vars["style"].get()
        cfg["table"]["header_align"] = self.table_vars["header_align"].get()
        cfg["table"]["header_bold"] = bool(self.table_vars["header_bold"].get())

        for level in ("h1", "h2", "h3"):
            v = self.heading_vars[level]
            cfg["headings"][level] = {
                "font_zh": v["font_zh"].get().strip() or DEFAULT_CONFIG["headings"][level]["font_zh"],
                "font_size": self._get_numeric(v["font_size"], f"{level.upper()} 字号"),
                "bold": bool(v["bold"].get()),
                "align": v["align"].get(),
                "space_before": self._get_numeric(v["space_before"], f"{level.upper()} 段前间距"),
                "space_after": self._get_numeric(v["space_after"], f"{level.upper()} 段后间距"),
            }

        return deep_merge(DEFAULT_CONFIG, cfg)

    def _convert(self):
        md_path = self.md_var.get().strip()
        if not md_path:
            self._log("请先选择 Markdown 文件。")
            return
        out_path = self.out_var.get().strip() or None

        try:
            config = self._collect_config()
        except ValueError as exc:
            self._log(f"配置错误: {exc}")
            return

        self.convert_btn.config(state=tk.DISABLED)
        self._log(f"开始转换: {md_path}")

        def worker():
            from .converter import Converter
            try:
                if not out_path:
                    output = os.path.splitext(os.path.abspath(md_path))[0] + ".docx"
                else:
                    output = out_path
                Converter(config).convert_file(md_path, output)
                self.log_queue.put(f"完成: {output}")
            except Exception as exc:
                self.log_queue.put(f"转换失败: {exc}")
            finally:
                self.root.after(0, lambda: self.convert_btn.config(state=tk.NORMAL))

        threading.Thread(target=worker, daemon=True).start()
