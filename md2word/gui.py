"""Tkinter GUI：选择 Markdown/输出文件，在界面内直接编辑样式配置，并支持保存/加载配置。"""

from __future__ import annotations

import os
import queue
import threading
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, ttk

import yaml

from .config import DEFAULT_CONFIG, deep_merge


# Word 中文字号 ↔ 磅值 对照表
CHINESE_FONT_SIZES = [
    ("初号", 42), ("小初", 36), ("一号", 26), ("小一", 24),
    ("二号", 22), ("小二", 18), ("三号", 16), ("小三", 15),
    ("四号", 14), ("小四", 12), ("五号", 10.5), ("小五", 9),
    ("六号", 7.5), ("小六", 6.5), ("七号", 5.5), ("八号", 5),
]
CHINESE_SIZE_NAMES = [name for name, _ in CHINESE_FONT_SIZES]
PT_TO_SIZE_NAME = {pt: name for name, pt in CHINESE_FONT_SIZES}
NAME_TO_PT = {name: pt for name, pt in CHINESE_FONT_SIZES}


def _default_config_path() -> str:
    return os.path.join(os.path.expanduser("~"), ".md2word.yaml")


def _get_font_families() -> list[str]:
    try:
        families = sorted({f for f in tkfont.families() if not f.startswith("@")})
    except Exception:
        families = []
    if not families:
        families = [
            "Arial", "Times New Roman", "Calibri", "Consolas", "Courier New",
            "宋体", "黑体", "微软雅黑", "楷体", "仿宋",
        ]
    return families


def run_gui() -> None:
    root = tk.Tk()
    App(root)
    root.mainloop()


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("md2word - Markdown 转 Word")
        root.geometry("820x640")

        self.md_var = tk.StringVar()
        self.out_var = tk.StringVar()
        self.log_queue: queue.Queue[str] = queue.Queue()

        self.font_families = _get_font_families()

        pad = {"padx": 10, "pady": 4}

        # 文件选择
        file_frame = ttk.LabelFrame(root, text="文件", padding=6)
        file_frame.pack(fill=tk.X, **pad)
        self._row(file_frame, 0, "Markdown 文件:", self.md_var, self._pick_md)
        self._row(file_frame, 1, "输出文件:", self.out_var, self._pick_out)

        # 配置区
        cfg_frame = ttk.LabelFrame(root, text="样式配置", padding=6)
        cfg_frame.pack(fill=tk.BOTH, expand=True, **pad)
        self._build_config_editor(cfg_frame)

        # 配置保存/加载按钮
        btn_frame = ttk.Frame(root)
        btn_frame.pack(fill=tk.X, padx=10, pady=2)
        ttk.Button(btn_frame, text="保存配置", command=self._save_config).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_frame, text="加载配置", command=self._load_config).pack(side=tk.LEFT)

        self.convert_btn = ttk.Button(root, text="开始转换", command=self._convert)
        self.convert_btn.pack(pady=6)

        self.log = tk.Text(root, height=10, state=tk.DISABLED, wrap=tk.WORD)
        self.log.pack(fill=tk.BOTH, expand=False, padx=10, pady=(0, 8))

        self.root.after(100, self._poll_log)

        # 自动加载默认配置
        default_path = _default_config_path()
        if os.path.exists(default_path):
            try:
                with open(default_path, "r", encoding="utf-8") as f:
                    cfg = yaml.safe_load(f) or {}
                self._apply_config(cfg)
                self._log(f"已加载默认配置: {default_path}")
            except Exception as exc:
                self._log(f"加载默认配置失败: {exc}")

    # ------------------------------------------------------------------ 文件选择

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

    # ------------------------------------------------------------------ 配置编辑器

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

    def _label(self, parent, row, col, text, **kwargs):
        ttk.Label(parent, text=text, **kwargs).grid(row=row, column=col, sticky=tk.W, padx=3, pady=2)

    def _font_combo(self, parent, row, label, var, width=22):
        self._label(parent, row, 0, label)
        cb = ttk.Combobox(parent, textvariable=var, values=self.font_families, width=width, state="readonly")
        cb.grid(row=row, column=1, sticky=tk.W, padx=4, pady=2)
        return cb

    def _size_combo(self, parent, row, label, var, width=10):
        self._label(parent, row, 0, label)
        cb = ttk.Combobox(parent, textvariable=var, values=CHINESE_SIZE_NAMES, width=width, state="readonly")
        cb.grid(row=row, column=1, sticky=tk.W, padx=4, pady=2)
        return cb

    def _numeric_entry(self, parent, row, label, var, width=10, unit=""):
        self._label(parent, row, 0, label)
        ttk.Entry(parent, textvariable=var, width=width).grid(row=row, column=1, sticky=tk.W, padx=4, pady=2)
        if unit:
            self._label(parent, row, 2, unit)

    def _combo(self, parent, row, label, var, values, width=12):
        self._label(parent, row, 0, label)
        cb = ttk.Combobox(parent, textvariable=var, values=values, width=width, state="readonly")
        cb.grid(row=row, column=1, sticky=tk.W, padx=4, pady=2)
        return cb

    def _check(self, parent, row, label, var):
        ttk.Checkbutton(parent, text=label, variable=var).grid(row=row, column=0, columnspan=3, sticky=tk.W, pady=3)

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
        self._numeric_entry(frame, 1, "宽度 (cm)", vars_["width_cm"])
        self._numeric_entry(frame, 2, "高度 (cm)", vars_["height_cm"])
        self._numeric_entry(frame, 3, "上边距 (cm)", vars_["margin_top_cm"])
        self._numeric_entry(frame, 4, "下边距 (cm)", vars_["margin_bottom_cm"])
        self._numeric_entry(frame, 5, "左边距 (cm)", vars_["margin_left_cm"])
        self._numeric_entry(frame, 6, "右边距 (cm)", vars_["margin_right_cm"])
        return vars_

    def _build_body_tab(self, notebook):
        frame = self._add_tab(notebook, "正文")
        defaults = DEFAULT_CONFIG["body"]
        vars_ = {
            "font_en": tk.StringVar(value=defaults["font_en"]),
            "font_zh": tk.StringVar(value=defaults["font_zh"]),
            "font_size": tk.StringVar(value=PT_TO_SIZE_NAME.get(defaults["font_size"], "小四")),
            "line_spacing": tk.StringVar(value=str(defaults["line_spacing"])),
            "first_line_indent": tk.BooleanVar(value=defaults["first_line_indent"]),
            "color": tk.StringVar(value=defaults["color"]),
        }
        self._font_combo(frame, 0, "英文字体", vars_["font_en"])
        self._font_combo(frame, 1, "中文字体", vars_["font_zh"])
        self._size_combo(frame, 2, "字号", vars_["font_size"])
        self._numeric_entry(frame, 3, "行距", vars_["line_spacing"], unit="倍")
        self._numeric_entry(frame, 4, "颜色 (RRGGBB)", vars_["color"], width=16)
        self._check(frame, 5, "段落首行缩进两字符", vars_["first_line_indent"])
        return vars_

    def _build_headings_tab(self, notebook):
        frame = self._add_tab(notebook, "标题")
        defaults = DEFAULT_CONFIG["headings"]

        headers = ["级别", "中文字体", "字号", "加粗", "对齐", "段前空行", "段后空行", "行距"]
        for col, text in enumerate(headers):
            self._label(frame, 0, col, text)

        vars_ = {}
        for i, level in enumerate(("h1", "h2", "h3", "h4"), start=1):
            d = defaults[level]
            v = {
                "font_zh": tk.StringVar(value=d["font_zh"]),
                "font_size": tk.StringVar(value=PT_TO_SIZE_NAME.get(d["font_size"], "三号")),
                "bold": tk.BooleanVar(value=d.get("bold", True)),
                "align": tk.StringVar(value=d.get("align", "left")),
                "space_before_lines": tk.StringVar(value=str(d.get("space_before_lines", 0))),
                "space_after_lines": tk.StringVar(value=str(d.get("space_after_lines", 0))),
                "line_spacing": tk.StringVar(value=str(d.get("line_spacing", 1.5))),
            }
            vars_[level] = v

            self._label(frame, i, 0, level.upper(), width=6)
            ttk.Combobox(frame, textvariable=v["font_zh"], values=self.font_families, width=14, state="readonly").grid(row=i, column=1, padx=3, pady=2)
            ttk.Combobox(frame, textvariable=v["font_size"], values=CHINESE_SIZE_NAMES, width=8, state="readonly").grid(row=i, column=2, padx=3, pady=2)
            ttk.Checkbutton(frame, variable=v["bold"]).grid(row=i, column=3)
            ttk.Combobox(frame, textvariable=v["align"], values=["left", "center", "right", "justify"], width=8, state="readonly").grid(row=i, column=4, padx=3, pady=2)
            ttk.Entry(frame, textvariable=v["space_before_lines"], width=6).grid(row=i, column=5, padx=3, pady=2)
            ttk.Entry(frame, textvariable=v["space_after_lines"], width=6).grid(row=i, column=6, padx=3, pady=2)
            ttk.Entry(frame, textvariable=v["line_spacing"], width=6).grid(row=i, column=7, padx=3, pady=2)
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
            "font_size": tk.StringVar(value=PT_TO_SIZE_NAME.get(code_defaults["font_size"], "五号")),
            "line_spacing": tk.StringVar(value=str(code_defaults["line_spacing"])),
            "color": tk.StringVar(value=code_defaults["color"]),
        }
        self._font_combo(code, 0, "英文字体", code_vars["font_en"])
        self._font_combo(code, 1, "中文字体", code_vars["font_zh"])
        self._size_combo(code, 2, "字号", code_vars["font_size"])
        self._numeric_entry(code, 3, "行距", code_vars["line_spacing"], unit="倍")
        self._numeric_entry(code, 4, "颜色 (RRGGBB)", code_vars["color"], width=16)

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

    # ------------------------------------------------------------------ 配置保存/加载

    def _apply_config(self, cfg: dict):
        """把配置字典同步到界面控件。"""
        cfg = deep_merge(DEFAULT_CONFIG, cfg)

        page = cfg.get("page", {})
        self.page_vars["size"].set(page.get("size", DEFAULT_CONFIG["page"]["size"]))
        self.page_vars["width_cm"].set(str(page.get("width_cm", DEFAULT_CONFIG["page"]["width_cm"])))
        self.page_vars["height_cm"].set(str(page.get("height_cm", DEFAULT_CONFIG["page"]["height_cm"])))
        self.page_vars["margin_top_cm"].set(str(page.get("margin_top_cm", DEFAULT_CONFIG["page"]["margin_top_cm"])))
        self.page_vars["margin_bottom_cm"].set(str(page.get("margin_bottom_cm", DEFAULT_CONFIG["page"]["margin_bottom_cm"])))
        self.page_vars["margin_left_cm"].set(str(page.get("margin_left_cm", DEFAULT_CONFIG["page"]["margin_left_cm"])))
        self.page_vars["margin_right_cm"].set(str(page.get("margin_right_cm", DEFAULT_CONFIG["page"]["margin_right_cm"])))

        body = cfg.get("body", {})
        self.body_vars["font_en"].set(body.get("font_en", DEFAULT_CONFIG["body"]["font_en"]))
        self.body_vars["font_zh"].set(body.get("font_zh", DEFAULT_CONFIG["body"]["font_zh"]))
        self.body_vars["font_size"].set(PT_TO_SIZE_NAME.get(body.get("font_size"), "小四"))
        self.body_vars["line_spacing"].set(str(body.get("line_spacing", DEFAULT_CONFIG["body"]["line_spacing"])))
        self.body_vars["first_line_indent"].set(bool(body.get("first_line_indent", DEFAULT_CONFIG["body"]["first_line_indent"])))
        self.body_vars["color"].set(body.get("color", DEFAULT_CONFIG["body"]["color"]))

        code = cfg.get("code", {})
        self.code_vars["font_en"].set(code.get("font_en", DEFAULT_CONFIG["code"]["font_en"]))
        self.code_vars["font_zh"].set(code.get("font_zh", DEFAULT_CONFIG["code"]["font_zh"]))
        self.code_vars["font_size"].set(PT_TO_SIZE_NAME.get(code.get("font_size"), "五号"))
        self.code_vars["line_spacing"].set(str(code.get("line_spacing", DEFAULT_CONFIG["code"]["line_spacing"])))
        self.code_vars["color"].set(code.get("color", DEFAULT_CONFIG["code"]["color"]))

        table = cfg.get("table", {})
        self.table_vars["style"].set(table.get("style", DEFAULT_CONFIG["table"]["style"]))
        self.table_vars["header_align"].set(table.get("header_align", DEFAULT_CONFIG["table"]["header_align"]))
        self.table_vars["header_bold"].set(bool(table.get("header_bold", DEFAULT_CONFIG["table"]["header_bold"])))

        headings = cfg.get("headings", {})
        for level in ("h1", "h2", "h3", "h4"):
            d = DEFAULT_CONFIG["headings"][level]
            h = headings.get(level, {})
            v = self.heading_vars[level]
            v["font_zh"].set(h.get("font_zh", d["font_zh"]))
            v["font_size"].set(PT_TO_SIZE_NAME.get(h.get("font_size", d["font_size"]), "三号"))
            v["bold"].set(bool(h.get("bold", d.get("bold", True))))
            v["align"].set(h.get("align", d.get("align", "left")))
            v["space_before_lines"].set(str(h.get("space_before_lines", d.get("space_before_lines", 0))))
            v["space_after_lines"].set(str(h.get("space_after_lines", d.get("space_after_lines", 0))))
            v["line_spacing"].set(str(h.get("line_spacing", d.get("line_spacing", 1.5))))

    def _save_config(self):
        try:
            cfg = self._collect_config()
        except ValueError as exc:
            self._log(f"配置未保存: {exc}")
            return

        path = filedialog.asksaveasfilename(
            title="保存配置",
            initialfile="md2word_config.yaml",
            defaultextension=".yaml",
            filetypes=[("YAML 配置文件", "*.yaml *.yml"), ("所有文件", "*.*")],
        )
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                yaml.safe_dump(cfg, f, allow_unicode=True, sort_keys=False)
            self._log(f"配置已保存: {path}")
        except Exception as exc:
            self._log(f"保存配置失败: {exc}")

    def _load_config(self):
        path = filedialog.askopenfilename(
            title="加载配置",
            defaultextension=".yaml",
            filetypes=[("YAML 配置文件", "*.yaml *.yml"), ("所有文件", "*.*")],
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                cfg = yaml.safe_load(f) or {}
            self._apply_config(cfg)
            self._log(f"配置已加载: {path}")
        except Exception as exc:
            self._log(f"加载配置失败: {exc}")

    # ------------------------------------------------------------------ 日志

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

    # ------------------------------------------------------------------ 转换

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

        body_font_size_name = self.body_vars["font_size"].get()
        if body_font_size_name not in NAME_TO_PT:
            raise ValueError("正文字号选择无效")
        cfg["body"]["font_en"] = self.body_vars["font_en"].get().strip() or DEFAULT_CONFIG["body"]["font_en"]
        cfg["body"]["font_zh"] = self.body_vars["font_zh"].get().strip() or DEFAULT_CONFIG["body"]["font_zh"]
        cfg["body"]["font_size"] = NAME_TO_PT[body_font_size_name]
        cfg["body"]["line_spacing"] = self._get_numeric(self.body_vars["line_spacing"], "正文行距")
        cfg["body"]["first_line_indent"] = bool(self.body_vars["first_line_indent"].get())
        cfg["body"]["color"] = self.body_vars["color"].get().strip() or DEFAULT_CONFIG["body"]["color"]

        code_font_size_name = self.code_vars["font_size"].get()
        if code_font_size_name not in NAME_TO_PT:
            raise ValueError("代码字号选择无效")
        cfg["code"]["font_en"] = self.code_vars["font_en"].get().strip() or DEFAULT_CONFIG["code"]["font_en"]
        cfg["code"]["font_zh"] = self.code_vars["font_zh"].get().strip() or DEFAULT_CONFIG["code"]["font_zh"]
        cfg["code"]["font_size"] = NAME_TO_PT[code_font_size_name]
        cfg["code"]["line_spacing"] = self._get_numeric(self.code_vars["line_spacing"], "代码行距")
        cfg["code"]["color"] = self.code_vars["color"].get().strip() or DEFAULT_CONFIG["code"]["color"]

        cfg["table"]["style"] = self.table_vars["style"].get()
        cfg["table"]["header_align"] = self.table_vars["header_align"].get()
        cfg["table"]["header_bold"] = bool(self.table_vars["header_bold"].get())

        for level in ("h1", "h2", "h3", "h4"):
            v = self.heading_vars[level]
            size_name = v["font_size"].get()
            if size_name not in NAME_TO_PT:
                raise ValueError(f"{level.upper()} 字号选择无效")
            cfg["headings"][level] = {
                "font_zh": v["font_zh"].get().strip() or DEFAULT_CONFIG["headings"][level]["font_zh"],
                "font_size": NAME_TO_PT[size_name],
                "bold": bool(v["bold"].get()),
                "align": v["align"].get(),
                "space_before_lines": self._get_numeric(v["space_before_lines"], f"{level.upper()} 段前空行"),
                "space_after_lines": self._get_numeric(v["space_after_lines"], f"{level.upper()} 段后空行"),
                "line_spacing": self._get_numeric(v["line_spacing"], f"{level.upper()} 行距"),
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
