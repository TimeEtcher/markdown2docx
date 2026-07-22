"""Tkinter GUI：选 Markdown 文件、选配置（可空）、选输出路径，一键转换。"""

from __future__ import annotations

import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, ttk


def run_gui() -> None:
    root = tk.Tk()
    App(root)
    root.mainloop()


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("md2word - Markdown 转 Word")
        root.geometry("640x420")

        self.md_var = tk.StringVar()
        self.cfg_var = tk.StringVar()
        self.out_var = tk.StringVar()
        self.log_queue: queue.Queue[str] = queue.Queue()

        pad = {"padx": 8, "pady": 4}
        frame = ttk.Frame(root)
        frame.pack(fill=tk.X, **pad)

        self._row(frame, 0, "Markdown 文件:", self.md_var, self._pick_md)
        self._row(frame, 1, "配置文件 (可选):", self.cfg_var, self._pick_cfg)
        self._row(frame, 2, "输出文件:", self.out_var, self._pick_out)

        self.convert_btn = ttk.Button(root, text="开始转换", command=self._convert)
        self.convert_btn.pack(pady=6)

        self.log = tk.Text(root, height=12, state=tk.DISABLED, wrap=tk.WORD)
        self.log.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        self.root.after(100, self._poll_log)

    def _row(self, parent, row, label, var, command):
        ttk.Label(parent, text=label, width=16).grid(row=row, column=0, sticky=tk.W, pady=3)
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

    def _pick_cfg(self):
        path = filedialog.askopenfilename(
            title="选择 YAML 配置文件",
            filetypes=[("YAML", "*.yaml *.yml"), ("所有文件", "*.*")])
        if path:
            self.cfg_var.set(path)

    def _pick_out(self):
        path = filedialog.asksaveasfilename(
            title="选择输出位置", defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx")])
        if path:
            self.out_var.set(path)

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

    def _convert(self):
        md_path = self.md_var.get().strip()
        if not md_path:
            self._log("请先选择 Markdown 文件。")
            return
        cfg_path = self.cfg_var.get().strip() or None
        out_path = self.out_var.get().strip() or None
        self.convert_btn.config(state=tk.DISABLED)
        self._log(f"开始转换: {md_path}")
        if cfg_path:
            self._log(f"使用配置: {cfg_path}")
        else:
            self._log("使用默认配置。")

        def worker():
            from .cli import convert
            try:
                out = convert(md_path, cfg_path, out_path)
                self.log_queue.put(f"完成: {out}")
            except Exception as exc:
                self.log_queue.put(f"转换失败: {exc}")
            finally:
                self.root.after(0, lambda: self.convert_btn.config(state=tk.NORMAL))

        threading.Thread(target=worker, daemon=True).start()
