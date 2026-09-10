#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gui.py —— restore_zhongzhuhao 的图形界面。

功能：选择源 docx（含着重号/字体颜色），
     读取（或由程序调用 pandoc 生成）Markdown，
     把标注以内联 HTML 补回，输出 Markdown 或 HTML。

运行：python gui.py
"""

from __future__ import annotations

import io
import os
import threading
import traceback
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from tkinter import font as tkfont

from .core import (
    DEFAULT_EMPHASIS_CLASS,
    DEFAULT_IGNORE_COLORS,
    Options,
    process,
    run_pandoc,
)

# ------------------------------------------------------------------ 配色/字体 --
BG = "#eef1f8"
CARD = "#ffffff"
ACCENT = "#2f6fed"
ACCENT_DARK = "#1f56cc"
TEXT = "#20242c"
MUTED = "#6b7280"
BORDER = "#dce1ec"


def pick_family(root: tk.Misc) -> str:
    families = set(tkfont.families(root))
    for name in ("Microsoft YaHei UI", "Microsoft YaHei", "Segoe UI", "Arial"):
        if name in families:
            return name
    return "TkDefaultFont"


class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        root.title("着重号 / 字体颜色 修复工具")
        root.configure(bg=BG)
        root.minsize(720, 540)
        self.family = pick_family(root)

        self.docx_var = tk.StringVar()
        self.md_var = tk.StringVar()
        self.out_var = tk.StringVar()
        self.pandoc_var = tk.StringVar(value="pandoc")
        self.use_pandoc_var = tk.BooleanVar(value=False)
        self.emphasis_var = tk.BooleanVar(value=True)
        self.color_var = tk.BooleanVar(value=True)
        self.style_var = tk.BooleanVar(value=True)
        self.class_var = tk.StringVar(value=DEFAULT_EMPHASIS_CLASS)
        self.ignore_colors_var = tk.StringVar(value=",".join(DEFAULT_IGNORE_COLORS))
        self.format_var = tk.StringVar(value="markdown")
        self.status_var = tk.StringVar(value="就绪")

        self._init_style()
        self._build()

    # ------------------------------------------------------------- 样式 --
    def _init_style(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        f = self.family
        style.configure(".", background=BG, foreground=TEXT, font=(f, 10))
        style.configure("Card.TFrame", background=CARD)
        style.configure("Card.TLabel", background=CARD, foreground=TEXT, font=(f, 10))
        style.configure("Muted.TLabel", background=CARD, foreground=MUTED, font=(f, 9))
        style.configure("Section.TLabel", background=CARD, foreground=ACCENT,
                        font=(f, 11, "bold"))
        style.configure("Card.TCheckbutton", background=CARD, foreground=TEXT, font=(f, 10))
        style.map("Card.TCheckbutton", background=[("active", CARD)])
        style.configure("Card.TRadiobutton", background=CARD, foreground=TEXT, font=(f, 10))
        style.map("Card.TRadiobutton", background=[("active", CARD)])
        style.configure("TButton", font=(f, 10), padding=(10, 6))
        style.configure("Accent.TButton", background=ACCENT, foreground="white",
                        font=(f, 10, "bold"), borderwidth=0, focuscolor=ACCENT,
                        padding=(16, 9))
        style.map("Accent.TButton",
                  background=[("active", ACCENT_DARK), ("pressed", ACCENT_DARK),
                              ("disabled", "#a9bcf0")],
                  foreground=[("disabled", "#eef1f8")])
        style.configure("TEntry", padding=5, fieldbackground=CARD)
        style.configure("Header.TFrame", background=ACCENT)
        style.configure("Title.TLabel", background=ACCENT, foreground="white",
                        font=(f, 17, "bold"))
        style.configure("Sub.TLabel", background=ACCENT, foreground="#dbe6ff", font=(f, 9))

    # --------------------------------------------------------------- UI --
    def _card(self, parent, title: str):
        card = ttk.Frame(parent, style="Card.TFrame", padding=10)
        card.pack(fill="x", pady=(0, 8))
        ttk.Label(card, text=title, style="Section.TLabel").pack(anchor="w", pady=(0, 6))
        return card

    def _file_row(self, parent, label, var, command, hint=""):
        row = ttk.Frame(parent, style="Card.TFrame")
        row.pack(fill="x", pady=2)
        ttk.Label(row, text=label, style="Card.TLabel", width=14, anchor="e").pack(side="left")
        ttk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True, padx=(8, 8))
        ttk.Button(row, text="浏览…", command=command).pack(side="left")
        if hint:
            ttk.Label(row, text=hint, style="Muted.TLabel").pack(side="left", padx=(8, 0))
        return row

    def _build(self) -> None:
        # 顶部标题栏
        header = tk.Frame(self.root, bg=ACCENT)
        header.pack(fill="x")
        tk.Label(header, text="着重号 / 字体颜色 修复工具", bg=ACCENT, fg="white",
                 font=(self.family, 16, "bold")).pack(anchor="w", padx=20, pady=(12, 1))
        tk.Label(header, text="读回 docx 的 w:em 与 w:color，在 Markdown / HTML 中以内联 <span> 补回标注",
                 bg=ACCENT, fg="#dbe6ff", font=(self.family, 9)).pack(anchor="w", padx=20, pady=(0, 12))

        # 可滚动内容区：小屏幕也能访问到全部控件（含按钮与日志）
        outer = ttk.Frame(self.root)
        outer.pack(fill="both", expand=True)
        outer.rowconfigure(0, weight=1)
        outer.columnconfigure(0, weight=1)
        self.canvas = tk.Canvas(outer, bg=BG, highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        vbar = ttk.Scrollbar(outer, orient="vertical", command=self.canvas.yview)
        vbar.grid(row=0, column=1, sticky="ns")
        self.canvas.configure(yscrollcommand=vbar.set)

        body = ttk.Frame(self.canvas, padding=12)
        win = self.canvas.create_window((0, 0), window=body, anchor="nw")
        body.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )
        self.canvas.bind(
            "<Configure>",
            lambda e: self.canvas.itemconfigure(win, width=e.width),
        )
        self.root.bind("<MouseWheel>", self._on_wheel)

        # 文件卡片
        c1 = self._card(body, "文件")
        self._file_row(c1, "源 docx", self.docx_var, self._pick_docx, "必选")
        self._file_row(c1, "输入 Markdown", self.md_var, self._pick_md)
        self._file_row(c1, "输出文件", self.out_var, self._pick_out)

        opt_row = ttk.Frame(c1, style="Card.TFrame")
        opt_row.pack(fill="x", pady=(8, 0))
        ttk.Checkbutton(opt_row, text="由 pandoc 生成 Markdown（忽略上面的输入）",
                        variable=self.use_pandoc_var,
                        style="Card.TCheckbutton").pack(side="left")
        ttk.Label(opt_row, text="  pandoc：", style="Card.TLabel").pack(side="left")
        self.pandoc_entry = ttk.Entry(opt_row, textvariable=self.pandoc_var, width=18)
        self.pandoc_entry.pack(side="left")

        fmt_row = ttk.Frame(c1, style="Card.TFrame")
        fmt_row.pack(fill="x", pady=(8, 0))
        ttk.Label(fmt_row, text="输出格式：", style="Card.TLabel").pack(side="left")
        ttk.Radiobutton(fmt_row, text="Markdown (.md)", value="markdown",
                        variable=self.format_var, command=self._on_format,
                        style="Card.TRadiobutton").pack(side="left", padx=4)
        ttk.Radiobutton(fmt_row, text="HTML (.html)", value="html",
                        variable=self.format_var, command=self._on_format,
                        style="Card.TRadiobutton").pack(side="left", padx=4)
        ttk.Label(fmt_row, text="（HTML 会用 pandoc 转成独立网页）",
                  style="Muted.TLabel").pack(side="left", padx=(8, 0))

        # 选项卡片
        c2 = self._card(body, "处理选项")
        sw = ttk.Frame(c2, style="Card.TFrame")
        sw.pack(fill="x")
        ttk.Checkbutton(sw, text="着重号（字下加点）", variable=self.emphasis_var,
                        style="Card.TCheckbutton").pack(side="left", padx=(0, 14))
        ttk.Checkbutton(sw, text="字体颜色", variable=self.color_var,
                        style="Card.TCheckbutton").pack(side="left", padx=(0, 14))
        ttk.Checkbutton(sw, text="注入样式", variable=self.style_var,
                        style="Card.TCheckbutton").pack(side="left", padx=(0, 14))
        ttk.Label(sw, text="着重号 class：", style="Card.TLabel").pack(side="left")
        ttk.Entry(sw, textvariable=self.class_var, width=12).pack(side="left")

        sw2 = ttk.Frame(c2, style="Card.TFrame")
        sw2.pack(fill="x", pady=(8, 0))
        ttk.Label(sw2, text="忽略颜色（默认黑/白，逗号分隔）：", style="Card.TLabel").pack(side="left")
        ttk.Entry(sw2, textvariable=self.ignore_colors_var, width=24).pack(side="left")

        # 效果图例（并入选项卡片，紧凑一行）
        legend = ttk.Frame(c2, style="Card.TFrame")
        legend.pack(fill="x", pady=(8, 0))
        canvas = tk.Canvas(legend, height=32, bg=CARD, highlightthickness=0)
        canvas.pack(fill="x")
        self._draw_legend(canvas)

        # 按钮
        btns = ttk.Frame(body)
        btns.pack(fill="x", pady=(0, 8))
        self.run_btn = ttk.Button(btns, text="开始处理", style="Accent.TButton",
                                  command=self._run)
        self.run_btn.pack(side="left")
        ttk.Button(btns, text="试运行", command=lambda: self._run(dry_run=True)).pack(
            side="left", padx=8)
        self.open_btn = ttk.Button(btns, text="打开输出位置", command=self._open_out,
                                   state="disabled")
        self.open_btn.pack(side="left", padx=8)
        ttk.Button(btns, text="清空日志", command=self._clear_log).pack(side="left")

        # 日志卡片
        c4 = ttk.Frame(body, style="Card.TFrame", padding=10)
        c4.pack(fill="both", expand=True)
        ttk.Label(c4, text="日志", style="Section.TLabel").pack(anchor="w", pady=(0, 6))
        self.log = scrolledtext.ScrolledText(
            c4, height=8, wrap="word", relief="flat", borderwidth=0,
            bg="#fbfcfe", fg=TEXT, insertbackground=TEXT, font=(self.family, 9))
        self.log.pack(fill="both", expand=True)
        self.log.configure(state="disabled")

        # 状态栏
        status = tk.Label(self.root, textvariable=self.status_var, bg=BG, fg=MUTED,
                          anchor="w", font=(self.family, 9), padx=18, pady=6)
        status.pack(fill="x")

    def _draw_legend(self, canvas: tk.Canvas) -> None:
        y = 22
        canvas.create_text(12, y, anchor="w", text="着重号：", fill=MUTED,
                           font=(self.family, 10))
        x = 78
        for ch in "着重号":
            canvas.create_text(x, y - 6, anchor="w", text=ch, fill=TEXT,
                               font=(self.family, 13))
            canvas.create_oval(x + 6, y + 8, x + 10, y + 12, fill=ACCENT, outline=ACCENT)
            x += 26
        canvas.create_text(x + 24, y, anchor="w", text="字体颜色：", fill=MUTED,
                           font=(self.family, 10))
        canvas.create_text(x + 112, y, anchor="w", text="红色", fill="#FF0000",
                           font=(self.family, 13, "bold"))
        canvas.create_text(x + 158, y, anchor="w", text="  普通文字", fill=TEXT,
                           font=(self.family, 12))

    # ------------------------------------------------------------ 交互 --
    def _on_format(self) -> None:
        out = self.out_var.get().strip()
        if out:
            ext = ".html" if self.format_var.get() == "html" else ".md"
            self.out_var.set(str(Path(out).with_suffix(ext)))

    def _pick_docx(self) -> None:
        path = filedialog.askopenfilename(
            title="选择源 docx", filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")])
        if path:
            self.docx_var.set(path)
            if not self.out_var.get():
                ext = ".html" if self.format_var.get() == "html" else ".md"
                self.out_var.set(str(Path(path).with_suffix(ext)))

    def _pick_md(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 Markdown",
            filetypes=[("Markdown", "*.md *.markdown"), ("所有文件", "*.*")])
        if path:
            self.md_var.set(path)
            if not self.out_var.get():
                ext = ".html" if self.format_var.get() == "html" else ".md"
                self.out_var.set(str(Path(path).with_suffix(ext)))

    def _pick_out(self) -> None:
        ext = ".html" if self.format_var.get() == "html" else ".md"
        path = filedialog.asksaveasfilename(
            title="保存为", defaultextension=ext,
            filetypes=[("HTML", "*.html"), ("Markdown", "*.md"), ("所有文件", "*.*")])
        if path:
            self.out_var.set(path)

    def _open_out(self) -> None:
        path = self.out_var.get().strip()
        if not path:
            return
        folder = Path(path).parent
        try:
            if os.name == "nt":
                os.startfile(str(folder))  # type: ignore[attr-defined]
            else:
                import subprocess
                subprocess.Popen(["xdg-open", str(folder)])
        except Exception:
            messagebox.showinfo("提示", "输出目录：\n%s" % folder)

    def _log(self, text: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", text if text.endswith("\n") else text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _clear_log(self) -> None:
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _on_wheel(self, event) -> None:
        # 指针在日志上时让日志自己滚动，其余位置滚动整个页面
        log = getattr(self, "log", None)
        w = self.root.winfo_containing(event.x_root, event.y_root)
        while w is not None:
            if w is log:
                return
            w = getattr(w, "master", None)
        delta = -1 if event.delta > 0 else 1
        self.canvas.yview_scroll(delta, "units")

    def _set_running(self, running: bool) -> None:
        self.run_btn.configure(state="disabled" if running else "normal")
        self.status_var.set("处理中…" if running else "就绪")

    # ------------------------------------------------------------ 运行 --
    def _run(self, dry_run: bool = False) -> None:
        docx = self.docx_var.get().strip()
        if not docx or not Path(docx).exists():
            messagebox.showerror("错误", "请选择有效的源 docx 文件。")
            return

        use_pandoc = self.use_pandoc_var.get()
        md_path = self.md_var.get().strip()
        if not use_pandoc and (not md_path or not Path(md_path).exists()):
            messagebox.showerror("错误", "请选择有效的 Markdown 文件，或勾选“由 pandoc 生成”。")
            return

        out_path = self.out_var.get().strip()
        if not dry_run and not out_path:
            messagebox.showerror("错误", "请指定输出文件。")
            return

        ignore_colors = []
        for c in self.ignore_colors_var.get().split(","):
            c = c.strip().lstrip("#")
            if c:
                ignore_colors.append("#" + c.upper())

        opts = Options(
            emphasis=self.emphasis_var.get(),
            color=self.color_var.get(),
            emphasis_class=self.class_var.get().strip() or DEFAULT_EMPHASIS_CLASS,
            inject_style=self.style_var.get(),
            ignore_colors=tuple(ignore_colors),
        )
        fmt = self.format_var.get()
        pandoc_bin = self.pandoc_var.get().strip() or "pandoc"

        self._set_running(True)
        self._log("=" * 64)
        self._log("开始处理（%s，输出 %s）…" % ("试运行" if dry_run else "写出文件", fmt))
        threading.Thread(
            target=self._worker,
            args=(docx, use_pandoc, md_path, out_path, opts, fmt, pandoc_bin, dry_run),
            daemon=True,
        ).start()

    def _worker(self, docx, use_pandoc, md_path, out_path, opts, fmt, pandoc_bin, dry_run):
        buf = io.StringIO()
        ok = False
        try:
            with redirect_stdout(buf), redirect_stderr(buf):
                if use_pandoc:
                    print("调用 pandoc 生成 Markdown …")
                    source = run_pandoc(docx, pandoc_bin, "markdown")
                else:
                    source = Path(md_path).read_text(encoding="utf-8")

                result, stats = process(
                    docx, source, opts,
                    output_format=fmt, pandoc_bin=pandoc_bin, dry_run=dry_run,
                )
                print("已包裹 %d 个（着重号 %d / 颜色 %d），未命中 %d 个。" % (
                    stats["wrapped"], stats["by_type"]["emphasis"],
                    stats["by_type"]["color"], len(stats["missed"])))
                if stats["missed"]:
                    print("未命中的短语：")
                    for t in stats["missed"]:
                        print("  - %s" % t)

                if not dry_run:
                    Path(out_path).write_text(result, encoding="utf-8")
                    print("已写入：%s" % out_path)
                    ok = True
        except BaseException:
            # 注意：core 在 pandoc 失败时会抛 SystemExit，必须一并捕获并显示
            buf.write("\n处理失败：\n")
            buf.write(traceback.format_exc())
        finally:
            text = buf.getvalue()
            self.root.after(0, lambda: self._finish(text, dry_run, ok))

    def _finish(self, text: str, dry_run: bool, ok: bool) -> None:
        self._log(text)
        self._set_running(False)
        if ok:
            self.status_var.set("完成")
            self.open_btn.configure(state="normal")
            messagebox.showinfo("完成", "处理完成，已写出结果文件。")
        elif not dry_run:
            self.status_var.set("失败")
        else:
            self.status_var.set("试运行结束")


def _enable_dpi_awareness() -> None:
    """让窗口在高 DPI 下按物理像素布局，避免被系统放大后超出屏幕。"""
    if os.name != "nt":
        return
    try:
        from ctypes import windll
        try:
            windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def _work_area():
    """返回 Windows 工作区（不含任务栏）(left, top, right, bottom)，失败返回 None。"""
    if os.name != "nt":
        return None
    try:
        import ctypes
        from ctypes import wintypes

        class RECT(ctypes.Structure):
            _fields_ = [
                ("left", wintypes.LONG),
                ("top", wintypes.LONG),
                ("right", wintypes.LONG),
                ("bottom", wintypes.LONG),
            ]

        r = RECT()
        SPI_GETWORKAREA = 0x0030
        if ctypes.windll.user32.SystemParametersInfoW(
            SPI_GETWORKAREA, 0, ctypes.byref(r), 0
        ):
            return r.left, r.top, r.right, r.bottom
    except Exception:
        pass
    return None


def main() -> None:
    _enable_dpi_awareness()
    root = tk.Tk()
    App(root)
    root.update_idletasks()

    area = _work_area()
    if area:
        left, top, right, bottom = area
        avail_w, avail_h = right - left, bottom - top
    else:
        left, top = 0, 0
        avail_w, avail_h = root.winfo_screenwidth(), root.winfo_screenheight()

    w = min(860, max(600, avail_w - 50))
    h = min(720, max(440, avail_h - 40))
    x = left + max(0, (avail_w - w) // 2)
    y = top + max(0, (avail_h - h) // 2)
    root.geometry("%dx%d+%d+%d" % (w, h, x, y))
    root.mainloop()


if __name__ == "__main__":
    main()
