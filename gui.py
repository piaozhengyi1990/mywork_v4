#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PDF 转 Word 图形界面

功能: 拖拽/选择文件夹/单文件、批量勾选、后台线程转换(可停止)、
      逐文件+OCR逐页进度、结果表格、失败重试、错误详情悬停/点击查看、
      实时日志面板、路径记忆、工具栏图标、支持 PyInstaller 单文件打包。
"""

import json
import logging
import os
import queue
import shutil
import sys
import threading
import tkinter as tk
from logging.handlers import RotatingFileHandler
from tkinter import filedialog, messagebox, scrolledtext, ttk

from pdf_to_word_converter import PDFToWordConverter

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _HAS_DND = True
except ImportError:
    DND_FILES = None
    TkinterDnD = None
    _HAS_DND = False


def _app_dir():
    """打包为 exe 后返回 exe 所在目录, 源码运行时返回脚本目录"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _resource_path(name):
    """获取资源文件路径: 打包后从 _MEIPASS 解压目录查找, 源码运行时从脚本目录查找"""
    if getattr(sys, "frozen", False):
        base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(sys.executable)))
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, name)


def _enable_dpi_awareness():
    """Windows 高分屏下启用 DPI 感知, 避免界面发虚、过小"""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def _parse_dnd_paths(data):
    """解析 Windows 拖放路径: 支持 {带空格路径} 与普通空格分隔"""
    paths, i, n = [], 0, len(data)
    while i < n:
        if data[i] == '{':
            j = data.find('}', i)
            if j == -1:
                break
            paths.append(data[i + 1:j])
            i = j + 1
        elif data[i] == ' ':
            i += 1
        else:
            j = i
            while j < n and data[j] != ' ':
                j += 1
            paths.append(data[i:j])
            i = j
    return [p.strip('"') for p in paths if p]


class _ToolTip:
    """轻量悬停提示 (用于状态列错误详情)"""

    def __init__(self, widget):
        self._widget = widget
        self._tip = None
        widget.bind("<Motion>", self._on_motion, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _on_motion(self, event):
        try:
            iid = self._widget.identify_row(event.y)
            col = self._widget.identify_column(event.x)
        except tk.TclError:
            return
        if not iid or col != "#2":
            self._hide()
            return
        detail = self._widget.set(iid, "status")
        # 只对带详情的状态显示悬停提示
        if not detail or len(detail) < 12:
            self._hide()
            return
        self._show(event, detail)

    def _show(self, event, text):
        self._hide()
        tw = tk.Toplevel(self._widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{event.x_root + 14}+{event.y_root + 18}")
        frame = tk.Frame(tw, bg="#111827", padx=8, pady=6,
                         highlightbackground="#374151", highlightthickness=1)
        frame.pack()
        # 折行显示, 限制宽度
        wrap = text if len(text) < 480 else text[:480] + "…"
        tk.Label(frame, text=wrap, bg="#111827", fg="#F9FAFB",
                 font=("Microsoft YaHei UI", 9), justify="left",
                 wraplength=460).pack()
        self._tip = tw

    def _hide(self, _event=None):
        if self._tip is not None:
            try:
                self._tip.destroy()
            except tk.TclError:
                pass
            self._tip = None


APP_NAME = "PDF转Word - PDF 批量转换 Word 文档"
APP_VERSION = "1.2.0"
COPYRIGHT_TEXT = "内部使用.禁止外传与商业使用"
CONFIG_DIR = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "PdfToWord")
CONFIG_FILE = os.path.join(CONFIG_DIR, "config.json")
COLUMNS = ("check", "status", "filename", "pages", "output")
COL_WIDTHS = {"check": 40, "status": 220, "filename": 320, "pages": 64, "output": 240}

# ================= 设计令牌 (Design Tokens) =================
C = {
    "primary": "#2B579A",
    "primary_dark": "#1E3F73",
    "primary_light": "#E8EEF8",
    "accent": "#E5484D",

    "header_bg": "#1A2332",
    "header_fg": "#FFFFFF",
    "header_sub": "#8B9BB4",

    "bg": "#F5F7FA",
    "card": "#FFFFFF",
    "card_border": "#E8ECF1",

    "text": "#1F2937",
    "text_sec": "#6B7280",
    "text_muted": "#9CA3AF",

    "ok": "#10B981",
    "ok_bg": "#ECFDF5",
    "err": "#EF4444",
    "err_bg": "#FEF2F2",
    "run": "#3B82F6",
    "run_bg": "#EFF6FF",
    "pending": "#9CA3AF",
    "pending_bg": "#F9FAFB",
    "warn": "#F59E0B",
    "warn_bg": "#FFFBEB",

    "toolbar_bg": "#FFFFFF",
    "toolbar_border": "#E5E7EB",
    "btn_secondary_bg": "#F3F4F6",
    "btn_secondary_hover": "#E5E7EB",
    "btn_danger": "#EF4444",
    "btn_danger_hover": "#DC2626",

    "statusbar_bg": "#1A2332",
    "statusbar_fg": "#D1D5DB",
}

FONT = "Microsoft YaHei UI"

logger = logging.getLogger("converter")


def _load_toolbar_icons():
    """运行时生成工具栏图标 -> {name: PhotoImage}"""
    icons = {}
    try:
        from PIL import Image, ImageDraw
        import io
    except ImportError:
        return icons

    S = 16
    GRAY = (75, 85, 99, 255)
    WHITE = (255, 255, 255, 255)

    def canvas():
        im = Image.new("RGBA", (S, S), (0, 0, 0, 0))
        return im, ImageDraw.Draw(im)

    def to_photo(im):
        buf = io.BytesIO()
        im.save(buf, format="PNG")
        return tk.PhotoImage(data=buf.getvalue())

    try:
        im, d = canvas()
        d.line([(3, 8), (13, 8)], fill=GRAY, width=2)
        d.line([(8, 3), (8, 13)], fill=GRAY, width=2)
        icons["add"] = to_photo(im)

        im, d = canvas()
        d.polygon([(2, 5), (6, 5), (7, 3), (12, 3), (13, 5), (14, 5),
                   (14, 13), (2, 13)], outline=GRAY, width=1)
        d.line([(2, 6), (14, 6)], fill=GRAY, width=1)
        icons["folder"] = to_photo(im)

        im, d = canvas()
        d.rectangle([(2, 2), (14, 14)], outline=GRAY, width=1)
        d.line([(4, 8), (7, 11), (12, 5)], fill=GRAY, width=2)
        icons["check"] = to_photo(im)

        im, d = canvas()
        d.rectangle([(2, 2), (14, 14)], outline=GRAY, width=1)
        icons["uncheck"] = to_photo(im)

        im, d = canvas()
        d.polygon([(4, 2), (14, 8), (4, 14)], fill=WHITE)
        icons["play"] = to_photo(im)

        im, d = canvas()
        d.rectangle([(3, 3), (13, 13)], fill=WHITE)
        icons["stop"] = to_photo(im)

        im, d = canvas()
        d.arc([(3, 3), (13, 13)], 40, 320, fill=GRAY, width=2)
        d.polygon([(10, 1), (15, 5), (10, 6)], fill=GRAY)
        icons["retry"] = to_photo(im)

        im, d = canvas()
        d.line([(3, 4), (13, 4)], fill=GRAY, width=1)
        d.line([(6, 2), (10, 2)], fill=GRAY, width=1)
        d.rectangle([(4, 5), (12, 14)], outline=GRAY, width=1)
        d.line([(7, 7), (7, 12)], fill=GRAY, width=1)
        d.line([(9, 7), (9, 12)], fill=GRAY, width=1)
        icons["trash"] = to_photo(im)
    except Exception:
        return icons
    return icons


class ConverterApp:
    def __init__(self):
        if _HAS_DND:
            self._root = TkinterDnD.Tk()
        else:
            self._root = tk.Tk()
        self._root.title(APP_NAME)
        self._root.minsize(960, 640)
        self._root.geometry("1120x740")
        self._root.configure(bg=C["bg"])
        self._root.protocol("WM_DELETE_WINDOW", self._on_close)

        self._items = []
        self._thread = None
        self._stop = threading.Event()
        self._ui_queue = queue.Queue()
        self._log_collapsed = False
        self._icons = _load_toolbar_icons()
        self._last_mode = "—"
        self._last_elapsed = None
        self._status_tip = None

        self._set_window_icon()
        self._build_ui()
        self._bind_dnd()
        self._root.after(50, self._drain_ui_queue)
        self._load_config()
        self._bind_shortcuts()
        self._refresh_status_right()
        logger.info("就绪 - 拖拽 PDF 到窗口, 或添加文件/扫描输入目录")

    def _set_window_icon(self):
        for name in ("icon.png", "icon.ico"):
            p = _resource_path(name)
            if not os.path.isfile(p):
                continue
            try:
                if name.endswith(".png"):
                    img = tk.PhotoImage(file=p)
                    self._root.iconphoto(True, img)
                    self._icon_img = img
                else:
                    self._root.iconbitmap(p)
                break
            except tk.TclError:
                continue

    # ================================================================== UI

    def _build_ui(self):
        self._setup_styles()
        self._build_header()
        self._build_statusbar()

        content = ttk.Frame(self._root, style="Bg.TFrame")
        content.pack(fill="both", expand=True)
        content.columnconfigure(0, weight=1)
        content.rowconfigure(1, weight=1)

        # ---- 目录设置卡片 ----
        dir_card = self._card(content, padding=(16, 14))
        dir_card.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 0))
        dir_card.columnconfigure(1, weight=1)

        in_frame = ttk.Frame(dir_card, style="Card.TFrame")
        in_frame.grid(row=0, column=0, columnspan=3, sticky="ew")
        in_frame.columnconfigure(1, weight=1)

        ttk.Label(in_frame, text="输入目录", style="Label.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 10))
        self._var_in = tk.StringVar()
        ttk.Entry(in_frame, textvariable=self._var_in, font=(FONT, 9),
                  style="Entry.TEntry").grid(row=0, column=1, sticky="ew", padx=(0, 8))
        self._btn_browse_in = ttk.Button(in_frame, text="浏览", style="Browse.TButton",
                                         command=lambda: self._pick_dir("in"))
        self._btn_browse_in.grid(row=0, column=2)

        out_frame = ttk.Frame(dir_card, style="Card.TFrame")
        out_frame.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        out_frame.columnconfigure(1, weight=1)

        ttk.Label(out_frame, text="输出目录", style="Label.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 10))
        self._var_out = tk.StringVar()
        ttk.Entry(out_frame, textvariable=self._var_out, font=(FONT, 9),
                  style="Entry.TEntry").grid(row=0, column=1, sticky="ew", padx=(0, 8))
        self._btn_browse_out = ttk.Button(out_frame, text="浏览", style="Browse.TButton",
                                          command=lambda: self._pick_dir("out"))
        self._btn_browse_out.grid(row=0, column=2)

        # ---- 文件列表卡片 ----
        list_card = self._card(content, padding=(12, 10))
        list_card.grid(row=1, column=0, sticky="nsew", padx=12, pady=(10, 0))
        list_card.columnconfigure(0, weight=1)
        list_card.rowconfigure(0, weight=1)

        tree_frame = ttk.Frame(list_card, style="Card.TFrame")
        tree_frame.grid(row=0, column=0, sticky="nsew")
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        self._tree = ttk.Treeview(tree_frame, columns=COLUMNS, show="headings",
                                  selectmode="browse", style="Treeview")
        for c in COLUMNS:
            opts = {"text": {
                "check": "  ", "status": "状态", "filename": "文件名",
                "pages": "页数", "output": "输出文件"
            }[c], "anchor": "center" if c in ("check", "status", "pages") else "w"}
            if c == "check":
                opts["command"] = self._on_header_check
            self._tree.heading(c, **opts)
            self._tree.column(c, width=COL_WIDTHS[c], minwidth=40,
                              anchor="center" if c in ("check", "status", "pages") else "w")
        self._tree.column("filename", anchor="w", stretch=True)

        self._tree.tag_configure("ok", foreground=C["ok"], background=C["ok_bg"])
        self._tree.tag_configure("err", foreground=C["err"], background=C["err_bg"])
        self._tree.tag_configure("run", foreground=C["run"], background=C["run_bg"])
        self._tree.tag_configure("pending", foreground=C["pending"], background=C["pending_bg"])
        self._tree.tag_configure("warn", foreground=C["warn"], background=C["warn_bg"])

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        self._tree.bind("<Button-1>", self._on_click)
        self._tree.bind("<Double-Button-1>", self._on_dbl)
        self._tree.bind("<Button-3>", self._on_right_click)
        self._tree.bind("<Delete>", self._on_delete_key)
        _ToolTip(self._tree)

        # 空状态引导
        self._empty_state = tk.Frame(tree_frame, bg=C["card"])
        tk.Label(self._empty_state, text="还没有文件",
                 bg=C["card"], fg=C["text_sec"],
                 font=(FONT, 13, "bold")).pack()
        dnd_hint = "将 PDF 文件直接拖拽到本窗口，" if _HAS_DND else ""
        tk.Label(self._empty_state,
                 text=f"{dnd_hint}或点击「添加文件」选择 PDF，\n"
                      f"也可「扫描目录」一键导入输入目录下的全部 PDF。",
                 bg=C["card"], fg=C["text_muted"], font=(FONT, 9),
                 justify="center").pack(pady=(8, 0))
        self._refresh_empty_state()

        # ---- 工具栏 ----
        toolbar = tk.Frame(list_card, bg=C["toolbar_bg"], height=44)
        toolbar.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        toolbar.pack_propagate(False)

        left_btns = tk.Frame(toolbar, bg=C["toolbar_bg"])
        left_btns.pack(side="left", fill="y")

        self._list_btns = []
        self._list_btns.append(self._make_tool_btn(
            left_btns, "添加文件", self._add_files, icon="add"))
        self._make_separator(left_btns)
        self._list_btns.append(self._make_tool_btn(
            left_btns, "扫描目录", self._scan_dir, icon="folder"))
        self._make_separator(left_btns)
        self._list_btns.append(self._make_tool_btn(
            left_btns, "全选", lambda: self._set_all(True), icon="check"))
        self._list_btns.append(self._make_tool_btn(
            left_btns, "取消全选", lambda: self._set_all(False), icon="uncheck"))

        right_btns = tk.Frame(toolbar, bg=C["toolbar_bg"])
        right_btns.pack(side="right", fill="y")

        self._btn_start = self._make_primary_btn(
            right_btns, "开始转换", self._start, icon="play")
        self._make_separator(right_btns)
        self._btn_stop = self._make_danger_btn(
            right_btns, "停止", self._stop_convert, icon="stop")
        self._btn_retry = self._make_tool_btn(
            right_btns, "重试失败", self._retry, icon="retry")
        self._btn_clear = self._make_tool_btn(
            right_btns, "清空列表", self._clear, icon="trash")
        self._list_btns.append(self._btn_clear)

        self._btn_stop.configure(state="disabled")
        self._btn_retry.configure(state="disabled")

        # ---- 进度条 ----
        progress_frame = ttk.Frame(list_card, style="Card.TFrame")
        progress_frame.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        progress_frame.columnconfigure(0, weight=1)

        self._bar = ttk.Progressbar(progress_frame, mode="determinate",
                                    style="Green.Horizontal.TProgressbar")
        self._bar.grid(row=0, column=0, sticky="ew", padx=(0, 12))

        self._var_st = tk.StringVar(value="就绪")
        ttk.Label(progress_frame, textvariable=self._var_st,
                  style="Status.TLabel").grid(row=0, column=1, sticky="e")

        # ---- 日志卡片 ----
        log_card = self._card(content, padding=(12, 10))
        log_card.grid(row=2, column=0, sticky="nsew", padx=12, pady=(10, 10))
        log_card.columnconfigure(0, weight=1)
        log_card.rowconfigure(1, weight=1)

        log_header = tk.Frame(log_card, bg=C["card"])
        log_header.grid(row=0, column=0, sticky="ew")
        log_header.columnconfigure(0, weight=1)

        ttk.Label(log_header, text="运行日志", style="Label.TLabel").grid(
            row=0, column=0, sticky="w")

        log_actions = tk.Frame(log_header, bg=C["card"])
        log_actions.grid(row=0, column=1, sticky="e")

        self._btn_collapse = ttk.Button(log_actions, text="收起", style="Link.TButton",
                                        command=self._toggle_log)
        self._btn_collapse.pack(side="left", padx=(0, 8))
        ttk.Button(log_actions, text="打开输出目录", style="Link.TButton",
                   command=self._open_out).pack(side="left", padx=(0, 8))
        ttk.Button(log_actions, text="打开日志文件", style="Link.TButton",
                   command=self._open_log).pack(side="left")

        self._log = scrolledtext.ScrolledText(
            log_card, height=8, state="disabled", wrap="word",
            font=("Consolas", 9), bg="#0F172A", fg="#CBD5E1",
            insertbackground="#CBD5E1", relief="flat",
            borderwidth=0, highlightthickness=0
        )
        self._log.grid(row=1, column=0, sticky="nsew", pady=(8, 0))
        self._log.tag_configure("INFO", foreground="#CBD5E1")
        self._log.tag_configure("WARNING", foreground="#FCD34D")
        self._log.tag_configure("ERROR", foreground="#FCA5A5")

    def _btn_kwargs(self, icon=None):
        kw = {}
        if icon and icon in self._icons:
            kw["image"] = self._icons[icon]
            kw["compound"] = "left"
        return kw

    def _make_tool_btn(self, parent, text, command, icon=None):
        btn = tk.Button(
            parent, text=text, command=command,
            font=(FONT, 9), fg=C["text"], bg=C["btn_secondary_bg"],
            activebackground=C["btn_secondary_hover"],
            activeforeground=C["text"],
            relief="flat", borderwidth=0, padx=10, pady=4,
            cursor="hand2", **self._btn_kwargs(icon)
        )
        btn.pack(side="left", padx=2)
        return btn

    def _make_primary_btn(self, parent, text, command, icon=None):
        btn = tk.Button(
            parent, text=text, command=command,
            font=(FONT, 9, "bold"), fg="#FFFFFF", bg=C["primary"],
            activebackground=C["primary_dark"],
            activeforeground="#FFFFFF",
            relief="flat", borderwidth=0, padx=14, pady=5,
            cursor="hand2", **self._btn_kwargs(icon)
        )
        btn.pack(side="right", padx=(4, 0))
        return btn

    def _make_danger_btn(self, parent, text, command, icon=None):
        btn = tk.Button(
            parent, text=text, command=command,
            font=(FONT, 9), fg="#FFFFFF", bg=C["btn_danger"],
            activebackground=C["btn_danger_hover"],
            activeforeground="#FFFFFF",
            relief="flat", borderwidth=0, padx=10, pady=4,
            cursor="hand2", **self._btn_kwargs(icon)
        )
        btn.pack(side="right", padx=(4, 0))
        return btn

    def _make_separator(self, parent):
        sep = tk.Frame(parent, bg=C["toolbar_border"], width=1)
        sep.pack(side="left", fill="y", padx=6, pady=6)

    def _setup_styles(self):
        s = ttk.Style()
        try:
            s.theme_use("clam")
        except tk.TclError:
            pass

        s.configure(".", font=(FONT, 9), foreground=C["text"])
        s.configure("Card.TFrame", background=C["card"])
        s.configure("Bg.TFrame", background=C["bg"])

        s.configure("Card.TLabel", background=C["card"], foreground=C["text"])
        s.configure("Label.TLabel", background=C["card"], foreground=C["text"],
                    font=(FONT, 9, "bold"))
        s.configure("Sec.TLabel", background=C["card"], foreground=C["text_sec"])
        s.configure("Status.TLabel", background=C["card"], foreground=C["text_sec"],
                    font=(FONT, 8))

        s.configure("Entry.TEntry", font=(FONT, 9))
        s.configure("TButton", padding=(10, 4), font=(FONT, 9), relief="flat")

        s.configure("Browse.TButton", background=C["accent"], foreground="#FFFFFF",
                    bordercolor=C["accent"], lightcolor=C["accent"], darkcolor=C["accent"],
                    font=(FONT, 9), padding=(10, 4))
        s.map("Browse.TButton",
              background=[("active", "#C73E3D"), ("disabled", "#FCA5A5")],
              bordercolor=[("active", "#C73E3D"), ("disabled", "#FCA5A5")])

        s.configure("Link.TButton", background=C["card"], foreground=C["primary"],
                    bordercolor=C["card"], lightcolor=C["card"], darkcolor=C["card"],
                    font=(FONT, 8), padding=(6, 2))
        s.map("Link.TButton", foreground=[("active", C["primary_dark"])])

        s.configure("Ghost.TButton", background=C["card"], foreground=C["primary"],
                    bordercolor=C["card_border"], lightcolor=C["card"],
                    darkcolor=C["card"], font=(FONT, 9), padding=(12, 6))
        s.map("Ghost.TButton",
              background=[("active", C["primary_light"]), ("disabled", C["card"])],
              foreground=[("active", C["primary_dark"])],
              bordercolor=[("active", C["primary_dark"])])

        s.configure("Treeview", rowheight=30, font=(FONT, 9),
                    background=C["card"], fieldbackground=C["card"],
                    bordercolor=C["card_border"], lightcolor=C["card"],
                    darkcolor=C["card"])
        s.configure("Treeview.Heading", font=(FONT, 9, "bold"),
                    background="#F9FAFB", foreground=C["text"],
                    bordercolor=C["card_border"], relief="flat")
        s.map("Treeview.Heading", background=[("active", "#F3F4F6")])

        s.configure("Green.Horizontal.TProgressbar",
                    background=C["primary"], troughcolor="#E5E7EB",
                    bordercolor="#E5E7EB", lightcolor=C["primary"],
                    darkcolor=C["primary"])

        s.configure("Vertical.TScrollbar", background="#D1D5DB", troughcolor=C["card"],
                    bordercolor=C["card"], arrowcolor=C["text_sec"])

    def _card(self, parent, padding):
        return ttk.Frame(parent, style="Card.TFrame", padding=padding)

    def _build_header(self):
        hd = tk.Frame(self._root, bg=C["header_bg"], height=64)
        hd.pack(fill="x")
        hd.pack_propagate(False)

        inner = tk.Frame(hd, bg=C["header_bg"])
        inner.pack(anchor="w", padx=16, pady=10)

        icon_path = _resource_path("icon.png")
        if os.path.isfile(icon_path):
            try:
                self._header_img = tk.PhotoImage(file=icon_path).subsample(12, 12)
                tk.Label(inner, image=self._header_img, bg=C["header_bg"]).pack(
                    side="left", padx=(0, 10))
            except tk.TclError:
                self._header_img = None

        txt = tk.Frame(inner, bg=C["header_bg"])
        txt.pack(side="left")
        tk.Label(txt, text="PDF转Word", bg=C["header_bg"], fg=C["header_fg"],
                 font=(FONT, 13, "bold")).pack(anchor="w")
        tk.Label(txt,
                 text=f"PdfToWord v{APP_VERSION} · PDF 批量转换 Word · 离线可用"
                      + ("" if _HAS_DND else " · (拖拽组件未安装)"),
                 bg=C["header_bg"], fg=C["header_sub"], font=(FONT, 8)).pack(anchor="w")

    def _build_statusbar(self):
        sb = tk.Frame(self._root, bg=C["statusbar_bg"], height=28)
        sb.pack(side="bottom", fill="x")
        sb.pack_propagate(False)
        sb.columnconfigure(0, weight=1)
        sb.columnconfigure(1, weight=0)
        sb.columnconfigure(2, weight=1)

        self._status_lbl = tk.Label(sb, text="准备就绪", bg=C["statusbar_bg"],
                                    fg=C["statusbar_fg"], font=(FONT, 8), anchor="w")
        self._status_lbl.grid(row=0, column=0, sticky="w", padx=12)

        # 版权居中: 窄窗口时自动截断省略, 不挤压两侧
        self._status_copy = tk.Label(
            sb, text=COPYRIGHT_TEXT, bg=C["statusbar_bg"],
            fg=C["header_sub"], font=(FONT, 8), anchor="center")
        self._status_copy.grid(row=0, column=1, sticky="")

        self._status_right = tk.Label(sb, text="", bg=C["statusbar_bg"],
                                      fg=C["statusbar_fg"], font=(FONT, 8),
                                      anchor="e")
        self._status_right.grid(row=0, column=2, sticky="e", padx=12)

    def _refresh_status_right(self):
        parts = [f"v{APP_VERSION}", f"模式: {self._last_mode}"]
        if self._last_elapsed is not None:
            parts.append(f"耗时: {self._last_elapsed}s")
        if _HAS_DND:
            parts.append("拖拽可用")
        self._status_right.configure(text="  |  ".join(parts))

    # ============================================================ 拖拽

    def _bind_dnd(self):
        if not _HAS_DND:
            return
        try:
            self._root.drop_target_register(DND_FILES)
            self._root.dnd_bind('<<Drop>>', self._on_drop)
            self._tree.drop_target_register(DND_FILES)
            self._tree.dnd_bind('<<Drop>>', self._on_drop)
        except Exception as e:
            logger.warning(f"拖拽注册失败: {e}")

    def _on_drop(self, event):
        paths = _parse_dnd_paths(event.data)
        n = skip = 0
        for p in paths:
            if not os.path.isfile(p):
                continue
            if p.lower().endswith(".pdf"):
                if self._insert(p):
                    n += 1
                else:
                    skip += 1
            else:
                skip += 1
        msg = f"已拖入 {n} 个 PDF"
        if skip:
            msg += f" (跳过 {skip} 个非PDF/重复文件)"
        self._set_status(msg)
        logger.info(msg)

    # ============================================================ 目录/文件

    def _pick_dir(self, which):
        cur = (self._var_in if which == "in" else self._var_out).get()
        d = filedialog.askdirectory(initialdir=cur or None, title="选择目录")
        if not d:
            return
        if which == "in":
            self._var_in.set(d)
            self._load_dir_items(d)
        else:
            self._var_out.set(d)
            self._set_status(f"输出目录: {d}")

    def _load_dir_items(self, d):
        self._clear()
        if not os.path.isdir(d):
            return
        n = 0
        for f in sorted(os.listdir(d)):
            if f.lower().endswith(".pdf") and self._insert(os.path.join(d, f)):
                n += 1
        msg = f"已加载 {n} 个 PDF 文件" if n else "该目录中没有 PDF 文件"
        self._set_status(msg)
        logger.info(msg)

    def _add_files(self):
        for p in filedialog.askopenfilenames(
            title="选择 PDF 文件", filetypes=[("PDF", "*.pdf"), ("所有", "*.*")],
            initialdir=self._var_in.get() or None,
        ):
            self._insert(p)

    def _scan_dir(self):
        d = self._var_in.get()
        if not d or not os.path.isdir(d):
            messagebox.showwarning("提示", "请先选择有效的输入目录")
            return
        n = 0
        for f in sorted(os.listdir(d)):
            if f.lower().endswith(".pdf") and self._insert(os.path.join(d, f)):
                n += 1
        logger.info(f"扫描到 {n} 个 PDF" if n else "输入目录中没有 PDF 文件")

    def _insert(self, path):
        norm = os.path.normpath(path)
        if any(it["path"] == norm for it in self._items):
            return False
        iid = self._tree.insert("", "end", values=("☑", "待转换", os.path.basename(path), "", ""))
        self._items.append({"iid": iid, "path": norm, "checked": True,
                            "status": "待转换", "stats": {}})
        self._refresh_empty_state()
        return True

    def _clear(self):
        self._tree.delete(*self._tree.get_children())
        self._items.clear()
        self._refresh_empty_state()

    def _refresh_empty_state(self):
        if self._tree.get_children():
            self._empty_state.place_forget()
        else:
            self._empty_state.place(relx=0.5, rely=0.5, anchor="center")
        self._refresh_header_check()

    @staticmethod
    def _status_tag(status):
        return {"成功": "ok", "失败": "err", "转换中": "run", "警告": "warn"}.get(status, "pending")

    def _set_all(self, val):
        mark = "☑" if val else "☐"
        for it in self._items:
            it["checked"] = val
            vals = list(self._tree.item(it["iid"], "values"))
            vals[0] = mark
            self._tree.item(it["iid"], values=vals,
                            tags=(self._status_tag(it["status"]),))
        self._refresh_header_check()

    def _refresh_header_check(self):
        if not self._items:
            text = "☐"
        else:
            n = sum(1 for it in self._items if it["checked"])
            text = "☑" if n == len(self._items) else ("☐" if n == 0 else "⊟")
        self._tree.heading("check", text=text)

    def _on_header_check(self):
        if self._thread and self._thread.is_alive():
            return
        if not self._items:
            return
        all_checked = all(it["checked"] for it in self._items)
        self._set_all(not all_checked)

    def _update_row(self, iid, status, stats=None):
        for it in self._items:
            if it["iid"] == iid:
                it["status"] = status
                if stats:
                    it["stats"] = stats
                break
        else:
            return
        mark = "☑" if it["checked"] else "☐"
        pages = stats.get("pages", "") if stats else ""
        output = stats.get("output", "") if stats else ""
        tag = self._status_tag(status)
        display_status = status
        if status == "失败" and stats and stats.get("error"):
            display_status = f"失败: {stats['error']}"
        elif status == "警告" and stats and stats.get("warning"):
            display_status = f"警告: {stats['warning']}"
        self._tree.item(iid, values=(mark, display_status, os.path.basename(it["path"]),
                                     pages, output), tags=(tag,))

    # ============================================================ 点击

    def _on_click(self, event):
        iid = self._tree.identify_row(event.y)
        if not iid:
            return
        col = self._tree.identify_column(event.x)
        if col == "#1":
            for it in self._items:
                if it["iid"] == iid:
                    it["checked"] = not it["checked"]
                    vals = list(self._tree.item(iid, "values"))
                    vals[0] = "☑" if it["checked"] else "☐"
                    self._tree.item(iid, values=vals,
                                    tags=(self._status_tag(it["status"]),))
                    self._refresh_header_check()
                    break
        elif col == "#2":
            # 点击状态列: 失败/警告弹出完整详情
            it = self._by_iid(iid)
            if not it:
                return
            stats = it.get("stats") or {}
            detail = stats.get("error") or stats.get("warning")
            if detail:
                title = "转换失败详情" if it["status"] == "失败" else "转换警告详情"
                messagebox.showinfo(
                    title,
                    f"文件: {os.path.basename(it['path'])}\n\n{detail}")

    def _on_dbl(self, event):
        iid = self._tree.identify_row(event.y)
        if not iid:
            return
        it = self._by_iid(iid)
        if it:
            self._try_open(it)

    def _try_open(self, it):
        if it["status"] in ("成功", "警告") and it["stats"].get("output"):
            path = it["stats"].get("abs_output") or os.path.join(
                self._var_out.get(), it["stats"]["output"])
            if os.path.isfile(path):
                os.startfile(path)
                return
            messagebox.showinfo("提示", f"找不到输出文件:\n{path}")
        elif it["status"] == "失败":
            stats = it.get("stats") or {}
            detail = stats.get("error") or "未知错误"
            messagebox.showinfo("转换失败详情",
                                f"文件: {os.path.basename(it['path'])}\n\n{detail}")
        else:
            self._set_status(f"{os.path.basename(it['path'])} 尚未转换成功, 无法打开")

    # ============================================================ 右键菜单

    def _by_iid(self, iid):
        for it in self._items:
            if it["iid"] == iid:
                return it
        return None

    def _on_right_click(self, event):
        iid = self._tree.identify_row(event.y)
        if not iid:
            return
        it = self._by_iid(iid)
        if not it:
            return

        menu = tk.Menu(self._root, tearoff=0)
        openable = it["status"] in ("成功", "警告") and it["stats"].get("output")
        stats = it.get("stats") or {}
        has_detail = bool(stats.get("error") or stats.get("warning"))

        menu.add_command(label="打开 Word 文档",
                         command=lambda iid=iid: self._try_open(self._by_iid(iid)),
                         state="normal" if openable else "disabled")
        menu.add_command(label="打开所在目录",
                         command=lambda iid=iid: self._open_file_dir(iid))
        if has_detail:
            menu.add_command(label="查看详情",
                             command=lambda iid=iid: self._show_detail(self._by_iid(iid)))
        menu.add_separator()
        menu.add_command(label="取消勾选" if it["checked"] else "勾选",
                         command=lambda iid=iid: self._toggle_check(iid))
        menu.add_separator()
        menu.add_command(label="从列表移除",
                         command=lambda iid=iid: self._remove_item(iid))
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _show_detail(self, it):
        if not it:
            return
        stats = it.get("stats") or {}
        detail = stats.get("error") or stats.get("warning") or "无详情"
        title = "转换失败详情" if it["status"] == "失败" else "详情"
        messagebox.showinfo(title,
                            f"文件: {os.path.basename(it['path'])}\n\n{detail}")

    def _toggle_check(self, iid):
        it = self._by_iid(iid)
        if not it:
            return
        it["checked"] = not it["checked"]
        vals = list(self._tree.item(iid, "values"))
        vals[0] = "☑" if it["checked"] else "☐"
        self._tree.item(iid, values=vals, tags=(self._status_tag(it["status"]),))
        self._refresh_header_check()

    def _remove_item(self, iid):
        it = self._by_iid(iid)
        if not it:
            return
        self._tree.delete(iid)
        self._items.remove(it)
        self._set_status(f"已移除 {os.path.basename(it['path'])}")
        self._refresh_empty_state()

    def _on_delete_key(self, event):
        sel = self._tree.selection()
        if not sel:
            return
        if self._thread and self._thread.is_alive():
            self._set_status("转换进行中, 暂不能移除文件")
            return
        for iid in sel:
            self._remove_item(iid)

    def _open_file_dir(self, iid):
        it = self._by_iid(iid)
        if not it:
            return
        d = os.path.dirname(it["path"])
        if os.path.isdir(d):
            os.startfile(d)
        else:
            self._set_status("目录不存在")

    # ============================================================ 转换

    def _start(self):
        checked = [it for it in self._items if it["checked"]]
        if not checked:
            messagebox.showinfo("提示", "请至少勾选一个 PDF 文件")
            return
        out_dir = self._var_out.get().strip()
        if not out_dir:
            messagebox.showwarning("提示", "请输入有效的输出目录")
            return
        out_dir = os.path.abspath(os.path.expanduser(out_dir))
        self._var_out.set(out_dir)
        for it in checked:
            if not os.path.isfile(it["path"]):
                messagebox.showwarning("提示", f"文件不存在: {os.path.basename(it['path'])}")
                return

        self._save_config()
        self._stop.clear()
        self._set_buttons_running(True)
        self._bar["value"] = 0
        self._last_mode = "—"
        self._last_elapsed = None
        self._refresh_status_right()
        self._var_st.set(f"转换中… 0/{len(checked)}")
        self._set_status(f"开始转换 {len(checked)} 个文件")
        for it in checked:
            self._update_row(it["iid"], "转换中")

        self._thread = threading.Thread(target=self._worker,
                                         args=(out_dir, checked), daemon=True)
        self._thread.start()

    def _worker(self, out_dir, items):
        total = len(items)
        conv = PDFToWordConverter(output_dir=out_dir, logger=logger)

        for i, it in enumerate(items, 1):
            if self._stop.is_set():
                self._post_ui(self._update_row, it["iid"], "停止", None)
                self._post_ui(self._tick, i, total)
                continue

            def on_page(cur, tot, _i=i):
                self._post_ui(self._tick_page, _i, total, cur, tot)

            try:
                results = conv.convert_files(
                    [it["path"]], on_page_progress=on_page)
                if results and results[0]["status"] == "ok":
                    r = results[0]
                    stats = {
                        "pages": r.get("pages", ""),
                        "output": r.get("output", ""),
                        "abs_output": os.path.join(out_dir, r.get("output", "")),
                        "warning": r.get("warning"),
                        "elapsed": r.get("elapsed"),
                        "method": r.get("method"),
                    }
                    status = "警告" if r.get("warning") else "成功"
                    self._post_ui(self._update_row, it["iid"], status, stats)
                    self._post_ui(self._set_mode, r.get("method") or "—",
                                  r.get("elapsed"))
                    logger.info(f"[{i}/{total}] 成功: {os.path.basename(it['path'])} "
                                f"(mode={r.get('method')})")
                elif results:
                    err = results[0].get("error", "未知错误")
                    self._post_ui(self._update_row, it["iid"], "失败", {"error": err})
                    logger.warning(f"[{i}/{total}] 失败: {os.path.basename(it['path'])} ({err})")
                else:
                    self._post_ui(self._update_row, it["iid"], "失败", {"error": "无输出"})
            except Exception as e:
                self._post_ui(self._update_row, it["iid"], "失败", {"error": str(e)})
                logger.error(f"[{i}/{total}] 异常: {os.path.basename(it['path'])}: {e}")

            self._post_ui(self._tick, i, total)

        self._post_ui(self._done)

    def _set_mode(self, method, elapsed):
        mode_map = {"ocr": "OCR", "layout": "版式", "hybrid": "混合"}
        self._last_mode = mode_map.get(method, "—")
        if elapsed is not None:
            self._last_elapsed = elapsed
        self._refresh_status_right()

    def _tick(self, cur, total):
        self._bar["value"] = cur / total * 100 if total else 0
        self._var_st.set(f"转换中… {cur}/{total}")

    def _tick_page(self, file_i, file_n, page_c, page_n):
        """OCR 页级进度: 状态栏显示当前文件页码, 进度条按页推进"""
        page_frac = (page_c / page_n) if page_n else 1.0
        overall = ((file_i - 1) + page_frac) / file_n * 100 if file_n else 0
        self._bar["value"] = overall
        self._var_st.set(
            f"转换中… {file_i}/{file_n} · OCR 第 {page_c}/{page_n} 页")

    def _done(self):
        self._set_buttons_running(False)
        ok = [it for it in self._items if it["status"] in ("成功", "警告")]
        fail = [it for it in self._items if it["status"] == "失败"]
        stopped = [it for it in self._items if it["status"] == "停止"]
        warned = [it for it in self._items if it["status"] == "警告"]
        self._btn_retry.configure(state="normal" if (fail or stopped) else "disabled")

        total_pages = sum(it["stats"].get("pages", 0) or 0 for it in ok)
        total_el = sum(it["stats"].get("elapsed", 0) or 0 for it in ok)
        parts = [f"{len(ok)} 成功", f"{len(fail)} 失败", f"{len(stopped)} 停止"]
        if stopped:
            self._var_st.set(f"已停止 ({' / '.join(parts)})")
            self._set_status(f"已停止: {' / '.join(parts)}")
        else:
            self._bar["value"] = 100
            self._var_st.set(f"完成 ({len(ok)} 成功 / {len(fail)} 失败)")
            msg = (f"完成: {len(ok)} 成功 / {len(fail)} 失败, "
                   f"共 {total_pages} 页, 总耗时 {total_el:.1f}s")
            if warned:
                msg += f" (其中 {len(warned)} 个有警告)"
            self._set_status(msg)
        if total_el:
            self._last_elapsed = round(total_el, 1)
            self._refresh_status_right()
        if fail or stopped:
            logger.warning(f"完成: {' / '.join(parts)}, 共 {total_pages} 页")
        else:
            logger.info(f"全部完成: {len(ok)} 文件, 共 {total_pages} 页, "
                        f"总耗时 {total_el:.1f}s")

    def _set_status(self, text):
        if hasattr(self, "_status_lbl"):
            self._status_lbl.configure(text=text)

    def _log_to_text(self, message, level):
        if not hasattr(self, "_log"):
            return
        try:
            self._post_ui(self._append_log, message, level)
        except (tk.TclError, RuntimeError):
            pass

    def _append_log(self, message, level):
        self._log.configure(state="normal")
        self._log.insert("end", message + "\n", level)
        self._log.configure(state="disabled")
        self._log.see("end")

    # ============================================================ 线程安全 UI 投递

    def _post_ui(self, fn, *args):
        q = getattr(self, "_ui_queue", None)
        if q is not None:
            q.put((fn, args))

    def _drain_ui_queue(self):
        while True:
            try:
                fn, args = self._ui_queue.get_nowait()
            except queue.Empty:
                break
            try:
                fn(*args)
            except Exception:
                logger.exception("UI 回调执行失败")
        try:
            self._root.after(50, self._drain_ui_queue)
        except tk.TclError:
            pass

    def _set_buttons_running(self, running):
        state = "disabled" if running else "normal"
        self._btn_start.configure(state=state)
        self._btn_stop.configure(state="normal" if running else "disabled")
        self._btn_retry.configure(state="disabled" if running else "normal")
        self._btn_browse_in.configure(state=state)
        self._btn_browse_out.configure(state=state)
        for btn in self._list_btns:
            btn.configure(state=state)

    def _stop_convert(self):
        self._stop.set()
        self._var_st.set("正在停止…")
        self._set_status("正在停止…")
        logger.info("用户请求停止转换")

    def _retry(self):
        self._set_all(False)
        for it in self._items:
            if it["status"] in ("失败", "停止"):
                it["checked"] = True
                self._update_row(it["iid"], "待转换")
        self._refresh_header_check()
        logger.info("已重置失败/停止文件, 仅重试这些文件")
        self._start()

    # ============================================================ 折叠/打开

    def _toggle_log(self):
        self._log_collapsed = not self._log_collapsed
        if self._log_collapsed:
            self._log.grid_remove()
            self._btn_collapse.configure(text="展开")
        else:
            self._log.grid()
            self._btn_collapse.configure(text="收起")

    def _open_out(self):
        d = self._var_out.get()
        if d and os.path.isdir(d):
            os.startfile(d)
        else:
            messagebox.showinfo("提示", "输出目录不存在")

    def _open_log(self):
        p = os.path.join(_app_dir(), "logs", "converter.log")
        if os.path.isfile(p):
            os.startfile(p)
        else:
            messagebox.showinfo("提示", "日志文件不存在")

    # ============================================================ 配置

    def _load_config(self):
        default_out = os.path.join(_app_dir(), "output")
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            self._var_in.set(cfg.get("input_dir", ""))
            self._var_out.set(cfg.get("output_dir") or default_out)
        except (FileNotFoundError, json.JSONDecodeError):
            self._var_out.set(default_out)

    def _save_config(self):
        try:
            os.makedirs(CONFIG_DIR, exist_ok=True)
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump({"input_dir": self._var_in.get(),
                           "output_dir": self._var_out.get()}, f, ensure_ascii=False)
        except OSError:
            pass

    # ============================================================ 生命周期

    def _bind_shortcuts(self):
        self._root.bind("<Control-o>", lambda e: self._add_files())
        self._root.bind("<Control-Return>", lambda e: self._start())

    def run(self):
        self._root.mainloop()

    def _on_close(self):
        if self._thread and self._thread.is_alive():
            if not messagebox.askyesno("确认", "转换仍在进行中, 确定关闭?"):
                return
            self._stop.set()
        self._save_config()
        self._root.destroy()


# ========== 日志桥接 ==========
class _GUILogHandler(logging.Handler):
    _app = None

    @classmethod
    def set_app(cls, app):
        cls._app = app

    def emit(self, record):
        if self._app:
            try:
                self._app._log_to_text(self.format(record), record.levelname)
            except Exception:
                pass


def main():
    _enable_dpi_awareness()
    log_dir = os.path.join(_app_dir(), "logs")
    os.makedirs(log_dir, exist_ok=True)
    fmt = "%(asctime)s - %(levelname)s - %(message)s"

    fh = RotatingFileHandler(
        os.path.join(log_dir, "converter.log"),
        maxBytes=2 * 1024 * 1024, backupCount=5, encoding="utf-8"
    )
    fh.setFormatter(logging.Formatter(fmt))

    gh = _GUILogHandler()
    gh.setFormatter(logging.Formatter("%(message)s"))

    logger.setLevel(logging.INFO)
    logger.addHandler(fh)
    logger.addHandler(gh)

    p2d = logging.getLogger("pdf2docx")
    p2d.setLevel(logging.WARNING)
    p2d.addHandler(fh)
    p2d.addHandler(gh)

    app = ConverterApp()
    _GUILogHandler.set_app(app)
    app.run()


if __name__ == "__main__":
    main()
