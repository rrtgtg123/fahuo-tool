# -*- coding: utf-8 -*-
"""
发货数据填入工具 (GUI 版)
从文本框粘贴多行文本，解析出 收件人 / 手机 / 地址和品类及数量，
按表头写入 Excel 模板的对应列，保留模板原有格式。
"""
import json
import os
import re
import time
import tkinter as tk
import zipfile

import openpyxl
import PySimpleGUI as sg

import ai_split

# ---------------- 解析规则 ----------------
# 手机号：11 位，或以“-”连接 4 位短号的虚拟号码
PHONE_RE = re.compile(r"1[3-9]\d{9}(?:-\d{4})?")
# 文件名里误带的表格后缀，统一剥掉，后缀跟随模板
STRIP_EXT_RE = re.compile(r"\.(xlsx|xlsm|xltx|xltm|xls)$", re.I)

HEADERS = ("收件人", "手机", "地址和品类及数量")
PREVIEW_LIMIT = 500  # 预览最多显示的行数，避免大数据卡顿

# 数据框右键菜单（菜单项文字会作为事件返回）
RC_PASTE, RC_SELECT_ALL, RC_COPY, RC_CLEAR = "粘贴", "全选", "复制选中", "清空全部"
RC_MENU = ["", [RC_PASTE, RC_SELECT_ALL, RC_COPY, RC_CLEAR]]

# ---------------- 视觉规范 ----------------
FONT_FAMILY = "Microsoft YaHei UI"
MONO = "Consolas"
F_TITLE = (FONT_FAMILY, 15, "bold")
F_BODY = (FONT_FAMILY, 10)
F_SMALL = (FONT_FAMILY, 9)
F_MONO = (MONO, 10)
F_BTN = (FONT_FAMILY, 10, "bold")

# 中性配色，沿用系统浅灰主题，不引入品牌色
HEADER_BG = "#F0F0F0"
HEADER_TEXT = "#333333"
SURFACE = "#FFFFFF"
TEXT = "#333333"
MUTED = "#6B6B6B"
BTN_PRIMARY = ("#FFFFFF", "#3C3C3C")
BTN_PRIMARY_HOVER = ("#FFFFFF", "#2A2A2A")
BTN_SECONDARY = ("#000000", "#E8E8E8")
BTN_SECONDARY_HOVER = ("#000000", "#DBDBDB")
TABLE_HEADER_BG = "#E8E8E8"
TABLE_ALT = "#F7F7F7"
TABLE_SELECTED = ("#000000", "#D6D6D6")
OK = "#2E7D32"
WARN = "#B00020"

sg.theme("Default1")
sg.set_options(font=F_BODY, border_width=1, element_padding=(6, 4))


# ---------------- 解析 ----------------
def _trim(text: str) -> str:
    """移除字符串头尾的空格、制表符、全角空格以及中英文逗号。"""
    return text.strip(" \t\r\n\u3000，,、;；")


def parse_line(text: str) -> dict:
    text = text.strip()
    m = PHONE_RE.search(text)
    if not m:
        # 没找到手机号时，整行当作收件人兜底
        return {"收件人": _trim(text), "手机": "", "地址和品类及数量": ""}
    return {
        "收件人": _trim(text[: m.start()]),
        "手机": m.group(0),
        "地址和品类及数量": _trim(text[m.end():]),
    }


def parse_text(raw: str) -> list:
    lines = raw.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    return [parse_line(ln) for ln in (l.strip() for l in lines) if ln]


# ---------------- Excel 写入 ----------------
def locate_header(ws):
    """找出表头行号，以及三个目标字段所在的列号。"""
    header_row, header_cols = None, {}
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v is None:
                continue
            sv = str(v)
            if "收件人" in sv and header_row is None:
                header_row = r
            if "收件人" in sv:
                header_cols["收件人"] = c
            elif "手机" in sv:
                header_cols["手机"] = c
            elif "地址" in sv and "地址和品类及数量" not in header_cols:
                header_cols["地址和品类及数量"] = c
    return header_row, header_cols


def write_excel(template_path: str, rows: list, out_path: str) -> int:
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    header_row, header_cols = locate_header(ws)
    if header_row is None:
        raise RuntimeError("模板中未找到『收件人』表头")
    missing = [h for h in HEADERS if h not in header_cols]
    if missing:
        raise RuntimeError(f"模板中缺少表头：{'、'.join(missing)}")

    start = header_row + 1  # 模板预置格式的第一行数据位置
    for i, row in enumerate(rows):
        r = start + i
        for key, c in header_cols.items():
            ws.cell(r, c, row[key])  # 只写值，不碰样式

    wb.save(out_path)
    return len(rows)


# ---------------- 工具函数 ----------------
def _cn_popup(msg: str, title: str, ok_text: str, is_error: bool):
    """中文弹窗：按钮居中、固定宽度不换行、无 OK/Error 英文。"""
    btn_color = ("#FFFFFF", "#B00020") if is_error else ("#FFFFFF", "#3C3C3C")
    layout = [
        [sg.Text(msg, font=F_BODY, pad=((16, 16), (14, 8)))],
        [sg.Column([[sg.Button(ok_text, key="-OK-", size=(10, 1), border_width=0,
                               button_color=btn_color, font=F_BTN, pad=(0, 0))]],
                   element_justification="center", expand_x=True, pad=(0, (4, 14)))],
    ]
    win = sg.Window(title, layout, modal=True, keep_on_top=True,
                    finalize=True, disable_minimize=True)
    win["-OK-"].set_focus()
    while True:
        ev, _ = win.read()
        if ev in (sg.WIN_CLOSED, "-OK-", ok_text):
            break
    win.close()


def info_popup(msg: str, title: str = "提示"):
    _cn_popup(msg, title, "知道了", False)


def err_popup(msg: str, title: str = "出错了"):
    _cn_popup(msg, title, "知道了", True)


def install_autohide_scrollbars(el, idle_ms: int = 1500):
    """滚动条默认隐藏，滚动时出现、停手后自动隐藏，且【不改变外层框体的尺寸】。

    实现要点：
    - 滚动条用 pack 会占布局空间，出现/消失会把外层 Frame 顶高顶宽。
      这里把滚动条从 pack 里永久摘除，改用 place 叠加在内容之上（覆盖式滚动条），
      因此显示与否完全不影响几何布局。
    - 绝不隐藏内容控件本身。
    - 只在滚动时触发（<MouseWheel> / <Button-4,5>），悬停不触发。
    - 内容未超出可视范围时，对应滚动条不出现。
    """
    main_w = el.Widget
    bars = [w for w in (getattr(el, "vsb", None), getattr(el, "hsb", None)) if w is not None]
    if not bars:
        return

    geom = {}
    for w in bars:
        vertical = str(w.cget("orient")) == "vertical"
        # 记录厚度：竖向记宽度，横向记高度（尚未映射时退回请求尺寸）
        size = (w.winfo_width() or w.winfo_reqwidth()) if vertical else \
               (w.winfo_height() or w.winfo_reqheight())
        geom[w] = (vertical, max(size, 8))
        w.pack_forget()              # 永久移出 pack，不再参与空间分配

    timer = {"id": None}

    def _needed(w) -> bool:
        """内容没有超出可视范围时，这条滚动条没必要出现。"""
        try:
            lo, hi = w.get()
        except Exception:
            return True
        return not (float(lo) <= 0.0 and float(hi) >= 1.0)

    def hide(_e=None):
        if timer["id"]:
            try:
                main_w.after_cancel(timer["id"])
            except Exception:
                pass
            timer["id"] = None
        for w in geom:
            w.place_forget()

    def schedule(_e=None):
        if timer["id"]:
            main_w.after_cancel(timer["id"])
        timer["id"] = main_w.after(idle_ms, hide)

    def show(_e=None):
        for w, (vertical, size) in geom.items():
            if not _needed(w):
                w.place_forget()
                continue
            if vertical:
                w.place(relx=1.0, x=0, rely=0.0, relheight=1.0,
                        width=size, anchor="ne")
            else:
                w.place(rely=1.0, y=0, relx=0.0, relwidth=1.0,
                        height=size, anchor="sw")
            w.tkraise()
        schedule()

    # 只在真正滚动时出现：滚轮 / 触控板 / 拖动滚动条
    main_w.bind("<MouseWheel>", show, add="+")           # Windows / macOS
    main_w.bind("<Button-4>", show, add="+")             # Linux 上滚
    main_w.bind("<Button-5>", show, add="+")             # Linux 下滚
    for w in bars:
        w.bind("<B1-Motion>", schedule, add="+")         # 拖动滚动条期间持续续期
        w.bind("<ButtonRelease-1>", schedule, add="+")
    hide()
    main_w.after_idle(hide)  # 窗口真正显示后再藏一次，防止首帧未映射导致漏网


def install_log_hover(window, get_history, width=110, height=12, idle_ms=500):
    """鼠标悬停『最新日志』时，在其下方浮出完整历史；离开即隐藏。"""
    target = window["-LASTLOG-"].Widget
    top = tk.Toplevel(window.TKroot)
    top.overrideredirect(True)
    top.withdraw()

    txt = tk.Text(top, width=width, height=height, bg="#FFFFF1", fg="#5A5A5A",
                  font=("Consolas", 9), wrap="none", borderwidth=1, relief="solid")
    sb = tk.Scrollbar(top, command=txt.yview)
    txt.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    txt.pack(side="left", fill="both", expand=True)
    txt.config(state="disabled")

    state = {"hide_id": None}

    def refresh():
        txt.config(state="normal")
        txt.delete("1.0", "end")
        txt.insert("1.0", "\n".join(get_history()))
        txt.see("end")
        txt.config(state="disabled")

    def show(_e=None):
        if state["hide_id"]:
            top.after_cancel(state["hide_id"])
            state["hide_id"] = None
        refresh()
        x = target.winfo_rootx()
        y = target.winfo_rooty() + target.winfo_height() + 2
        top.deiconify()
        top.geometry(f"+{x}+{y}")
        top.lift()

    def schedule_hide(_e=None):
        if state["hide_id"]:
            top.after_cancel(state["hide_id"])
        state["hide_id"] = top.after(idle_ms, hide)

    def hide(_e=None):
        state["hide_id"] = None
        top.withdraw()

    def cancel_hide(_e=None):
        if state["hide_id"]:
            top.after_cancel(state["hide_id"])
            state["hide_id"] = None

    target.bind("<Enter>", show)
    target.bind("<Leave>", schedule_hide)
    top.bind("<Enter>", cancel_hide)
    top.bind("<Leave>", schedule_hide)
    txt.bind("<Motion>", cancel_hide)
    return top


def ai_settings_dialog(parent, cfg: dict) -> dict:
    """AI 设置弹窗：填 OpenAI 兼容的接口地址 / API Key / 模型名。

    返回新的配置 dict；用户取消则原样返回传入的 cfg。
    """
    layout = [
        [sg.Text("接口地址", size=(9, 1)),
         sg.Input(cfg.get("base_url", ""), key="-URL-", expand_x=True, size=(46, 1)),
         sg.Text("OpenAI 兼容，形如 https://api.deepseek.com/v1", font=F_SMALL,
                 text_color=MUTED)],
        [sg.Text("API Key", size=(9, 1)),
         sg.Input(cfg.get("api_key", ""), key="-KEY-", expand_x=True,
                  password_char="*", size=(46, 1)),
         sg.Text("本地模型（Ollama 等）可留空", font=F_SMALL, text_color=MUTED)],
        [sg.Text("模型名", size=(9, 1)),
         sg.Input(cfg.get("model", ai_split.DEFAULT_MODEL), key="-MODEL-",
                  expand_x=True, size=(46, 1)),
         sg.Text("例：gpt-4o-mini / deepseek-chat / qwen-plus", font=F_SMALL,
                 text_color=MUTED)],
        [sg.Text("", font=F_SMALL, text_color=MUTED, expand_x=True)],
        [sg.Column([[
            sg.Button("保存", key="-SAVE-", size=(10, 1), font=F_BTN, border_width=0,
                      button_color=BTN_PRIMARY, mouseover_colors=BTN_PRIMARY_HOVER),
            sg.Button("取消", key="-CANCEL-", size=(10, 1), border_width=0,
                      button_color=BTN_SECONDARY, mouseover_colors=BTN_SECONDARY_HOVER),
        ]], element_justification="center", expand_x=True, pad=(0, (4, 4)))],
    ]
    win = sg.Window("AI 分词设置", layout, modal=True, keep_on_top=True,
                    finalize=True, disable_minimize=True, font=F_BODY)
    win["-URL-"].set_focus()
    result = cfg
    while True:
        ev, vals = win.read()
        if ev in (sg.WIN_CLOSED, "-CANCEL-", "取消"):
            break
        if ev in ("-SAVE-", "保存"):
            result = {"base_url": vals["-URL-"].strip(),
                      "api_key": vals["-KEY-"].strip(),
                      "model": vals["-MODEL-"].strip() or ai_split.DEFAULT_MODEL}
            break
    win.close()
    return result


def _resolve_sep_text(split_by: str) -> str:
    """分隔符转成肉眼可读的展示文字。"""
    m = {" ": "空格", "\t": "Tab", "": "无（按手机号切）"}
    return m.get(split_by, f"“{split_by}”")


def _resolve_line_text(line_sep: str) -> str:
    m = {"\n": "换行", "\n\n": "空行", "\t\t": "双 Tab"}
    return m.get(line_sep, f"“{line_sep}”")


def human_error(e: Exception) -> str:
    """把常见异常翻译成人话 + 给出解决办法。"""
    s = str(e)
    if isinstance(e, PermissionError) or "Errno 13" in s or "Permission denied" in s:
        return "文件正在使用中，无法写入。请先关闭可能打开它的程序（Excel/WPS/预览窗口），再重新生成。"
    if "Errno 28" in s:
        return "磁盘空间不足。请清理输出目录所在磁盘后再试。"
    if isinstance(e, FileNotFoundError) or re.search(r"Errno 2\]", s):
        return "找不到文件或目录。请检查模板文件和输出目录是否存在、路径是否正确。"
    if isinstance(e, zipfile.BadZipFile) or "BadZipFile" in s or "InvalidFileException" in s:
        return "模板文件不是有效的 .xlsx 格式（可能是老版 .xls 或文件已损坏）。请用 Excel 打开后另存为 .xlsx 再试。"
    if "Errno 5" in s or "I/O error" in s:
        return "文件读写失败。请确认文件未被其他程序占用（如网盘正在同步），再重试。"
    return f"{s or e.__class__.__name__}（若反复出现，请把此信息截图反馈）"


APP_NAME = "发货数据填入工具"


def config_path() -> str:
    """用户级配置文件位置（免管理员权限，打包成 exe 也能写）。"""
    base = os.getenv("APPDATA") or os.path.expanduser("~")
    return os.path.join(base, APP_NAME, "config.json")


def load_config() -> dict:
    try:
        with open(config_path(), encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_config(data: dict):
    try:
        p = config_path()
        os.makedirs(os.path.dirname(p), exist_ok=True)
        with open(p, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass  # 配置写不进去不影响主流程


def remember_template(path: str):
    """记住上次使用的模板文件（仅在确实是文件时才记）。"""
    path = (path or "").strip()
    if not path or not os.path.isfile(path):
        return
    cfg = load_config()
    if cfg.get("template") == path:
        return
    cfg["template"] = path
    save_config(cfg)


def load_ai_config() -> dict:
    """读取 AI 接口配置（地址 / Key / 模型），缺项用默认值补齐。"""
    cfg = load_config().get("ai", {})
    if not isinstance(cfg, dict):
        cfg = {}
    return {
        "base_url": cfg.get("base_url", ai_split.DEFAULT_BASE_URL) or "",
        "api_key": cfg.get("api_key", ai_split.DEFAULT_API_KEY) or "",
        "model": cfg.get("model", ai_split.DEFAULT_MODEL) or ai_split.DEFAULT_MODEL,
    }


def save_ai_config(base_url: str, api_key: str, model: str):
    cfg = load_config()
    cfg["ai"] = {"base_url": (base_url or "").strip(),
                 "api_key": (api_key or "").strip(),
                 "model": (model or "").strip() or ai_split.DEFAULT_MODEL}
    save_config(cfg)


def template_ext(path: str) -> str:
    """输出文件的后缀跟随模板，兜底 .xlsx。"""
    ext = os.path.splitext(path or "")[1]
    return ext.lower() if ext.lower() in (".xlsx", ".xlsm", ".xltx", ".xltm", ".xls") else ".xlsx"


def build_out_path(outdir: str, name: str, ext: str) -> str:
    name = STRIP_EXT_RE.sub("", (name or "").strip()) or "结果"
    return os.path.join(outdir, name + ext)


def label(text: str, width=8) -> sg.Text:
    return sg.Text(text, size=(width, 1), text_color=TEXT)


def secondary_btn(text: str, key: str, size=(12, 1)) -> sg.Button:
    return sg.Button(text, key=key, size=size, border_width=0,
                     button_color=BTN_SECONDARY, mouseover_colors=BTN_SECONDARY_HOVER)


# ---------------- 界面 ----------------
def make_layout(template: str = "") -> list:
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")

    header = [sg.Column([[
        sg.Text("发货数据填入工具", font=F_TITLE, text_color=HEADER_TEXT,
                background_color=HEADER_BG, pad=((14, 10), 12)),
        sg.Text("粘贴文本 → 自动解析 → 按模板生成 Excel", font=F_SMALL,
                text_color=MUTED, background_color=HEADER_BG,
                expand_x=True, justification="right", pad=(0, 12)),
    ]], background_color=HEADER_BG, expand_x=True, pad=(0, 0))]

    file_rows = [
        [label("模板文件"),
         sg.Input(template, key="-TEMPLATE-", expand_x=True, enable_events=True,
                  tooltip="自动记住上次选择的模板；留空可手动粘贴路径"),
         sg.FileBrowse("浏览", size=(8, 1), target="-TEMPLATE-", button_color=BTN_SECONDARY)],
        [label("输出目录"),
         sg.Input(desktop, key="-OUTDIR-", expand_x=True),
         sg.FolderBrowse("浏览", size=(8, 1), target="-OUTDIR-", button_color=BTN_SECONDARY)],
        [label("文件名"),
         sg.Input("结果", key="-FILENAME-", expand_x=True, tooltip="无需输入后缀，自动与模板一致"),
         sg.Text(".xlsx", key="-EXT-", size=(8, 1), text_color=MUTED)],
    ]

    data_rows = [
        [sg.Multiline(size=(80, 7), key="-DATA-", font=F_MONO, expand_x=True,
                      enable_events=True, border_width=1, right_click_menu=RC_MENU,
                      tooltip="每行一条：收件人 手机号 地址和品类及数量（支持右键粘贴）")],
        [sg.Text("", key="-COUNT-", font=F_SMALL, text_color=TEXT, expand_x=True),
         secondary_btn("AI 分词", "-AISPLIT-", size=(10, 1)),
         secondary_btn("AI 设置", "-AICFG-", size=(10, 1)),
         secondary_btn("从剪贴板粘贴", "-PASTE-", size=(14, 1)),
         secondary_btn("清空", "-CLEAR-", size=(8, 1))],
    ]

    preview_rows = [[
        sg.Table(values=[], headings=list(HEADERS), key="-TABLE-",
                 col_widths=[14, 18, 72], auto_size_columns=False,
                 num_rows=6, justification="left",
                 vertical_scroll_only=False,  # 允许横向滚动
                 # 注意：Table 不能开 expand_x，列宽会跟随窗口拉伸导致横向滚动失效
                 header_background_color=TABLE_HEADER_BG, header_text_color="#000000",
                 header_font=(FONT_FAMILY, 10, "bold"),
                 alternating_row_color=TABLE_ALT, row_height=24,
                 selected_row_colors=TABLE_SELECTED, background_color=SURFACE),
    ]]

    # 日志：单行显示最新一条，鼠标悬停浮出完整历史（带时间戳）
    log_section = [sg.Frame("", [[
        sg.Text("日志", font=F_SMALL, text_color=MUTED, pad=((10, 4), (4, 4))),
        sg.Text("", key="-LASTLOG-", font=F_SMALL, text_color=TEXT,
                expand_x=True, size=(64, 1), pad=(4, 4),
                tooltip="鼠标悬停可查看全部日志"),
        sg.Text("悬停查看全部", font=F_SMALL, text_color=MUTED, pad=((0, 10), (4, 4))),
    ]], expand_x=True, pad=((12, 12), (0, 10)), relief=sg.RELIEF_FLAT, border_width=1)]

    layout = [
        header,
        [sg.Frame("文件设置", file_rows, expand_x=True, pad=((12, 12), (12, 4)),
                  title_color=TEXT, relief=sg.RELIEF_FLAT, border_width=1)],
        [sg.Frame("待处理数据", data_rows, expand_x=True,
                  pad=((12, 12), (4, 4)), title_color=TEXT, relief=sg.RELIEF_FLAT)],
        [sg.Frame("解析预览", preview_rows, expand_x=True,
                  pad=((12, 12), (4, 4)), title_color=TEXT, relief=sg.RELIEF_FLAT)],
        [sg.Button("生成 Excel", key="-RUN-", size=(14, 1), font=F_BTN, border_width=0,
                   button_color=BTN_PRIMARY, mouseover_colors=BTN_PRIMARY_HOVER),
         secondary_btn("打开输出目录", "-OPEN-", size=(14, 1)),
         sg.Text("", key="-STATUS-", font=F_BTN, expand_x=True, justification="right"),
         sg.ProgressBar(100, size=(18, 16), key="-BAR-", visible=False, relief=sg.RELIEF_FLAT)],
        log_section,
    ]
    return layout


def main():
    # 窗口尺寸跟随屏幕，避免低分辨率下按钮被挤出可视区
    sw, sh = sg.Window.get_screen_size()
    win_w = max(700, min(920, sw - 60))
    win_h = max(520, min(700, sh - 90))

    # 载入上次使用的模板（文件已不存在则留空，避免填一条坏路径）
    last_tpl = load_config().get("template", "") or ""
    if last_tpl and not os.path.isfile(last_tpl):
        last_tpl = ""

    window = sg.Window("发货数据填入工具 v2.3", make_layout(last_tpl),
                       size=(win_w, win_h), resizable=True, finalize=True)
    window.set_min_size((660, 480))
    if last_tpl:
        window["-EXT-"].update(template_ext(last_tpl))
    window.bind("<Control-Return>", "-RUN-")  # Ctrl+Enter 直接生成

    # 预览表格与数据输入框的滚动条：默认隐藏，滚动/悬停时出现
    install_autohide_scrollbars(window["-TABLE-"])
    install_autohide_scrollbars(window["-DATA-"])

    log_history = []

    # AI 分词状态：active 为 True 时预览/生成都用 AI 拆出来的行
    ai_state = {"active": False, "rows": [], "rule": None}

    def log(msg: str):
        log_history.append(f"[{time.strftime('%H:%M:%S')}] {msg}")
        window["-LASTLOG-"].update(msg)

    # 悬停『最新日志』浮出完整历史
    install_log_hover(window, lambda: log_history)

    def set_status(msg: str, color=TEXT):
        window["-STATUS-"].update(msg, text_color=color)

    def current_text() -> str:
        return window["-DATA-"].Widget.get("1.0", "end-1c")

    def refresh_all() -> list:
        rows = refresh_preview(current_text())
        set_status(f"待处理 {len(rows)} 条" if rows else "")
        return rows

    def paste_from_clipboard():
        try:
            clip = window.TKroot.clipboard_get()
        except Exception:
            clip = ""
        if not clip:
            log("[提示] 剪贴板为空")
            return
        wgt = window["-DATA-"].Widget
        try:
            if wgt.tag_ranges("sel"):        # 有选中内容则替换
                wgt.delete("sel.first", "sel.last")
            wgt.insert("insert", clip)
        except Exception:
            window["-DATA-"].update(current_text() + clip)
        refresh_all()                        # update() 不产生事件，必须手动刷新

    def clear_data():
        window["-DATA-"].Widget.delete("1.0", "end")
        refresh_all()

    def select_all():
        wgt = window["-DATA-"].Widget
        wgt.tag_add("sel", "1.0", "end")
        wgt.focus_set()

    def copy_selection():
        wgt = window["-DATA-"].Widget
        try:
            sel = wgt.get("sel.first", "sel.last")
        except Exception:
            sel = ""
        if sel:
            window.TKroot.clipboard_clear()
            window.TKroot.clipboard_append(sel)

    def refresh_preview(raw: str) -> list:
        rows = ai_state["rows"] if ai_state["active"] else parse_text(raw)
        shown = rows[:PREVIEW_LIMIT]
        window["-TABLE-"].update([[r[h] for h in HEADERS] for r in shown])
        no_phone = sum(1 for r in rows if not r["手机"])
        if not rows:
            window["-COUNT-"].update("")
        else:
            tip = f"共 {len(rows)} 条"
            if ai_state["active"]:
                tip += " · AI 分词"
            if no_phone:
                tip += f" · {no_phone} 条未识别手机号"
            if len(rows) > PREVIEW_LIMIT:
                tip += f" · 预览前 {PREVIEW_LIMIT} 条"
            window["-COUNT-"].update(tip)
        return rows

    def run_ai_split():
        """调模型拿拆分规则，按规则重切数据并覆盖预览。"""
        raw = current_text()
        if not raw.strip():
            err_popup("还没有数据。\n\n请先在『待处理数据』框中粘贴内容，再点『AI 分词』。")
            return
        cfg = load_ai_config()
        if not cfg["base_url"]:
            err_popup("还没配置 AI 接口。\n\n点击『AI 设置』，填入 OpenAI 兼容的接口地址"
                      "（例如 https://api.deepseek.com/v1）和 API Key。")
            return

        set_status("AI 分析中…", TEXT)
        window["-AISPLIT-"].update(disabled=True)
        window["-BAR-"].update(0, visible=True)
        window.refresh()
        try:
            rows, rule = ai_split.split_with_ai(
                raw, cfg["base_url"], cfg["api_key"], cfg["model"])
        except ai_split.AIError as e:
            window["-BAR-"].update(visible=False)
            window["-AISPLIT-"].update(disabled=False)
            set_status("AI 分词失败", WARN)
            log(f"[AI失败] {e}")
            err_popup(str(e))
            return
        except Exception as e:  # 兜底，不让异常炸掉窗口
            window["-BAR-"].update(visible=False)
            window["-AISPLIT-"].update(disabled=False)
            set_status("AI 分词失败", WARN)
            log(f"[AI失败] {human_error(e)}")
            err_popup(human_error(e))
            return
        window["-BAR-"].update(100, visible=False)
        window["-AISPLIT-"].update(disabled=False)

        if not rows:
            set_status("AI 未拆出数据", WARN)
            log("[AI提示] 按模型给的规则没有拆出任何数据，请检查原始文本。")
            err_popup("按模型给出的规则没有拆出任何数据。\n\n请确认粘贴的内容是否为收货人数据。")
            return

        ai_state["active"] = True
        ai_state["rows"] = rows
        ai_state["rule"] = rule
        refresh_preview(raw)
        set_status(f"AI 已拆 {len(rows)} 条", OK)

        if rule.get("line_sep_re"):
            line_desc = "正则模式（见日志）"
        else:
            line_desc = _resolve_line_text(ai_split._resolve_line_sep(rule["line_sep"]))
        rule_desc = (f"分行：{line_desc}"
                     f" · 分列：{_resolve_sep_text(ai_split._resolve_sep(rule['split_by']))}"
                     f" · 列序：{'/'.join(rule['columns'])}")
        log(f"[AI] {rule_desc}")
        if rule.get("line_sep_re"):
            log(f"[AI] 分条正则：{rule['line_sep_re']}")
        if rule.get("strip_prefix"):
            log(f"[AI] 剔除前缀：{rule['strip_prefix']}")
        if rule["reason"]:
            log(f"[AI] 依据：{rule['reason']}（把握 {rule['confidence']:.0%}）")
        no_phone = sum(1 for r in rows if not r["手机"])
        if no_phone:
            log(f"[AI提示] 其中 {no_phone} 条未识别到手机号，请人工核对。")

        info_popup(f"已用 AI 规则重新拆分，共 {len(rows)} 条。\n\n"
                   f"{rule_desc}\n\n"
                   f"判断依据：{rule['reason'] or '—'}\n"
                   f"（把握 {rule['confidence']:.0%}）\n\n"
                   f"可直接点『生成 Excel』，或继续修改文本后自动恢复本地解析。", title="AI 分词完成")

    log("就绪。粘贴数据（Ctrl+V 或右键菜单），确认预览无误后生成 Excel。")
    if last_tpl:
        log(f"[记录] 已自动填入上次使用的模板：{os.path.basename(last_tpl)}")

    while True:
        event, values = window.read()
        if event in (sg.WIN_CLOSED, "-EXIT-"):
            break

        # 文本框内容变化 → 实时刷新预览（手动改动即放弃 AI 结果，回到本地解析）
        if event == "-DATA-":
            if ai_state["active"]:
                ai_state["active"] = False
                ai_state["rows"] = []
                ai_state["rule"] = None
                log("[提示] 文本已修改，恢复本地解析规则。可再次点『AI 分词』。")
            rows = refresh_preview(values["-DATA-"])
            set_status(f"待处理 {len(rows)} 条" if rows else "")
            continue

        # 模板变化 → 同步输出后缀
        if event == "-TEMPLATE-":
            window["-EXT-"].update(template_ext(values["-TEMPLATE-"]))
            remember_template(values["-TEMPLATE-"])  # 记住这次选的模板
            continue

        if event == RC_PASTE:
            paste_from_clipboard()
            continue

        if event == RC_SELECT_ALL:
            select_all()
            continue

        if event == RC_COPY:
            copy_selection()
            continue

        if event == RC_CLEAR:
            clear_data()
            continue

        if event == "-PASTE-":
            paste_from_clipboard()
            continue

        if event == "-AISPLIT-":
            run_ai_split()
            continue

        if event == "-AICFG-":
            new_cfg = ai_settings_dialog(window, load_ai_config())
            save_ai_config(new_cfg["base_url"], new_cfg["api_key"], new_cfg["model"])
            log(f"[AI设置] 已保存：{new_cfg['base_url'] or '（未填地址）'} · "
                f"{new_cfg['model']}")
            continue

        if event == "-CLEAR-":
            clear_data()
            continue

        if event == "-OPEN-":
            d = values["-OUTDIR-"]
            if d and os.path.isdir(d):
                os.startfile(d)
            else:
                set_status("输出目录不存在", WARN)
                err_popup("输出目录不存在，请重新选择。")
            continue

        if event == "-RUN-":
            tpl = values["-TEMPLATE-"].strip()
            outdir = values["-OUTDIR-"].strip()
            ext = template_ext(tpl)

            if not tpl or not os.path.isfile(tpl):
                set_status("模板无效", WARN)
                err_popup("模板文件无效。\n\n请点击『模板文件』右侧的浏览按钮，选择一个 .xlsx 模板。")
                continue
            if not outdir or not os.path.isdir(outdir):
                set_status("输出目录无效", WARN)
                err_popup("输出目录不存在。\n\n请点击『输出目录』右侧的浏览按钮，重新选择一个存在的文件夹。")
                continue

            remember_template(tpl)  # 生成成功前再记一次，覆盖手动输入路径的情况

            rows = ai_state["rows"] if ai_state["active"] else parse_text(values["-DATA-"])
            if not rows:
                set_status("没有数据", WARN)
                err_popup("还没有数据。\n\n请先在『待处理数据』框中粘贴内容（每行一条），确认预览无误后再生成。")
                continue

            out_path = build_out_path(outdir, values["-FILENAME-"], ext)
            window["-BAR-"].update(0, visible=True)
            window.refresh()
            try:
                n = write_excel(tpl, rows, out_path)
            except Exception as e:
                window["-BAR-"].update(visible=False)
                msg = human_error(e)
                set_status("生成失败", WARN)
                log(f"[失败] {msg}")
                log(f"[详情] {e}")
                err_popup(msg)
                continue
            window["-BAR-"].update(100, visible=False)

            no_phone = sum(1 for r in rows if not r["手机"])
            set_status(f"已生成 {n} 条", OK)
            log(f"[完成] 写入 {n} 条 → {out_path}")
            if no_phone:
                log(f"[提示] 其中 {no_phone} 条未识别到手机号，请核对。")
            info_popup(f"已生成 {n} 条数据\n\n文件位置：\n{out_path}", title="完成")
            window["-BAR-"].update(visible=False)

    window.close()


if __name__ == "__main__":
    main()
