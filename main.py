# -*- coding: utf-8 -*-
"""
发货数据填入工具 (CustomTkinter 版)
从文本框粘贴多行文本，解析出 收件人 / 手机 / 地址和品类及数量，
按表头写入 Excel 模板的对应列，保留模板原有格式。
"""
import ctypes
import json
import os
import queue
import re
import sys
import threading
import time
import tkinter as tk
import zipfile

import customtkinter as ctk
import openpyxl

import ai_split

# ---------------- 解析规则 ----------------
# 手机号：11 位，或以“-”连接 4 位短号的虚拟号码
PHONE_RE = re.compile(r"1[3-9]\d{9}(?:-\d{4})?")
# 文件名里误带的表格后缀，统一剥掉，后缀跟随模板
STRIP_EXT_RE = re.compile(r"\.(xlsx|xlsm|xltx|xltm|xls)$", re.I)

HEADERS = ("收件人", "手机", "地址和品类及数量")
# 预览表列头（与导出用的 HEADERS 一致，仅列宽/对齐单独控制）
PREVIEW_HEADERS = HEADERS
# v3.7：字号整体调大后，收件人/手机两列同步放宽，避免内容被过多截断
PREVIEW_WIDTHS = (150, 172, 0)
PREVIEW_ALIGNS = ("w", "center", "w")
PREVIEW_LIMIT = 500  # 预览最多显示的行数，避免大数据卡顿
ADDR_MIN_W = 320     # 地址列最小宽度：再窄就把地址挤没了，此时改为横向滚动
LOG_POPUP_W, LOG_POPUP_H = 680, 280   # 日志浮层尺寸
# 「AI 分析中」的点点转圈动画帧：盲文点阵，视觉上就是一个小点在绕圈跑
LOAD_FRAMES = ("⠁", "⠉", "⠙", "⠚", "⠒", "⠂", "⠒", "⠲", "⠴", "⠦", "⠖", "⠐")
LOAD_FRAME_MS = 100                    # 每帧时长（毫秒）

# ---------------- 版本信息 ----------------
APP_NAME = "发货数据填入工具"
# 数据目录名（英文，避免中文路径在个别环境/命令行下的编码麻烦）。
# 与 APP_NAME 解耦：改目录名不影响界面显示。
DATA_DIR_NAME = "fahuo"
APP_VERSION = "v2.10"
APP_AUTHOR = "—"

# 更新日志只保留两个大版本节点（详细的小版本记录不再展示）。
CHANGELOG = [
    ("v2", "重构软件，引入 AI 辅助拆分"),
    ("v1", "使用代码来进行拆分"),
]

# ---------------- 视觉规范（跟随系统主题）----------------
ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

FONT_FAMILY = "Microsoft YaHei UI"
MONO = "Consolas"

# 字号规范（集中管理，方便统一调整）
FS_TITLE = 13      # 区块标题（待处理数据 / 解析预览 / 操作按钮）
FS_LINK = 11       # 右上角链接（AI 设置 / 版本信息 / 引导链接）
FS_BODY = 11       # 正文标签、按钮
FS_SMALL = 10      # 辅助说明、日志
FS_HINT = 10       # 小提示
FS_TABLE = 12      # 预览表格数据（重点放大）
FS_TABLE_HEAD = 12 # 预览表格表头
FS_INPUT = 12      # 待处理数据输入框
FS_DIALOG_T = 14   # 弹窗标题
FS_DIALOG = 11     # 弹窗正文
FS_VERSION = 16    # 版本号大标题

# 语义色（深浅色下都可读，具体在组件里按主题选）
OK = "#2E9E5B"
WARN = "#D64545"
LINK = "#2E86DE"


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


# ---------------- 配置 ----------------
def data_dir() -> str:
    """用户级数据目录：**配置与日志统一放这里**，便于查看、备份、清理。

    - 优先 `%LOCALAPPDATA%\\fahuo`（免管理员权限，打包成 exe 也能写）
    - 取不到时退回 `%APPDATA%\\fahuo`，再退回 `~/.fahuo`
    - **绝不放在 exe 所在目录**：那里可能是只读位置（如 Program Files），
      也会导致换台电脑/挪动 exe 时数据丢失。

    打开方式：应用内「版本信息」→「打开数据目录」。
    """
    base = os.getenv("LOCALAPPDATA") or os.getenv("APPDATA")
    if base:
        d = os.path.join(base, DATA_DIR_NAME)
    else:
        d = os.path.join(os.path.expanduser("~"), "." + DATA_DIR_NAME)
    try:
        os.makedirs(d, exist_ok=True)
    except Exception:
        pass
    return d


def legacy_data_dirs() -> list:
    """历史版本用过的数据目录（目录名曾是中文的 APP_NAME）。"""
    out = []
    for base in (os.getenv("LOCALAPPDATA"), os.getenv("APPDATA")):
        if base:
            out.append(os.path.join(base, APP_NAME))
    out.append(os.path.join(os.path.expanduser("~"), "." + APP_NAME))
    out.append(os.path.join(os.path.expanduser("~"), ".fahuo"))
    return out


_migrated = False


def migrate_legacy_data():
    """把旧目录里的数据一次性搬到新目录（只在目标文件缺失时复制，不覆盖）。

    只搬配置文件与日志，搬完保留旧目录（不删，避免误删用户数据）。
    """
    global _migrated
    if _migrated:
        return
    _migrated = True
    dst = data_dir()
    for old in legacy_data_dirs():
        try:
            if os.path.normcase(os.path.abspath(old)) == \
               os.path.normcase(os.path.abspath(dst)):
                continue
            if not os.path.isdir(old):
                continue
            for name in ("config.json", "run.log", "crash.log"):
                src = os.path.join(old, name)
                target = os.path.join(dst, name)
                if os.path.isfile(src) and not os.path.exists(target):
                    try:
                        import shutil
                        shutil.copy2(src, target)
                    except Exception:
                        pass
        except Exception:
            pass


def config_path() -> str:
    """用户级配置文件位置（免管理员权限，打包成 exe 也能写）。"""
    return os.path.join(data_dir(), "config.json")


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


def remember_output(dir_path=None, filename=None):
    """记住输出目录 / 文件名；传 None 表示该项不动。

    只在值非空且确实变化时才写盘，避免频繁 IO。
    """
    cfg = load_config()
    changed = False
    if dir_path is not None:
        dir_path = (dir_path or "").strip()
        if dir_path and cfg.get("outdir") != dir_path:
            cfg["outdir"] = dir_path
            changed = True
    if filename is not None:
        filename = (filename or "").strip()
        if filename and cfg.get("filename") != filename:
            cfg["filename"] = filename
            changed = True
    if changed:
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


def save_ai_from_dialog(cfg: dict) -> bool:
    """把设置弹窗的返回值写盘（仅当用户点了「保存」）。返回是否已保存。"""
    if not cfg.get("_saved"):
        return False
    save_ai_config(cfg.get("base_url", ""), cfg.get("api_key", ""),
                   cfg.get("model", ""))
    return True


def template_ext(path: str) -> str:
    """输出文件的后缀跟随模板，兜底 .xlsx。"""
    ext = os.path.splitext(path or "")[1]
    return ext.lower() if ext.lower() in (".xlsx", ".xlsm", ".xltx", ".xltm", ".xls") else ".xlsx"


def out_name(name: str, ext: str) -> str:
    """规范化用户填写的文件名：剥掉误带的表格后缀，空则用「结果」。"""
    return (STRIP_EXT_RE.sub("", (name or "").strip()) or "结果") + ext


def unique_path(path: str) -> str:
    """同名文件已存在时，给**新文件**换一个不冲突的名字（绝不覆盖原文件）。

    命名规则与资源管理器一致：
    `结果.xlsx` → `结果 (2).xlsx` → `结果 (3).xlsx` …
    """
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    i = 2
    while os.path.exists(f"{base} ({i}){ext}"):
        i += 1
    return f"{base} ({i}){ext}"


def build_out_path(outdir: str, name: str, ext: str) -> str:
    """输出文件的完整路径；目标已存在同名文件时自动换名，避免覆盖。"""
    return unique_path(os.path.join(outdir, out_name(name, ext)))


def default_outdir() -> str:
    """默认输出目录：系统**真实的**「桌面」。

    不能想当然用 `~/Desktop` —— 开启 OneDrive/第三方同步后桌面会被重定向
    （本机就是 `C:\\Users\\<用户>\\OneDrive\\桌面`，且是中文名），
    `~/Desktop` 往往并不存在，新用户一上来生成就会报「输出目录不存在」。
    """
    # Windows：直接问系统要桌面路径（能正确处理重定向、中文名、漫游）
    try:
        buf = ctypes.create_unicode_buffer(260)
        # CSIDL_DESKTOPDIRECTORY = 0x0010
        if ctypes.windll.shell32.SHGetFolderPathW(None, 0x0010, None, 0, buf) == 0:
            p = buf.value.strip()
            if p and os.path.isdir(p):
                return p
    except Exception:
        pass
    # 兜底：~/Desktop 存在就用它，否则退回用户主目录
    home = os.path.expanduser("~")
    legacy = os.path.join(home, "Desktop")
    return legacy if os.path.isdir(legacy) else home


# ---------------- 应用图标 ----------------
ICON_FILE = "app.ico"


def resource_path(name: str) -> str:
    """取随程序一起分发的资源路径。

    打包（PyInstaller）后资源被解到临时目录 `sys._MEIPASS`，
    开发态则与 main.py 同目录。
    """
    base = getattr(sys, "_MEIPASS", None) or os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, name)


def apply_icon(win):
    """给窗口设置应用图标（标题栏 + 任务栏）。

    失败不影响使用（例如资源缺失或非 Windows 平台），静默忽略。
    """
    try:
        path = resource_path(ICON_FILE)
        if os.path.isfile(path):
            win.iconbitmap(path)
    except Exception:
        pass


def human_error(e: Exception) -> str:
    """把异常转成给用户看的一两句人话。"""
    s = str(e).strip()
    if s:
        return s
    name = e.__class__.__name__
    if "Permission" in name:
        return "没有权限写入该文件，请确认文件未被 Excel 打开、或换一个输出目录。"
    if "Zip" in name or "BadZip" in name:
        return "模板文件已损坏或不是有效的 Excel 文件。"
    return f"{name}（若反复出现，请把此信息截图反馈）"


def _resolve_sep_text(split_by: str) -> str:
    """分隔符转成肉眼可读的展示文字。"""
    m = {" ": "空格", "\t": "Tab", "": "无（按手机号切）"}
    return m.get(split_by, f"“{split_by}”")


def _resolve_line_text(line_sep: str) -> str:
    m = {"\n": "换行", "\n\n": "空行", "\t": "Tab", "": "无"}
    return m.get(line_sep, f"“{line_sep}”")


# ---------------- 主题辅助 ----------------
def is_dark() -> bool:
    return ctk.get_appearance_mode().lower() == "dark"


def c_text() -> str:
    return "#E8E8E8" if is_dark() else "#1A1A1A"


def c_muted() -> str:
    return "#9A9A9A" if is_dark() else "#6B6B6B"


def c_surface() -> str:
    return "#2B2B2B" if is_dark() else "#FFFFFF"


def c_panel() -> str:
    return "#333333" if is_dark() else "#F5F6F8"


def c_border() -> str:
    return "#3F3F3F" if is_dark() else "#E0E2E6"


def c_head_bg() -> str:
    return "#3A3A3A" if is_dark() else "#EDEFF2"


def c_bg() -> str:
    """页面底色：比卡片略深，形成层次。"""
    return "#1E1E1E" if is_dark() else "#F0F2F5"


def c_row_alt() -> str:
    """斑马纹底色。"""
    return "#303030" if is_dark() else "#F6F7F9"


def c_divider() -> str:
    """表格内列间竖线。比容器边框深一档，浅色底上也能看清。"""
    return "#5A5A5A" if is_dark() else "#C2C7CF"


def c_bar() -> str:
    """滚动条颜色：浅色主题下用浅灰，避免突兀的重色块。"""
    return "#4A4A4A" if is_dark() else "#C9CED6"


def c_bar_hover() -> str:
    return "#5A5A5A" if is_dark() else "#AEB4BD"


# ---------------- 通用弹窗 ----------------
class BaseDialog(ctk.CTkToplevel):
    """模态对话框基类：居中、置顶、Esc 关闭。

    注意：绝不能使用 self._w / self._h，CustomTkinter 的 BaseWidget 内部
    用 self._w 保存 tkinter 控件路径名，覆盖它会让所有子控件创建失败。
    """

    # 关掉 CTk 的「Windows 标题栏着色」机制：它在构造/`resizable()` 时会
    # withdraw → update() →（5ms 后）deiconify，而且执行时机在内容构建完之后，
    # 既会让窗口先以默认尺寸闪一下，又可能在定位之后把弹窗藏回去（闪一下就消失）。
    # 标题栏深浅色改由 `_apply_titlebar_color()` 自己一次性设置，行为完全可控。
    _deactivate_windows_window_header_manipulation = True

    def __init__(self, parent, title: str, width=420, height=200):
        super().__init__(parent)
        # 内容构建期间彻底隐藏（alpha=0 + withdraw），定位完成后再一次性显示
        self._ready_to_show = False
        self._grab_pending = False
        try:
            self.attributes("-alpha", 0.0)
        except tk.TclError:
            pass
        self.withdraw()

        self.title(title)
        apply_icon(self)               # 弹窗标题栏也用应用图标
        self.result = None
        self.resizable(False, False)
        self.transient(parent)
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        self.bind("<Escape>", lambda e: self._on_cancel())
        self.configure(fg_color=c_panel())

        self._dlg_w, self._dlg_h = width, height
        # 记下定时器，销毁前取消，避免窗口没了回调还在跑
        self._jobs = []
        self._jobs.append(self.after(40, self._center))
        self.bind("<Destroy>", lambda e: self._cancel_jobs(e), add="+")

    def _apply_titlebar_color(self):
        """自己设置 Windows 标题栏深浅色（替代 CTk 那套会闪/会藏窗口的实现）。

        只调用 DwmSetWindowAttribute，不做 withdraw/update，因此不会影响显示时序。
        """
        if not sys.platform.startswith("win"):
            return
        try:
            dark = ctk.get_appearance_mode().lower() == "dark"
            value = ctypes.c_int(1 if dark else 0)
            hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
            for attr in (20, 19):   # 20 = 新版本；19 = 20H1 之前
                if ctypes.windll.dwmapi.DwmSetWindowAttribute(
                        hwnd, attr, ctypes.byref(value),
                        ctypes.sizeof(value)) == 0:
                    break
        except Exception:
            pass

    def _revert_withdraw_after_windows_set_titlebar_color(self):
        """保险：万一 CTk 的标题栏流程仍然把窗口藏了回去，这里补一次显示。"""
        try:
            super()._revert_withdraw_after_windows_set_titlebar_color()
        except Exception:
            pass
        if getattr(self, "_ready_to_show", False):
            self._show_now()

    def _show_now(self):
        """把弹窗显示出来（幂等，可被多次调用）。"""
        try:
            if not self.winfo_exists():
                return
            self.deiconify()
            self.attributes("-alpha", 1.0)
            self.lift()
        except tk.TclError:
            return
        if not getattr(self, "_grab_pending", False):
            self._grab_pending = True
            jobs = getattr(self, "_jobs", None)
            if jobs is not None:
                jobs.append(self.after(40, self._safe_grab))

    def _cancel_jobs(self, event=None):
        # 只在自身被销毁时处理（子控件销毁也会冒泡 <Destroy>）
        if event is not None and event.widget is not self:
            return
        for j in getattr(self, "_jobs", []):
            try:
                self.after_cancel(j)
            except Exception:
                pass
        self._jobs = []

    def _safe_grab(self):
        try:
            if self.winfo_exists():
                self.grab_set()
        except Exception:
            pass

    def _center(self):
        try:
            self.update_idletasks()
            pw = self.master.winfo_width()
            ph = self.master.winfo_height()
            px = self.master.winfo_rootx()
            py = self.master.winfo_rooty()
            # ⚠️ 单位陷阱（高 DPI 下必踩）：
            #   winfo_rootx/rooty/width/height 返回的是**物理像素**，
            #   而 _dlg_w/_dlg_h 是我们传入的**逻辑值** —— CTkToplevel.geometry()
            #   会把宽高乘上 window scaling(1.75)，但位置 +x+y 原样交给 Tk。
            #   所以必须先把逻辑宽高换算成物理像素，再与 winfo_* 相减，
            #   否则弹窗会明显偏右、偏下（表现为「不在应用里居中」）。
            #   同理 winfo_screenwidth/height 也是逻辑值，要乘回缩放。
            try:
                sc = self._get_window_scaling()
            except Exception:
                sc = 1.0
            dw = int(round(self._dlg_w * sc))
            dh = int(round(self._dlg_h * sc))
            sw = int(round(self.winfo_screenwidth() * sc))
            sh = int(round(self.winfo_screenheight() * sc))
            # 相对主窗口水平居中；主窗口位置异常时退回屏幕居中
            if pw > 100 and ph > 100:
                x = px + (pw - dw) // 2
                y = py + (ph - dh) // 3
                target_cx = px + pw / 2
            else:
                x = (sw - dw) // 2
                y = (sh - dh) // 3
                target_cx = sw / 2
            x = max(0, min(x, sw - dw))
            y = max(0, min(y, sh - dh))
            self.geometry(f"{self._dlg_w}x{self._dlg_h}+{x}+{y}")
        except tk.TclError:
            return   # 窗口已被销毁时忽略
        self._ready_to_show = True
        self._apply_titlebar_color()

        # —— 水平位置二次校准 ——
        # Tk 的 `geometry +x` 定位的是「窗口**外框**」，而 winfo_rootx 读的是
        # 「**客户区**」，两者相差一个左边框宽（本机 175% 缩放下约 12px）——
        # 不校准的话弹窗整体偏右半个边框，看着就是「没在应用里居中」。
        # 做法：先 deiconify 让 Tk 真正落位（此时 alpha 仍为 0，肉眼完全看不见），
        # 实测真实坐标把误差扣掉，最后再亮出来 —— 全程无可见跳动。
        try:
            self.deiconify()
            self.update_idletasks()
            err = (self.winfo_rootx() + self.winfo_width() / 2) - target_cx
            if abs(err) >= 1:
                nx = max(0, min(x - int(round(err)), sw - dw))
                self.geometry(f"{self._dlg_w}x{self._dlg_h}+{nx}+{y}")
                self.update_idletasks()
        except Exception:
            pass

        self._show_now()
        # 兜底看门狗：CTk 万一还有别的路径把窗口藏起来，1.2s 内纠正
        deadline = time.time() + 1.2
        self._jobs.append(self.after(150, lambda: self._watch_visible(deadline)))

    def _watch_visible(self, deadline):
        try:
            if not self.winfo_exists():
                return
            # 只在「被藏起来」时纠正；用户主动最小化（iconic）不动它
            if self.state() == "withdrawn":
                self._show_now()
        except tk.TclError:
            return
        if time.time() < deadline:
            self._jobs.append(
                self.after(150, lambda: self._watch_visible(deadline)))

    def _on_cancel(self):
        self.result = None
        self.destroy()


def info_popup(parent, msg: str, title: str = "提示"):
    _msg_dialog(parent, msg, title, is_error=False)


def err_popup(parent, msg: str, title: str = "出错了"):
    _msg_dialog(parent, msg, title, is_error=True)


def _msg_dialog(parent, msg, title, is_error):
    dlg = BaseDialog(parent, title, width=420, height=230)
    color = WARN if is_error else OK
    ctk.CTkLabel(dlg, text=title, font=(FONT_FAMILY, FS_DIALOG_T, "bold"),
                 text_color=color).pack(pady=(22, 8))
    box = ctk.CTkTextbox(dlg, wrap="word", font=(FONT_FAMILY, FS_DIALOG),
                         fg_color="transparent", height=90)
    box.pack(fill="both", expand=True, padx=24)
    box.insert("1.0", msg)
    box.configure(state="disabled")
    attach_edit_menu(box, readonly=True)   # 右键可复制报错信息
    ctk.CTkButton(dlg, text="知道了", width=110, height=32,
                  command=dlg.destroy).pack(pady=(4, 20))
    dlg.wait_window()


# ---------------- AI 设置 ----------------
def ai_settings_dialog(parent, cfg: dict) -> dict:
    """AI 设置弹窗：填 OpenAI 兼容的接口地址 / API Key / 模型名。

    返回新的配置 dict（含 _saved 标记）；用户取消则原样返回传入的 cfg。
    """
    dlg = BaseDialog(parent, "AI 拆分设置", width=580, height=262)
    dlg.result = dict(cfg)
    dlg.result["_saved"] = False

    body = ctk.CTkFrame(dlg, fg_color="transparent")
    body.pack(fill="x", padx=24, pady=(22, 0))

    def field(row, label, hint, value, show=None):
        ctk.CTkLabel(body, text=label, width=70, anchor="w",
                     font=(FONT_FAMILY, FS_BODY)).grid(
            row=row, column=0, sticky="w", pady=(0, 12))
        e = ctk.CTkEntry(body, width=260, height=32, font=(FONT_FAMILY, FS_BODY))
        e.grid(row=row, column=1, sticky="w", pady=(0, 12))
        e.insert(0, value or "")
        if show:
            e.configure(show=show)
        attach_edit_menu(e)
        ctk.CTkLabel(body, text=hint, text_color=c_muted(),
                     font=(FONT_FAMILY, FS_SMALL), anchor="w").grid(
            row=row, column=2, sticky="w", padx=(12, 0), pady=(0, 12))
        return e

    e_url = field(0, "接口地址", "形如 https://api.deepseek.com/v1",
                  cfg.get("base_url", ""))
    e_key = field(1, "API Key", "本地模型（Ollama 等）可留空",
                  cfg.get("api_key", ""), show="*")
    e_model = field(2, "模型名", "例：gpt-4o-mini / deepseek-chat",
                    cfg.get("model", ai_split.DEFAULT_MODEL))

    def on_save():
        dlg.result = {
            "base_url": e_url.get().strip(),
            "api_key": e_key.get().strip(),
            "model": e_model.get().strip() or ai_split.DEFAULT_MODEL,
            "_saved": True,
        }
        dlg.destroy()

    btns = ctk.CTkFrame(dlg, fg_color="transparent")
    btns.pack(pady=(4, 20))
    ctk.CTkButton(btns, text="保存", width=104, height=34,
                  font=(FONT_FAMILY, FS_TITLE, "bold"), command=on_save).pack(
        side="left", padx=8)
    ctk.CTkButton(btns, text="取消", width=104, height=34,
                  font=(FONT_FAMILY, FS_BODY), fg_color="transparent",
                  border_width=1, text_color=("#1A1A1A", "#E8E8E8"),
                  command=dlg._on_cancel).pack(side="left", padx=8)

    dlg._jobs.append(dlg.after(80, e_url.focus_set))
    dlg.wait_window()
    return dlg.result


# ---------------- 输入控件：右键编辑菜单 ----------------
def _inner_input(widget):
    """取出 CTk 输入控件内部真正的原生 tkinter 控件。

    CTkEntry → tkinter.Entry（`_entry`）；CTkTextbox → tkinter.Text（`_textbox`）。
    编辑类操作（增删文字、取选区）都要落在这个原生控件上。
    """
    for attr in ("_entry", "_textbox"):
        inner = getattr(widget, attr, None)
        if inner is not None:
            return inner
    return widget


def attach_edit_menu(widget, clear_cmd=None, readonly=False):
    """给输入框挂右键菜单：粘贴 / 复制 / 剪切 / 全选（可选「清空全部」）。

    - 同时兼容 CTkEntry（单行）与 CTkTextbox（多行），内部各自映射到
      原生 tkinter.Entry / tkinter.Text 上做编辑。
    - **右键点在哪里，光标就落到哪里** —— 粘贴插在点击处，而不是上次的光标处。
    - `readonly=True` 时只保留「复制 / 全选」，供只读展示框使用。
    - `clear_cmd`：额外加一个「清空全部」。

    注意：CTkEntry.bind() / CTkTextbox.bind() 都会转发到内部原生控件，
    所以这里直接绑 `<Button-3>` 即可，点控件本身就能触发。
    """
    inner = _inner_input(widget)
    is_text = isinstance(inner, tk.Text)
    menu = tk.Menu(widget, tearoff=0)
    pos = {}
    n = 0

    def _add(label, cmd):
        nonlocal n
        pos[label] = n
        menu.add_command(label=label, command=cmd)
        n += 1

    def _sep():
        nonlocal n
        menu.add_separator()
        n += 1

    def _set_state(label, enabled):
        i = pos.get(label)
        if i is not None:
            menu.entryconfigure(i, state="normal" if enabled else "disabled")

    def sel_range():
        """当前选区 (first, last)；没有选中内容时返回 None。"""
        try:
            first, last = inner.index("sel.first"), inner.index("sel.last")
            return (first, last) if first != last else None
        except Exception:
            return None

    def sel_text():
        sel = sel_range()
        if is_text:
            return inner.get(*sel) if sel else ""
        # Entry 是单行字段：有选区取选区，没选区取全文（更符合直觉）
        if sel:
            a, b = (int(float(x)) for x in sel)
            return inner.get()[a:b]
        return inner.get()

    def do_paste():
        try:
            text = inner.clipboard_get()
        except Exception:
            return
        if not text:
            return
        sel = sel_range()
        if sel:
            try:
                inner.delete(*sel)
            except Exception:
                pass
        try:
            inner.insert("insert", text)
        except Exception:
            pass

    def do_copy():
        text = sel_text()
        if not text:
            return
        try:
            inner.clipboard_clear()
            inner.clipboard_append(text)
        except Exception:
            pass

    def do_cut():
        sel = sel_range()
        if not sel:
            return
        do_copy()
        try:
            inner.delete(*sel)
        except Exception:
            pass

    def do_select_all():
        try:
            if is_text:
                inner.tag_add("sel", "1.0", "end-1c")
                inner.mark_set("insert", "end-1c")
            else:
                inner.select_range(0, "end")
                inner.icursor("end")
        except Exception:
            pass

    if readonly:
        _add("复制", do_copy)
    else:
        _add("粘贴", do_paste)
        _add("复制", do_copy)
        _add("剪切", do_cut)
        _sep()
    _add("全选", do_select_all)
    if clear_cmd is not None:
        _sep()
        _add("清空全部", clear_cmd)

    def popup(e):
        sel = sel_range()
        _set_state("复制", bool(sel) or not is_text)
        if not readonly:
            try:
                has_clip = bool(inner.clipboard_get())
            except Exception:
                has_clip = False
            _set_state("粘贴", has_clip)
            _set_state("剪切", bool(sel))
        # 关键：把光标挪到鼠标点击处，粘贴才会插在这儿
        try:
            if is_text:
                inner.mark_set("insert", "@%d,%d" % (e.x, e.y))
            else:
                inner.icursor("@%d" % e.x)
        except Exception:
            pass
        try:
            widget.focus_set()
        except Exception:
            try:
                inner.focus_set()
            except Exception:
                pass
        try:
            menu.tk_popup(e.x_root, e.y_root)
        finally:
            menu.grab_release()

    widget.bind("<Button-3>", popup, add="+")
    return menu


# ---------------- 版本信息 ----------------
def readonly_textbox(box):
    """把 CTkTextbox 变成「只读但可选中/复制」。

    不用 `state="disabled"` —— 那样用户连选中内容都做不到。
    改为放行导航键与 Ctrl+C / Ctrl+A，其余按键一律拦下。
    """
    def on_key(e):
        if e.state & 0x0004 and e.keysym.lower() in ("c", "a", "insert"):
            return None
        if e.keysym in ("Left", "Right", "Up", "Down", "Home", "End",
                        "Prior", "Next", "Shift_L", "Shift_R",
                        "Control_L", "Control_R"):
            return None
        return "break"

    box.bind("<Key>", on_key)


def version_dialog(parent=None):
    """版本信息弹窗：版本号 + 功能简介 + 更新日志。"""
    dlg = BaseDialog(parent, f"关于 {APP_NAME}", width=620, height=330)

    ctk.CTkLabel(dlg, text=f"{APP_NAME}  {APP_VERSION}",
                 font=(FONT_FAMILY, FS_VERSION, "bold")).pack(pady=(22, 2))
    ctk.CTkLabel(dlg, text="粘贴收货人文本 → 自动解析成三列 → 按 Excel 模板生成发货单",
                 text_color=c_muted(), font=(FONT_FAMILY, FS_BODY)).pack(pady=(0, 12))

    ctk.CTkFrame(dlg, height=1, fg_color=c_border()).pack(fill="x", padx=24)
    ctk.CTkLabel(dlg, text="更新日志", font=(FONT_FAMILY, FS_DIALOG_T, "bold"),
                 anchor="w").pack(fill="x", padx=24, pady=(12, 6))

    # 更新日志用「一个文本框 + 制表位 + 悬挂缩进」渲染，只占 1 个控件。
    # 条目很少，固定高度即可，不需要撑满。
    log = ctk.CTkTextbox(dlg, wrap="word", fg_color="transparent", height=72,
                         font=(FONT_FAMILY, FS_BODY))
    log.pack(fill="x", padx=20)
    inner = log._textbox          # 内部就是原生 tk.Text，可直接配 tag/制表位
    inner.configure(tabs=(56,), padx=10, pady=6)
    inner.tag_configure("ver", font=(MONO, FS_BODY, "bold"),
                        foreground=c_muted(),
                        lmargin1=0, lmargin2=56, spacing1=8)
    inner.tag_configure("note", font=(FONT_FAMILY, FS_BODY))
    for ver, note in CHANGELOG:
        inner.insert("end", f"{ver}\t", "ver")
        inner.insert("end", note + "\n", "note")
    readonly_textbox(log)

    # 数据目录：配置与日志都在这里，给用户一个明确的入口
    ctk.CTkFrame(dlg, height=1, fg_color=c_border()).pack(fill="x", padx=24,
                                                          pady=(6, 0))
    foot = ctk.CTkFrame(dlg, fg_color="transparent")
    foot.pack(fill="x", padx=24, pady=(8, 0))
    ctk.CTkLabel(foot, text="数据目录（配置与日志）", anchor="w",
                 font=(FONT_FAMILY, FS_BODY),
                 text_color=c_muted()).pack(side="left")
    ctk.CTkLabel(foot, text=data_dir(), anchor="e",
                 font=(MONO, FS_SMALL), text_color=c_muted()).pack(side="right")

    btns = ctk.CTkFrame(dlg, fg_color="transparent")
    btns.pack(pady=(12, 18))
    ctk.CTkButton(btns, text="打开数据目录", width=136, height=34,
                  font=(FONT_FAMILY, FS_BODY), fg_color="transparent",
                  border_width=1, border_color=c_border(), text_color=c_text(),
                  command=lambda: open_path(data_dir())).pack(side="left", padx=6)
    ctk.CTkButton(btns, text="关闭", width=110, height=34,
                  command=dlg.destroy).pack(side="left", padx=6)
    dlg.wait_window()


# ---------------- 自绘表格 ----------------
class DataTable(ctk.CTkFrame):
    """轻量表格：表头 + 可滚动数据区。

    CustomTkinter 没有表格控件，这里用 CTkScrollableFrame + 网格自绘。
    - 列有固定宽度，内容过长时截断加省略号，避免撑破布局。
    - 表头文字**居中**，列与列之间画**竖分隔线**。
    - 每列可单独指定对齐方式（`aligns`），例如占比列右对齐。
    """

    def __init__(self, master, columns, widths, height=240, row_h=38,
                 aligns=None, **kw):
        super().__init__(master, fg_color=c_surface(), corner_radius=8,
                         border_width=1, border_color=c_border(), **kw)
        self.columns = columns
        self.widths = widths
        self.row_h = row_h
        # 每列对齐：默认首列左对齐、其余居中；可被 aligns 覆盖
        self.aligns = tuple(aligns) if aligns else tuple(
            "w" if i == 0 else "center" for i in range(len(columns)))
        self._rows = []
        self._header = None
        self._row_widgets = []

        # 用原生 Canvas 承载表格：
        # CTkScrollableFrame 的内部 frame 会被超宽内容顶大（请求宽度一路上传），
        # 导致整个控件被撑开、内容画到可视区外。原生 Canvas 用 scrollregion
        # 显式控制可滚区域，才做得出「列宽固定 + 横向拖动」。
        self.canvas = tk.Canvas(self, highlightthickness=0, bd=0,
                                bg=c_surface())
        self.vbar = ctk.CTkScrollbar(self, orientation="vertical",
                                     command=self.canvas.yview,
                                     button_color=c_bar(),
                                     button_hover_color=c_bar_hover())
        self.hbar = ctk.CTkScrollbar(self, orientation="horizontal",
                                     command=self.canvas.xview,
                                     button_color=c_bar(),
                                     button_hover_color=c_bar_hover())
        self.canvas.configure(yscrollcommand=self.vbar.set,
                              xscrollcommand=self.hbar.set)
        self.vbar.grid(row=0, column=1, sticky="ns", pady=2, padx=(0, 2))
        self.hbar.grid(row=1, column=0, sticky="ew", pady=(0, 2))
        self.canvas.grid(row=0, column=0, sticky="nsew", padx=(2, 0),
                         pady=(2, 0))
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # canvas 内部的容器：表头 + 数据行自上而下排列
        self._inner = tk.Frame(self.canvas, bg=c_surface())
        self._inner_id = self.canvas.create_window(
            (0, 0), window=self._inner, anchor="nw")
        self._inner.bind("<Configure>", self._on_inner_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        # 鼠标滚轮：纵向滚动；Shift+滚轮 → 横向
        self.canvas.bind("<MouseWheel>", self._on_wheel)
        self._inner.bind("<MouseWheel>", self._on_wheel)

        self._make_header()
        self._rows_host = tk.Frame(self._inner, bg=c_surface())
        self._rows_host.pack(fill="x", padx=0, pady=(2, 4))
        self.empty_label = ctk.CTkLabel(
            self._rows_host, text="（暂无数据）", text_color=c_muted(),
            font=(FONT_FAMILY, FS_SMALL))
        self.empty_label.pack(pady=24)

    # ---------- 滚动 / 尺寸 ----------
    def _on_inner_configure(self, event=None):
        """内部尺寸变化 → 更新可滚动区域。"""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        """视口变化：内容比视口窄时拉满，宽时保持原宽以触发横向滚动。"""
        total = self._total_w()
        w = max(event.width, total)
        self.canvas.itemconfigure(self._inner_id, width=w)

    def _on_wheel(self, event):
        """滚轮：默认纵向；按住 Shift 横向滚动。"""
        if event.state & 0x0001:      # Shift
            self.canvas.xview_scroll(-1 if event.delta > 0 else 1, "units")
        else:
            self.canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")
        return "break"

    @staticmethod
    def _ellipsis(s: str, width_px: int) -> str:
        """按像素宽度粗略截断并加省略号（中文按 2 字符宽估算）。"""
        s = str(s or "")
        budget = max(4, int(width_px / 8.6))
        out, used = [], 0
        for ch in s:
            need = 2 if ord(ch) > 127 else 1
            if used + need > budget:
                return "".join(out) + "…"
            out.append(ch)
            used += need
        return s

    def _divider(self, parent):
        """列间竖分隔线。

        用原生 tk.Frame 而不是 CTkFrame：CustomTkinter 的 CTkFrame 是
        canvas 绘制的圆角矩形，width=1 时会被裁掉、根本看不见。
        原生 Frame 能画出实打实的 1px 直线。
        """
        return tk.Frame(parent, width=1, bg=c_divider(),
                        highlightthickness=0, bd=0)

    def _total_w(self) -> int:
        """所有列宽合计（含列间 1px 分隔线）。表头/数据行都按这个宽度排。"""
        return sum(self.widths) + max(0, len(self.widths) - 1) * 3

    def _make_cell(self, parent, text, px_w, align, bg, fg, pad, size=None):
        """定宽单元格。原生 Label 的 width 单位是字符，所以用外层 Frame 定宽。"""
        cell = tk.Frame(parent, width=px_w, bg=bg, height=self.row_h)
        cell.pack(side="left")
        cell.pack_propagate(False)
        tk.Label(cell, text=text, anchor=align, bg=bg, fg=fg,
                 font=(FONT_FAMILY, size or FS_TABLE)).pack(
            fill="both", expand=True, padx=pad)
        return cell

    def set_rows(self, rows):
        """rows 为 list[tuple]，按 columns 顺序给值。"""
        host = self._rows_host
        for w in host.winfo_children():
            w.destroy()
        self._rows = rows or []
        self._row_widgets = []
        if not self._rows:
            ctk.CTkLabel(host, text="（暂无数据）", text_color=c_muted(),
                         font=(FONT_FAMILY, FS_SMALL)).pack(pady=24)
            self._refresh_scrollregion()
            return
        total = self._total_w()
        for idx, row in enumerate(self._rows):
            bg = c_row_alt() if idx % 2 else c_surface()
            line = tk.Frame(host, bg=bg, height=self.row_h, width=total)
            line.pack(fill="x", pady=1)
            line.pack_propagate(False)
            self._row_widgets.append(line)
            for i, w in enumerate(self.widths):
                if i:
                    self._divider(line).pack(side="left", fill="y", pady=4)
                val = row[i] if i < len(row) else ""
                align = self.aligns[i] if i < len(self.aligns) else "w"
                pad = self._cell_pad(i)
                self._make_cell(
                    line, self._ellipsis(val, w - pad[0] - pad[1]),
                    w, align, bg, c_text(), pad)
            # 滚轮绑到行上，鼠标停在数据行也能滚
            line.bind("<MouseWheel>", self._on_wheel)
            for c in line.winfo_children():
                c.bind("<MouseWheel>", self._on_wheel)
                for cc in c.winfo_children():
                    cc.bind("<MouseWheel>", self._on_wheel)
        self._refresh_scrollregion()

    def _refresh_scrollregion(self):
        self.canvas.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def relayout(self):
        """列宽变化后重绘（表头 + 数据）。"""
        self._destroy_header()
        # 先让数据容器让位，重画表头后再补回，保证表头仍在最上面
        self._rows_host.pack_forget()
        self._make_header()
        self._rows_host.pack(fill="x", padx=0, pady=(2, 4))
        self.set_rows(self._rows)

    def _destroy_header(self):
        w = getattr(self, "_header", None)
        if w is not None:
            try:
                w.destroy()
            except Exception:
                pass
            self._header = None

    def _cell_pad(self, i: int) -> tuple:
        """第 i 列的左右 padx。表头与数据行必须一致，否则列对不齐。"""
        align = self.aligns[i] if i < len(self.aligns) else "w"
        return (4, 4) if align == "center" else (10, 6)

    def _make_header(self):
        """表头：与数据行同处一个滚动容器，横向滚动时一起移动。"""
        head = tk.Frame(self._inner, bg=c_head_bg(), height=38,
                        width=self._total_w())
        head.pack(fill="x", pady=(0, 2))
        head.pack_propagate(False)
        for i, (col, w) in enumerate(zip(self.columns, self.widths)):
            if i:
                self._divider(head).pack(side="left", fill="y", pady=6)
            pad = self._cell_pad(i)
            self._make_cell(head, col, w, "center",
                            c_head_bg(), c_text(), pad, size=FS_TABLE_HEAD)
        head.bind("<MouseWheel>", self._on_wheel)
        self._header = head
        return head

    def clear(self):
        self.set_rows([])


# ---------------- 主应用 ----------------
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        migrate_legacy_data()          # 旧数据目录 → 新目录（仅在缺失时复制）
        self.title(f"{APP_NAME} {APP_VERSION}")
        apply_icon(self)               # 应用图标：标题栏 + 任务栏

        # —— 主窗口居中（必须按物理像素算） ——
        # ⚠️ 与弹窗同源的坑：CTk 的 geometry() 会把**宽高**乘上 window scaling(1.75)，
        #   但位置 +x+y 原样交给 Tk（即物理像素）；而 winfo_screenwidth/height
        #   返回的是**逻辑值**。直接拿逻辑屏宽减逻辑窗宽算 x，在 175% 缩放的屏上
        #   会明显偏左（实测左边距 383 / 右边距 957）。这里统一换算成物理像素再居中。
        sw_l, sh_l = self.winfo_screenwidth(), self.winfo_screenheight()
        sc = self._window_scale()
        # v3.7：字号整体调大，窗口同步放大一档
        w = max(760, min(880, sw_l - 120))
        h = max(660, min(900, sh_l - 100))
        self._win_wh = (w, h)
        self._screen_px = (int(round(sw_l * sc)), int(round(sh_l * sc)))
        x = (self._screen_px[0] - int(round(w * sc))) // 2
        y = max(0, (self._screen_px[1] - int(round(h * sc))) // 2 - 20)
        self._win_pos = (x, y)
        self.withdraw()            # 先藏好：等校准完成再一次性显示，避免露脸跳动
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.minsize(720, 600)
        self.configure(fg_color=c_bg())

        # 状态
        self.log_history = []
        self.ai_state = {"active": False, "rows": [], "rule": None,
                         "hint_shown": False, "hint_clicked": False,
                         "loading": False}   # AI 分析中：此期间禁止显示引导链接
        self._anim_job = None
        # AI 拆分：模型调用放后台线程，结果用队列交回主线程处理。
        # （在主线程里直接发请求会把 Tk 事件循环卡死，「AI 分析中」动画一帧都动不了）
        self._split_q = queue.Queue()
        self._split_pending = 0        # 还在后台跑的请求数
        self._split_polling = False    # 结果轮询是否已挂起
        self._split_poll_job = None
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._build_ui()
        self._load_initial()
        # 首次显示：等界面构建完、窗口真正映射后，实测一次坐标做校准再亮出来
        self.after(0, self._first_show)

    def _window_scale(self) -> float:
        """CTk 的 window scaling（高 DPI 下 1.75）。取不到就退回 1.0。"""
        try:
            return float(self._get_window_scaling())
        except Exception:
            return 1.0

    def _first_show(self):
        """首次显示主窗口：趁全透明时实测一次真实坐标，扣掉 Tk 的边框误差。

        两个坑叠在一起（与弹窗同源）：
          ① CTk 的 geometry() 把宽高 ×scaling，位置 +x+y 却是物理像素原样透传；
          ② Tk 的 `geometry +x` 定位窗口**外框**，winfo_rootx 读的是**客户区**，
             两者相差一个左边框（本机 175% 缩放下约 12px）。
        做法：alpha=0（肉眼完全看不见）→ deiconify 让 Tk 真正落位 → 实测中心偏差
        → 修正 → 恢复不透明。全程无可见跳动。
        """
        w, h = self._win_wh
        sw, sh = self._screen_px
        x, y = self._win_pos
        try:
            self.attributes("-alpha", 0.0)
            self.deiconify()
            self.update_idletasks()
            err = (self.winfo_rootx() + self.winfo_width() / 2) - sw / 2
            if abs(err) >= 1:
                nx = max(0, min(x - int(round(err)), int(sw - w * self._window_scale())))
                self.geometry(f"{w}x{h}+{nx}+{y}")
                self._win_pos = (nx, y)
                self.update_idletasks()
        except Exception:
            pass
        finally:
            try:
                if not self.winfo_viewable():     # 兜底：万一没显示出来
                    self.deiconify()
                self.attributes("-alpha", 1.0)
                self.lift()
            except Exception:
                pass

    def _on_close(self):
        """关闭前先停掉动画/轮询定时器，避免回调在窗口销毁后触发报错。"""
        self._stop_load_anim()
        self._stop_split_poll()
        # 兜底再存一次：万一用户改完文件名/目录直接关窗（没触发失焦）
        try:
            self.on_output_change()
        except Exception:
            pass
        self.destroy()

    # ---------- 界面搭建 ----------
    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)   # 预览区吃掉多余高度

        self._build_file_section()
        self._build_data_section()
        self._build_preview_section()
        self._build_action_bar()
        self._build_log_bar()

    def _build_file_section(self):
        wrap = ctk.CTkFrame(self, fg_color=c_surface(), corner_radius=10,
                            border_width=1, border_color=c_border())
        wrap.grid(row=0, column=0, sticky="ew", padx=14, pady=(10, 6))
        wrap.grid_columnconfigure(0, weight=1)
        self.file_wrap = wrap

        # 标题行：左侧「文件设置 ▾」，右侧两个链接
        # 箭头直接拼进按钮文字里，永远紧贴「文件设置」右边，不会被列宽拉开
        ROW_H = 30
        head = ctk.CTkFrame(wrap, fg_color="transparent", height=ROW_H + 12)
        head.grid(row=0, column=0, sticky="ew", padx=10, pady=(4, 0))
        head.grid_propagate(False)
        head.grid_columnconfigure(2, weight=1)

        self.file_expanded = True
        self._file_autocollapsed = False   # 首次出现数据后自动折叠，只触发一次
        self.file_title = ctk.CTkButton(
            head, text=self._file_title_text(), anchor="w", height=ROW_H,
            width=112, corner_radius=8,
            fg_color="transparent", hover_color=c_panel(),
            text_color=c_text(), font=(FONT_FAMILY, FS_TITLE, "bold"),
            command=self.toggle_file)
        self.file_title.grid(row=0, column=0, sticky="w", pady=6)

        # 右侧链接；从右往左依次是「版本信息」「AI 设置」
        ctk.CTkButton(head, text="版本信息", width=62, height=ROW_H,
                      fg_color="transparent", hover_color=c_panel(),
                      text_color=LINK, font=(FONT_FAMILY, FS_LINK),
                      command=lambda: version_dialog(self)).grid(
            row=0, column=4, sticky="e", padx=(0, 6), pady=6)
        ctk.CTkButton(head, text="AI 设置", width=52, height=ROW_H,
                      fg_color="transparent", hover_color=c_panel(),
                      text_color=LINK, font=(FONT_FAMILY, FS_LINK),
                      command=self.on_ai_config).grid(
            row=0, column=3, sticky="e", padx=(0, 14), pady=6)

        # 内容区
        self.file_body = ctk.CTkFrame(wrap, fg_color="transparent")
        self.file_body.grid(row=1, column=0, sticky="ew", padx=14, pady=(2, 12))
        self.file_body.grid_columnconfigure(1, weight=1)

        self.var_template = tk.StringVar()
        # 默认输出目录 = 系统真实桌面（OneDrive 重定向后 ~/Desktop 可能不存在）
        self.var_outdir = tk.StringVar(value=default_outdir())
        self.var_filename = tk.StringVar(value="结果")

        self._file_row(0, "模板文件", self.var_template, browse="file")
        self._file_row(1, "输出目录", self.var_outdir, browse="dir")
        self._file_row(2, "文件名", self.var_filename, suffix=".xlsx")

        self._sync_file_title()

    def _file_row(self, row, label, var, browse=None, suffix=None):
        ctk.CTkLabel(self.file_body, text=label, width=64, anchor="w",
                     font=(FONT_FAMILY, FS_BODY), text_color=c_text()).grid(
            row=row, column=0, sticky="w", pady=5, padx=(0, 8))
        ent = ctk.CTkEntry(self.file_body, textvariable=var, height=32,
                           font=(FONT_FAMILY, FS_BODY))
        ent.grid(row=row, column=1, sticky="ew", pady=5)
        attach_edit_menu(ent)   # 右键：粘贴 / 复制 / 剪切 / 全选
        if label == "模板文件":
            ent.bind("<FocusOut>", lambda e: self.on_template_change())
            self.ent_template = ent
        elif label in ("输出目录", "文件名"):
            # 手动改完失焦即记住（浏览按钮走的路径另外单独记）
            ent.bind("<FocusOut>", lambda e: self.on_output_change())
        if browse:
            cb = (lambda: self.pick_file()) if browse == "file" else (
                lambda: self.pick_dir())
            ctk.CTkButton(self.file_body, text="浏览", width=64, height=32,
                          font=(FONT_FAMILY, FS_BODY), command=cb).grid(
                row=row, column=2, padx=(8, 0), pady=5)
        elif suffix:
            self.lbl_ext = ctk.CTkLabel(self.file_body, text=suffix, width=64,
                                        anchor="w", text_color=c_muted(),
                                        font=(FONT_FAMILY, FS_BODY))
            self.lbl_ext.grid(row=row, column=2, padx=(8, 0), pady=5)

    def _build_data_section(self):
        wrap = ctk.CTkFrame(self, fg_color=c_surface(), corner_radius=10,
                            border_width=1, border_color=c_border())
        wrap.grid(row=1, column=0, sticky="ew", padx=14, pady=6)
        wrap.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(wrap, text="待处理数据", font=(FONT_FAMILY, FS_TITLE, "bold"),
                     anchor="w", text_color=c_text()).grid(
            row=0, column=0, sticky="w", padx=14, pady=(10, 4))

        self.txt_data = ctk.CTkTextbox(wrap, height=116, font=(MONO, FS_INPUT),
                                       wrap="none", corner_radius=8)
        self.txt_data.grid(row=1, column=0, sticky="ew", padx=14)
        self.txt_data.bind("<KeyRelease>", lambda e: self.on_data_change())
        attach_edit_menu(self.txt_data, clear_cmd=self.on_clear)

        foot = ctk.CTkFrame(wrap, fg_color="transparent")
        foot.grid(row=2, column=0, sticky="ew", padx=14, pady=(6, 12))
        foot.grid_columnconfigure(0, weight=1)
        self.lbl_count = ctk.CTkLabel(foot, text="", anchor="w",
                                      font=(FONT_FAMILY, FS_BODY), text_color=c_muted())
        self.lbl_count.grid(row=0, column=0, sticky="w")
        ctk.CTkButton(foot, text="从剪贴板粘贴", width=124, height=34,
                      font=(FONT_FAMILY, FS_BODY), command=self.on_paste).grid(
            row=0, column=1, padx=(6, 0))
        ctk.CTkButton(foot, text="清空", width=72, height=34,
                      font=(FONT_FAMILY, FS_BODY), fg_color="transparent",
                      border_width=1, text_color=("#1A1A1A", "#E8E8E8"),
                      command=self.on_clear).grid(row=0, column=2, padx=(6, 0))

    def _build_preview_section(self):
        wrap = ctk.CTkFrame(self, fg_color=c_surface(), corner_radius=10,
                            border_width=1, border_color=c_border())
        wrap.grid(row=2, column=0, sticky="nsew", padx=14, pady=6)
        wrap.grid_columnconfigure(0, weight=1)
        wrap.grid_rowconfigure(1, weight=1)
        self.preview_wrap = wrap

        head = ctk.CTkFrame(wrap, fg_color="transparent", height=32)
        head.grid(row=0, column=0, sticky="ew", padx=14, pady=(8, 2))
        head.grid_propagate(False)
        self.preview_head = head
        ctk.CTkLabel(head, text="解析预览", font=(FONT_FAMILY, FS_TITLE, "bold"),
                     text_color=c_text()).pack(side="left")
        # AI 引导链接 / 加载提示：两者互斥，同一时刻只显示一个
        self.btn_ai_hint = ctk.CTkButton(
            head, text="分的不对？试试 AI 拆分", height=26,
            fg_color="transparent", hover_color=c_panel(), width=150,
            text_color=LINK, font=(FONT_FAMILY, FS_LINK, "underline"),
            command=self.on_ai_split)
        self.lbl_ai_load = ctk.CTkLabel(head, text="", font=(FONT_FAMILY, FS_BODY),
                                        text_color=c_muted())
        # 复制按钮固定靠右
        self.btn_copy = ctk.CTkButton(
            head, text="复制表格", width=80, height=28,
            fg_color="transparent", hover_color=c_panel(),
            border_width=1, border_color=c_border(),
            text_color=c_text(), font=(FONT_FAMILY, FS_BODY),
            command=self.on_copy_table)
        self.btn_copy.pack(side="right")

        # 表格宽度按窗口自适应：最后一列吃掉剩余
        self.table = DataTable(wrap, PREVIEW_HEADERS, PREVIEW_WIDTHS,
                               aligns=PREVIEW_ALIGNS)
        self.table.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 12))
        wrap.bind("<Configure>", self._on_preview_resize)

    def _set_preview_hint(self, mode: str):
        """统一管理预览标题右侧的显示态，避免两个控件叠在一起。

        mode: "hint" 显示 AI 引导链接 | "loading" 显示分析中 | "none" 都不显示
        同时维护 ai_state["loading"]，让 refresh_preview 知道现在正处于分析中。
        """
        self.ai_state["loading"] = (mode == "loading")
        # 先全部撤下，再按需挂上，保证互斥
        self.btn_ai_hint.pack_forget()
        self.lbl_ai_load.pack_forget()
        self._stop_load_anim()
        if mode == "hint":
            self.btn_ai_hint.pack(side="left", padx=(12, 0))
        elif mode == "loading":
            self.lbl_ai_load.pack(side="left", padx=(12, 0))
            self._start_load_anim()

    def _allow_ai_hint(self):
        """让「分的不对？试试 AI 拆分」引导链接重新可用。

        链接只在「数据非空 & 没点过 AI 拆分 & 没在用 AI 结果」时显示；
        `hint_clicked` 一旦为 True 就再也不出现。换了文本、或这次拆分没成时
        必须复位，否则用户找不到入口（重贴一批数据后按钮消失就是这个原因）。
        拆分成功后复位也是安全的 —— 那时 `active=True`，本来就轮不到链接显示。
        """
        self.ai_state["hint_clicked"] = False

    # ---------- 「AI 分析中」动画 ----------
    def _start_load_anim(self):
        """启动旋转小动画，让用户看得出程序还在跑。"""
        self._anim_i = 0
        self._tick_load_anim()

    def _tick_load_anim(self):
        if not self.ai_state.get("loading"):
            return
        # 点点转圈：盲文点阵字符循环，看起来就是一个小点在绕圈跑
        f = LOAD_FRAMES[self._anim_i % len(LOAD_FRAMES)]
        # 圆点数随时间循环 1→2→3，配合前面的转圈点
        dots = "." * (self._anim_i // 3 % 3 + 1)
        self.lbl_ai_load.configure(text=f"AI 分析中{dots} {f}")
        self._anim_i += 1
        self._anim_job = self.after(LOAD_FRAME_MS, self._tick_load_anim)

    def _stop_load_anim(self):
        """停掉动画定时器，避免窗口关掉后回调还在跑。"""
        job = getattr(self, "_anim_job", None)
        if job:
            try:
                self.after_cancel(job)
            except Exception:
                pass
            self._anim_job = None

    def _on_preview_resize(self, event):
        """末列在有富余宽度时吃满剩余空间；不够时保持最小宽度并允许横向滚动。

        以前是「无条件把末列压到剩余宽度」，导致内容永远不会超出可视宽度，
        横向滚动条自然永远滚不动 —— 这就是横向滑动"丢失"的原因。
        """
        if not hasattr(self, "table"):
            return
        # event.width 是外层预览区（wrap）的宽度，要扣掉从 wrap 到表格内部的
        # 全部横向损耗（实测值，含 wrap 左右 padx、纵向滚动条、表格内缩等），
        # 再扣掉列间竖线（n 列共 n-1 条，每条 3px），剩下的才给末列。
        fixed = sum(PREVIEW_WIDTHS[:-1])
        seps = (len(PREVIEW_WIDTHS) - 1) * 3
        avail = event.width - 76 - fixed - seps - 8
        # 末列最小宽度：太小会把地址挤成一两个字，这里给个下限
        avail = max(ADDR_MIN_W, avail)
        if abs(self.table.widths[-1] - avail) > 8:
            self.table.widths = tuple(PREVIEW_WIDTHS[:-1]) + (avail,)
            self.table.relayout()

    def _build_action_bar(self):
        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.grid(row=3, column=0, sticky="ew", padx=14, pady=(2, 4))
        bar.grid_columnconfigure(2, weight=1)

        ctk.CTkButton(bar, text="生成 Excel", width=132, height=42,
                      font=(FONT_FAMILY, FS_TITLE, "bold"),
                      command=self.on_run).grid(row=0, column=0, sticky="w")
        # 与「生成 Excel」并排
        ctk.CTkButton(bar, text="打开输出目录", width=124, height=42,
                      font=(FONT_FAMILY, FS_BODY), fg_color="transparent",
                      border_width=1, text_color=("#1A1A1A", "#E8E8E8"),
                      command=self.on_open_dir).grid(row=0, column=1, sticky="w",
                                                     padx=(8, 0))

        # 状态文字靠右
        self.lbl_status = ctk.CTkLabel(bar, text="", anchor="e",
                                       font=(FONT_FAMILY, FS_BODY), text_color=c_muted())
        self.lbl_status.grid(row=0, column=2, sticky="e", padx=10)

    def _build_log_bar(self):
        bar = ctk.CTkFrame(self, fg_color=c_panel(), corner_radius=0, height=38)
        bar.grid(row=4, column=0, sticky="ew")
        bar.grid_propagate(False)
        ctk.CTkLabel(bar, text="日志", font=(FONT_FAMILY, FS_SMALL),
                     text_color=c_muted()).pack(side="left", padx=(14, 6))
        self.lbl_log = ctk.CTkLabel(bar, text="", anchor="w", cursor="hand2",
                                    font=(FONT_FAMILY, FS_SMALL), text_color=c_text())
        self.lbl_log.pack(side="left", fill="x", expand=True)

        self.lbl_log_hint = ctk.CTkLabel(bar, text="点击查看全部", anchor="e",
                                         cursor="hand2",
                                         font=(FONT_FAMILY, FS_SMALL),
                                         text_color=c_muted())
        self.lbl_log_hint.pack(side="right", padx=(8, 14))

        # v3.9：改为点击触发（原来是悬停触发，鼠标扫过就弹，容易误触）
        # 两个控件共用同一套状态机，否则会各自弹出独立浮层。
        self._bind_log_click(self.lbl_log, self.lbl_log_hint)

    # ---------- 小工具 ----------
    # 右键编辑菜单统一走模块级 attach_edit_menu()（同时支持 CTkEntry 与
    # CTkTextbox）；原来只为文本框写的 _install_paste_menu / _select_all /
    # _copy_sel 已被取代删除。

    def _bind_log_click(self, *widgets):
        """点击日志栏弹出完整日志历史。

        v3.9.1：**彻底移除「鼠标移出即隐藏」那一套**（原来靠 `after` 轮询鼠标位置，
        鼠标一离开日志栏/弹窗范围就自动收回）。现在开与关只由点击决定：

          ① 点日志栏（或右侧「点击查看全部」）→ 开 / 关 切换
          ② 点弹窗右上角的 × → 关闭

        除此之外，移动鼠标、点界面别的地方都不会让它消失 —— 这样可以在弹窗里
        从容地选中、复制日志，不会因为手一抖移开就被收走。

        `*widgets`：触发区可以是多个控件（日志文字 + 提示文字），
        它们共用同一个状态机，不会各自弹出独立浮层。
        """
        state = {"win": None, "btn": None}

        def has_win() -> bool:
            w = state["win"]
            try:
                return w is not None and bool(w.winfo_exists())
            except Exception:
                return False

        def hide(_e=None):
            w = state["win"]
            state["win"] = None
            state["btn"] = None
            if w is not None:
                try:
                    w.destroy()
                except Exception:
                    pass

        def _readonly_key(e):
            """弹窗内容只读但可选中复制（方便把日志发给别人看）。"""
            if e.state & 0x0004 and e.keysym.lower() in ("c", "a", "insert"):
                return None                    # 放行 Ctrl+C / Ctrl+A
            if e.keysym in ("Left", "Right", "Up", "Down", "Home", "End",
                            "Prior", "Next", "Shift_L", "Shift_R",
                            "Control_L", "Control_R"):
                return None                    # 放行选择与光标移动
            return "break"

        def build():
            if has_win() or not self.log_history:
                return
            w = ctk.CTkToplevel(self)
            state["win"] = w
            w.overrideredirect(True)
            w.attributes("-topmost", True)
            w.configure(fg_color=c_border())    # 1px 外框，浅色底上也看得清

            # 用一个容器承载「顶部小条（放关闭按钮）+ 日志文本框」
            wrap = ctk.CTkFrame(w, fg_color=c_surface(), corner_radius=0)
            wrap.pack(fill="both", expand=True, padx=1, pady=1)

            # 顶部小条：右侧放 × 关闭按钮。
            # 不用「浮在文本上」的做法 —— 那会盖住第一行日志的末尾。
            bar_top = ctk.CTkFrame(wrap, fg_color=c_panel(), corner_radius=0,
                                   height=30)
            bar_top.pack(fill="x")
            bar_top.pack_propagate(False)

            btn_close = ctk.CTkButton(bar_top, text="×", width=30, height=22,
                                      corner_radius=4, fg_color="transparent",
                                      hover_color=c_border(), border_width=0,
                                      text_color=c_text(),
                                      font=(FONT_FAMILY, FS_TITLE),
                                      command=hide)
            btn_close.pack(side="right", padx=6, pady=4)
            state["btn"] = btn_close

            box = ctk.CTkTextbox(wrap,
                                 width=LOG_POPUP_W,
                                 height=max(80, LOG_POPUP_H - 32),
                                 font=(MONO, FS_SMALL), corner_radius=0)
            box.pack(fill="both", expand=True)
            box.insert("1.0", "\n".join(self.log_history[-200:]))
            box.configure(state="normal")       # 只读但可选中复制
            box.bind("<Key>", _readonly_key)
            attach_edit_menu(box, readonly=True)   # 右键：复制 / 全选

            # 关键：overrideredirect 窗口必须先映射再定尺寸，
            # 否则 geometry 会被忽略、窗口按内容撑成巨大矩形。
            w.update_idletasks()
            w.deiconify()
            w.update_idletasks()

            self._place_log_popup(w, widgets[0])

        def on_trigger_click(_e=None):
            """点日志栏 / 提示文字：开 / 关（切换）。

            弹窗开着时再点一次 = 关闭，也就是「点最下面的日志行关闭」。"""
            if has_win():
                hide()
            else:
                build()

        # 触发与关闭都只绑在触发控件上；不再有任何全局监听
        for _w in widgets:
            _w.bind("<Button-1>", on_trigger_click, add="+")
            # 控件被销毁时一并清理，避免残留浮层
            _w.bind("<Destroy>", lambda e: hide(), add="+")

    def _place_log_popup(self, w, anchor):
        """把日志浮层摆到日志栏正上方。

        v3.9.1 起不再需要考虑「避开鼠标」—— 浮层改为点击触发、且只由 ×
        或再点日志行关闭，已经没有「盖住光标 → <Leave> → 隐藏」的闪烁路径，
        原来那套碰撞规避（甚至把窗口挪出屏幕外）反而会让浮层莫名看不见。

        注意 DPI：geometry 收的是逻辑像素，winfo_* 返回的是物理像素，
        本机缩放约 1.75x。所以尺寸统一用实测的 winfo_width/height 定。
        """
        sw, _sh = self.winfo_screenwidth(), self.winfo_screenheight()
        try:
            lx = anchor.winfo_rootx()
            ly = anchor.winfo_rooty()
        except Exception:
            lx, ly = 0, 0

        # 用请求尺寸先定下大小，再读回真实渲染尺寸（已含 DPI 缩放）
        w.geometry(f"{LOG_POPUP_W}x{LOG_POPUP_H}+0+0")
        w.update_idletasks()
        pw, ph = w.winfo_width(), w.winfo_height()
        if pw <= 1 or ph <= 1:
            pw, ph = LOG_POPUP_W, LOG_POPUP_H

        # 水平方向与日志栏左对齐，但不许超出屏幕右缘
        x = max(0, min(lx, sw - pw - 8))
        # 竖直方向：日志栏正上方；上方空间不够就贴屏幕顶
        y = max(0, ly - ph - 8)
        w.geometry(f"{LOG_POPUP_W}x{LOG_POPUP_H}+{x}+{y}")

    # ---------- 业务逻辑 ----------
    def log(self, msg: str):
        self.log_history.append(f"[{time.strftime('%H:%M:%S')}] {msg}")
        self.lbl_log.configure(text=msg)

    def set_status(self, msg: str, color=None):
        self.lbl_status.configure(text=msg,
                                  text_color=color or c_muted())

    def _load_initial(self):
        cfg = load_config()
        last = cfg.get("template", "") or ""
        if last and os.path.isfile(last):
            self.var_template.set(last)
            self.lbl_ext.configure(text=template_ext(last))
        # 输出目录 / 文件名：恢复上次记忆的值（没记过就用默认）
        saved_dir = (cfg.get("outdir") or "").strip()
        if saved_dir:
            self.var_outdir.set(saved_dir)
        saved_name = (cfg.get("filename") or "").strip()
        if saved_name:
            self.var_filename.set(saved_name)
        # 文件设置：记住上次的展开/折叠状态
        #   ① 已有 file_expanded → 直接沿用
        #   ② 旧配置只有 file_set_seen → 兼容：打开过就折叠
        #   ③ 全新用户（两键都没有）→ 默认展开，方便先设模板/输出目录
        if "file_expanded" in cfg:
            self.file_expanded = bool(cfg.get("file_expanded"))
        else:
            self.file_expanded = not cfg.get("file_set_seen")
        self._apply_file_state()
        if not cfg.get("file_set_seen"):
            c = load_config()
            c["file_set_seen"] = True
            save_config(c)

        self.log("就绪。粘贴数据（Ctrl+V 或右键菜单），确认预览无误后生成 Excel。")
        if last and os.path.isfile(last):
            self.log(f"[记录] 已自动填入上次使用的模板：{os.path.basename(last)}")

    # ---------- 折叠 ----------
    def _file_title_text(self) -> str:
        """标题文字（含箭头），箭头紧跟「文件设置」右侧。"""
        return "文件设置 ▾" if self.file_expanded else "文件设置 ▸"

    def _sync_file_title(self):
        self.file_title.configure(text=self._file_title_text())

    def _apply_file_state(self):
        if self.file_expanded:
            self.file_body.grid()
        else:
            self.file_body.grid_remove()
        self._sync_file_title()

    def toggle_file(self):
        self.file_expanded = not self.file_expanded
        self._apply_file_state()
        self._save_file_state()

    def _save_file_state(self):
        """把展开/折叠状态写进配置，下次打开保持一致。"""
        try:
            c = load_config()
            if c.get("file_expanded") != bool(self.file_expanded):
                c["file_expanded"] = bool(self.file_expanded)
                save_config(c)
        except Exception:
            pass

    def _collapse_file_section(self):
        """收起「文件设置」并记住状态（内部调用，不改变其他内容）。"""
        self.file_expanded = False
        self._apply_file_state()
        self._save_file_state()

    # ---------- 文件选择 ----------
    def on_template_change(self):
        p = self.var_template.get().strip()
        self.lbl_ext.configure(text=template_ext(p))
        remember_template(p)

    def on_output_change(self):
        """输出目录 / 文件名变化后记住（失焦或选目录后触发）。"""
        remember_output(self.var_outdir.get(), self.var_filename.get())

    def pick_file(self):
        from tkinter import filedialog
        p = filedialog.askopenfilename(
            title="选择 Excel 模板",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm *.xltx *.xltm *.xls"),
                       ("全部文件", "*.*")])
        if p:
            self.var_template.set(p)
            self.on_template_change()

    def pick_dir(self):
        from tkinter import filedialog
        cur = self.var_outdir.get().strip()
        # 从当前输出目录开始浏览（目录不存在时交给系统默认位置）
        kw = {"initialdir": cur} if os.path.isdir(cur) else {}
        p = filedialog.askdirectory(title="选择输出目录", **kw)
        if p:
            self.var_outdir.set(p)
            self.on_output_change()   # 选完立刻记住

    # ---------- 数据区事件 ----------
    def current_text(self) -> str:
        return self.txt_data.get("1.0", "end-1c")

    def on_data_change(self):
        # 手动改动即放弃 AI 结果，回到本地解析
        if self.ai_state["active"]:
            self.ai_state["active"] = False
            self.ai_state["rows"] = []
            self.ai_state["rule"] = None
            self.log("[提示] 文本已修改，恢复本地解析规则。可再次点『AI 拆分』。")
        # 换了文本 = 换了数据，上一批的「已点过 AI 拆分」不该继续生效，
        # 复位后引导链接会重新出现（粘贴 / 清空 / 手改都走这里）
        self._allow_ai_hint()
        rows = self.refresh_preview(self.current_text())
        self.set_status(f"待处理 {len(rows)} 条" if rows else "")
        # 用户开始粘贴/录入待处理数据后，自动收起「文件设置」，
        # 把纵向空间让给解析预览；只自动收一次，之后尊重用户手动展开
        if rows and not self._file_autocollapsed:
            self._file_autocollapsed = True
            if self.file_expanded:
                self._collapse_file_section()

    def on_paste(self):
        _trace("on_paste 进入")
        try:
            clip = self.clipboard_get()
        except Exception as e:
            _trace(f"on_paste 读剪贴板失败：{e}")
            clip = ""
        if not clip:
            self.log("[提示] 剪贴板为空")
            _trace("on_paste 剪贴板为空，返回")
            return
        _trace(f"on_paste 读到 {len(clip)} 字符")
        try:
            if self.txt_data.tag_ranges("sel"):
                self.txt_data.delete("sel.first", "sel.last")
        except Exception:
            pass
        self.txt_data.insert("insert", clip)
        _trace("on_paste 已插入文本，准备解析")
        self.on_data_change()
        _trace("on_paste 全部完成")

    def on_clear(self):
        self.txt_data.delete("1.0", "end")
        self.on_data_change()

    # ---------- 预览 ----------
    def on_copy_table(self):
        """把当前预览复制成表格（制表符分隔），可直接粘进 Excel。"""
        # table._rows 已是三列的显示元组，直接用
        rows = self.table._rows
        if not rows:
            self.log("[提示] 预览为空，没有可复制的内容")
            self.set_status("预览为空", WARN)
            return
        lines = ["\t".join(str(c) for c in HEADERS)]
        for r in rows:
            # 制表符/换行会破坏表格结构，先替换成空格
            cells = [str(c or "").replace("\t", " ").replace("\n", " ")
                     for c in r]
            lines.append("\t".join(cells))
        text = "\n".join(lines)
        try:
            self.clipboard_clear()
            self.clipboard_append(text)
            self.update()   # 确保剪贴板内容对其它程序可见
        except Exception as e:
            self.log(f"[错误] 复制失败：{e}")
            self.set_status("复制失败", WARN)
            return
        self.log(f"[复制] 已复制 {len(rows)} 行（含表头）为表格，可直接粘贴到 Excel。")
        self.set_status(f"已复制 {len(rows)} 行 · 可直接粘到 Excel", OK)

    def _build_preview_rows(self, rows) -> list:
        """把解析结果拼成预览行（三列，与导出一致）。"""
        return [(r["收件人"], r["手机"], r["地址和品类及数量"]) for r in rows]

    def refresh_preview(self, raw: str) -> list:
        rows = self.ai_state["rows"] if self.ai_state["active"] else parse_text(raw)
        shown = rows[:PREVIEW_LIMIT]
        self.table.set_rows(self._build_preview_rows(shown))
        no_phone = sum(1 for r in rows if not r["手机"])
        if not rows:
            self.lbl_count.configure(text="")
        else:
            tip = f"共 {len(rows)} 条"
            if self.ai_state["active"]:
                tip += " · AI 拆分"
            if no_phone:
                tip += f" · {no_phone} 条未识别手机号"
            if len(rows) > PREVIEW_LIMIT:
                tip += f" · 预览前 {PREVIEW_LIMIT} 条"
            self.lbl_count.configure(text=tip)
        # 数据非空且未用过 AI 拆分时，显示引导链接；
        # AI 正在分析中时，刷新不得干扰「AI 分析中…」，也不要把它撤掉
        if self.ai_state["loading"]:
            return rows
        show_hint = bool(rows) and not self.ai_state["active"] \
            and not self.ai_state["hint_clicked"]
        if show_hint:
            self._set_preview_hint("hint")
            if not self.ai_state["hint_shown"]:
                self.ai_state["hint_shown"] = True
                self.log("[提示] 若本地解析结果不对，可点击『解析预览』右侧的『AI 拆分』。")
        else:
            self._set_preview_hint("none")
        return rows

    # ---------- AI 配置 ----------
    def on_ai_config(self):
        new_cfg = ai_settings_dialog(self, load_ai_config())
        if save_ai_from_dialog(new_cfg):
            self.log(f"[AI设置] 已保存并记住：{new_cfg['base_url'] or '（未填地址）'}"
                     f" · {new_cfg['model']}")
            self.set_status("AI 设置已保存", OK)
        else:
            self.log("[AI设置] 已取消，配置未改动。")

    # ---------- AI 拆分 ----------
    def on_ai_split(self):
        if self.ai_state["loading"]:
            self.set_status("AI 正在分析中，请稍候…", WARN)
            return
        raw = self.current_text()
        if not raw.strip():
            err_popup(self, "还没有数据。\n\n请先在『待处理数据』框中粘贴内容，"
                            "再点『AI 拆分』。")
            return
        self.ai_state["hint_clicked"] = True
        cfg = load_ai_config()

        # 首次使用、配置为空 → 打开设置页引导填写
        if not cfg["base_url"]:
            self.log("[AI] 还没配置 AI 接口，已打开 AI 设置。")
            self.set_status("请先完成 AI 设置", WARN)
            cfg = ai_settings_dialog(self, cfg)
            if not cfg["base_url"]:
                self.log("[AI] 未完成设置，已取消拆分。")
                self.set_status("")
                self._allow_ai_hint()      # 没真正拆 → 链接要回来，用户能再点
                self.refresh_preview(raw)
                return
            if save_ai_from_dialog(cfg):
                self.log("[AI设置] 已保存并记住。")

        self._start_split(raw, cfg)

    # ---------- AI 拆分的异步执行 ----------
    # 为什么不用「主线程直接发请求」：Tk 的事件循环是单线程的，请求期间
    # after() 排队的动画帧要等请求回来才轮得上 —— 表现就是「AI 分析中」
    # 纹丝不动（用户看到的「动画缺失」）。所以改成：
    #   后台线程只负责「网络 + 解析」，一个字都不写界面；
    #   主线程专心跑动画，用队列把结果取回来处理。
    def _start_split(self, raw, cfg, retry: bool = False):
        if retry:
            self.log("[AI] 配置已更新，正在重试…")
            cfg = dict(cfg, _retried=True)
        self.set_status("AI 分析中…")
        self._set_preview_hint("loading")
        # 用 update() 而不是 update_idletasks()：后者只处理布局，不会把
        # 「隐藏引导链接、显示分析中」这一步真正重绘到屏幕上。
        self.update()
        self._split_pending += 1
        if not self._split_polling:
            self._split_polling = True
            self._split_poll_job = self.after(80, self._poll_split_queue)
        threading.Thread(target=self._split_worker, args=(raw, cfg),
                         daemon=True).start()

    def _split_worker(self, raw, cfg):
        """后台线程：只调模型、只算数据，绝不触碰任何界面控件。"""
        try:
            rows, rule = ai_split.split_with_ai(
                raw, cfg["base_url"], cfg["api_key"], cfg["model"])
            item = ("ok", raw, cfg, rows, rule)
        except ai_split.AIConfigError as e:
            item = ("cfg", raw, cfg, e, None)
        except ai_split.AIError as e:
            item = ("aierr", raw, cfg, e, None)
        except Exception as e:
            item = ("ex", raw, cfg, e, None)
        self._split_q.put(item)

    def _poll_split_queue(self):
        """主线程：把后台结果取回来处理（Tk 只允许主线程操作界面）。"""
        self._split_poll_job = None
        while True:
            try:
                item = self._split_q.get_nowait()
            except queue.Empty:
                break
            self._split_pending -= 1
            self._on_split_done(item)
        if self._split_pending > 0:
            self._split_poll_job = self.after(80, self._poll_split_queue)
        else:
            self._split_polling = False

    def _stop_split_poll(self):
        """停掉结果轮询（关窗时调用）。"""
        job = self._split_poll_job
        self._split_poll_job = None
        self._split_polling = False
        if job:
            try:
                self.after_cancel(job)
            except Exception:
                pass

    def _on_split_done(self, item):
        """主线程：拿到结果后在界面上收尾（原来的同步逻辑搬到这里）。"""
        kind, raw, cfg, a, b = item
        self._set_preview_hint("none")
        # 这次拆分已经有了结论 → 重新允许引导链接：
        # 成功的靠 active 挡住（本来就不显示），失败的则让链接回来，用户能再试一次
        self._allow_ai_hint()

        if kind == "ok":
            rows, rule = a, b
            if not rows:
                self.set_status("AI 未拆出数据", WARN)
                self.log("[AI提示] 按模型给的规则没有拆出任何数据，请检查原始文本。")
                err_popup(self, "按模型给出的规则没有拆出任何数据。\n\n"
                                "请确认粘贴的内容是否为收货人数据。")
                return
            # 分析期间用户改了文本 → 结果已对不上，丢弃并说明
            if self.current_text() != raw:
                self.log("[AI提示] 分析期间文本已被修改，本次结果已丢弃，请重新拆分。")
                self.set_status("文本已修改，AI 结果已丢弃", WARN)
                self.refresh_preview(self.current_text())
                return
            self.ai_state["active"] = True
            self.ai_state["rows"] = rows
            self.ai_state["rule"] = rule
            self.refresh_preview(raw)
            self.set_status(f"AI 已拆 {len(rows)} 条", OK)

            if rule.get("line_sep_re"):
                line_desc = "正则模式（见日志）"
            else:
                line_desc = _resolve_line_text(ai_split._resolve_line_sep(rule["line_sep"]))
            rule_desc = (f"分行：{line_desc}"
                         f" · 分列：{_resolve_sep_text(ai_split._resolve_sep(rule['split_by']))}"
                         f" · 列序：{'/'.join(rule['columns'])}")
            self.log(f"[AI] {rule_desc}")
            if rule.get("line_sep_re"):
                self.log(f"[AI] 分条正则：{rule['line_sep_re']}")
            if rule.get("strip_prefix"):
                self.log(f"[AI] 剔除前缀：{rule['strip_prefix']}")
            if rule["reason"]:
                self.log(f"[AI] 依据：{rule['reason']}（把握 {rule['confidence']:.0%}）")
            no_phone = sum(1 for r in rows if not r["手机"])
            if no_phone:
                self.log(f"[AI提示] 其中 {no_phone} 条未识别到手机号，请人工核对。")
            self.log(f"[AI] 拆分完成，共 {len(rows)} 条。")
            self.set_status(f"AI 拆分完成 · 共 {len(rows)} 条", OK)
            return

        if kind == "cfg":
            # 已经重试过一次仍连不上 → 不再弹设置页，避免来回死循环
            if cfg.get("_retried"):
                self.set_status("AI 设置有误，请重新设置", WARN)
                self.log("[AI设置] 重设后仍连接失败，已停止。")
                err_popup(self, f"仍然连接不上：{a}\n\n"
                                f"请确认接口地址和网络是否正常，稍后再试。")
                self.refresh_preview(raw)
                return
            self.set_status("AI 设置有误，请重新设置", WARN)
            self.log(f"[AI设置] 连接失败：{a} 请重新设置。")
            err_popup(self, f"AI 设置有误：{a}\n\n"
                            f"请重新填写接口地址、API Key 和模型名。\n"
                            f"（常见原因：地址填错、Key 失效、网络不通）")
            new_cfg = ai_settings_dialog(self, cfg)
            if save_ai_from_dialog(new_cfg):
                self.log("[AI设置] 已更新并记住。")
            if not new_cfg["base_url"]:
                self.log("[AI] 未完成设置，已取消拆分。")
                self.set_status("")
                self.refresh_preview(raw)
                return
            # 用户已改好配置 → 自动重试一次
            self._start_split(raw, new_cfg, retry=True)
            return

        if kind == "aierr":
            self.set_status("AI 拆分失败", WARN)
            self.log(f"[AI失败] {a}")
            err_popup(self, str(a))
            self.refresh_preview(raw)
            return

        self.set_status("AI 拆分失败", WARN)
        self.log(f"[AI失败] {human_error(a)}")
        err_popup(self, human_error(a))
        self.refresh_preview(raw)

    # ---------- 生成 Excel ----------
    def on_run(self):
        tpl = self.var_template.get().strip()
        if not tpl or not os.path.isfile(tpl):
            err_popup(self, "模板文件不存在。\n\n请在上方『文件设置』里选择模板。")
            return
        rows = self.ai_state["rows"] if self.ai_state["active"] \
            else parse_text(self.current_text())
        if not rows:
            err_popup(self, "没有可生成的数据。\n\n请先粘贴内容，确认预览有数据。")
            return
        outdir = self.var_outdir.get().strip() or "."
        if not os.path.isdir(outdir):
            err_popup(self, f"输出目录不存在：\n{outdir}")
            return
        # 生成时顺手落盘：避免「改完目录/文件名但没失焦就直接关窗」导致没记住
        self.on_output_change()
        ext = template_ext(tpl)
        out = build_out_path(outdir, self.var_filename.get(), ext)
        renamed = os.path.basename(out) != out_name(self.var_filename.get(), ext)
        try:
            n = write_excel(tpl, rows, out)
        except Exception as e:
            err_popup(self, human_error(e), title="生成失败")
            self.log(f"[失败] {human_error(e)}")
            return
        self.set_status(f"已生成 {n} 条", OK)
        self.log(f"[完成] 已生成 {n} 条 → {out}"
                 + ("（已有同名文件，已另存为新文件）" if renamed else ""))
        msg = f"已生成 {n} 条数据。\n\n{out}"
        if renamed:
            msg += "\n\n目标文件夹中已有同名文件，本次已另存为新文件，原文件未改动。"
        info_popup(self, msg, title="生成成功")

    def on_open_dir(self):
        outdir = self.var_outdir.get().strip()
        if not os.path.isdir(outdir):
            err_popup(self, "输出目录不存在。")
            return
        open_path(outdir)


def open_path(path: str):
    """用系统默认方式打开文件或目录（Windows / 其它平台兜底）。"""
    try:
        os.startfile(path)          # Windows
        return
    except Exception:
        pass
    import subprocess
    opener = "explorer" if os.name == "nt" else (
        "open" if sys.platform == "darwin" else "xdg-open")
    try:
        subprocess.Popen([opener, path])
    except Exception:
        pass


def _crash_log_path() -> str:
    """崩溃日志路径（与配置同目录，pythonw 无控制台时也能留痕）。"""
    return os.path.join(data_dir(), "crash.log")


def _install_excepthook():
    """全局异常兜底：pythonw 下没有控制台，崩溃信息会丢失，这里落盘。"""
    import traceback
    import datetime

    def hook(exc_type, exc_val, exc_tb):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_val, exc_tb)
            return
        text = "".join(traceback.format_exception(exc_type, exc_val, exc_tb))
        stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            with open(_crash_log_path(), "a", encoding="utf-8") as f:
                f.write(f"\n===== {stamp} =====\n{text}")
        except Exception:
            pass
        sys.__excepthook__(exc_type, exc_val, exc_tb)

    sys.excepthook = hook


def _trace(msg: str):
    """运行痕迹日志。用于区分「程序自己退出」与「被外部强杀」：
    若只有 start 没有 exit，说明进程是被外部直接终止的。"""
    import datetime
    try:
        p = os.path.join(data_dir(), "run.log")
        with open(p, "a", encoding="utf-8") as f:
            f.write(f"[{os.getpid()}] "
                    f"{datetime.datetime.now():%Y-%m-%d %H:%M:%S} {msg}\n")
    except Exception:
        pass


def main():
    _install_excepthook()
    _trace("start 启动")
    try:
        app = App()
        app.mainloop()
    finally:
        _trace("exit  正常退出（mainloop 结束）")


if __name__ == "__main__":
    main()
