# -*- coding: utf-8 -*-
"""
AI 智能分词模块（OpenAI 兼容接口）

设计思路
--------
用户粘贴的数据格式千奇百怪，纯正则搞不定。把问题拆成两步：

  第一步：这些数据按什么分行？（默认按换行）
  第二步：每一行按什么拆成 收件人 / 手机 / 地址 三列？

模型不直接返回拆好的数据，而是只返回这两个"拆分规则"参数，
本地再用规则去切数据。好处：
  - 省 token、快、便宜
  - 结果可复算、可校验、可缓存
  - 用户能看懂模型给了什么规则，不信任时可以手动改

接口地址与 API Key 默认留空，走 OpenAI 兼容协议（/chat/completions），
适配 OpenAI / DeepSeek / 通义 / 硅基流动 / Ollama / vLLM 等一切兼容端点。
"""
import json
import re
import urllib.error
import urllib.request

# ---------------- 默认配置（留空，由用户在界面填写） ----------------
DEFAULT_BASE_URL = ""   # 例：https://api.deepseek.com/v1  或  http://localhost:11434/v1
DEFAULT_API_KEY = ""    # 例：sk-xxxxxxxx
DEFAULT_MODEL = "gpt-4o-mini"

REQUEST_TIMEOUT = 45    # 秒


# ---------------- Prompt ----------------
SYSTEM_PROMPT = """你是一个数据格式分析助手。用户会给你一段从剪贴板粘贴的收货人数据，\
你需要分析出「如何把这段文本拆成一条条记录」以及「每条记录如何拆成三列」。

拆解为两步思考：

【第一步】按什么分行？
  观察这段文本，每条独立的收货记录是用什么分隔的。常见情况：
  - "\\n"        每条一行（最常见）
  - "\\n\\n"      空行分隔（一条记录可能自身占多行）
  - ";"  "；"    分号分隔
  - "|"          竖线分隔
  - "\\t\\t"      连续制表符分隔
  如果一条记录本身会跨多行（比如地址换行），要识别出真正的"记录分隔符"，
  而不是简单地用换行。

【第二步】每行按什么拆成三列？
  观察单条记录内部，收件人、手机号、地址之间是用什么隔开的。常见情况：
  - "空格"       最宽松，按空白字符切，且只切前两段，剩下的全归地址
  - "\\t"        制表符
  - ","  "，"    逗号（中文或英文）
  - "|"          竖线
  - "无"         没有明显分隔符，只能靠手机号位置来切

  注意：如果记录内没有稳定分隔符，但收件人在最前、手机号紧随其后，
  可以把 split_by 设为 "无"，届时程序会用手机号正则来切分。

【列的顺序】
  绝大多数情况是「收件人 手机号 地址」。
  如果手机号在第一位，或地址在中间，要用 columns 字段说明真实顺序。

严格只输出一个 JSON 对象，不要任何解释文字、不要 markdown 代码块。格式：

{
  "line_sep": "\\n",
  "split_by": "空格",
  "columns": ["收件人", "手机", "地址"],
  "confidence": 0.9,
  "reason": "一句话说明你的判断依据"
}

字段约束：
- line_sep：记录之间的分隔符，用转义写法。可选 "\\n"、"\\n\\n"、";"、"|"、"\\t\\t" 等。
- split_by：列之间的分隔符。可选 "空格"、"\\t"、","、"|"、"无" 等。
- columns：长度必须是 3，元素从 ["收件人", "手机", "地址"] 中取，代表三段内容依次对应哪一列。
- confidence：0 到 1 的小数，你对这次判断的把握。
- reason：不超过 40 字的中文说明。"""

USER_PROMPT_TEMPLATE = """请分析下面这段数据，给出分行规则和分列规则。

--- 数据开始 ---
{sample}
--- 数据结束 ---

注意：
1. 如果数据超过 20 条，上面只截取了前 20 条，按这个样本判断规则即可。
2. 地址中可能包含空格、逗号等字符，判断分隔符时要选真正稳定的那个。
3. 严格只输出 JSON。"""


# ---------------- 异常 ----------------
class AIError(Exception):
    """AI 分词相关错误，message 已经是人话，可直接展示给用户。"""


# ---------------- 接口调用 ----------------
def call_model(messages: list, base_url: str, api_key: str, model: str,
               timeout: int = REQUEST_TIMEOUT, response_format: bool = True) -> str:
    """调用 OpenAI 兼容的 /chat/completions 接口，返回 assistant 的文本内容。

    response_format 为 True 时请求 JSON 输出（部分端点不支持，失败会自动重试一次）。
    """
    base_url = (base_url or "").strip().rstrip("/")
    if not base_url:
        raise AIError("未配置接口地址。请点击『AI 设置』填写 OpenAI 兼容的接口地址（例如 "
                      "https://api.deepseek.com/v1）。")
    if not (base_url.startswith("http://") or base_url.startswith("https://")):
        raise AIError("接口地址格式不对，需要以 http:// 或 https:// 开头。")

    url = base_url + "/chat/completions"
    payload = {"model": model or DEFAULT_MODEL, "messages": messages, "temperature": 0}
    if response_format:
        payload["response_format"] = {"type": "json_object"}
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")

    headers = {"Content-Type": "application/json"}
    if (api_key or "").strip():
        headers["Authorization"] = "Bearer " + api_key.strip()

    req = urllib.request.Request(url, data=body, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            data = json.loads(resp.read().decode("utf-8", "replace"))
    except urllib.error.HTTPError as e:
        detail = ""
        try:
            detail = e.read().decode("utf-8", "replace")[:300]
        except Exception:
            pass
        if e.code == 401:
            raise AIError("接口认证失败（401）。请检查 API Key 是否正确、是否已过期。")
        if e.code == 404:
            raise AIError("接口地址不存在（404）。请确认地址是否为 ……/v1 形式，"
                          "有些服务不需要 /v1 后缀。")
        if e.code == 429:
            raise AIError("请求过于频繁或额度不足（429）。稍后再试，或检查账户余额。")
        # 502 / 503 / 504 多为网关或代理没连上真实端点
        if e.code in (502, 503, 504):
            raise AIError(f"连接不上接口（HTTP {e.code}）。请确认接口地址填写正确、"
                          f"服务可访问；若走代理，检查代理是否拦截了该域名。")
        # 部分端点不支持 response_format，去掉后重试一次
        if response_format and e.code == 400:
            return call_model(messages, base_url, api_key, model, timeout, False)
        raise AIError(f"接口返回错误（HTTP {e.code}）。{detail}")
    except urllib.error.URLError as e:
        raise AIError(f"网络连接失败：{e.reason}。请检查接口地址是否可达、网络或代理是否正常。")
    except json.JSONDecodeError:
        raise AIError("接口返回的不是合法 JSON，请确认该地址是 OpenAI 兼容的 chat 接口。")

    try:
        return data["choices"][0]["message"]["content"] or ""
    except (KeyError, IndexError, TypeError):
        raise AIError(f"接口返回结构不符合 OpenAI 规范：{str(data)[:200]}")


def _extract_json(text: str) -> dict:
    """从模型输出里抠出 JSON 对象（容忍 ```json 包裹和前后废话）。"""
    text = (text or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```[a-zA-Z]*\s*", "", text)
        text = re.sub(r"\s*```$", "", text).strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass
    # 兜底：取第一个 { 到最后一个 } 之间的内容
    start, end = text.find("{"), text.rfind("}")
    if start != -1 and end > start:
        try:
            return json.loads(text[start:end + 1])
        except json.JSONDecodeError:
            pass
    raise AIError("模型没有返回可解析的 JSON，请重试，或换一个模型试试。")


# ---------------- 第一步：取样本 ----------------
def build_sample(raw: str, max_lines: int = 20, max_chars: int = 3000) -> str:
    """截取前若干行作为分析样本，避免 token 浪费。"""
    text = (raw or "").replace("\r\n", "\n").replace("\r", "\n").strip()
    if not text:
        raise AIError("待处理数据是空的，请先粘贴内容。")
    lines = text.split("\n")[:max_lines]
    sample = "\n".join(lines)
    if len(sample) > max_chars:
        sample = sample[:max_chars]
    return sample


# ---------------- 主入口：让模型给出拆分参数 ----------------
VALID_COLUMNS = ("收件人", "手机", "地址")
# 模型可能返回的各种别名，统一映射到内部列名
_COLUMN_ALIAS = {
    "收件人": "收件人", "姓名": "收件人", "名字": "收件人", "收货人": "收件人",
    "手机": "手机", "手机号": "手机", "电话": "手机", "手机号码": "手机",
    "联系电话": "手机", "号码": "手机", "phone": "手机", "tel": "手机",
    "地址": "地址", "收货地址": "地址", "详细地址": "地址",
    "地址和品类及数量": "地址", "商品": "地址", "品类": "地址",
}


def normalize_rule(data: dict) -> dict:
    """校验并规范化模型返回的规则，非法值一律回退到安全默认。"""
    if not isinstance(data, dict):
        raise AIError("模型返回的内容不是 JSON 对象。")

    line_sep = data.get("line_sep", "\n")
    if not isinstance(line_sep, str) or line_sep == "":
        line_sep = "\n"

    split_by = data.get("split_by", "空格")
    if not isinstance(split_by, str) or not split_by:
        split_by = "空格"

    cols = data.get("columns", ["收件人", "手机", "地址"])
    if not isinstance(cols, list) or len(cols) != 3:
        cols = ["收件人", "手机", "地址"]
    normed = []
    for c in cols:
        key = str(c).strip().lower()
        normed.append(_COLUMN_ALIAS.get(str(c).strip()) or _COLUMN_ALIAS.get(key) or "地址")
    # 去重校验：三列必须是三个不同角色，否则回退默认顺序
    if sorted(normed) != sorted(VALID_COLUMNS):
        normed = ["收件人", "手机", "地址"]

    try:
        confidence = float(data.get("confidence", 0.0))
    except (TypeError, ValueError):
        confidence = 0.0
    confidence = min(max(confidence, 0.0), 1.0)

    reason = str(data.get("reason", "") or "")[:80]
    return {"line_sep": line_sep, "split_by": split_by,
            "columns": normed, "confidence": confidence, "reason": reason}


def ask_split_rule(raw: str, base_url: str, api_key: str, model: str,
                   timeout: int = REQUEST_TIMEOUT) -> dict:
    """让模型分析文本，返回拆分规则 dict：

    {"line_sep": "\\n", "split_by": "空格", "columns": [...],
     "confidence": 0.9, "reason": "..."}
    """
    sample = build_sample(raw)
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": USER_PROMPT_TEMPLATE.format(sample=sample)},
    ]
    content = call_model(messages, base_url, api_key, model, timeout)
    return normalize_rule(_extract_json(content))


# ---------------- 第二步：用规则把数据切出来 ----------------
_SEP_ALIAS = {
    "空格": " ", "空 格": " ", "space": " ", "whitespace": " ", "空白": " ",
    "tab": "\t", "\\t": "\t", "制表符": "\t", "制表": "\t",
    "逗号": ",", "，": "，", "comma": ",", ",": ",",
    "分号": ";", "；": ";", "semicolon": ";",
    "竖线": "|", "管道符": "|", "pipe": "|", "|": "|",
    "无": "", "none": "", "": "", "null": "",
    "\\n": "\n", "换行": "\n", "newline": "\n",
}
_NO_SPLIT = {"", "无", "none", "null", "无分隔符", "没有", "不存在", "靠手机号"}


def _resolve_sep(split_by: str) -> str:
    """把模型给的分隔符描述翻译成真实字符；'' 表示无分隔符。"""
    s = (split_by or "").strip()
    low = s.lower()
    if low in _NO_SPLIT:
        return ""
    if s in _SEP_ALIAS:
        return _SEP_ALIAS[s]
    if low in _SEP_ALIAS:
        return _SEP_ALIAS[low]
    # 形如 "\t" 这类字面转义，还原成真实字符
    if s.startswith("\\") and len(s) >= 2:
        try:
            return s.encode("utf-8").decode("unicode_escape")
        except Exception:
            pass
    return s  # 其它情况当作字面分隔符


def _resolve_line_sep(line_sep: str) -> str:
    """分行符同理，空值兜底为换行。"""
    s = (line_sep or "").strip()
    if s == "" or s.lower() in ("none", "null", "无"):
        return "\n"
    if s == "\\n":
        return "\n"
    if s == "\\r\\n":
        return "\r\n"
    if s == "\\t\\t":
        return "\t\t"
    if s.startswith("\\") and len(s) >= 2:
        try:
            return s.encode("utf-8").decode("unicode_escape")
        except Exception:
            pass
    return s


def _split_columns(line: str, sep: str, cols: list) -> dict:
    """把单行文本按分隔符切成三段，并按 cols 顺序映射到 收件人/手机/地址。"""
    line = line.strip()
    parts = []
    if sep == " ":
        # 空格模式：只切前两段，剩下的全部归第三列（地址常含空格）
        parts = re.split(r"\s+", line, maxsplit=2)
    elif sep:
        parts = [p.strip() for p in line.split(sep)]
        # 分隔符模式：多余段合并进最后一段（地址里可能也有同种符号）
        if len(parts) > 3:
            parts = parts[:2] + [sep.join(parts[2:])]
    else:
        parts = [line]

    parts = [p.strip(" \t，,;；、") for p in parts]
    if len(parts) == 1:
        # 只有一段：退回手机号定位的老逻辑
        return _fallback_by_phone(line)

    # 补齐到三段
    while len(parts) < 3:
        parts.append("")

    row = {"收件人": "", "手机": "", "地址": ""}
    for name, val in zip(cols, parts):
        row[name] = val
    return row


def _fallback_by_phone(line: str) -> dict:
    """无分隔符时的兜底：用手机号正则定位，切成 前/手机/后。"""
    m = re.search(r"1[3-9]\d{9}(?:-\d{4})?", line)
    if not m:
        return {"收件人": line.strip(" \t，,;；、"), "手机": "", "地址": ""}
    return {
        "收件人": line[:m.start()].strip(" \t，,;；、"),
        "手机": m.group(0),
        "地址": line[m.end():].strip(" \t，,;；、"),
    }


def apply_rule(raw: str, rule: dict) -> list:
    """接收模型给出的两个参数（line_sep / split_by），生成拆好的数据列表。

    返回 [{"收件人": ..., "手机": ..., "地址和品类及数量": ...}, ...]
    与 parse_text() 的输出结构完全一致，可直接喂给 write_excel()。
    """
    text = (raw or "").replace("\r\n", "\n").replace("\r", "\n")
    line_sep = _resolve_line_sep(rule.get("line_sep", "\n"))
    sep = _resolve_sep(rule.get("split_by", "空格"))
    cols = rule.get("columns") or ["收件人", "手机", "地址"]

    records = text.split(line_sep) if line_sep else [text]
    rows = []
    for rec in records:
        # 一条记录内部若还有换行（如地址换行），拉平成空格，避免写进单元格带换行符
        rec = rec.replace("\r", " ").replace("\n", " ").strip()
        rec = re.sub(r"[ \t]{2,}", " ", rec)
        if not rec:
            continue
        row = _split_columns(rec, sep, cols)
        # 内部统一用模板表头名，方便直接写 Excel
        rows.append({
            "收件人": row.get("收件人", ""),
            "手机": row.get("手机", ""),
            "地址和品类及数量": row.get("地址", ""),
        })
    return rows


def split_with_ai(raw: str, base_url: str, api_key: str, model: str,
                  timeout: int = REQUEST_TIMEOUT) -> tuple:
    """一步到位：调模型拿规则 → 用规则切数据。

    返回 (rows, rule)，rule 里含 confidence / reason，可展示给用户。
    """
    rule = ask_split_rule(raw, base_url, api_key, model, timeout)
    rows = apply_rule(raw, rule)
    return rows, rule
