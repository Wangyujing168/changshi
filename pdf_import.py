# -*- coding: utf-8 -*-
"""广联达 PDF 报表解析：只提取每个专业的「单位工程造价汇总表」。

从每页报表中取：
  - 工程名称（如「道路工程清表」「管线工程雨水工程」）
  - 工程造价（汇总表最后一行金额，单位元，含税）
  - 设备费（A5设备费 行金额，单位元）

金额转万元（原值不四舍五入，显示精度由 Excel 单元格格式控制）。

工程名称切分规则：
  - 按专业关键词做最长前缀匹配（智能交通工程须排在交通工程之前）；
  - 剩余部分含「（X）」→ 括号前为节（单位工程），括号内为细目；
  - 剩余部分无括号 → 整体作为叶子名称（如 管线工程雨水工程 → 专业=管线工程，
    叶子=雨水工程）。
"""

import io
import re

# 专业关键词（最长优先匹配）
SPECIALTY_KEYWORDS = [
    "智能交通工程", "电力通信工程", "桥涵工程", "桥梁工程", "道路工程",
    "土方工程", "交通工程", "照明工程", "绿化工程", "管线工程", "拆除工程",
    "外电工程", "景观工程", "隧道工程", "河道工程", "给水工程", "再生水工程",
    "污水工程", "雨水工程", "燃气工程", "热力工程", "电气工程",
]

# 归入「土方工程」专业的节/细目名（清表、填方、挖方均属土方）
_TUFANG_NAMES = {"清表", "填方", "挖方"}

# 总造价放「合计」列、设备费单独提取、安装列=合计-设备费 的专业
ELECTRICAL_SPECIALTIES = {"照明工程", "智能交通工程", "电气工程"}

_BLOCK_HEADER = "单位工程造价汇总表"  # 与压缩空白后的文本比对
_NUM_RE = re.compile(r"[\d,]+\.\d{2}")
_CN_DIGITS = {"零": 0, "壹": 1, "贰": 2, "叁": 3, "肆": 4, "伍": 5, "陆": 6,
              "柒": 7, "捌": 8, "玖": 9, "一": 1, "二": 2, "三": 3, "四": 4,
              "五": 5, "六": 6, "七": 7, "八": 8, "九": 9, "两": 2}
_CN_UNITS = {"拾": 10, "佰": 100, "仟": 1000, "万": 10000, "亿": 100000000}


def _parse_cn_amount(s: str) -> float | None:
    """解析大写金额：壹佰零叁万零陆佰贰拾壹元柒角伍分 → 1030621.75。"""
    m = re.search(r"([零壹贰叁肆伍陆柒捌玖拾佰仟万亿两元角分整]+)", s or "")
    if not m:
        return None
    s = m.group(1)
    if "元" in s:
        yuan_part, tail = s.split("元", 1)
    else:
        yuan_part, tail = s, ""
    total = section = num = 0.0
    for ch in yuan_part:
        if ch in _CN_DIGITS:
            num = _CN_DIGITS[ch]
        elif ch in ("拾", "佰", "仟"):
            section += (num or 1) * _CN_UNITS[ch]
            num = 0
        elif ch == "万":
            section += num
            total += section * 10000
            section = num = 0
        elif ch == "亿":
            section += num
            total += section * 100000000
            section = num = 0
    total += section + num
    jiao = fen = 0
    if "角" in tail:
        i = tail.index("角")
        jiao = _CN_DIGITS.get(tail[i - 1], 0) if i > 0 else 0
    if "分" in tail:
        i = tail.index("分")
        fen = _CN_DIGITS.get(tail[i - 1], 0) if i > 0 else 0
    return total + jiao * 0.1 + fen * 0.01


def _unglue(numstr: str) -> float:
    """去掉与金额粘连的费率前缀（广联达导出常见「100123992.32」=费率100+金额）。"""
    s = numstr.replace(",", "")
    if len(s) >= 11 and s[:3] in ("100", "103", "110") and s[3] != ".":
        return float(s[3:])
    return float(s)


def split_engine_name(full: str):
    """把 工程名称 切成 (专业, 节, 细目)。

    「道路工程路面工程（行车道、路缘带）」→ ("道路工程", "路面工程", "行车道、路缘带")
    「管线工程雨水工程」→ ("管线工程", "", "雨水工程")
    「道路工程清表」→ ("土方工程", "", "清表")（清表/填方/挖方 归土方工程）
    """
    name = (full or "").strip()
    zhuanye = ""
    rest = name
    for kw in sorted(SPECIALTY_KEYWORDS, key=len, reverse=True):
        if name.startswith(kw):
            zhuanye = kw
            rest = name[len(kw):].strip()
            break
    if not zhuanye:
        zhuanye, rest = name, ""
    jie, ximu = "", ""
    if rest:
        m = re.match(r"^(.*?)（(.+?)）\s*$", rest)
        if m:
            jie = m.group(1).strip()
            ximu = m.group(2).strip()
        else:
            ximu = rest
    if jie in _TUFANG_NAMES or ximu in _TUFANG_NAMES:
        zhuanye = "土方工程"  # 清表/填方/挖方 归土方工程
    return zhuanye, jie, ximu


def _extract_zaojia(text: str) -> float | None:
    """工程造价（含税）：优先解析「含税工程造价：壹佰…」大写金额行；
    缺失时回退为「工程造价」块后的首个两位小数（含费率粘连剥离）。"""
    idx = text.find("含税工程造价")
    if idx >= 0:
        val = _parse_cn_amount(text[idx:])
        if val is not None and val > 0:
            return val
    idx = text.find("工程造价")
    if idx < 0:
        return None
    tail = text[idx:]
    m = _NUM_RE.search(tail)
    if not m:
        return None
    return _unglue(m.group(0))


def _extract_device(text: str) -> float:
    """A5设备费 行的金额（无此行则为 0）。"""
    m = re.search(r"A5\s*设备费[\s\S]{0,80}?([\d,]+\.\d{2})", text)
    if not m:
        return 0.0
    return _unglue(m.group(1))


def parse_pdf(source) -> list[dict]:
    """解析广联达 PDF，返回叶子清单（每个「单位工程造价汇总表」一条）。

    source：文件路径 或 bytes。返回元素：
      {full_name, zhuanye, jie, ximu, amount_wan, device_wan,
       electrical, group_no, page}
    分组（group_no）：按页序扫描，同一名称再次出现视为进入下一批
    （一份 PDF 含多个标段时，同名单位工程各出现一次），供网页编辑调整。
    """
    from pypdf import PdfReader
    if isinstance(source, (bytes, bytearray)):
        reader = PdfReader(io.BytesIO(source))
    else:
        reader = PdfReader(source)

    leaves = []
    batch_no = 1
    batch_seen: set = set()
    for page_no, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""
        if _BLOCK_HEADER not in re.sub(r"\s+", "", text):
            continue
        m = re.search(r"工程名称[:：]\s*(.+?)\s*第\s*\d+\s*页\s*共\s*\d+\s*页",
                      text, re.DOTALL)
        if m:
            full_name = re.sub(r"\s+", "", m.group(1))
        else:
            m2 = re.search(r"工程名称[:：]\s*(\S+)", text)
            if not m2:
                continue
            full_name = m2.group(1)
        amount_yuan = _extract_zaojia(text)
        if amount_yuan is None:
            continue
        device_yuan = _extract_device(text)
        if full_name in batch_seen:  # 同名再现 → 进入下一批
            batch_no += 1
            batch_seen = set()
        batch_seen.add(full_name)
        zhuanye, jie, ximu = split_engine_name(full_name)
        leaves.append({
            "full_name": full_name,
            "zhuanye": zhuanye,
            "jie": jie,
            "ximu": ximu,
            "amount_wan": amount_yuan / 10000.0,  # 原值，不四舍五入
            "device_wan": device_yuan / 10000.0,
            "electrical": zhuanye in ELECTRICAL_SPECIALTIES,
            "group_no": batch_no,
            "page": page_no,
        })
    return leaves


def _s(v) -> str:
    """安全转字符串（NaN/None → ""）。"""
    if v is None:
        return ""
    try:
        if isinstance(v, float) and v != v:  # NaN
            return ""
    except TypeError:
        pass
    return str(v).strip()


def _f(v) -> float:
    """安全转 float（NaN/None/空 → 0）。"""
    try:
        x = float(v)
        return 0.0 if x != x else x
    except (TypeError, ValueError):
        return 0.0


# 中文列名（st.data_editor 输出）→ 内部英文键；两种键都接受
_ROW_KEYS = {"组号": "group_no", "组名": "group_name", "专业": "zhuanye",
             "节": "jie", "细目": "ximu", "造价(万元)": "amount_wan",
             "设备费(万元)": "device_wan", "电气专业": "electrical"}


def build_groups(rows: list[dict]) -> list[dict]:
    """把（用户编辑后的）叶子行组装成 组→专业→叶子 结构，供导出与预览。

    rows 元素需含：group_no(int)、group_name(str)、zhuanye、jie、ximu、
    amount_wan、device_wan、electrical(bool)；也接受中文列名键
    （组号/组名/专业/节/细目/造价(万元)/设备费(万元)/电气专业）。
    返回 [{name, items: [{zhuanye, electrical, leaves: [...]}]}]，
    叶子名称 = 有细目用细目、否则用节（无节则叶子名本身已含在 ximu/jie 中）。
    完全空白的行（无专业且无金额）跳过。
    """
    groups: dict[int, dict] = {}
    for raw in rows:
        row = {en: raw.get(en, raw.get(cn)) for cn, en in _ROW_KEYS.items()}
        zy_raw = _s(row.get("zhuanye"))
        amount_raw = _f(row.get("amount_wan"))
        if not zy_raw and not amount_raw:
            continue  # 空白行（编辑器新增未填）
        try:
            gno = int(row.get("group_no"))
        except (TypeError, ValueError):
            gno = 0
        g = groups.setdefault(gno, {
            "name": _s(row.get("group_name")),
            "items": {},
        })
        zy = zy_raw or "未分类"
        it = g["items"].setdefault(zy, {
            "zhuanye": zy,
            "electrical": bool(row.get("electrical")),
            "leaves": [],
        })
        jie = _s(row.get("jie"))
        ximu = _s(row.get("ximu"))
        leaf_name = ximu or jie or zy
        it["leaves"].append({
            "name": leaf_name,
            "jie": jie,
            "amount": amount_raw,
            "device": _f(row.get("device_wan")) if it["electrical"] else 0.0,
        })
    out = []
    for gno in sorted(groups):
        g = groups[gno]
        items = []
        for zy in g["items"]:
            it = g["items"][zy]
            it["leaves"].sort(key=lambda x: (x["jie"] or x["name"], x["name"]))
            items.append(it)
        out.append({"name": g["name"], "items": items})
    return out


def parse_and_group(source) -> list[dict]:
    """解析 PDF 并按出现次序自动分组（组名留空待用户填写）。"""
    return build_groups([
        {**leaf, "group_name": ""} for leaf in parse_pdf(source)
    ])
