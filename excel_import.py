# -*- coding: utf-8 -*-
"""广联达 Excel 报表导入（单项造价分析 / 项目造价分析）。

序号层级：N=单项、N.M=单位工程（专业）、N.M.K=细项。三层表（项目造价分析）
按单项分组；两层表（单项造价分析，顶层即专业）整体一个组。金额(元)→万元；
细项行的「工程规模」「单位造价」直接成为该细项的技术经济指标（导出表
单位/数量/单位价值列）；大项（专业）行指标由 app 层按口径聚合（道路=
m2 路面面积+m 道路长度提问、照明=m 道路长度、信号灯=路口，其余同单位
求和+金额÷数量折算）。

单位按名称关键词推断（用户 2026-11-11 确认口径）：
  土方/路基→m3；路面/拆除道路/桥涵/绿化→m2；拆除电杆→根；乔木迁移→株；
  井/拆除井→座；车止石/树穴石/井圈加固→个；路缘石→m；照明/信号灯及电子
  警察（智能交通）子项→套；各类管线/电力通信→m。
"""
import io
import re

import pandas as pd

_LEVEL_RE = re.compile(r"^(\d+)(?:\.(\d+))?(?:\.(\d+))?$")

# 单位关键词推断（有序，先匹配先得；叶行推断为空时回退父专业推断）
_UNIT_RULES = [
    (re.compile(r"挖方|填方|路基|土方"), "m3"),
    (re.compile(r"拆除电杆"), "根"),
    (re.compile(r"乔木|砍树|伐树|移植"), "株"),
    (re.compile(r"井"), "座"),
    (re.compile(r"车止石|树穴石|井圈加固"), "个"),
    (re.compile(r"路缘石|缘石"), "m"),
    (re.compile(r"照明|信号灯|电子警察|智能交通|智慧合杆|路灯|标志标线"), "套"),
    (re.compile(r"管线|给水|给排|雨水|污水|中水|燃气|热力|排管"
                r"|电力|通信|电缆|光缆|φ|DN"), "m"),
    (re.compile(r"路面|道路|桥|涵|绿化|清表|人行道|车行道|行车道"
                r"|拆除道路|土工布|格栅|罩面"), "m2"),
]


def infer_unit(name: str) -> str:
    """按名称关键词推断计量单位；无匹配返回 ""。"""
    n = str(name or "")
    for rx, u in _UNIT_RULES:
        if rx.search(n):
            return u
    return ""


def _level_of(v):
    """序号值 → 层级元组 (a, b, c)；无法识别返回 None。b/c 为 None 表示无。"""
    if v is None:
        return None
    if isinstance(v, float) and v != v:   # NaN
        return None
    s = str(v).strip()
    if isinstance(v, float) and float(v).is_integer():
        s = str(int(v))
    m = _LEVEL_RE.match(s)
    if not m:
        return None
    return tuple(int(x) if x else None for x in m.groups())


def _fnum(v):
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    return None if f != f else f


def _col_of(df_header_row, *prefixes):
    """在表头行里找列名（任一前缀匹配，取第一个命中）。"""
    for c in range(len(df_header_row)):
        v = str(df_header_row.iloc[c] or "").strip()
        for p in prefixes:
            if v.startswith(p):
                return c
    return None


def _leaf_row(gno, gname, zy, xm, amount, qty, uv, unit, parent_unit):
    """组装一行叶子（中文列名，与 pdf_rows 编辑表格列一致）。"""
    if amount is None:
        return None
    u = unit or parent_unit or ""
    row = {"组号": gno, "组名": gname or "", "专业": zy, "节": "",
           "细目": xm or "", "造价(万元)": round(amount / 10000.0, 4),
           "设备费(万元)": 0.0, "电气专业": False,
           "单位": u, "工程规模": qty, "单位造价(元)": uv}
    if qty is not None:
        row["工程规模"] = round(float(qty), 4)
    if uv is not None:
        row["单位造价(元)"] = round(float(uv), 4)
    return row


def parse_analysis(source, group_offset: int = 1, filename: str = "") -> list[dict]:
    """解析 单项造价分析 / 项目造价分析 Excel（.xls/.xlsx 字节或路径）。

    返回叶子行列表（组号从 group_offset 起）：组号/组名/专业/节/细目/
    造价(万元)/设备费(万元)/电气专业/单位/工程规模/单位造价(元)。
    三层表：组名=顶层名称（如 12#水系桥）；两层表：单组、组名留空。
    """
    src = io.BytesIO(source) if isinstance(source, (bytes, bytearray)) else source
    xl = pd.ExcelFile(src)
    df = None
    hdr = None
    for sn in xl.sheet_names:
        raw = xl.parse(sn, header=None)
        for r in range(min(raw.shape[0], 20)):
            row_vals = [str(raw.iat[r, c] or "") for c in range(raw.shape[1])]
            if ("序号" in row_vals and "名称" in row_vals
                    and any(v.startswith("金额") for v in row_vals)):
                df, hdr = raw, r
                break
        if df is not None:
            break
    if df is None:
        raise ValueError(
            f"未识别到「序号/名称/金额」表头：{filename or 'Excel 文件'}")

    c_no = _col_of(df.iloc[hdr], "序号")
    c_name = _col_of(df.iloc[hdr], "名称")
    c_amt = _col_of(df.iloc[hdr], "金额")
    c_qty = _col_of(df.iloc[hdr], "工程规模")
    c_uv = _col_of(df.iloc[hdr], "单位造价")
    if c_no is None or c_name is None or c_amt is None:
        raise ValueError(
            f"缺少 序号/名称/金额 列：{filename or 'Excel 文件'}")

    # 先扫一遍确定最大层级（3 层=含单项层；2 层=顶层即专业）
    max_depth = 1
    for r in range(hdr + 1, df.shape[0]):
        lv = _level_of(df.iat[r, c_no] if c_no < df.shape[1] else None)
        if lv:
            max_depth = max(max_depth, 1 + sum(1 for x in lv[1:] if x is not None))

    gno = group_offset
    rows: list[dict] = []
    cur_group = None          # (组号, 组名)
    cur_zy = None             # (专业名, 父专业推断单位)
    parent_pend = None        # 尚无子项的父行（有子项时清空；换父行时若有残留则发出）

    def _flush_parent():
        """换父行/结束时：若上一个父行一直没有子项且自身有金额 → 作为叶子发出。"""
        nonlocal parent_pend
        if parent_pend is not None:
            _g, _z, _amt, _qty, _uv, _unit = parent_pend
            row = _leaf_row(_g[0], _g[1], _z, _z, _amt, _qty, _uv,
                            _unit, infer_unit(_z))
            if row:
                rows.append(row)
            parent_pend = None

    for r in range(hdr + 1, df.shape[0]):
        lv = _level_of(df.iat[r, c_no])
        name = str(df.iat[r, c_name] or "").strip() if c_name < df.shape[1] else ""
        if not lv:
            continue                      # 合计行/空行/子表标题行
        if name in ("合计", "总计"):
            continue
        amt = _fnum(df.iat[r, c_amt]) if c_amt < df.shape[1] else None
        qty = _fnum(df.iat[r, c_qty]) if c_qty is not None and c_qty < df.shape[1] else None
        uv = _fnum(df.iat[r, c_uv]) if c_uv is not None and c_uv < df.shape[1] else None
        unit = infer_unit(name)

        if max_depth >= 3:
            if lv[1] is None:             # 顶层 = 单项（每个单项一组）
                _flush_parent()
                cur_group = (gno, name)
                gno += 1
                cur_zy = None
                continue
            if lv[2] is None:             # 中间层 = 专业
                _flush_parent()
                cur_zy = (name, unit)
                parent_pend = (cur_group, name, amt, qty, uv, unit) \
                    if cur_group is not None else None
                continue
            # 三层叶子：父专业已有子项，不再单独成行
            parent_pend = None
            if cur_group is not None and cur_zy is not None:
                row = _leaf_row(cur_group[0], cur_group[1], cur_zy[0], name,
                                amt, qty, uv, unit, cur_zy[1])
                if row:
                    rows.append(row)
        else:
            if lv[1] is None:             # 顶层 = 专业（两层表：整个文件一个组）
                _flush_parent()
                if cur_group is None:
                    cur_group = (gno, "")
                    gno += 1
                cur_zy = (name, unit)
                parent_pend = (cur_group, name, amt, qty, uv, unit)
            else:                          # 两层叶子
                parent_pend = None
                row = _leaf_row(cur_group[0], cur_group[1], cur_zy[0], name,
                                amt, qty, uv, unit, cur_zy[1] if cur_zy else "")
                if row:
                    rows.append(row)
    _flush_parent()
    return rows


# ── 大项（专业行）技术经济指标默认口径 ──

_ROAD_AREA_RE = re.compile(r"行车道|车行道|非机动车道|人行道")


def item_default_tech(zhuanye: str, leaves: list[dict],
                      road_len: float = 0.0, cross_n: float = 0.0) -> list[dict]:
    """大项技术经济指标默认值 → [{unit, qty, val}, ...]（金额÷数量折算）。

    leaves: 该专业的叶子列表（元素含 amount / tech{unit,qty,val}）。
    道路工程：m2=路面面积（行车道/车行道+非机动车道+人行道）+ m=道路长度
    （页面提问）；照明工程：m=道路长度；信号灯及电子警察（智能交通）：路口，
    智能交通未填路口数时默认 m=道路长度。
    其余专业：按专业名推断单位，同单位子项数量求和、val=大项金额÷数量。
    """
    zy = str(zhuanye or "")
    amt = sum(float(x.get("amount") or 0) for x in leaves or [])
    if not amt:
        return []

    def _grp(u, q):
        if not u or not q or q <= 0:
            return None
        # 单位价值(元/单位) = 金额(万元)×10000 ÷ 数量
        return {"unit": u, "qty": round(float(q), 4),
                "val": round(amt * 10000.0 / float(q), 2)}

    if "道路" in zy or "车行道" in zy or "道路工程" in zy:
        area = 0.0
        for x in leaves or []:
            t = x.get("tech") or {}
            nm = str(x.get("name") or "")
            if _ROAD_AREA_RE.search(nm) and str(t.get("unit") or "") == "m2":
                area += float(t.get("qty") or 0)
        out = []
        g = _grp("m2", area)
        if g:
            out.append(g)
        g = _grp("m", road_len)
        if g:
            out.append(g)
        return out
    if "照明" in zy:
        g = _grp("m", road_len)
        return [g] if g else []
    if "信号灯" in zy or "电子警察" in zy or "智能交通" in zy:
        g = _grp("路口", cross_n)
        if g:
            return [g]
        # 智能交通没有路口数时：默认单位 m、工程量=道路长度
        if "智能交通" in zy:
            g = _grp("m", road_len)
            return [g] if g else []
        return []
    u = infer_unit(zy)
    if not u:
        return []
    q = 0.0
    for x in leaves or []:
        t = x.get("tech") or {}
        if str(t.get("unit") or "") == u:
            q += float(t.get("qty") or 0)
    g = _grp(u, q)
    return [g] if g else []


def tech_to_text(groups) -> tuple[str, str, str]:
    """指标组列表 → (单位文本, 数量文本, 单位价值文本)；多组用「、」拼接。"""
    if not groups:
        return "", "", ""

    def _f(x):
        if x is None or x == "":
            return ""
        try:
            v = float(x)
        except (TypeError, ValueError):
            return str(x)
        return str(int(v)) if v == int(v) else str(round(v, 2))

    return ("、".join(str(g.get("unit") or "") for g in groups),
            "、".join(_f(g.get("qty")) for g in groups),
            "、".join(_f(g.get("val")) for g in groups))


def tech_from_text(u_s, q_s, v_s) -> list[dict]:
    """编辑器文本 → 指标组列表（单位按「、/」拆分；数量/单价按「、」）。"""
    u_s = str(u_s or "").strip()
    if not u_s:
        return []
    units = [p.strip() for p in re.split(r"[、,，/]", u_s) if p.strip()]
    qtys = [p.strip() for p in re.split(r"[、,，]", str(q_s or "")) if p.strip()]
    vals = [p.strip() for p in re.split(r"[、,，]", str(v_s or "")) if p.strip()]
    out = []
    for i, u in enumerate(units):
        q = qtys[i] if i < len(qtys) else ""
        v = vals[i] if i < len(vals) else ""
        try:
            q = float(q)
        except (TypeError, ValueError):
            q = None
        try:
            v = float(v)
        except (TypeError, ValueError):
            v = None
        out.append({"unit": u, "qty": q, "val": v})
    return out
