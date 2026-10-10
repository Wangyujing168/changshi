# -*- coding: utf-8 -*-
"""表格预处理：解析「主要工程数量总表.xls」→ 填充「表格.xlsx」信息输入表。

纯逻辑模块（不依赖 Streamlit），app.py 的「表格预处理」页面上传数量总表后调用：
- parse_xls(data: bytes) -> dict            解析 .xls，得到各部位工程量
- mapping_rows(m) -> list[(单元格,项目,值)]    生成核对预览行
- fill_template(tpl: bytes, m: dict) -> bytes  写入信息输入表，返回新 xlsx 字节

映射规则（用户 2026-10-08/09 拍板）：
- 桩基础混凝土按桩径分行（模板 B 列 1/1.2/1.5/1.8/2），模板没有的桩径提示不导入；
- 数值保留 2 位小数；工程数量表中没有的（承台HPB/各部位HPB等）留空；
- 承台 HRB = 工程数量表口径：桥台承台 80325 + 梯道承台 6144（弃用下部 sheet 旧版 76500）；
- 声测管 = 下部 sheet「超声波检测管」钢管φ57x3.5 + φ70x6.5 之和；
- 数量总表里模板没有的部位（如搭板）→ 在信息输入表末尾追加 七/八… 新分项。
"""
import io
from copy import copy as _copy
from pathlib import Path

import openpyxl
import xlrd

# 内置表格模板（与本模块同目录）
_TEMPLATE_PATH = Path(__file__).with_name("表格模板.xlsx")


def load_builtin_template() -> bytes:
    """读取内置表格模板 表格模板.xlsx（与本模块同目录）。"""
    if not _TEMPLATE_PATH.exists():
        raise QuantityImportError("内置表格模板缺失：表格模板.xlsx")
    return _TEMPLATE_PATH.read_bytes()


class QuantityImportError(Exception):
    """数量总表或模板不是预期格式（缺表/缺关键行/缺桩基类型区）。"""


def _num(v):
    return float(v) if isinstance(v, (int, float)) else 0.0


def _mat_row(sh, mat: str, sub: str, r0: int = 5, r1: int = 62) -> int:
    """按 B 列材料大类 + C 列细分找行（B 列为合并单元格，需跟踪当前大类）。"""
    cur = ""
    for r in range(r0, min(r1, sh.nrows)):
        b = str(sh.cell_value(r, 1)).strip()
        if b:
            cur = b
        if cur == mat and str(sh.cell_value(r, 2)).strip() == sub:
            return r
    raise QuantityImportError(f"数量总表未找到「{mat} / {sub}」行")


def _norm(s) -> str:
    return "".join(str(s).split())


def _group_cols(sh) -> dict:
    """r7 组名（桥台/桥墩/梯道下部，取第一次出现）→ (起始列, 结束列, 子列标签→列)。"""
    r7 = 6
    bounds = sorted(c for c in range(sh.ncols)
                    if _norm(sh.cell_value(r7, c)))
    first = {}
    for c in bounds:
        first.setdefault(_norm(sh.cell_value(r7, c)), c)
    groups = {}
    for g in ("桥台", "桥墩", "梯道下部"):
        s = first.get(g)
        if s is None:
            raise QuantityImportError(f"数量总表表头缺少「{g}」组")
        e = next((b for b in bounds if b > s), sh.ncols)
        cols = {}
        for c in range(s, e):
            lab = _norm(sh.cell_value(7, c))
            if lab:
                cols[lab] = c
        groups[g] = (s, e, cols)
    return groups


# 已被固定映射消费的 (组, 子列)；其余子列若带数据 → 追加为新分项
_CONSUMED = {("桥台", "台帽"), ("桥台", "台身"), ("桥台", "承台"), ("桥台", "桩基"),
             ("桥墩", "盖梁"), ("桥墩", "墩柱及柱系梁"), ("桥墩", "桩基"),
             ("梯道下部", "桥墩"), ("梯道下部", "承台"), ("梯道下部", "桩基")}


def _extra_cols(sh, groups) -> list[tuple[str, int]]:
    """组区间内未被消费的子列（含搭板列）→ [(名称, 列号)]。"""
    out = []
    for g, (s, e, cols) in groups.items():
        for lab, c in cols.items():
            if lab == "小计" or (g, lab) in _CONSUMED:
                continue
            out.append((f"{g}·{lab}", c))
    # 搭板：r8 无子列标签，r5 有（在桥台组起始列到梯道下部组结束列之间）
    g0 = groups["桥台"][0]
    g1 = groups["梯道下部"][1]
    for c in range(g0, g1):
        v5 = _norm(sh.cell_value(4, c))
        v8 = _norm(sh.cell_value(7, c))
        if v5 == "搭板" and not v8:
            out.append(("搭板", c))
    return out


def parse_xls(data: bytes) -> dict:
    """解析 .xls 数量总表，返回映射结果（桩列表/各部位工程量/新增分项/警告）。"""
    try:
        wb = xlrd.open_workbook(file_contents=data)
    except Exception as e:
        raise QuantityImportError(f"无法打开 .xls 文件：{e}")
    if "工程数量表" not in wb.sheet_names() or "下部" not in wb.sheet_names():
        raise QuantityImportError("缺少「工程数量表」或「下部」工作表")
    sh = wb.sheet_by_name("工程数量表")
    groups = _group_cols(sh)

    # 混凝土/钢材行（混凝土取全部等级行，钢筋取 HPB300/HRB400）
    conc_rows = []
    cur_mat = ""
    for r in range(5, sh.nrows):
        b = str(sh.cell_value(r, 1)).strip()
        if b:
            cur_mat = b
        c = str(sh.cell_value(r, 2)).strip()
        if cur_mat == "混凝土" and c and _norm(c) != "小计":
            conc_rows.append(r)
    hpb = _mat_row(sh, "钢材", "HPB300")
    hrb = _mat_row(sh, "钢材", "HRB400")

    def val_all(rows_idx, group, sub) -> float:
        c = groups[group][2].get(sub)
        if c is None:
            return 0.0
        return sum(_num(sh.cell_value(r, c)) for r in rows_idx)

    m = {"warnings": [], "extras": []}

    # ── 桩基类型（下部 sheet）：数量/桩径/桩长，按桩径聚合 ──
    lo = wb.sheet_by_name("下部")
    piles = {}
    in_sec = False
    for r in range(lo.nrows):
        v0 = str(lo.cell_value(r, 0)).strip()
        if v0 == "桩基类型":
            in_sec = True
        if in_sec:
            n, d, L = lo.cell_value(r, 3), lo.cell_value(r, 4), lo.cell_value(r, 5)
            if isinstance(n, (int, float)) and n and isinstance(d, (int, float)) \
                    and isinstance(L, (int, float)):
                k = float(d)
                p = piles.setdefault(k, {"d": k, "n": 0, "len": 0.0})
                p["n"] += int(n)
                p["len"] += int(n) * float(L)
            elif v0 and v0 != "桩基类型":
                break
    if not piles:
        raise QuantityImportError("「下部」表未找到桩基类型区")
    m["piles"] = [piles[k] for k in sorted(piles)]

    # ── 声测管 = 超声波检测管行的 钢管φ57x3.5 + φ70x6.5（可能分多段）──
    g57 = g70 = None
    sg = 0.0
    for r in range(lo.nrows):
        for c in range(lo.ncols):
            v = str(lo.cell_value(r, c)).strip()
            if v.startswith("钢管") and "φ57" in v:
                g57 = c
            elif v.startswith("钢管") and "φ70" in v:
                g70 = c
        if str(lo.cell_value(r, 0)).strip() == "超声波检测管":
            if g57 is not None:
                sg += _num(lo.cell_value(r, g57))
            if g70 is not None:
                sg += _num(lo.cell_value(r, g70))
    m["shengguan"] = sg

    # ── 固定映射部位（工程数量表口径）──
    m["pile_hpb"] = val_all([hpb], "桥台", "桩基") + val_all([hpb], "桥墩", "桩基") \
        + val_all([hpb], "梯道下部", "桩基")
    m["pile_hrb"] = val_all([hrb], "桥台", "桩基") + val_all([hrb], "桥墩", "桩基") \
        + val_all([hrb], "梯道下部", "桩基")
    m["chengtai_conc"] = val_all(conc_rows, "桥台", "承台") \
        + val_all(conc_rows, "梯道下部", "承台")
    m["chengtai_hpb"] = val_all([hpb], "桥台", "承台") \
        + val_all([hpb], "梯道下部", "承台")
    m["chengtai_hrb"] = val_all([hrb], "桥台", "承台") \
        + val_all([hrb], "梯道下部", "承台")
    m["qiaotai_conc"] = val_all(conc_rows, "桥台", "台帽") \
        + val_all(conc_rows, "桥台", "台身")
    m["qiaotai_hpb"] = val_all([hpb], "桥台", "台帽") \
        + val_all([hpb], "桥台", "台身")
    m["qiaotai_hrb"] = val_all([hrb], "桥台", "台帽") \
        + val_all([hrb], "桥台", "台身")
    m["qiaodun_conc"] = val_all(conc_rows, "桥墩", "墩柱及柱系梁") \
        + val_all(conc_rows, "梯道下部", "桥墩")
    m["qiaodun_hpb"] = val_all([hpb], "桥墩", "墩柱及柱系梁") \
        + val_all([hpb], "梯道下部", "桥墩")
    m["qiaodun_hrb"] = val_all([hrb], "桥墩", "墩柱及柱系梁") \
        + val_all([hrb], "梯道下部", "桥墩")
    m["dungailiang_conc"] = val_all(conc_rows, "桥墩", "盖梁")
    m["dungailiang_hpb"] = val_all([hpb], "桥墩", "盖梁")
    m["dungailiang_hrb"] = val_all([hrb], "桥墩", "盖梁")

    # ── 新增分项（如搭板）：未被固定映射消费且有数据的子列 ──
    for label, c in _extra_cols(sh, groups):
        conc = sum(_num(sh.cell_value(r, c)) for r in conc_rows)
        hp = _num(sh.cell_value(hpb, c))
        hr = _num(sh.cell_value(hrb, c))
        if conc or hp or hr:
            m["extras"].append({
                "name": label,
                "conc": round(conc, 2),
                "hpb": round(hp, 2) if hp else None,
                "hrb": round(hr, 2) if hr else None,
            })

    # ── 提示 ──
    upper = sum(_num(sh.cell_value(r, c)) for r in conc_rows for c in range(4, 9))
    if upper > 0:
        m["warnings"].append("上部构造（现浇箱梁/梯道/桥面系）不在信息输入表范围内，未导入。")
    return m


# 预览行：固定单元格 → 映射键（信息输入表 D 列）
_CELLS = {
    "D11": ("桩基础 HPB 钢筋(kg)", "pile_hpb"),
    "D12": ("桩基础 HRB 钢筋(kg)", "pile_hrb"),
    "D13": ("声测管(m)", "shengguan"),
    "D15": ("承台 混凝土(m3)", "chengtai_conc"),
    "D16": ("承台 HPB 钢筋(kg)", "chengtai_hpb"),
    "D17": ("承台 HRB 钢筋(kg)", "chengtai_hrb"),
    "D19": ("桥台 混凝土(m3)", "qiaotai_conc"),
    "D20": ("桥台 HPB 钢筋(kg)", "qiaotai_hpb"),
    "D21": ("桥台 HRB 钢筋(kg)", "qiaotai_hrb"),
    "D27": ("桥墩 混凝土(m3)", "qiaodun_conc"),
    "D28": ("桥墩 HPB 钢筋(kg)", "qiaodun_hpb"),
    "D29": ("桥墩 HRB 钢筋(kg)", "qiaodun_hrb"),
    "D31": ("墩盖梁 混凝土(m3)", "dungailiang_conc"),
    "D32": ("墩盖梁 HPB 钢筋(kg)", "dungailiang_hpb"),
    "D33": ("墩盖梁 HRB 钢筋(kg)", "dungailiang_hrb"),
}

_PILE_CELL = {1.0: 6, 1.2: 7, 1.5: 8, 1.8: 9, 2.0: 10}


def mapping_rows(m: dict) -> list[tuple[str, str, object]]:
    """生成核对预览行 (单元格, 项目, 值)。"""
    out = []
    for p in m.get("piles") or []:
        r = _PILE_CELL.get(float(p["d"]))
        if r is None:
            continue
        out.append((f"D{r}", f"桩基础 φ{float(p['d']):g} 桩长(m)",
                    round(float(p["len"]), 2)))
        out.append((f"F{r}", f"桩基础 φ{float(p['d']):g} 根数", int(p["n"])))
    for cell, (name, key) in _CELLS.items():
        v = m.get(key)
        out.append((cell, name, round(v, 2) if v else None))
    for ex in m.get("extras") or []:
        out.append(("新增", f"{ex['name']} 混凝土(m3)", ex["conc"]))
        out.append(("新增", f"{ex['name']} HPB 钢筋(kg)", ex["hpb"]))
        out.append(("新增", f"{ex['name']} HRB 钢筋(kg)", ex["hrb"]))
    return out


_CN = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十",
       "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八"]


def fill_template(tpl: bytes, m: dict) -> bytes:
    """按映射结果写入模板「信息输入表」，返回新 xlsx 字节。"""
    wb = openpyxl.load_workbook(io.BytesIO(tpl))
    if "信息输入表" not in wb.sheetnames:
        raise QuantityImportError("表格模板缺少「信息输入表」工作表")
    ws = wb["信息输入表"]

    # ── 桩基础：按 B 列桩径匹配模板行 6-10，其余桩径行清空 ──
    pile_rows = {}
    for r in range(6, 11):
        b = ws.cell(row=r, column=2).value
        if isinstance(b, (int, float)):
            pile_rows[float(b)] = r
    for p in m.get("piles") or []:
        d = float(p["d"])
        r = pile_rows.get(d)
        if r is None:
            w = f"桩径 φ{d:g} 在模板中没有对应行，未导入。"
            if w not in (m.get("warnings") or []):
                m.setdefault("warnings", []).append(w)
            continue
        length = float(p["len"])
        ws.cell(row=r, column=4).value = int(length) if length.is_integer() \
            else round(length, 2)
        ws.cell(row=r, column=6).value = int(p["n"])
    for d, r in pile_rows.items():
        if not any(float(p["d"]) == d for p in m.get("piles") or []):
            ws.cell(row=r, column=4).value = None
            ws.cell(row=r, column=6).value = None

    # ── 固定 D 列（没有数据 → 留空）──
    for cell, (_, key) in _CELLS.items():
        v = m.get(key)
        ws[cell] = round(v, 2) if v else None
    # 台盖梁：数量总表无来源 → 留空
    for cell in ("D23", "D24", "D25"):
        ws[cell] = None

    # ── 新增分项：追加在现有内容之后（样式照抄承台段 14-17 行）──
    start = 34
    for r in range(34, 70):
        if ws.cell(row=r, column=1).value not in (None, "") \
                or ws.cell(row=r, column=2).value not in (None, ""):
            start = r + 1
    for i, ex in enumerate(m.get("extras") or []):
        r = start
        ws.cell(row=r, column=1).value = _CN[6 + i]
        ws.cell(row=r, column=2).value = ex["name"]
        ws.cell(row=r + 1, column=1).value = "混凝土"
        ws.cell(row=r + 1, column=3).value = "m3"
        ws.cell(row=r + 1, column=4).value = ex["conc"]
        ws.cell(row=r + 2, column=1).value = "钢筋"
        ws.cell(row=r + 2, column=2).value = "HPB"
        ws.cell(row=r + 2, column=3).value = "kg"
        ws.cell(row=r + 2, column=4).value = ex["hpb"]
        ws.cell(row=r + 3, column=2).value = "HRB"
        ws.cell(row=r + 3, column=3).value = "kg"
        ws.cell(row=r + 3, column=4).value = ex["hrb"]
        for src_row, dst_row in ((14, r), (15, r + 1), (16, r + 2), (17, r + 3)):
            for col in range(1, 8):
                sc = ws.cell(row=src_row, column=col)
                dc = ws.cell(row=dst_row, column=col)
                dc.font = _copy(sc.font)
                dc.border = _copy(sc.border)
                dc.alignment = _copy(sc.alignment)
                dc.number_format = sc.number_format
        start = r + 4

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
