# -*- coding: utf-8 -*-
"""公路桥梁定额匹配引擎（公路工程概预算定额库，模板列结构导出）。

库格式（.xls，xlrd 读取，15 列）：项/目/节/细目代号、费率编号、定额代号、名称、
单位、工程量1、工程量2、定额调整情况、说明、代号、默认项。
换算（钢筋抽换 / 实际运距 / 实际桩径 / 配合比调整）集中在「定额调整情况」+
「说明」两列，表述方式与套好的模板（桥梁工程(高沙岭互通立交D匝道桥).xls）一致。

匹配结构镜像模板：QL01 基础工程（灌注桩基础 / 钢围堰 / 承台(系梁) / 土方工程）→
QL02 下部构造（桥墩 / 桥台）→ QL03 上部构造（预应力混凝土小箱梁 / 钢箱梁）→
QL04 桥面铺装（沥青混凝土铺装 / 水泥混凝土铺装 / 桥面防水层）→
QL05 桥梁附属结构（支座 / 伸缩缝 / 护栏与护网 / 搭板 / 锥坡 / 防抛网）→
QL06 其他工程（SX，不计费）。
"""
import math
import os
import re
from io import BytesIO

import xlrd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from bridge_match import BridgeMatchError, parse_quantity_excel

# 内置默认公路定额库（概算+预算混合，带 [部2018预] 前缀的为预算定额条目）
_DEFAULT_LIB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                 "quota_base", "公路定额+预算.xls")
_default_lib_cache = None

ROAD_HEADERS = ["项的代号", "目的代号", "节的代号", "细目代号", "费率编号",
                "定额代号", "项或目或节或细目或定额的名称", "单位",
                "工程量 1", "工程量 2", "定额调整情况", "说明", "代号", "默认项"]

DEFAULT_PARAMS = {
    "hun_km": 8.0,      # 混凝土搅拌运输车运距 km
    "gang_km": 5.0,     # 钢材运输运距 km
    "tu_km": 15.0,      # 土方外运运距 km
    "long_km": 5.0,     # 钢筋笼运输运距 km
    "liq_km": 5.0,      # 沥青混合料运输运距 km
    "nijiang_dj": 25.0, # 泥浆处置费单价 元/m3
    "shatu_pct": 20.0,  # 钻孔砂土比例 %
    "yashi": 1.16,      # 填方压实换算系数（余土外运 = 挖方 − 填方×系数）
    "ganghu_t": 0.0,    # 钢护筒总量 t（0=按定额说明(16)表内插×每根2m估算并提示）
    "gjx_shu": 40.0,    # 钢绞线束长档 m（20/40/80）
    "qiaochang": 0.0,   # 桥梁长度 m（护栏节统计量）
    "weiyan_t": 0.0,    # 钢围堰量 t（0=从数量表自动识别）
}

# 库缺口兜底子目（模板口径；新库里有同编号时以库为准）
_BORROWED = {
    "001001008005": ("01", "斗容量1.0m3以内挖掘机挖装普通土", "1000m3天然密实方"),
    "001001010009": ("03", "装载质量15t以内自卸汽车运土", "1000m3天然密实方"),
    "001001010013": ("03", "装载质量30t以内自卸汽车运土", "1000m3天然密实方"),
    "001004003002": ("06", "预制、铺砌混凝土预制块护坡", "10m3"),
    "005001001004": ("06", "现浇钢筋混凝土防撞护栏混凝土", "10m3实体"),
    "005001001005": ("10", "现浇钢筋混凝土防撞护栏钢筋", "1t"),
    "005001001006": ("06", "现浇钢筋混凝土防撞护栏制作、安装铸铁柱及栏杆", "1t"),
    "005001005006": ("14", "防落(物)网(钢管立柱)", "100m"),
    "[部2018预]41105006": ("07", "基础混凝土垫层", "10m3实体"),
    "[部2018预]90105071": ("03", "装载质量20t以内载货汽车运输钢材", "100t"),
    "[部2018预]40613003": ("04", "桥面行车道铺装水泥混凝土面层泵送", "10m3"),
}

# 公路钢护筒单位质量表（定额说明(16)）：桩径cm → 成品质量 kg/m（含加劲肋/法兰盘）
CASING_KG_M = {100: 267, 120: 390, 150: 568, 200: 919,
               250: 1504, 300: 1961, 350: 2576}


def _casing_unit_kg(dia_m):
    """钢护筒单位质量 kg/m：按桩径(m)在 CASING_KG_M 线性内插，超界取端点值。"""
    cm = dia_m * 100
    pts = sorted(CASING_KG_M)
    if cm <= pts[0]:
        return float(CASING_KG_M[pts[0]])
    if cm >= pts[-1]:
        return float(CASING_KG_M[pts[-1]])
    for i in range(len(pts) - 1):
        a, b = pts[i], pts[i + 1]
        if a <= cm <= b:
            wa, wb = CASING_KG_M[a], CASING_KG_M[b]
            return wa + (wb - wa) * (cm - a) / (b - a)
    return float(CASING_KG_M[pts[-1]])  # unreachable



# ============================ 定额库 ============================

def parse_road_library(data):
    """解析公路定额库 .xls（或 .xlsx）。条目行定额代号带 ,DJ 后缀，章节行定额代号=DJ。"""
    try:
        wb = xlrd.open_workbook(file_contents=data)
        sh = wb.sheet_by_index(0)
        entries = {}
        for r in range(1, sh.nrows):
            code_raw = str(sh.cell_value(r, 5)).strip()
            name = str(sh.cell_value(r, 6)).strip()
            if not code_raw or not name or code_raw == "DJ":
                continue
            code = code_raw.split(",")[0].strip()
            if code in entries:
                continue
            entries[code] = {
                "code": code,
                "rate": str(sh.cell_value(r, 4)).strip(),
                "name": name,
                "unit": str(sh.cell_value(r, 7)).strip(),
            }
        return {"entries": entries, "n": len(entries)}
    except Exception as e:
        # .xlsx 回退（新库可能是 xlsx）
        try:
            wb = load_workbook(BytesIO(data), read_only=True, data_only=True)
            ws = wb.worksheets[0]
            entries = {}
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or len(row) < 8:
                    continue
                code_raw = str(row[5] or "").strip()
                name = str(row[6] or "").strip()
                if not code_raw or not name or code_raw == "DJ":
                    continue
                code = code_raw.split(",")[0].strip()
                if code in entries:
                    continue
                entries[code] = {
                    "code": code,
                    "rate": str(row[4] or "").strip(),
                    "name": name,
                    "unit": str(row[7] or "").strip(),
                }
            return {"entries": entries, "n": len(entries)}
        except Exception as e2:
            raise BridgeMatchError(f"定额库解析失败：{e2}")


def default_road_library():
    """内置默认公路定额库（quota_base/公路定额+预算.xls），解析结果模块级缓存。"""
    global _default_lib_cache
    if _default_lib_cache is not None:
        return _default_lib_cache
    if os.path.exists(_DEFAULT_LIB_PATH):
        with open(_DEFAULT_LIB_PATH, "rb") as f:
            _default_lib_cache = parse_road_library(f.read())
    else:
        _default_lib_cache = {"entries": {}, "n": 0}
    return _default_lib_cache


def _find_entry(code, lib):
    """库优先，其次模板内置兜底子目。"""
    lib = lib or {}
    e = lib.get("entries", lib).get(code) if isinstance(lib, dict) else None
    if e:
        return e
    if code in _BORROWED:
        rate, name, unit = _BORROWED[code]
        return {"code": code, "rate": rate, "name": name, "unit": unit}
    return None


# ============================ 数量表辅助 ============================

def _scan_quantity_extra(data):
    """扫描数量总表：支座规格×数量、钢围堰量、工程名称。"""
    extra = {"bearings": [], "weiyan_t": None, "title": ""}
    try:
        wb = load_workbook(BytesIO(data), read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        # 工程名称：第 1 行 / 第 2 行含「桥梁工程数量表」
        title = ""
        for row in ws.iter_rows(min_row=1, max_row=3, values_only=True):
            for v in row:
                if isinstance(v, str) and "数量表" in v and len(v) > 4:
                    title = re.sub(r"\s*桥梁工程数量表.*$", "", v).strip()
                    break
            if title:
                break
        extra["title"] = title
        # 支座块：行首为「支座」的行起，收集 (规格, 数量) 对
        bearing = False
        for row in ws.iter_rows(values_only=True):
            if not row:
                continue
            a, b = str(row[0] or "").strip(), str(row[1] or "").strip()
            if a == "支座" or "支座" in a:
                bearing = True
                if a == "支座":
                    pairs = [(row[3], row[6]), (row[7], row[10])]
                    for spec, n in pairs:
                        if isinstance(spec, str) and isinstance(n, (int, float)) and n:
                            extra["bearings"].append((spec.strip(), float(n)))
                    continue
            if bearing:
                spec = a if isinstance(a, str) and a else str(row[3] or "").strip()
                if spec and re.match(r"^(GBZY|LRB|GPZ|GJZ)", spec):
                    n = row[6]
                    if isinstance(n, (int, float)) and n:
                        extra["bearings"].append((spec, float(n)))
                    continue
            if bearing and a and a != "支座":
                bearing = False
            # 钢围堰量：备注类文本
            for v in row:
                if isinstance(v, str) and "钢围堰" in v:
                    m = re.search(r"钢围堰[^0-9]*(\d+(?:\.\d+)?)\s*t", v)
                    if m:
                        extra["weiyan_t"] = float(m.group(1))
    except Exception:
        pass
    return extra


def _sum_by(items, fx, mat):
    """按 分项+材料 求和。"""
    s, n = 0.0, 0
    for it in items:
        if (it.get("fenxiang") or "").strip() == fx and str(it.get("mat") or "").strip() == mat:
            v = it.get("value")
            if isinstance(v, (int, float)):
                s += v
                n += 1
    return s, n


def _r(v, nd=4):
    if v is None:
        return None
    v = round(float(v), nd)
    return int(v) if v == int(v) else v


# ============================ 匹配 ============================

def _quota(rate, code, name, unit, q1, adj="", note="", q2=""):
    return {"t": "quota", "d": "", "rate": rate, "code": code + ",DJ",
            "name": name, "unit": unit, "q1": q1, "q2": q2,
            "adj": adj, "note": note}


def _sec(rows, level, code, name, unit, q1="", q2="", code_cell="DJ", adj="", note=""):
    row = {"t": level, "d": code, "rate": "", "code": code_cell, "name": name,
           "unit": unit, "q1": q1, "q2": q2, "adj": adj, "note": note}
    rows.append(row)
    return row


def _dist_name(base, km=None):
    """运距类子目：库内基准名去掉『第一个1km』/尾缀 km，接实际运距。"""
    if km is None:
        return base
    name = re.sub(r"第一个1km|\d+(?:\.\d+)?km", "", base).rstrip()
    return f"{name}{_r(km, 1)}km".replace(".0km", "km")


def _adj_yunju(name, km):
    return (f"[实际运距(km)#实际运距(km)@{_r(km, 1)}],12,{name},0"
            .replace("@5.0", "@5").replace("@8.0", "@8").replace("@15.0", "@15")), \
           f"实际运距(km): {_r(km, 1)}km".replace(".0km", "km")


_GANG_NOTES = {"HPB300": "钢筋抽换：[2001002] 换 [2001001]",
               "HRB400": "钢筋抽换：[2001001] 换 [2001002]"}


def match_road_quotas(items, lib, qty_data=None, params=None):
    """items: parse_quantity_excel 结果；lib: parse_road_library 结果（可为 None）。
    qty_data: 数量表原始字节（支座/围堰/名称扫描）。返回
    {rows, warnings, skipped}；rows 为模板列结构行列表。"""
    p = dict(DEFAULT_PARAMS)
    if params:
        p.update(params)
    lib = lib or {}
    entries = lib.get("entries", lib) if isinstance(lib, dict) else {}
    warnings, skipped = [], []
    rows = []

    # ---- 基础聚合 ----
    def s2(fx, mat):  # 两部位合计
        return _sum_by(items, fx, mat)[0]

    pile_conc = s2("桥墩 桩基", "C35水下") + s2("桥台 桩基", "C35水下")
    pile_hpb = s2("桥墩 桩基", "HPB300") + s2("桥台 桩基", "HPB300")
    pile_hrb = s2("桥墩 桩基", "HRB400") + s2("桥台 桩基", "HRB400")
    pile_dias = []
    for it in items:
        if it.get("kind") == "桩基参数":
            pile_dias.append((float(it.get("dia")), float(it.get("count")),
                              float(it.get("depth"))))
    pile_dias.sort(key=lambda x: x[0])
    pile_len_total = sum(c * d for _, c, d in pile_dias)

    chengtai = s2("桥墩 系梁承台", "C35") + s2("桥台 系梁承台", "C35")
    dianceng = s2("桥墩 系梁承台", "C20") + s2("桥台 系梁承台", "C20")
    chengtai_hrb = s2("桥墩 系梁承台", "HRB400") + s2("桥台 系梁承台", "HRB400")

    dun_conc = s2("桥墩 盖梁", "C40") + s2("桥墩 盖梁", "C50") + s2("桥墩 墩柱", "C40")
    dun_hrb = s2("桥墩 盖梁", "HRB400") + s2("桥墩 墩柱", "HRB400")
    tai_conc = s2("桥台 台帽", "C40") + s2("桥台 台帽", "C50") + s2("桥台 耳背墙", "C40")
    tai_hrb = s2("桥台 台帽", "HRB400") + s2("桥台 耳背墙", "HRB400")

    xl_conc = s2("预应力小箱梁", "C50")
    xl_hpb = s2("预应力小箱梁", "HPB300")
    xl_hrb = s2("预应力小箱梁", "HRB400")
    gjx_t = s2("15.2钢绞线", "15.2钢绞线")
    gxq345 = s2("钢箱梁", "Q345钢材")

    area = s2("桥梁面积", "桥梁面积")
    hl_conc = s2("护栏地袱", "C50")
    hl_hrb = s2("护栏地袱", "HRB400")
    hl_q235 = s2("护栏地袱", "Q235钢材")
    dab_conc = s2("搭板", "C35")
    dab_hrb = s2("搭板", "HRB400")

    sma = s2("桥面铺装", "SMA13SBS改性沥青")
    hmac = s2("桥面铺装", "高模量中粒式沥青混凝土HMAC-20C，SBS改性沥青")
    fs_conc = s2("桥面铺装", "C50防水混凝土")
    uhpc = s2("桥面铺装", "UHPC超高韧性混凝土")
    pm_crb = s2("桥面铺装", "CRB550")
    pm_hpb = s2("桥面铺装", "HPB300")
    pm_hrb = s2("桥面铺装", "HRB400")
    fangshui = s2("桥面防水层", "桥面防水层")
    nianceng = s2("沥青粘层油", "沥青粘层油")

    ss_80 = s2("伸缩缝", "80型伸缩缝长度")
    ss_160 = s2("伸缩缝", "160型伸缩缝长度")
    ss_conc = s2("伸缩缝", "C50混凝土")
    ss_hrb_kg = s2("伸缩缝", "HRB400")

    llk = s2("C30预制混凝土六棱块", "C30预制混凝土六棱块")
    wafang = s2("挖方", "挖方")
    tianfang = s2("填方", "填方")
    fpw_len = s2("防抛网", "总长度")

    extra = _scan_quantity_extra(qty_data or b"")
    weiyan_t = p.get("weiyan_t") or 0
    if not weiyan_t and extra["weiyan_t"]:
        weiyan_t = extra["weiyan_t"]
    title = (p.get("title") or "").strip() or extra["title"] or "公路桥梁工程"
    qiaochang = p.get("qiaochang") or 0

    hun_km, gang_km = p.get("hun_km") or 8, p.get("gang_km") or 5
    tu_km, long_km = p.get("tu_km") or 15, p.get("long_km") or 5
    liq_km = p.get("liq_km") or 5
    nijiang_dj = p.get("nijiang_dj") or 25
    shatu = (p.get("shatu_pct") or 20) / 100.0
    yashi = p.get("yashi") or 1.16

    # ---- 总行 ----
    _sec(rows, "xiang", "", title, "m/座",
         q1=_r(area) if area else "", q2=_r(qiaochang) if qiaochang else "")

    # ================= QL01 基础工程 =================
    _sec(rows, "xiang", "QL01", "基础工程", "m3")
    _sec(rows, "mu", "QL0102", "桩基础", "m3/m",
         q1=_r(pile_conc) if pile_conc else "",
         q2=_r(pile_len_total) if pile_len_total else "")
    _sec(rows, "jie", "QL010201", "灌注桩基础", "m3",
         q1=_r(pile_conc) if pile_conc else "")

    # ---- 钻孔（砂土/粘土 × 桩径档）----
    # 口径（用户确认 2026-11-10 最终版）：桩径档与实际桩径一致（库无该桩径档才上靠
    # 一档+实际桩径换算）；同一桩径不区分桩长，孔深档按桩长向上取档（数量表只有
    # 合计米数→按平均桩长=合计m÷根数；80m/100m 等档按实际桩长选用，不是固定60m），
    # 超过该桩径最大档按最大档套用；砂土=总桩长×0.2、粘土×0.8。
    DIA_DEPTH_GRID = {100: [30, 40], 120: [40, 50, 60], 150: [40, 60, 80],
                      200: [40, 60, 80, 100], 250: [40, 60, 80, 100],
                      300: [40, 60, 80, 100, 130], 350: [40, 60, 80, 100, 130]}
    DRILL_FALLBACK = {
        (100, 30, "砂土"): "004002015001", (100, 30, "粘土"): "004002015002",
        (100, 40, "砂土"): "004002015009", (100, 40, "粘土"): "004002015010",
        (120, 40, "砂土"): "004002015017", (120, 40, "粘土"): "004002015018",
        (120, 50, "砂土"): "004002015025", (120, 50, "粘土"): "004002015026",
        (120, 60, "砂土"): "004002015033", (120, 60, "粘土"): "004002015034",
        (150, 40, "砂土"): "004002015041", (150, 40, "粘土"): "004002015042",
        (150, 60, "砂土"): "004002015049", (150, 60, "粘土"): "004002015050",
        (150, 80, "砂土"): "004002015057", (150, 80, "粘土"): "004002015058",
        (200, 40, "砂土"): "004002015065", (200, 40, "粘土"): "004002015066",
        (200, 60, "砂土"): "004002015073", (200, 60, "粘土"): "004002015074",
        (200, 80, "砂土"): "004002015081", (200, 80, "粘土"): "004002015082",
        (200, 100, "砂土"): "004002015089", (200, 100, "粘土"): "004002015090",
        (250, 40, "砂土"): "004002015097", (250, 40, "粘土"): "004002015098",
        (250, 60, "砂土"): "004002015105", (250, 60, "粘土"): "004002015106",
        (250, 80, "砂土"): "004002015113", (250, 80, "粘土"): "004002015114",
        (250, 100, "砂土"): "004002015121", (250, 100, "粘土"): "004002015122",
        (300, 40, "砂土"): "004002015129", (300, 40, "粘土"): "004002015130",
        (300, 60, "砂土"): "004002015137", (300, 60, "粘土"): "004002015138",
        (300, 80, "砂土"): "004002015145", (300, 80, "粘土"): "004002015146",
        (300, 100, "砂土"): "004002015153", (300, 100, "粘土"): "004002015154",
        (300, 130, "砂土"): "004002015161", (300, 130, "粘土"): "004002015162",
        (350, 40, "砂土"): "004002015169", (350, 40, "粘土"): "004002015170",
        (350, 60, "砂土"): "004002015177", (350, 60, "粘土"): "004002015178",
        (350, 80, "砂土"): "004002015185", (350, 80, "粘土"): "004002015186",
        (350, 100, "砂土"): "004002015193", (350, 100, "粘土"): "004002015194",
        (350, 130, "砂土"): "004002015201", (350, 130, "粘土"): "004002015202",
    }
    if pile_dias and pile_conc > 0:
        for dia, cnt, depth in pile_dias:
            d_cm = int(round(dia * 100))
            # 桩径档：与实际桩径一致（无该档则上靠最小档，套实际桩径换算）
            d_lv = d_cm if d_cm in DIA_DEPTH_GRID else \
                next((d for d in sorted(DIA_DEPTH_GRID) if d >= d_cm), None)
            if d_lv is None:
                d_lv = max(DIA_DEPTH_GRID)
                warnings.append(f"桩径 φ{_r(dia, 1)} 超过库内最大桩径档"
                                f"（{max(DIA_DEPTH_GRID)}cm），已按最大档套用")
            # 孔深档（用户口径 2026-11-10 最终版）：同一桩径不区分桩长，按桩长
            # 向上取档（80m/100m 等档按实际桩长选用，不是固定60m）；
            # 超过该桩径最大档按最大档套用
            hs = DIA_DEPTH_GRID[d_lv]
            h_lv = next((h for h in hs if h >= depth), None)
            if h_lv is None:
                h_lv = hs[-1]
            zong = cnt * depth
            sa = zong * shatu
            nian = zong - sa
            for soil, q in (("砂土", sa), ("粘土", nian)):
                pat = f"陆地上回旋钻机钻孔桩径{d_lv}cm以内孔深{h_lv}m以内{soil}"
                code = name = unit = None
                for c, e in entries.items():  # 库内按名称找档
                    if e["name"] == pat:
                        code, name, unit = c, e["name"], e["unit"]
                        break
                if code is None:
                    code = DRILL_FALLBACK.get((d_lv, h_lv, soil), "")
                    name, unit = pat, "10m"
                    if not code:
                        warnings.append(f"库内无钻孔子目（{pat}），请补全定额库")
                adj = note = ""
                if d_lv != d_cm:  # 库档≠实际桩径 → 实际桩径换算（桩径保持一致）
                    name = name.replace(f"{d_lv}cm以内", f"{d_cm}cm以内")
                    adj = f"[实际桩径(cm)#{d_cm}cm],12,{name},0"
                    note = f"实际桩径(cm): {d_cm}cm"
                rows.append(_quota("07", code, name, unit, _r(q / 10), adj, note))
    elif pile_dias:
        warnings.append("桩基块无混凝土量，钻孔/灌注混凝土/泥浆处置未生成")

    # ---- 钢护筒（公路口径：按定额说明(16)桩径-单位质量表内插 × 每根2m）----
    ganghu = p.get("ganghu_t") or 0
    if pile_dias and (ganghu or pile_conc > 0):
        if not ganghu:
            ganghu = sum(_casing_unit_kg(dia) * 2.0 * cnt
                         for dia, cnt, _ in pile_dias) / 1000
            warnings.append(f"钢护筒量按定额说明(16)表内插×每根2m估算得 "
                            f"{_r(ganghu, 4)}t（φ1.0→267/φ1.5→568/φ1.8→778.6kg/m），"
                            f"设计已给定护筒总质量时请填参数框覆盖")
        e = _find_entry("004002020006", lib)
        rows.append(_quota("10-1", "004002020006", e["name"] if e else "干处埋设钢护筒",
                           e["unit"] if e else "1t", _r(ganghu, 4)))

    # ---- 灌注桩混凝土（按桩长×截面积理论量比例拆分到桩径，用户确认口径）----
    if pile_dias and pile_conc > 0:
        w = [c * d * math.pi * (dia / 2) ** 2 for dia, c, d in pile_dias]
        wt = sum(w)
        w = [x / wt for x in w]
        conc_by = [pile_conc * x for x in w]
        DIA_CONC = {1.0: "004002019008", 1.5: "004002019011", 1.8: "004002019014"}
        for (dia, _, _), q in zip(pile_dias, conc_by):
            key = 1.0 if dia <= 1.2 else 1.5 if dia <= 1.65 else 1.8
            code = DIA_CONC[key]
            e = _find_entry(code, lib)
            rows.append(_quota("07", code, e["name"] if e else "灌注桩混凝土",
                               e["unit"] if e else "10m3实体", _r(q / 10)))

    # ---- 桩基钢筋 ----
    for t, q in (("HPB300", pile_hpb), ("HRB400", pile_hrb)):
        if q <= 0:
            continue
        e = _find_entry("004005001013", lib)
        base = e["name"] if e else "集中加工现浇混凝土灌注桩钢筋"
        name = f"{base} {t}"
        rows.append(_quota("10-1", "004005001013", name,
                           e["unit"] if e else "1t钢筋", _r(q),
                           f"[钢筋抽换#{t}钢筋],12,{name},0",
                           _GANG_NOTES.get(t, "")))

    # ---- 钢筋笼运输 ----
    gang_t = pile_hpb + pile_hrb
    if gang_t > 0:
        e = _find_entry("004005001032", lib)
        base = e["name"] if e else "平板拖车运输钢筋笼"
        name = _dist_name(base, long_km)
        adj, note = _adj_yunju(name, long_km)
        rows.append(_quota("03", "004005001032", name,
                           e["unit"] if e else "100t钢筋", _r(gang_t / 100), adj, note))

    # ---- 拌和 + 搅拌车（桩基 吊斗灌注 ×1.0）----
    def _banhe(rate, conc, coef):
        e = _find_entry("004006001016", lib)
        rows.append(_quota(rate, "004006001016",
                           e["name"] if e else "生产能力90m3/h以内混凝土搅拌站(楼)拌和",
                           e["unit"] if e else "100m3", _r(conc * coef / 100)))

    def _jiaoban(conc, coef, km=None):
        km = km if km is not None else hun_km
        e = _find_entry("004006001028", lib)
        base = e["name"] if e else "容量10m3以内混凝土搅拌运输车运输混凝土"
        name = _dist_name(base, km)
        adj, note = _adj_yunju(name, km)
        rows.append(_quota("03", "004006001028", name,
                           e["unit"] if e else "100m3", _r(conc * coef / 100), adj, note))

    if pile_conc > 0:
        _banhe("07", pile_conc, 1.0)
        _jiaoban(pile_conc, 1.0)
        # ---- 钻渣挖装 + 15t 运土 + 泥浆处置费（模板口径：按桩基混凝土量）----
        e = _find_entry("001001008005", lib)
        rows.append(_quota("01", "001001008005", e["name"] if e else "挖掘机挖装普通土",
                           e["unit"] if e else "1000m3天然密实方", _r(pile_conc / 1000)))
        e = _find_entry("001001010009", lib)
        base = e["name"] if e else "装载质量15t以内自卸汽车运土"
        name = _dist_name(base, tu_km)
        adj, note = _adj_yunju(name, tu_km)
        rows.append(_quota("03", "001001010009", name,
                           e["unit"] if e else "1000m3天然密实方",
                           _r(pile_conc / 1000), adj, note))
        # 泥浆处置费：同望计算项（定额库无此子目，人为添加，代号 S,DJ）。
        # 模板口径：工程量1=桩基混凝土量(m3)，工程量2=机械单价（文本，如"25.0"），
        # 定额调整情况="0.0，100.0"、说明="B-NJCZ"。桩基工程必须保留此项。
        rows.append(_quota("12", "S", "泥浆处置费", "m3", _r(pile_conc),
                           "0.0，100.0", "B-NJCZ", q2=f"{float(nijiang_dj):.1f}"))

    # ---- QL0104 钢围堰 ----
    if weiyan_t:
        _sec(rows, "mu", "QL0104", "钢围堰", "t", q1=_r(weiyan_t))
        e = _find_entry("004002002004", lib)
        rows.append(_quota("10-1", "004002002004",
                           e["name"] if e else "钢套箱围堰无底模",
                           e["unit"] if e else "10t钢套箱", _r(weiyan_t / 10)))

    # ---- QL0105 承台(系梁) ----
    if chengtai > 0:
        _sec(rows, "mu", "QL0105", "承台(系梁)", "m3", q1=_r(chengtai + dianceng))
        e = _find_entry("004002022002", lib)
        rows.append(_quota("07", "004002022002",
                           e["name"] if e else "混凝土输送泵混凝土承台",
                           e["unit"] if e else "10m3实体", _r(chengtai / 10)))
        if dianceng > 0:
            e = _find_entry("[部2018预]41105006", lib)
            rows.append(_quota("07", "[部2018预]41105006",
                               e["name"] if e else "基础混凝土垫层",
                               e["unit"] if e else "10m3实体", _r(dianceng / 10),
                               "11,2,1503030,0,1503032,10.2,0",
                               "普C10-32.5-4 换 普C20-32.5-4"))
        if chengtai_hrb > 0:
            e = _find_entry("004005001014", lib)
            base = e["name"] if e else "集中加工现浇混凝土承台钢筋"
            name = f"{base} HRB400"
            rows.append(_quota("10-1", "004005001014", name,
                               e["unit"] if e else "1t钢筋", _r(chengtai_hrb),
                               f"12,{name},0"))
            _gang_yun(rows, lib, chengtai_hrb, gang_km)
        _banhe("07", chengtai + dianceng, 1.02)
        _jiaoban(chengtai + dianceng, 1.02)

    # ---- QL0107 土方工程 ----
    if wafang > 0:
        _sec(rows, "mu", "QL0107", "土方工程", "m3", q1=_r(wafang))
        e = _find_entry("004002003007", lib)
        rows.append(_quota("07", "004002003007",
                           e["name"] if e else "机械开挖基坑挖掘机挖土方",
                           e["unit"] if e else "1000m3", _r(wafang / 1000)))
        yutu = max(0.0, wafang - tianfang * yashi)
        if yutu > 0:
            e = _find_entry("001001010013", lib)
            base = e["name"] if e else "装载质量30t以内自卸汽车运土"
            name = _dist_name(base, tu_km)
            adj, note = _adj_yunju(name, tu_km)
            rows.append(_quota("03", "001001010013", name,
                               e["unit"] if e else "1000m3天然密实方",
                               _r(yutu / 1000), adj, note))
        if tianfang > 0:
            skipped.append(f"填方 {_r(tianfang)}m3：按利用回填处理（余土外运已扣除"
                           f"填方×{_r(yashi)} 换算），模板未单列填方子目")

    # ================= QL02 下部构造 =================
    if dun_conc or tai_conc:
        _sec(rows, "xiang", "QL02", "下部构造", "m3", q1=_r(dun_conc + tai_conc))
    if dun_conc > 0:
        _sec(rows, "mu", "QL0201", "桥墩", "m3", q1=_r(dun_conc))
        e = _find_entry("004003005003", lib)
        rows.append(_quota("07", "004003005003",
                           e["name"] if e else "梁板桥桥墩实体式混凝土墩高20m以内",
                           e["unit"] if e else "10m3实体", _r(dun_conc / 10)))
        if dun_hrb > 0:
            e = _find_entry("004005001016", lib)
            base = e["name"] if e else "集中加工现浇混凝土墩台钢筋"
            name = f"{base} HRB400"
            rows.append(_quota("10-1", "004005001016", name,
                               e["unit"] if e else "1t钢筋", _r(dun_hrb),
                               f"12,{name},0"))
            _gang_yun(rows, lib, dun_hrb, gang_km)
        _banhe("07", dun_conc, 1.02)
        _jiaoban(dun_conc, 1.02)
    if tai_conc > 0:
        _sec(rows, "mu", "QL0202", "桥台", "m3", q1=_r(tai_conc))
        e = _find_entry("004003002009", lib)
        rows.append(_quota("07", "004003002009",
                           e["name"] if e else "梁板桥肋型埋置式混凝土桥台高度8m以内",
                           e["unit"] if e else "10m3实体", _r(tai_conc / 10)))
        if tai_hrb > 0:
            e = _find_entry("004005001016", lib)
            base = e["name"] if e else "集中加工现浇混凝土墩台钢筋"
            name = f"{base} HRB400"
            rows.append(_quota("10-1", "004005001016", name,
                               e["unit"] if e else "1t钢筋", _r(tai_hrb),
                               f"12,{name},0"))
            _gang_yun(rows, lib, tai_hrb, gang_km)
        _banhe("07", tai_conc, 1.02)
        _jiaoban(tai_conc, 1.02)

    # ================= QL03 上部构造 =================
    if xl_conc or gxq345:
        _sec(rows, "xiang", "QL03", "上部构造", "m2", q1=_r(area) if area else "")
    if xl_conc > 0:
        _sec(rows, "mu", "QL0304", "预应力混凝土小箱梁", "m3", q1=_r(xl_conc))
        e = _find_entry("004004005004", lib)
        rows.append(_quota("07", "004004005004",
                           e["name"] if e else "预制、安装预应力混凝土等截面连续箱梁上部构造泵送",
                           e["unit"] if e else "10m3实体", _r(xl_conc / 10)))
        for t, q in (("HPB300", xl_hpb), ("HRB400", xl_hrb)):
            if q <= 0:
                continue
            e = _find_entry("004005001031", lib)
            base = e["name"] if e else "集中加工预制混凝土梁钢筋"
            name = f"{base} {t}"
            rows.append(_quota("10-1", "004005001031", name,
                               e["unit"] if e else "1t钢筋", _r(q),
                               f"[钢筋抽换#{t}钢筋],12,{name},0",
                               _GANG_NOTES.get(t, "")))
        if xl_hpb + xl_hrb > 0:
            _gang_yun(rows, lib, xl_hpb + xl_hrb, gang_km)
        if gjx_t > 0:
            _gangjiaoxian(rows, lib, gjx_t, items, p, warnings)
        _banhe("06", xl_conc, 1.02)
        _jiaoban(xl_conc, 1.02)
    if gxq345 > 0:
        _sec(rows, "mu", "QL0305", "钢箱梁", "t", q1=_r(gxq345))
        e = _find_entry("004004025003", lib)
        rows.append(_quota("10-1", "004004025003",
                           e["name"] if e else "钢箱梁汽车起重机吊装",
                           e["unit"] if e else "10t", _r(gxq345 / 10)))

    # ================= QL04 桥面铺装 =================
    if sma or hmac or fs_conc or uhpc or fangshui:
        _sec(rows, "xiang", "QL04", "桥面铺装", "m2", q1=_r(area) if area else "")
    if sma or hmac:
        _sec(rows, "mu", "QL0401", "沥青混凝土铺装", "m3", q1=_r(sma + hmac))
        if sma > 0:
            e = _find_entry("002002012081", lib)
            rows.append(_quota("04", "002002012081",
                               e["name"] if e else "拌和及铺筑沥青玛蹄脂碎石混合料路面",
                               e["unit"] if e else "1000m3路面实体", _r(sma / 1000),
                               "11,3,5505017,0,5505024,1236.02,3001002,0,0",
                               "[5505017] 换 [5505024]\n删:[3001002]"))
        if hmac > 0:
            e = _find_entry("002002012071", lib)
            rows.append(_quota("04", "002002012071",
                               e["name"] if e else "拌和及铺筑中粒式改性沥青混凝土路面",
                               e["unit"] if e else "1000m3路面实体", _r(hmac / 1000)))
        e = _find_entry("002002012096", lib)
        base = e["name"] if e else "装载质量20t以内自卸汽车运沥青混合料"
        name = _dist_name(base, liq_km)
        adj, note = _adj_yunju(name, liq_km)
        rows.append(_quota("03", "002002012096", name,
                           e["unit"] if e else "1000m3路面实体",
                           _r((sma + hmac) / 1000), adj, note))
    if nianceng > 0:
        e = _find_entry("002002014007", lib)
        rows.append(_quota("04", "002002014007",
                           e["name"] if e else "改性乳化沥青沥青层黏层",
                           e["unit"] if e else "1000m2", _r(nianceng / 1000)))
    if fs_conc or uhpc:
        _sec(rows, "mu", "QL0402", "水泥混凝土铺装", "m3", q1=_r(fs_conc + uhpc))
        if fs_conc > 0:
            e = _find_entry("004004013004", lib)
            rows.append(_quota("04", "004004013004",
                               e["name"] if e else "行车道桥面铺装防水混凝土垫层",
                               e["unit"] if e else "10m3实体", _r(fs_conc / 10)))
        if uhpc > 0:
            e = _find_entry("[部2018预]40613003", lib)
            rows.append(_quota("04", "[部2018预]40613003",
                               e["name"] if e else "桥面行车道铺装水泥混凝土面层泵送",
                               e["unit"] if e else "10m3", _r(uhpc / 10)))
        e = _find_entry("004005001011", lib)
        base = e["name"] if e else "现场加工桥面铺装水泥混凝土钢筋"
        if pm_crb > 0:
            rows.append(_quota("10-1", "004005001011", base,
                               e["unit"] if e else "1t钢筋", _r(pm_crb),
                               "11,2,2001001,0,2001003,1.025,0",
                               "[2001001] 换 [2001003]"))
        if pm_hpb > 0:
            name = f"{base} HPB300"
            rows.append(_quota("10-1", "004005001011", name,
                               e["unit"] if e else "1t钢筋", _r(pm_hpb),
                               f"12,{name},0"))
        if pm_hrb > 0:
            name = f"{base} HRB400"
            rows.append(_quota("10-1", "004005001011", name,
                               e["unit"] if e else "1t钢筋", _r(pm_hrb),
                               f"[钢筋抽换#钢筋抽换],12,{name},0",
                               _GANG_NOTES.get("HRB400", "")))
        e = _find_entry("004004013008", lib)
        rows.append(_quota("04", "004004013008",
                           e["name"] if e else "行车道桥面混凝土表面抛丸处理",
                           e["unit"] if e else "1000m2", _r((fs_conc + uhpc) / 100)))
        _banhe("06", fs_conc + uhpc, 1.02)
        _jiaoban(fs_conc + uhpc, 1.02)
    if fangshui > 0:
        _sec(rows, "mu", "QL0404", "桥面防水层", "m2", q1=_r(fangshui))
        e = _find_entry("004004013009", lib)
        rows.append(_quota("04", "004004013009",
                           e["name"] if e else "桥面防水剂",
                           e["unit"] if e else "1000m2", _r(fangshui / 1000)))

    # ================= QL05 桥梁附属结构 =================
    bearings = extra["bearings"]
    if bearings or ss_80 or ss_160 or hl_conc or dab_conc or llk or fpw_len:
        _sec(rows, "xiang", "QL05", "桥梁附属结构", "m2", q1=_r(area) if area else "")

    # ---- QL0501 支座 ----
    def _bearing_dm3(spec, n):
        m = re.match(r"(GBZYH?|LRB)[^\d]*(\d+)x(\d+)", spec)
        if not m:
            return None
        d, h = float(m.group(2)) / 100, float(m.group(3)) / 100  # mm → dm
        return math.pi * (d / 2) ** 2 * h * n

    if bearings:
        ban = [(sp, n) for sp, n in bearings
               if sp.upper().startswith("GBZY") and not sp.upper().startswith("GBZYH")]
        sifu = [(sp, n) for sp, n in bearings if sp.upper().startswith("GBZYH")]
        lrb = [(sp, n) for sp, n in bearings if sp.upper().startswith("LRB")]
        total_n = sum(n for _, n in bearings)
        _sec(rows, "mu", "QL0501", "桥梁支座", "个", q1=_r(total_n))
        if ban:
            _sec(rows, "ximu", "QL050101", "板式橡胶支座", "个",
                 q1=_r(sum(n for _, n in ban)))
            e1 = _find_entry("004004027001", lib)
            q_ban = sum(_bearing_dm3(sp, n) or 0 for sp, n in ban)
            rows.append(_quota("10-1", "004004027001",
                               e1["name"] if e1 else "安装板式橡胶支座",
                               e1["unit"] if e1 else "1dm3", _r(q_ban)))
            if sifu:
                e2 = _find_entry("004004027002", lib)
                q_sifu = sum(_bearing_dm3(sp, n) or 0 for sp, n in sifu)
                rows.append(_quota("10-1", "004004027002",
                                   e2["name"] if e2 else "安装四氟板式橡胶组合支座",
                                   e2["unit"] if e2 else "1dm3", _r(q_sifu)))
        elif sifu:
            _sec(rows, "ximu", "QL050101", "板式橡胶支座", "个",
                 q1=_r(sum(n for _, n in sifu)))
            e2 = _find_entry("004004027002", lib)
            q_sifu = sum(_bearing_dm3(sp, n) or 0 for sp, n in sifu)
            rows.append(_quota("10-1", "004004027002",
                               e2["name"] if e2 else "安装四氟板式橡胶组合支座",
                               e2["unit"] if e2 else "1dm3", _r(q_sifu)))
        if lrb:
            _sec(rows, "ximu", "QL050103", "圆形铅芯隔震橡胶支座", "个",
                 q1=_r(sum(n for _, n in lrb)))
            e = _find_entry("004004027007", lib)
            for sp, n in sorted(lrb, key=lambda x: -x[1]):
                rows.append(_quota("10-1", "004004027007",
                                   e["name"] if e else "安装盆式橡胶支座反力10000kN",
                                   e["unit"] if e else "1个", _r(n)))
    elif any((it.get("fenxiang") or "").strip() == "支座" for it in items):
        skipped.append("支座：数量表中未解析到规格×数量，请检查支座块格式")

    # ---- QL0502 伸缩缝 ----
    if ss_80 or ss_160 or ss_conc:
        ss_total = ss_80 + ss_160
        _sec(rows, "mu", "QL0502", "伸缩缝", "m", q1=_r(ss_total))
        _sec(rows, "ximu", "QL050201", "模数式伸缩缝", "m", q1=_r(ss_total))
        e = _find_entry("004004027034", lib)
        SS_ADJ = {"80": ("11,2,6003004,0,6003001,1.0,0", "[6003004] 换 [6003001]"),
                  "160": ("11,2,6003004,0,6003003,1.0,0", "[6003004] 换 [6003003]")}
        for xing, q in (("80", ss_80), ("160", ss_160)):
            if q <= 0:
                continue
            adj, note = SS_ADJ.get(xing, ("", ""))
            if xing not in SS_ADJ:
                note = f"{xing}型伸缩缝：请核对伸缩量档及材料换算式"
            rows.append(_quota("10-1", "004004027034",
                               e["name"] if e else "安装模数式伸缩缝伸缩量480mm以内",
                               e["unit"] if e else "1m", _r(q), adj, note))
        if ss_conc > 0:
            e = _find_entry("004004027038", lib)
            rows.append(_quota("10-1", "004004027038",
                               e["name"] if e else "预留槽混凝土",
                               e["unit"] if e else "10m3", _r(ss_conc / 10)))
        if ss_hrb_kg > 0:
            e = _find_entry("004004027039", lib)
            base = e["name"] if e else "预留槽钢筋"
            name = f"{base} HRB400"
            rows.append(_quota("10-1", "004004027039", name,
                               e["unit"] if e else "1t", _r(ss_hrb_kg / 1000),
                               f"[钢筋抽换#HRB400钢筋],12,{name},0",
                               _GANG_NOTES.get("HRB400", "")))
        if ss_conc > 0:
            _banhe("06", ss_conc, 1.02)
            _jiaoban(ss_conc, 1.02)
        _sum, _n = _sum_by(items, "伸缩缝", "C型钢")
        if _n:
            skipped.append(f"伸缩缝 C型钢 {_r(_sum)}kg、钢纤维等：已含在伸缩缝安装子目中，不另套")

    # ---- QL0503 护栏与护网 ----
    if hl_conc > 0:
        _sec(rows, "mu", "QL0503", "护栏与护网", "m",
             q1=_r(qiaochang) if qiaochang else "")
        _sec(rows, "ximu", "QL050304", "桥梁混凝土防撞护栏", "m3", q1=_r(hl_conc))
        e = _find_entry("005001001004", lib)
        rows.append(_quota("06", "005001001004",
                           e["name"] if e else "现浇钢筋混凝土防撞护栏混凝土",
                           e["unit"] if e else "10m3实体", _r(hl_conc / 10)))
        if hl_hrb > 0:
            e = _find_entry("005001001005", lib)
            base = e["name"] if e else "现浇钢筋混凝土防撞护栏钢筋"
            name = f"{base} HRB400"
            rows.append(_quota("10", "005001001005", name,
                               e["unit"] if e else "1t", _r(hl_hrb),
                               f"[钢筋抽换#钢筋抽换],12,{name},0",
                               _GANG_NOTES.get("HRB400", "")))
        if hl_q235 > 0:
            e = _find_entry("005001001006", lib)
            rows.append(_quota("06", "005001001006",
                               e["name"] if e else "现浇钢筋混凝土防撞护栏制作、安装铸铁柱及栏杆",
                               e["unit"] if e else "1t", _r(hl_q235)))

    # ---- QL0504 搭板 ----
    if dab_conc > 0:
        _sec(rows, "mu", "QL0504", "搭板", "m3", q1=_r(dab_conc))
        e = _find_entry("004004013010", lib)
        rows.append(_quota("04", "004004013010",
                           e["name"] if e else "现浇桥头搭板混凝土",
                           e["unit"] if e else "10m3", _r(dab_conc / 10)))
        if dab_hrb > 0:
            e = _find_entry("004005001021", lib)
            rows.append(_quota("10-1", "004005001021",
                               e["name"] if e else "集中加工桥头搭板钢筋",
                               e["unit"] if e else "1t钢筋", _r(dab_hrb)))
            _gang_yun(rows, lib, dab_hrb, gang_km)
        _banhe("06", dab_conc, 1.0)
        _jiaoban(dab_conc, 1.0)

    # ---- QL0505 锥坡 ----
    if llk > 0:
        _sec(rows, "mu", "QL0505", "锥坡", "m3", q1=_r(llk))
        e = _find_entry("001004003002", lib)
        rows.append(_quota("06", "001004003002",
                           e["name"] if e else "预制、铺砌混凝土预制块护坡",
                           e["unit"] if e else "10m3", _r(llk / 10)))
        _sum, _n = _sum_by(items, "砂砾垫层", "砂砾垫层")
        if _n:
            skipped.append(f"砂砾垫层 {_r(_sum)}m3：模板未单列子目，请核对是否并入锥坡")

    # ---- QL0506 防抛网 ----
    if fpw_len > 0:
        _sec(rows, "mu", "QL0506", "防抛网", "m", q1=_r(fpw_len))
        e = _find_entry("005001005006", lib)
        rows.append(_quota("14", "005001005006",
                           e["name"] if e else "防落(物)网(钢管立柱)",
                           e["unit"] if e else "100m", _r(fpw_len / 100)))
        skipped.append("防抛网钢材/螺栓/焊网等材料：已含在防落(物)网子目中，不另套")

    # ================= QL06 其他工程 =================
    _sec(rows, "xiang", "QL06", "其他工程", "m", q1=_r(area) if area else "",
         code_cell="SX,DJ", adj="N", note="QL06")
    skipped.append("拆除/改造/排水系统/声测管/波纹管/锚具/防腐面积等：模板（QL06 不计费）未套定额，"
                   "如需计费请另行补充子目")

    return {"rows": rows, "warnings": warnings, "skipped": skipped,
            "title": title, "extra": extra}


def _gang_yun(rows, lib, gang_t, km):
    """钢材运输（承台/墩台/梁/搭板钢筋随同）。"""
    e = _find_entry("[部2018预]90105071", lib)
    base = e["name"] if e else "装载质量20t以内载货汽车运输钢材"
    name = _dist_name(base, km)
    adj, note = _adj_yunju(name, km)
    rows.append(_quota("03", "[部2018预]90105071", name,
                       e["unit"] if e else "100t", _r(gang_t / 100), adj, note))


def _gangjiaoxian(rows, lib, gjx_t, items, p, warnings):
    """钢绞线：按束长档+孔数档选子目（后张法 0040050020xx）。"""
    kong = None
    for it in items:
        m = re.search(r"15[-](\d{1,2})", str(it.get("mat") or ""))
        if m and (it.get("fenxiang") or "").strip() in ("GQY锚具", "锚具"):
            kong = int(m.group(1))
            break
    if kong is None:
        kong = 7
    shu = float(p.get("gjx_shu") or 40)
    cands = []
    lib_entries = lib.get("entries", lib) if isinstance(lib, dict) else {}
    for code, e in lib_entries.items():
        if not code.startswith("0040050020"):
            continue
        m = re.search(r"束长(\d+)m以内(\d+)孔", e["name"])
        if m and "后张法" in e["name"]:
            cands.append((int(m.group(1)), int(m.group(2)), code, e))
    cands.sort(key=lambda x: (x[0], x[1]))
    pick = None
    for cl, hl, code, e in cands:
        if cl >= shu and hl >= kong:
            pick = (cl, hl, code, e)
            break
    if not pick and cands:
        pick = cands[-1]
    if pick:
        _, _, code, e = pick
        name, unit = e["name"], e["unit"]
    else:
        code = "004005002017"
        e = _find_entry(code, lib)
        name = e["name"] if e else "后张法制作、张拉预应力钢绞线束长40m以内7孔每t3.82束"
        unit = e["unit"] if e else "1t钢绞线"
        warnings.append(f"库内未找到束长{_r(shu)}m/孔数{kong}档钢绞线子目，"
                        f"按 004005002017（40m/7孔）兜底，请核对")
    rows.append(_quota("10-1", code, name, unit, _r(gjx_t)))


# ============================ 导出 ============================

def build_road_excel(result, params=None):
    """导出 .xlsx（模板 15 列结构：项/目/节/细目代号+费率编号+定额代号+名称+单位+
    工程量1+工程量2+定额调整情况+说明+代号+默认项）。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "第一部分 建筑安装工程费"
    ws.append(ROAD_HEADERS)
    for row in result["rows"]:
        if row["t"] == "quota":
            ws.append(["", "", "", "", row["rate"], row["code"], row["name"],
                       row["unit"], row["q1"], row["q2"], row["adj"], row["note"],
                       "", ""])
        else:
            ws.append([row["d"] if row["t"] == "xiang" else "",
                       row["d"] if row["t"] == "mu" else "",
                       row["d"] if row["t"] == "jie" else "",
                       row["d"] if row["t"] == "ximu" else "",
                       "", row["code"], row["name"], row["unit"],
                       row["q1"], row["q2"], row["adj"], row["note"], "", ""])
    # 样式
    head_fill = PatternFill("solid", fgColor="D9E1F2")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = head_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    widths = [10, 10, 12, 12, 8, 20, 52, 13, 12, 12, 46, 34, 6, 6]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for r in range(2, ws.max_row + 1):
        for c in range(1, 15):
            cell = ws.cell(row=r, column=c)
            cell.border = border
            if c == 6:
                cell.alignment = Alignment(horizontal="left")
    ws.freeze_panes = "A2"
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()
