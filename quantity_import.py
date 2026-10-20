# -*- coding: utf-8 -*-
"""表格预处理：解析「主要工程数量总表.xls」→ 填充「表格.xlsx」信息输入表。

纯逻辑模块（不依赖 Streamlit），app.py 的「表格预处理」页面上传数量总表后调用：
- parse_xls(data: bytes) -> dict            解析 .xls，得到各部位工程量
- mapping_rows(m) -> list[(单元格,项目,值)]    生成核对预览行
- fill_template(tpl: bytes, m: dict) -> bytes  写入信息输入表，返回新 xlsx 字节

映射规则（用户 2026-10-08/09、2026-10-14 拍板）：
- 桩基础混凝土按桩径分行（模板 B 列 1/1.2/1.5/1.8/2），模板没有的桩径提示不导入；
- 数值保留 2 位小数；工程数量表中没有的（承台HPB/各部位HPB等）留空；
- 由多部分组成的单元格填 Excel 公式（如 =80325+6144、=36*60+24*70），
  明细保留 2 位，方便用户点开核对来源；
- 承台 HRB = 工程数量表口径：桥台承台 80325 + 梯道承台 6144（弃用下部 sheet 旧版 76500）；
- 声测管 = 下部 sheet「超声波检测管」钢管φ57x3.5 + φ70x6.5 之和；
- 数量总表里模板没有的部位 → 在信息输入表末尾追加 十一/十二… 新分项；
- 上部结构（模板 2026-10 更新版 七~十 节）：现浇箱梁/小箱梁/板梁按 r6 组名列取
  混凝土（全等级和）、HPB300、HRB400、钢绞线、锚具、波纹管；桥面铺装取细粒式/
  中粒式/C50防水混凝土/钢筋（HPB+HRB+冷轧带肋）/防水层；防撞护栏、搭板取混凝土
  +钢筋；伸缩缝/支座/钢管栏杆行读「材料统计」表合计行（支座细分板式/四氟时
  分别填 D67/D68，未细分时全量入 D67；钢管栏杆行数值与单位随数量总表，如
  人行栏杆 412m → D66=412、C66=m）；
  挖方/填方取数量总表单值行；梯道等无对应分项的上部组→在表格末尾追加为新分项。
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


def _group_cols(sh) -> tuple[dict, dict]:
    """r7 组名（取第一次出现）→ (起始列, 结束列, 子列标签→列)。
    返回 (下部组, 全部组)：下部组含桥台/桥墩/梯道下部（缺失报错）；
    全部组供上部结构（现浇箱梁/小箱梁/板梁/桥面系/梯道等）定位列。"""
    r7 = 6
    bounds = sorted(c for c in range(sh.ncols)
                    if _norm(sh.cell_value(r7, c)))
    first = {}
    for c in bounds:
        first.setdefault(_norm(sh.cell_value(r7, c)), c)
    all_groups = {}
    for g, s in first.items():
        e = next((b for b in bounds if b > s), sh.ncols)
        cols = {}
        for c in range(s, e):
            lab = _norm(sh.cell_value(7, c))
            if lab:
                cols[lab] = c
        all_groups[g] = (s, e, cols)
    groups = {}
    for g in ("桥台", "桥墩", "梯道下部"):
        if g not in all_groups:
            raise QuantityImportError(f"数量总表表头缺少「{g}」组")
        groups[g] = all_groups[g]
    return groups, all_groups


# 已被固定映射消费的 (组, 子列)；其余子列若带数据 → 追加为新分项
_CONSUMED = {("桥台", "台帽"), ("桥台", "台身"), ("桥台", "承台"), ("桥台", "桩基"),
             ("桥墩", "盖梁"), ("桥墩", "墩柱及柱系梁"), ("桥墩", "桩基"),
             ("梯道下部", "桥墩"), ("梯道下部", "承台"), ("梯道下部", "桩基")}


def _extra_cols(sh, groups) -> list[tuple[str, int]]:
    """组区间内未被消费的子列 → [(名称, 列号)]（搭板/防撞护栏已入固定映射）。"""
    out = []
    for g, (s, e, cols) in groups.items():
        for lab, c in cols.items():
            if lab == "小计" or (g, lab) in _CONSUMED:
                continue
            out.append((f"{g}·{lab}", c))
    return out


def _sum_parts(sh, mat: str, sub, col: int) -> list:
    """按 B 列材料大类 + C 列细分收集 col 列的非零明细行（供公式 =a+b+c）。
    sub 指定时只取该细分行；sub=None 时混凝土取全部等级行（除小计），
    其他大类（钢绞线/锚具/波纹管/桥面防水层等）取该大类全部行。"""
    out, cur = [], ""
    for r in range(5, sh.nrows):
        b = str(sh.cell_value(r, 1)).strip()
        if b:
            cur = b
        if cur != mat:
            continue
        c = str(sh.cell_value(r, 2)).strip()
        if _norm(c) == "小计":
            continue
        if sub is not None:
            if _norm(c) != _norm(sub):
                continue
        elif mat == "混凝土" and not c:
            continue
        v = _num(sh.cell_value(r, col))
        if v:
            out.append(v)
    return out


def _sum_by_mat(sh, mat: str, sub, col: int) -> float:
    """_sum_parts 的合计。"""
    return sum(_sum_parts(sh, mat, sub, col))


def _label_row_val(sh, label: str):
    """B 列大类 == label 的行，取该行首个非零数值单元格（挖方/填方等单值行）。"""
    for r in range(5, sh.nrows):
        if _norm(sh.cell_value(r, 1)) == label:
            for c in range(4, sh.ncols):
                v = sh.cell_value(r, c)
                if isinstance(v, (int, float)) and v:
                    return float(v)
            return None
    return None


def _tongji_item(sh, fname: str):
    """材料统计表分项名 fname 的首个「合计」行 → (合计值, 材料名, 单位)。"""
    for r in range(sh.nrows):
        if str(sh.cell_value(r, 0)).strip() == fname \
                and "合" in str(sh.cell_value(r, 1)):
            s = sum(_num(sh.cell_value(r, c)) for c in range(4, sh.ncols))
            val = round(s, 2) if s else None
            lab, unit = fname, ""
            for rr in range(r - 1, max(-1, r - 5), -1):
                if str(sh.cell_value(rr, 0)).strip() == "分项名":
                    lab = str(sh.cell_value(rr, 4)).strip()
                    unit = str(sh.cell_value(rr + 1, 4)).strip()
                    break
            return val, lab, unit
    return None, fname, ""


def _tongji_split(sh, fname: str, key1: str, key2: str):
    """材料统计表分项名 fname 的首个「合计」行按表头关键字拆分为两类：
    表头含 key2 的列 → 第二类；其余数值列（含表头含 key1 的列）并入第一类。
    未细分（表头无 key2 列）时第二类返回 None。
    返回 (第一类值, 第二类值, 第一类明细列, 第二类明细列)。"""
    for r in range(sh.nrows):
        if str(sh.cell_value(r, 0)).strip() != fname \
                or "合" not in str(sh.cell_value(r, 1)):
            continue
        hdr = None
        for rr in range(r - 1, max(-1, r - 5), -1):
            if str(sh.cell_value(rr, 0)).strip() == "分项名":
                hdr = rr
                break
        v1 = v2 = 0.0
        p1, p2 = [], []
        matched2 = False
        for c in range(4, sh.ncols):
            v = _num(sh.cell_value(r, c))
            if not v:
                continue
            lab = _norm(str(sh.cell_value(hdr, c))) if hdr is not None else ""
            if key2 in lab:
                v2 += v
                p2.append(round(v, 2))
                matched2 = True
            else:
                v1 += v
                p1.append(round(v, 2))
        return (round(v1, 2) or None), \
            ((round(v2, 2) or None) if matched2 else None), \
            p1, (p2 if matched2 else [])
    return None, None, [], []


def _tongji_parts(sh, fname: str) -> list:
    """材料统计表分项名 fname 首个「合计」行的非零数值列（公式明细）。"""
    for r in range(sh.nrows):
        if str(sh.cell_value(r, 0)).strip() == fname \
                and "合" in str(sh.cell_value(r, 1)):
            return [round(_num(sh.cell_value(r, c)), 2)
                    for c in range(4, sh.ncols)
                    if _num(sh.cell_value(r, c))]
    return []


def parse_xls(data: bytes) -> dict:
    """解析 .xls 数量总表，返回映射结果（桩列表/各部位工程量/新增分项/警告）。"""
    try:
        wb = xlrd.open_workbook(file_contents=data)
    except Exception as e:
        raise QuantityImportError(f"无法打开 .xls 文件：{e}")
    if "工程数量表" not in wb.sheet_names() or "下部" not in wb.sheet_names():
        raise QuantityImportError("缺少「工程数量表」或「下部」工作表")
    sh = wb.sheet_by_name("工程数量表")
    groups, all_groups = _group_cols(sh)

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

    m = {"warnings": [], "extras": [], "parts": {}}

    def _put(key, parts):
        """m[key] = 合计（保留 2 位）；m["parts"][key] = 非零明细（供公式 =a+b）。"""
        nz = [round(_num(p), 2) for p in parts if _num(p)]
        m["parts"][key] = nz
        m[key] = round(sum(_num(p) for p in parts), 2) or None

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
                p = piles.setdefault(k, {"d": k, "n": 0, "len": 0.0, "parts": []})
                p["n"] += int(n)
                p["len"] += int(n) * float(L)
                p["parts"].append((int(n), float(L)))
            elif v0 and v0 != "桩基类型":
                break
    if not piles:
        raise QuantityImportError("「下部」表未找到桩基类型区")
    m["piles"] = [piles[k] for k in sorted(piles)]

    # ── 声测管 = 超声波检测管行的 钢管φ57x3.5 + φ70x6.5（可能分多段）──
    col57 = col70 = None
    v57 = v70 = 0.0
    for r in range(lo.nrows):
        for c in range(lo.ncols):
            v = str(lo.cell_value(r, c)).strip()
            if v.startswith("钢管") and "φ57" in v:
                col57 = c
            elif v.startswith("钢管") and "φ70" in v:
                col70 = c
        if str(lo.cell_value(r, 0)).strip() == "超声波检测管":
            if col57 is not None:
                v57 += _num(lo.cell_value(r, col57))
            if col70 is not None:
                v70 += _num(lo.cell_value(r, col70))
    _put("shengguan", [v57, v70])

    # ── 固定映射部位（工程数量表口径）；多部分来源 → 记录明细供公式核对 ──
    _put("pile_hpb", [val_all([hpb], "桥台", "桩基"), val_all([hpb], "桥墩", "桩基"),
                      val_all([hpb], "梯道下部", "桩基")])
    _put("pile_hrb", [val_all([hrb], "桥台", "桩基"), val_all([hrb], "桥墩", "桩基"),
                      val_all([hrb], "梯道下部", "桩基")])
    _put("chengtai_conc", [val_all(conc_rows, "桥台", "承台"),
                           val_all(conc_rows, "梯道下部", "承台")])
    _put("chengtai_hpb", [val_all([hpb], "桥台", "承台"),
                          val_all([hpb], "梯道下部", "承台")])
    _put("chengtai_hrb", [val_all([hrb], "桥台", "承台"),
                          val_all([hrb], "梯道下部", "承台")])
    _put("qiaotai_conc", [val_all(conc_rows, "桥台", "台帽"),
                          val_all(conc_rows, "桥台", "台身")])
    _put("qiaotai_hpb", [val_all([hpb], "桥台", "台帽"),
                         val_all([hpb], "桥台", "台身")])
    _put("qiaotai_hrb", [val_all([hrb], "桥台", "台帽"),
                         val_all([hrb], "桥台", "台身")])
    _put("qiaodun_conc", [val_all(conc_rows, "桥墩", "墩柱及柱系梁"),
                          val_all(conc_rows, "梯道下部", "桥墩")])
    _put("qiaodun_hpb", [val_all([hpb], "桥墩", "墩柱及柱系梁"),
                         val_all([hpb], "梯道下部", "桥墩")])
    _put("qiaodun_hrb", [val_all([hrb], "桥墩", "墩柱及柱系梁"),
                         val_all([hrb], "梯道下部", "桥墩")])
    _put("dungailiang_conc", [val_all(conc_rows, "桥墩", "盖梁")])
    _put("dungailiang_hpb", [val_all([hpb], "桥墩", "盖梁")])
    _put("dungailiang_hrb", [val_all([hrb], "桥墩", "盖梁")])

    # ── 上部结构（模板 2026-10 更新版 七~十 节）──
    def _up_col(*names):
        """组名包含任一关键字的上部组列号；找不到返回 None。"""
        for g, (s, _e, _cols) in all_groups.items():
            if any(k in g for k in names):
                return s
        return None

    def _sec(key, col):
        """箱梁类节：混凝土(全等级和)/HPB300/HRB400/钢绞线/锚具/波纹管。"""
        if col is None:
            for suf in ("conc", "hpb", "hrb", "gsx", "maoju", "bwg"):
                m[f"{key}_{suf}"] = None
            return
        _put(f"{key}_conc", _sum_parts(sh, "混凝土", None, col))
        _put(f"{key}_hpb", [_sum_by_mat(sh, "钢材", "HPB300", col)])
        _put(f"{key}_hrb", [_sum_by_mat(sh, "钢材", "HRB400", col)])
        _put(f"{key}_gsx", _sum_parts(sh, "钢绞线", None, col))
        _put(f"{key}_maoju", _sum_parts(sh, "锚具", None, col))
        _put(f"{key}_bwg", _sum_parts(sh, "波纹管", None, col))

    _sec("xianjiao", _up_col("现浇箱梁"))
    _sec("xiaoxiangliang", _up_col("小箱梁"))
    _sec("banliang", _up_col("板梁"))

    # 桥面系：桥面铺装 + 防撞护栏
    qmx = all_groups.get("桥面系")
    if qmx:
        c = qmx[2].get("桥面铺装")
        if c is not None:
            _put("qiaomian_xili",
                 [_sum_by_mat(sh, "混凝土", "细粒式沥青混凝土", c)])
            _put("qiaomian_zhongli",
                 [_sum_by_mat(sh, "混凝土", "中粒式沥青混凝土", c)])
            _put("qiaomian_conc",
                 [_sum_by_mat(sh, "混凝土", "C50防水混凝土", c)])
            _put("qiaomian_rebar",
                 [_sum_by_mat(sh, "钢材", "冷轧带肋钢筋CRB550", c),
                  _sum_by_mat(sh, "钢材", "HPB300", c),
                  _sum_by_mat(sh, "钢材", "HRB400", c)])
            _put("qiaomian_fangshui", _sum_parts(sh, "桥面防水层", None, c))
        c = qmx[2].get("防撞护栏")
        if c is not None:
            _put("hulan_conc", _sum_parts(sh, "混凝土", None, c))
            _put("hulan_rebar", [_sum_by_mat(sh, "钢材", "HPB300", c),
                                 _sum_by_mat(sh, "钢材", "HRB400", c)])

    # 搭板：r5 组标签、r8 无子列标签的独立列
    db = None
    for c in range(sh.ncols):
        if _norm(sh.cell_value(4, c)) == "搭板" and not _norm(sh.cell_value(7, c)):
            db = c
    if db is not None:
        _put("daban_conc", _sum_parts(sh, "混凝土", None, db))
        _put("daban_rebar", [_sum_by_mat(sh, "钢材", "HPB300", db),
                             _sum_by_mat(sh, "钢材", "HRB400", db)])

    # 附属：挖方/填方（单值行）；伸缩缝/支座/钢管栏杆行读材料统计合计行
    # （钢管栏杆行数值与单位都随数量总表，如人行栏杆 412m → D66=412、C66=m；
    #   支座合计行细分板式/四氟时分别填 D67/D68，未细分时全量入 D67）
    _put("wafang", [_label_row_val(sh, "挖方")])
    _put("tianfang", [_label_row_val(sh, "填方")])
    if "材料统计" in wb.sheet_names():
        ts = wb.sheet_by_name("材料统计")
        _put("shensuofeng", _tongji_parts(ts, "伸缩缝"))
        zz1, zz2, zp1, zp2 = _tongji_split(ts, "支座", "板式", "四氟")
        _put("zhizuo_banshi", zp1)
        _put("zhizuo_sifu", zp2)
        qt, _qlab, qunit = _tongji_item(ts, "其他")
        if qt:
            _put("gangguan_langan", _tongji_parts(ts, "其他"))
            m["gangguan_langan_unit"] = qunit

    # ── 新增分项：未被固定映射消费且有数据的子列 ──
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

    # ── 模板无对应分项的上部组（如梯道）→ 在信息输入表末尾追加为新分项 ──
    handled = {_up_col("现浇箱梁"), _up_col("小箱梁"), _up_col("板梁")}
    if qmx:
        handled.add(qmx[0])
    for g, (s, _e, _cols) in all_groups.items():
        if s in handled or s >= groups["桥台"][0] or g in ("小计", "合计"):
            continue
        conc = _sum_by_mat(sh, "混凝土", None, s)
        hp = _sum_by_mat(sh, "钢材", "HPB300", s)
        hr = _sum_by_mat(sh, "钢材", "HRB400", s)
        if conc or hp or hr:
            m["extras"].append({
                "name": g,
                "conc": round(conc, 2),
                "hpb": round(hp, 2) if hp else None,
                "hrb": round(hr, 2) if hr else None,
            })
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
    # 上部结构（模板 2026-10 更新版新增 七~十 节）
    "D35": ("现浇箱梁 混凝土(m3)", "xianjiao_conc"),
    "D36": ("现浇箱梁 HPB 钢筋(kg)", "xianjiao_hpb"),
    "D37": ("现浇箱梁 HRB 钢筋(kg)", "xianjiao_hrb"),
    "D38": ("现浇箱梁 钢绞线(kg)", "xianjiao_gsx"),
    "D39": ("现浇箱梁 锚具(套)", "xianjiao_maoju"),
    "D40": ("现浇箱梁 波纹管(m)", "xianjiao_bwg"),
    "D42": ("小箱梁 混凝土(m3)", "xiaoxiangliang_conc"),
    "D43": ("小箱梁 HPB 钢筋(kg)", "xiaoxiangliang_hpb"),
    "D44": ("小箱梁 HRB 钢筋(kg)", "xiaoxiangliang_hrb"),
    "D45": ("小箱梁 钢绞线(kg)", "xiaoxiangliang_gsx"),
    "D46": ("小箱梁 锚具(套)", "xiaoxiangliang_maoju"),
    "D47": ("小箱梁 波纹管(m)", "xiaoxiangliang_bwg"),
    "D49": ("板梁 混凝土(m3)", "banliang_conc"),
    "D50": ("板梁 HPB 钢筋(kg)", "banliang_hpb"),
    "D51": ("板梁 HRB 钢筋(kg)", "banliang_hrb"),
    "D52": ("板梁 钢绞线(kg)", "banliang_gsx"),
    "D53": ("板梁 锚具(套)", "banliang_maoju"),
    "D54": ("板梁 波纹管(m)", "banliang_bwg"),
    "D56": ("桥面铺装 细粒式沥青混凝土(m3)", "qiaomian_xili"),
    "D57": ("桥面铺装 中粒式沥青混凝土(m3)", "qiaomian_zhongli"),
    "D58": ("桥面铺装 C50防水混凝土(m3)", "qiaomian_conc"),
    "D59": ("桥面铺装 钢筋(kg)", "qiaomian_rebar"),
    "D60": ("桥面防水层(m2)", "qiaomian_fangshui"),
    "D62": ("搭板 混凝土(m3)", "daban_conc"),
    "D63": ("搭板 钢筋(kg)", "daban_rebar"),
    "D64": ("防撞护栏 混凝土(m3)", "hulan_conc"),
    "D65": ("防撞护栏 钢筋(kg)", "hulan_rebar"),
    "D66": ("钢管栏杆(kg)", "gangguan_langan"),
    "D67": ("板式橡胶支座(个)", "zhizuo_banshi"),
    "D68": ("四氟板式橡胶支座(个)", "zhizuo_sifu"),
    "D69": ("伸缩缝(m)", "shensuofeng"),
    "D70": ("挖方(m3)", "wafang"),
    "D71": ("填方(m3)", "tianfang"),
}

_PILE_CELL = {1.0: 6, 1.2: 7, 1.5: 8, 1.8: 9, 2.0: 10}


def _fmt(p) -> str:
    """保留 2 位小数并去掉多余的 0（10857.34 / 97716.1 / 94）。"""
    s = f"{p:.2f}".rstrip("0").rstrip(".")
    return s or "0"


def _formula_of(m: dict, key: str, pile=None) -> str | None:
    """多部分来源 → '=a+b+…' 公式串（明细保留 2 位），单来源返回 None。"""
    if pile is not None:
        if len(pile.get("parts") or []) >= 2:
            return "=" + "+".join(f"{int(n)}*{_fmt(L)}" for n, L in pile["parts"])
        return None
    parts = (m.get("parts") or {}).get(key) or []
    if len(parts) >= 2:
        return "=" + "+".join(_fmt(p) for p in parts)
    return None


def mapping_rows(m: dict) -> list[tuple[str, str, object]]:
    """生成核对预览行 (单元格, 项目, 值)；多部分来源显示公式供核对。"""
    out = []
    for p in m.get("piles") or []:
        r = _PILE_CELL.get(float(p["d"]))
        if r is None:
            continue
        name = f"桩基础 φ{float(p['d']):g} 桩长(m)"
        out.append((f"D{r}", name, _formula_of(m, "", p) or round(float(p["len"]), 2)))
        out.append((f"F{r}", f"桩基础 φ{float(p['d']):g} 根数", int(p["n"])))
    for cell, (name, key) in _CELLS.items():
        v = m.get(key)
        if key == "gangguan_langan" and m.get("gangguan_langan_unit"):
            name = name.replace("(kg)", f"({m['gangguan_langan_unit']})")
        out.append((cell, name, _formula_of(m, key) or (round(v, 2) if v else None)))
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
        f = _formula_of(m, "", p)
        ws.cell(row=r, column=4).value = f or (
            int(length) if length.is_integer() else round(length, 2))
        ws.cell(row=r, column=6).value = int(p["n"])
    for d, r in pile_rows.items():
        if not any(float(p["d"]) == d for p in m.get("piles") or []):
            ws.cell(row=r, column=4).value = None
            ws.cell(row=r, column=6).value = None

    # ── 固定 D 列（没有数据 → 留空；多部分来源 → 公式 =a+b+…）──
    for cell, (_, key) in _CELLS.items():
        v = m.get(key)
        f = _formula_of(m, key)
        ws[cell] = f or (round(v, 2) if v else None)
    # 钢管栏杆行：单位随数量总表（模板默认 kg，数量总表为 m 时改写）
    if m.get("gangguan_langan_unit"):
        ws["C66"] = m["gangguan_langan_unit"]
    # 台盖梁：数量总表无来源 → 留空
    for cell in ("D23", "D24", "D25"):
        ws[cell] = None

    # ── 新增分项：追加在现有内容之后（样式照抄承台段 14-17 行）──
    start = 34
    for r in range(34, 80):
        if ws.cell(row=r, column=1).value not in (None, "") \
                or ws.cell(row=r, column=2).value not in (None, ""):
            start = r + 1
    # 分项序号接在模板已有分项之后（旧版模板 一~六 → 七 起；新版 一~十 → 十一 起）
    used = set()
    for r in range(5, start):
        v = str(ws.cell(row=r, column=1).value or "").strip()
        if v in _CN:
            used.add(_CN.index(v))
    base = (max(used) + 1) if used else 6
    for i, ex in enumerate(m.get("extras") or []):
        r = start
        ws.cell(row=r, column=1).value = _CN[base + i]
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
