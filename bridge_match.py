# -*- coding: utf-8 -*-
"""天津桥梁定额自动匹配（纯逻辑模块，不依赖 Streamlit）。

- parse_quota_library: 解析 天津桥梁.xlsx —— sheet1 定额库（编号/名称/单位/单价）
  + sheet2 套定额规则样例（分部/子分部 → 定额编号组合）。
- parse_quantity_excel: 解析海泰式数量总表（分项材料表：分项名/适用位置/备注/数量/材料列）。
- match_quotas: 规则驱动（分项→子分部→样例组合）+ 文本打分回退。
- ai_explain: DeepSeek 生成多候选的匹配理由/差异说明（仅解释，不改变入选结果）。
- build_result_excel: 输出结果工作簿。
"""
import io
import json
import re
import unicodedata

# ---------------------------------------------------------------------------
# 基础工具
# ---------------------------------------------------------------------------


class BridgeMatchError(Exception):
    """定额库/数量表格式不符合预期。"""


_UNIT_MAP = {
    "m3": "m3", "立方": "m3", "方": "m3", "m³": "m3",
    "m2": "m2", "平米": "m2", "平方米": "m2", "m²": "m2",
    "米": "m", "延米": "m", "m": "m",
    "吨": "t", "t": "t", "kg": "kg", "千克": "kg",
    "套": "套", "个": "个", "cm3": "cm3",
}
_COUNT_UNITS = {"套", "个", "只", "根", "块", "座", "处", "孔", "跨", "条", "道", "对"}


def _norm_text(s):
    """全角→半角、粘→黏、φ 变体统一、去全部空白。"""
    s = unicodedata.normalize("NFKC", str(s or ""))
    s = s.replace("粘", "黏").replace("ф", "φ").replace("Φ", "φ").replace("ϕ", "φ")
    s = re.sub(r"\s+", "", s)
    return s


def _norm_code(s):
    s = _norm_text(s).lower()
    s = re.sub(r"[—－]", "-", s)
    return s


def _unit_parts(unit):
    """定额单位 → (基类, 扩大倍数)："10m"→("m",10)、"100m2"→("m2",100)、"t"→("t",1)。"""
    u = _norm_text(unit).lower()
    u = u.replace("³", "3").replace("²", "2")
    m = re.match(r"^(\d+(?:\.\d+)?)\s*(.*)$", u)
    if m:
        return _UNIT_MAP.get(m.group(2), m.group(2)), float(m.group(1))
    return _UNIT_MAP.get(u, u), 1.0


# ---------------------------------------------------------------------------
# 解析：定额库（sheet1）+ 套定额规则样例（sheet2）
# ---------------------------------------------------------------------------

_HEADER_LABELS = {"序号": "seq", "编号": "code", "定额编号": "code", "项目名称": "name",
                  "名称": "name", "单位": "unit", "单价": "price", "工程量": "qty"}

_STD_CODE = re.compile(r"^\d+-\d+-\d+$")  # 标准定额子目编号（z 开头为材料明细，不参与匹配）


def _find_columns(cells, labels):
    """按表头标签找列位置 → {字段: 列号}。"""
    found = {}
    for j, c in enumerate(cells):
        key = _HEADER_LABELS.get(_norm_text(c))
        if key and key not in found:
            found[key] = j
    return found


def _to_float(v):
    try:
        return float(str(v).replace(",", "").replace("元", "").strip())
    except (TypeError, ValueError):
        return None


def parse_quota_library(data: bytes) -> dict:
    """解析天津桥梁.xlsx：定额库 + 套定额规则样例。

    返回 {"entries": [{code,name,unit,price,row}], "rules": {子分部名: [定额行...]},
          "duplicates": {编号: 条数}}
    """
    import openpyxl
    try:
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    except Exception as e:
        raise BridgeMatchError(f"无法读取定额库文件：{e}") from e

    lib_sheet = None
    rule_sheet = None
    for ws in wb.worksheets:
        t = _norm_text(ws.title)
        if "分部分项" in t and "自然单位" not in t:
            lib_sheet = ws
        elif "自然单位" in t or "分部小计" in t:
            rule_sheet = ws
    if lib_sheet is None:
        lib_sheet = wb.worksheets[0]

    # ---- sheet1：定额库 ----
    rows = list(lib_sheet.iter_rows(values_only=True))
    cols = None
    header_row = None
    for i, r in enumerate(rows[:15]):
        c = _find_columns(r, _HEADER_LABELS)
        if len(c) >= 3 and "name" in c:
            cols, header_row = c, i
            break
    if cols is None:
        raise BridgeMatchError("未在定额库中识别到表头（需含 编号/项目名称/单位/单价 列）。")
    # 两行表头（表7(1) 版：第二行是 单价/合价）：把缺的列从下一行补上
    if "price" not in cols and header_row + 1 < len(rows):
        c2 = _find_columns(rows[header_row + 1], _HEADER_LABELS)
        for k, v in c2.items():
            if k not in cols:
                cols[k] = v

    entries, counts = [], {}
    for i, r in enumerate(rows[header_row + 1:], start=header_row + 2):
        name = str(r[cols["name"]]).strip() if cols.get("name") is not None and r[cols["name"]] is not None else ""
        code = str(r[cols["code"]]).strip() if cols.get("code") is not None and r[cols["code"]] is not None else ""
        unit = str(r[cols["unit"]]).strip() if cols.get("unit") is not None and r[cols["unit"]] is not None else ""
        price = _to_float(r[cols["price"]]) if cols.get("price") is not None else None
        if not name:
            continue
        if any(k in name for k in ("合计", "小计", "总计")):
            continue
        if not _STD_CODE.match(code):
            continue  # 章节行（3-1）、z 开头材料明细行不参与匹配
        entries.append({"code": _norm_code(code), "name": name, "unit": unit,
                        "price": price, "row": i})
        if code:
            counts[code] = counts.get(code, 0) + 1
    if len(entries) < 50:
        raise BridgeMatchError(f"定额库条目过少（{len(entries)} 条），请确认选择的是天津桥梁定额库。")

    # ---- sheet2：套定额规则样例 ----
    rules = {}
    if rule_sheet is not None:
        rcols, rheader = None, None
        rrows = list(rule_sheet.iter_rows(values_only=True))
        for i, r in enumerate(rrows[:10]):
            c = _find_columns(r, _HEADER_LABELS)
            if "name" in c and ("code" in c or "unit" in c):
                rcols, rheader = c, i
                break
        if rcols is not None and "price" not in rcols and rheader + 1 < len(rrows):
            c2 = _find_columns(rrows[rheader + 1], _HEADER_LABELS)
            for k, v in c2.items():
                if k not in rcols:
                    rcols[k] = v
        if rcols is not None:
            parent_set = {"下部结构", "上部结构", "桥面系", "其他"}
            stack, group = [], None
            for i, r in enumerate(rrows[rheader + 1:], start=rheader + 2):
                name = str(r[rcols["name"]]).strip() if rcols.get("name") is not None and r[rcols["name"]] is not None else ""
                code = str(r[rcols["code"]]).strip() if rcols.get("code") is not None and r[rcols["code"]] is not None else ""
                unit = str(r[rcols["unit"]]).strip() if rcols.get("unit") is not None and r[rcols["unit"]] is not None else ""
                if not name:
                    continue
                if any(k in name for k in ("合计", "小计", "总计")):
                    continue
                if not code:
                    # 分部/子分部行
                    nm = _norm_text(name)
                    if nm in parent_set:
                        stack, group = [name], name
                    elif stack and stack[-1] in parent_set:
                        stack.append(name)
                        group = name
                    else:
                        stack, group = [name], name
                    continue
                # 明细行（带编号）
                if not code[0].isdigit():
                    continue
                qty = _to_float(r[rcols["qty"]]) if rcols.get("qty") is not None else None
                rules.setdefault(group, []).append(
                    {"code": _norm_code(code), "name": name, "unit": unit, "qty": qty})

    return {"entries": entries, "rules": rules,
            "duplicates": {k: v for k, v in counts.items() if v > 1}}


# ---------------------------------------------------------------------------
# 解析：海泰式数量总表（分项材料表）
# ---------------------------------------------------------------------------

_SKIP_SHEET_KEYS = ("使用说明", "工程数量表", "材料统计", "指标信息")

# 工程数量表里不分部位的材料行（值集中在首列，直接以材料名作为分项）
_MATRIX_STANDALONE = {"挖方", "填方", "试桩", "锥坡", "支座", "泄水孔", "伸缩缝",
                      "超声波检测管", "其他", "桩基类型"}


def _parse_matrix_sheet(rows):
    """「工程数量表」sheet → 部位×材料矩阵条目 [{fenxiang, mat, mat_unit, value}]。

    两级列头（构造组 + 子部位，如 桥台→台身），材料行（C50/HRB400…），交叉值=工程量。
    该表是海泰程序自动填表的汇总输出，作为权威工程量来源。
    """
    out = []
    hdr = None
    for i, r in enumerate(rows[:10]):
        if r and "工程材料" in " ".join(str(x) for x in r if x):
            hdr = i
            break
    if hdr is None:
        return out
    r6 = rows[hdr]
    r7 = rows[hdr + 1] if hdr + 1 < len(rows) else None
    # R6 上方一行是构造组行（上部构造/下部构造/…/搭板/合　计），
    # 它的「合计」列是矩阵右边界——其后是右侧「项/工程量/单价/综合单价」清单块，
    # 列头（盖梁/桩基…）与左侧重复，必须排除；搭板列在 R6/R7 无子头，取该行名称。
    grand = None
    for k in range(max(0, hdr - 3), hdr):
        rk = rows[k]
        if any("合计" in _norm_text(str(x)) or "搭板" in _norm_text(str(x)) for x in rk):
            grand = rk
            break
    cols, group = [], None
    for j in range(4, min(len(r6), 32)):
        gj = str(grand[j] or "").strip() if grand is not None and j < len(grand) else ""
        if gj and "合计" in _norm_text(gj):
            break  # 右侧清单块起，矩阵结束
        if gj and "搭板" in _norm_text(gj):
            group = "搭板"
        else:
            g = str(r6[j] or "").strip() if j < len(r6) else ""
            if g:
                group = g
        sub = str(r7[j] or "").strip() if r7 is not None and j < len(r7) else ""
        name = _norm_text(f"{group} {sub}".strip() if sub else group)
        if not name or any(k in name for k in ("小计", "合计")):
            continue
        cols.append((j, name))
    for i in range(hdr + 2, len(rows)):
        r = rows[i]
        mat = str(r[2] or r[1] or "").strip() if len(r) > 2 else ""
        if not mat or any(k in _norm_text(mat) for k in ("小计", "合计")):
            continue
        unit = str(r[3] or "").strip() if len(r) > 3 else ""
        mat_n = _norm_text(mat)
        for j, fx in cols:
            if len(r) <= j or r[j] in (None, ""):
                continue
            v = _to_float(r[j])
            if v is None or v == 0:
                continue
            if mat_n in _MATRIX_STANDALONE:
                fx = mat  # 挖方/填方等不分部位
            out.append({"fenxiang": fx, "mat": mat, "mat_unit": unit, "value": v})
    return out


def parse_quantity_excel(data: bytes) -> list[dict]:
    """解析海泰式数量总表：分项材料表块（表头行 分项名/适用位置/备注/数量/材料列…）。

    返回 [{sheet, kind:"材料"|"桩基参数", fenxiang, pos, remark, count,
           mat, mat_unit, value, row}] + 桩基参数 {dia, depth, 根数=count}。
    """
    # 格式嗅探：openpyxl 先试，失败回退 xlrd
    try:
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        sheets = []
        for ws in wb.worksheets:
            sheets.append((ws.title, [list(r) for r in ws.iter_rows(values_only=True)]))
        xls_mode = False
    except Exception:
        import xlrd
        try:
            wb = xlrd.open_workbook(file_contents=data)
        except Exception as e:
            raise BridgeMatchError(f"无法读取数量表文件：{e}") from e
        sheets = []
        for ws in wb.sheets():
            sheets.append((ws.name,
                           [[ws.cell_value(i, j) for j in range(ws.ncols)] for i in range(ws.nrows)]))
        xls_mode = True

    items = []
    skipped = []
    matrix_items = []
    for sname, rows in sheets:
        if "工程数量表" in _norm_text(sname):
            matrix_items = _parse_matrix_sheet(rows)  # 权威汇总矩阵，最后覆盖合并
            # 桩基类型汇总行（"桩径1m桩共：10 根; …桩长…合计：300 m; …"）
            # → 按桩径汇总的桩基参数条目（工程量以工程数量表为准）
            for r in rows:
                if r and _norm_text(str(r[1] or "")) == "桩基类型":
                    txt = str(r[4] or "") if len(r) > 4 else ""
                    groups = {}
                    for p in re.split(r"[;；]", txt):
                        m = re.search(r"桩径([\d.]+)m", p)
                        if not m:
                            continue
                        g = groups.setdefault(float(m.group(1)),
                                              {"dia": float(m.group(1)),
                                               "count": None, "total": 0.0})
                        if "桩共" in p:
                            mc = re.search(r"共[：:]\s*([\d.]+)\s*根", p)
                            if mc:
                                g["count"] = float(mc.group(1))
                        else:  # 桩长…合计：X m
                            mt = re.search(r"合计[：:]\s*([\d.]+)\s*m", p)
                            if mt:
                                g["total"] += float(mt.group(1))
                    for g in groups.values():
                        if g["count"] and g["total"]:
                            items.append({"sheet": "工程数量表", "kind": "桩基参数",
                                          "fenxiang": "桩基类型", "pos": "", "remark": "",
                                          "count": g["count"], "dia": g["dia"],
                                          "depth": g["total"] / g["count"], "row": 0})
            continue
        if any(k in _norm_text(sname) for k in _SKIP_SHEET_KEYS):
            skipped.append(sname)
            continue
        header_idx = [i for i, r in enumerate(rows)
                      if r and _norm_text(r[0]) == "分项名"]
        if not header_idx:
            continue
        header_idx.append(len(rows))  # 末哨兵
        for bi in range(len(header_idx) - 1):
            h = header_idx[bi]
            hend = header_idx[bi + 1]
            hdr = rows[h]
            # 单位行 = 表头下一行且首列为空
            unit_row = rows[h + 1] if h + 1 < hend and rows[h + 1] and not str(rows[h + 1][0] or "").strip() else None
            # 材料列（第 5 列起，即索引 4+）
            mat_cols = [(j, str(hdr[j]).strip(), _norm_text(str(hdr[j])))
                        for j in range(4, len(hdr)) if str(hdr[j] or "").strip()]
            pile_cols = {_norm_text(str(hdr[j])): j for j in range(4, len(hdr))}
            is_pile_block = ("桩径" in pile_cols) and ("桩长" in pile_cols)
            fenxiang = None
            for i in range(h + (2 if unit_row else 1), hend):
                r = rows[i]
                if not r:
                    continue
                c0 = str(r[0] or "").strip() if len(r) > 0 else ""
                if _norm_text(c0) == "分项名":
                    break  # 下一个块（安全兜底）
                if c0:
                    fenxiang = c0
                if not fenxiang or any(k in fenxiang for k in ("此行不填", "本行不填", "不填")):
                    continue
                pos = str(r[1] or "").strip() if len(r) > 1 else ""
                remark = str(r[2] or "").strip() if len(r) > 2 else ""
                count = _to_float(r[3]) if len(r) > 3 else None
                if is_pile_block:
                    dia = _to_float(r[pile_cols["桩径"]]) if "桩径" in pile_cols and len(r) > pile_cols["桩径"] else None
                    dep = _to_float(r[pile_cols["桩长"]]) if "桩长" in pile_cols and len(r) > pile_cols["桩长"] else None
                    if count is not None and dia is not None and dep is not None:
                        items.append({"sheet": sname, "kind": "桩基参数", "fenxiang": fenxiang,
                                      "pos": pos, "remark": remark, "count": count,
                                      "dia": dia, "depth": dep, "row": i + 1})
                    continue
                # 材料列取值
                row_has_value = False
                for j, mname, _ in mat_cols:
                    if len(r) <= j or r[j] in (None, ""):
                        continue
                    v = _to_float(r[j])
                    if v is None:
                        continue
                    if count is not None:
                        v = v * count  # 海泰口径：合计工程量 = 数量(倍数) × 材料量
                    mun = str(unit_row[j] or "").strip() if unit_row and len(unit_row) > j else ""
                    items.append({"sheet": sname, "kind": "材料", "fenxiang": fenxiang,
                                  "pos": pos, "remark": remark, "count": count,
                                  "mat": mname, "mat_unit": mun, "value": v, "row": i + 1})
                    row_has_value = True
                # 备注列带单位（如 泄水孔 的 套/m/处）
                if remark and _norm_text(remark) in _UNIT_MAP and count is not None:
                    items.append({"sheet": sname, "kind": "材料", "fenxiang": fenxiang,
                                  "pos": pos, "remark": remark, "count": None,
                                  "mat": f"{fenxiang}（{remark}）", "mat_unit": remark,
                                  "value": count, "row": i + 1})
                    row_has_value = True
                _ = row_has_value
    # 桩基参数以工程数量表桩基类型汇总为准（按桩径合并根数/桩长）
    if any(it.get("kind") == "桩基参数" and it.get("sheet") == "工程数量表" for it in items):
        items = [it for it in items
                 if it.get("kind") != "桩基参数" or it.get("sheet") == "工程数量表"]
    # 「工程数量表」矩阵为权威工程量：同部位同材料的明细条目以矩阵值为准
    if matrix_items:
        covered = {(_norm_text(m["fenxiang"]), _norm_text(m["mat"]))
                   for m in matrix_items}
        detail_by_key = {}
        for it in items:
            if it.get("kind") == "材料":
                detail_by_key.setdefault((_norm_text(it["fenxiang"]),
                                          _norm_text(it["mat"])), it)
        merged = [it for it in items
                  if it.get("kind") != "材料"
                  or (_norm_text(it["fenxiang"]), _norm_text(it["mat"])) not in covered]
        # 桩基"单桩用量"明细条目（海泰数量列=1 的小钢种小值）：
        # 矩阵已按 HPB300/HRB400 给出该部位钢筋笼总量 → 删除，避免重复/口径混乱
        pile_rebar_fx = {_norm_text(m["fenxiang"]) for m in matrix_items
                         if "桩基" in _norm_text(m["fenxiang"])
                         and re.search(r"HPB|HRB", _norm_text(m["mat"]))}
        merged = [it for it in merged
                  if not (it.get("kind") == "材料"
                          and _norm_text(it["fenxiang"]) in pile_rebar_fx
                          and re.search(r"HPB235|HRB335", _norm_text(it["mat"]))
                          and (_to_float(it.get("value")) or 0) <= 100)]
        for m in matrix_items:
            mu = m["mat_unit"]
            if not mu:  # 工程数量表钢筋行单位列为空 → 借用明细条目单位
                d = detail_by_key.get((_norm_text(m["fenxiang"]), _norm_text(m["mat"])))
                mu = (d or {}).get("mat_unit") or ""
            merged.append({"sheet": "工程数量表", "kind": "材料", "fenxiang": m["fenxiang"],
                           "pos": "", "remark": "", "count": None,
                           "mat": m["mat"], "mat_unit": mu,
                           "value": m["value"], "row": 0})
        items = merged
    if not items:
        raise BridgeMatchError(
            "未解析到任何分项材料条目。请上传海泰数量总表式文件"
            "（sheet 内含「分项名/适用位置/备注/数量/材料列」块）。")
    return items


def parse_text_lines(text: str) -> tuple[list[dict], list[str]]:
    """粘贴文本 → 条目列表 + 解析警告。每行「名称 数量 单位」。"""
    items, warns = [], []
    for ln, line in enumerate(str(text or "").splitlines(), 1):
        line = line.strip()
        if not line:
            continue
        if any(k in line for k in ("项目名称", "工程量", "序号")):
            continue
        m = re.match(r"^(.*?)[\s\t,，]+(\d+(?:\.\d+)?)[\s\t]*(\S{1,6})?$", line)
        if not m:
            warns.append(f"第 {ln} 行无法解析（缺少数量）：{line[:30]}")
            continue
        items.append({"sheet": "粘贴文本", "kind": "材料", "fenxiang": m.group(1).strip(),
                      "pos": "", "remark": "", "count": None,
                      "mat": "", "mat_unit": m.group(3) or "", "value": float(m.group(2)),
                      "row": ln})
    return items, warns


# ---------------------------------------------------------------------------
# 特征提取与打分
# ---------------------------------------------------------------------------

COMPONENT_GROUPS = [
    ("桩基础", ["钻孔灌注桩", "灌注桩", "钢护筒", "护筒", "桩基", "钢筋笼", "声测管",
                "超声波检测管", "检测管", "凿桩头", "桩头", "挖孔桩", "钢管桩",
                "预应力管桩", "预制桩"]),
    ("承台系", ["承台", "围堰", "垫层", "基坑", "垫石", "系梁"]),
    ("墩台系", ["墩身", "墩柱", "桥台", "台身", "台帽", "盖梁", "耳墙", "背墙",
                "柱式墩", "墩台", "墩"]),
    ("梁体", ["箱梁", "板梁", "T梁", "空心板", "实心板", "盖板", "预制梁", "现浇梁",
              "桥面板", "横隔板", "湿接缝", "梁板", "梁"]),
    ("预应力系", ["钢绞线", "预应力", "锚具", "波纹管"]),
    ("桥面系", ["桥面铺装", "铺装", "防水层", "防水混凝土", "伸缩缝", "搭板", "枕梁",
                "桥面连续", "沥青"]),
    ("护栏栏杆", ["防撞护栏", "护栏", "防撞墙", "栏杆", "人行道", "缘石"]),
    ("排水", ["泄水孔", "泄水管", "透水管", "排水管", "盲沟", "排水"]),
    ("支座", ["支座", "板式支座", "盆式支座", "球形支座"]),
    ("临时工程", ["支架", "脚手架", "模板", "操作平台", "便道"]),
    ("土方", ["土方", "开挖", "基坑", "回填", "填方", "挖方", "开槽", "运土", "废料"]),
    ("钢筋", ["钢筋", "HPB", "HRB", "CRB"]),
    ("混凝土", ["混凝土"]),
]

METHODS = ["钻孔", "成孔", "灌注", "埋设", "沉入", "打入", "锤击", "振动", "预制", "现浇",
           "安装", "制作", "绑扎", "张拉", "先张", "后张", "压浆", "运输", "吊装", "架设",
           "浇筑", "摊铺", "碾压", "拆除", "凿除", "抹面", "防水", "涂装", "焊接", "切割",
           "顶推", "悬浇"]
SOIL = ["砂土", "黏土", "砂砾", "砂卵石", "岩石", "风化岩", "淤泥"]
GRADES = ["细粒式", "中粒式", "粗粒式", "砂粒式"]


def _extract_features(text):
    """提取构件组/工艺/材料/规格/条件。"""
    t = _norm_text(text)
    comps, methods, materials, conds = set(), set(), set(), set()
    for group, kws in COMPONENT_GROUPS:
        for kw in sorted(kws, key=len, reverse=True):
            if kw in t:
                comps.add(group)
                break
    for kw in METHODS:
        if kw in t:
            methods.add(kw)
    if re.search(r"C\d{2}", t):
        materials.add(re.search(r"C\d{2}", t).group())
    for kw in ("混凝土", "钢筋", "钢绞线", "沥青", "浆砌", "片石", "碎石", "砂浆",
               "钢管", "铸铁管", "塑料管", "螺纹钢", "3号钢"):
        if kw in t:
            materials.add(kw)
    if re.search(r"PVC|UPVC", t):
        materials.add("塑料管")
    for kw in SOIL:
        if kw in t:
            conds.add(kw)
    for kw in GRADES:
        if kw in t:
            conds.add(kw)
    specs = {}
    m = re.search(r"[φΦ∅⌀]\s*≤\s*(\d+(?:\.\d+)?)", t)
    if m:
        specs["dia_le"] = float(m.group(1))
    else:
        # 注意 (?!\d*…)：防止 \d+ 回溯到 φ10以内 的 "1"
        m = re.search(r"[φΦ∅⌀]\s*(\d+(?:\.\d+)?)(?!\d*\s*(以内|以外))", t)
        if m:
            specs["dia"] = float(m.group(1))
    m = re.search(r"[hH]\s*≤\s*(\d+(?:\.\d+)?)\s*m?", t)
    if m:
        specs["depth_le"] = float(m.group(1))
    m = re.search(r"桩长\s*≤?\s*(\d+(?:\.\d+)?)", t)
    if m and "depth_le" not in specs:
        specs["depth_le"] = float(m.group(1))
    m = re.search(r"(\d+)\s*孔", t)
    if m:
        specs["holes"] = int(m.group(1))
    m = re.search(r"(\d+)\s*号钢", t)
    if m:
        specs["steel_no"] = m.group(1)
    # 钢种：HRB/CRB→螺纹钢，HPB/Q235→3号钢
    if re.search(r"HRB|CRB", t):
        specs["steel_type"] = "螺纹钢"
        materials.add("螺纹钢")
    elif re.search(r"HPB|Q235", t):
        specs["steel_type"] = "3号钢"
        materials.add("3号钢")
    # 伸缩缝型号数字（KF-80/TS160 等 → 80/160）
    m = re.search(r"(?:KF|TS|GQF|CD|F)[-]?(\d{2,4})", t)
    if m:
        specs["type_no"] = int(m.group(1))
    # 钢筋直径边界（φ10以内/φ14以外）
    m = re.search(r"[φΦ∅]\s*(\d+(?:\.\d+)?)\s*(以内|以外)", t)
    if m:
        specs["dia_bound"] = (float(m.group(1)), m.group(2))
    return {"comps": comps, "methods": methods, "materials": materials,
            "conds": conds, "specs": specs, "text": t}


# 分部 → 构件类别（用于「构件类别与分部一致」+5）
_GROUP_FAMILY = {
    "桩基础": "桩基础", "承台": "承台系", "桥台": "墩台系", "桥墩": "墩台系",
    "墩盖梁": "墩台系", "台盖梁": "墩台系", "预应力砼现浇箱梁": "梁体",
    "预应力混凝土小箱梁": "梁体", "预应力板梁": "梁体", "桥面铺装": "桥面系",
    "搭板": "桥面系", "防撞护栏": "护栏栏杆", "支座": "支座", "伸缩缝": "桥面系",
    "排水": "排水", "土方": "土方",
}
# 分部 → 全库补强检索的名称关键词（用于同分部变体定额）
_GROUP_FAMILY_KW = {
    "桩基础": "灌注桩|钻孔桩", "承台": "承台", "桥台": "台身|桥台|台帽",
    "桥墩": "墩|柱", "墩盖梁": "盖梁", "台盖梁": "盖梁",
    "预应力砼现浇箱梁": "箱梁", "预应力混凝土小箱梁": "小箱梁",
    "预应力板梁": "板梁|空心板", "桥面铺装": "桥面|铺装", "搭板": "搭板",
    "防撞护栏": "护栏|防撞", "支座": "支座", "伸缩缝": "伸缩缝",
    "排水": "泄水|透水|盲沟", "土方": "土方|开挖|回填",
}


def _bigrams(t):
    return {t[i:i + 2] for i in range(len(t) - 1)} if len(t) >= 2 else set()


def _score(item_feats, entry_feats, item_unit, entry_unit, item_group=""):
    """规则打分 → (分数, 命中理由列表, 单位是否兼容)。"""
    score, reasons, unit_ok = 0, [], True
    if item_feats["text"] and item_feats["text"] == entry_feats["text"]:
        return 100, ["名称完全一致"], True
    # 构件
    hit = item_feats["comps"] & entry_feats["comps"]
    comp_pts = min(len(hit) * 12, 24)
    score += comp_pts
    for c in hit:
        reasons.append(f"构件一致：{c}")
    # 工艺
    hit = item_feats["methods"] & entry_feats["methods"]
    meth_pts = min(len(hit) * 6, 12)
    score += meth_pts
    if hit:
        reasons.append("工艺一致：" + "、".join(sorted(hit)))
    # 材料（C 等级 + 材料词）
    hit = item_feats["materials"] & entry_feats["materials"]
    mat_pts = min(len(hit) * 5, 10)
    score += mat_pts
    if hit:
        reasons.append("材料一致：" + "、".join(sorted(hit)))
    # 规格
    isp, esp = item_feats["specs"], entry_feats["specs"]
    if "dia" in isp and "dia" in esp and isp["dia"] == esp["dia"]:
        score += 8
        reasons.append(f"桩径一致 φ{isp['dia']:g}")
    if "dia" in isp and "dia_le" in esp and isp["dia"] <= esp["dia_le"]:
        score += 6
        reasons.append(f"桩径 φ{isp['dia']:g} 在定额 φ≤{esp['dia_le']:g} 范围内")
    if "dia" in isp and "dia" in esp and isp["dia"] != esp["dia"]:
        score -= 6
        reasons.append(f"直径不同：{isp['dia']:g} vs {esp['dia']:g}")
    # 钢筋直径边界（φ10以内/φ14以外）
    if "dia" in isp and "dia_bound" in esp:
        v, mode = esp["dia_bound"]
        if (mode == "以内" and isp["dia"] <= v) or (mode == "以外" and isp["dia"] > v):
            score += 6
            reasons.append(f"直径 ∅{isp['dia']:g} 在定额 φ{v:g}{mode}范围内")
        else:
            score -= 6
            reasons.append(f"直径 ∅{isp['dia']:g} 不在定额 φ{v:g}{mode}范围")
    if "dia_bound" in isp and "dia" in esp:
        v, mode = isp["dia_bound"]
        if (mode == "以内" and esp["dia"] <= v) or (mode == "以外" and esp["dia"] > v):
            score += 6
            reasons.append(f"定额直径 φ{esp['dia']:g} 在清单 φ{v:g}{mode}范围内")
        else:
            score -= 6
            reasons.append(f"定额直径 φ{esp['dia']:g} 不在清单 φ{v:g}{mode}范围")
    if "depth_le" in isp and "depth_le" in esp and isp["depth_le"] == esp["depth_le"]:
        score += 6
        reasons.append(f"深度一致 h≤{isp['depth_le']:g}m")
    if "depth_le" in isp and "depth_le" in esp and isp["depth_le"] < esp["depth_le"]:
        score += 3
        reasons.append(f"深度 h≤{isp['depth_le']:g}m 在定额 h≤{esp['depth_le']:g}m 范围内")
    c_item = [m for m in isp_materials(item_feats) if re.match(r"C\d{2}$", m)]
    c_ent = [m for m in esp_materials(entry_feats) if re.match(r"C\d{2}$", m)]
    if c_item and c_ent:
        if c_item[0] == c_ent[0]:
            score += 6
            reasons.append(f"混凝土强度一致 {c_item[0]}")
        else:
            score -= 6
            reasons.append(f"混凝土强度不同：清单 {c_item[0]} vs 定额 {c_ent[0]}")
    for k in ("holes", "steel_no"):
        if k in isp and k in esp:
            if isp[k] == esp[k]:
                score += 5
                reasons.append(f"{k} 一致")
            else:
                score -= 6
                reasons.append(f"{k} 不同：{isp[k]} vs {esp[k]}")
    # 钢种（螺纹钢/3号钢）
    if "steel_type" in isp and "steel_type" in esp:
        if isp["steel_type"] == esp["steel_type"]:
            score += 6
            reasons.append(f"钢种一致 {isp['steel_type']}")
        else:
            score -= 6
            reasons.append(f"钢种不同：{isp['steel_type']} vs {esp['steel_type']}")
    # 伸缩缝型号（KF-80 ↔ TS80 伸缩量一致）
    if "type_no" in isp and "type_no" in esp:
        if isp["type_no"] == esp["type_no"]:
            score += 8
            reasons.append(f"伸缩量规格一致 {isp['type_no']}")
        else:
            score -= 6
            reasons.append(f"伸缩量不同：{isp['type_no']} vs {esp['type_no']}")
    # 地层
    hit = item_feats["conds"] & entry_feats["conds"]
    if hit:
        score += 4
        reasons.append("条件一致：" + "、".join(sorted(hit)))
    # 构件类别与分部一致
    if item_group:
        fam = _GROUP_FAMILY.get(_map_group(item_group))
        if fam and fam in entry_feats["comps"]:
            score += 5
            reasons.append(f"构件类别与分部一致（{fam}）")
    # bigram
    score += min(len(_bigrams(item_feats["text"]) & _bigrams(entry_feats["text"])), 15)
    # 名称包含
    a, b = item_feats["text"], entry_feats["text"]
    if len(a) >= 4 and len(b) >= 4 and (a in b or b in a):
        score += 8
        reasons.append("名称相互包含")
    # 单位
    ib, im = _unit_parts(item_unit) if item_unit else ("", 1.0)
    eb, em = _unit_parts(entry_unit)
    if item_unit and ib and eb:
        if ib == eb and im == em:
            score += 4
            reasons.append(f"单位一致 {item_unit}")
        elif ib == eb:
            score += 3
            reasons.append(f"单位同基类（{item_unit} ↔ {entry_unit}）")
        elif ib in _COUNT_UNITS and eb in _COUNT_UNITS:
            score += 3
            reasons.append("同为计数单位")
        elif {ib, eb} <= {"kg", "t"}:
            score += 3
            reasons.append(f"单位 {item_unit} ↔ {entry_unit}（kg/t）")
        else:
            score -= 10
            unit_ok = False
            reasons.append(f"单位基类不符（{item_unit} vs {entry_unit}）")
    return score, reasons, unit_ok


def isp_materials(feats):
    return feats["materials"]


def esp_materials(feats):
    return feats["materials"]


# ---------------------------------------------------------------------------
# 匹配主流程
# ---------------------------------------------------------------------------

TH_SINGLE = 16
MARGIN = 8
MIN_CANDIDATE = 10
MAX_CANDIDATES = 5

_GROUP_MAP = [
    ("桩基础", ["桩基", "凿桩头"]),
    ("承台", ["承台"]),
    ("桥台", ["台身", "桥台"]),
    ("桥墩", ["墩柱", "墩身", "桥墩", "系梁"]),
    ("预应力砼现浇箱梁", ["现浇箱梁", "箱梁"]),
    ("预应力混凝土小箱梁", ["小箱梁"]),
    ("预应力板梁", ["板梁", "空心板"]),
    ("桥面铺装", ["桥面铺装", "铺装"]),
    ("搭板", ["搭板"]),
    ("防撞护栏", ["防撞护栏", "护栏"]),
    ("支座", ["支座"]),
    ("伸缩缝", ["伸缩缝"]),
    ("排水", ["泄水孔", "泄水管", "透水管", "排水"]),
    ("土方", ["土方", "开槽", "回填", "挖方", "填方"]),
]


def _map_group(fenxiang):
    t = _norm_text(fenxiang)
    if "盖梁" in t:
        return "台盖梁" if "桥台" in t else "墩盖梁"
    for g, kws in _GROUP_MAP:
        for kw in sorted(kws, key=len, reverse=True):
            if kw in t:
                return g
    return None


# 混凝土→模板定额映射：不同部位混凝土随同套用相应模板定额，
# 模板量=混凝土量×系数。系数默认取参考表样例比例（模板样例量÷混凝土样例量），
# 页面可自行修改；梯道样例表无对应子分部，默认系数为用户指定。
# (fenxiang 关键词, 模板编号, 系数键, 默认系数)
_TEMPLATE_MAP = [
    ("承台", "3-7-26", "承台", 1.207),      # 692.6973/573.9
    ("台身", "3-7-28", "桥台", 1.499),      # 299.0505/199.5
    ("台盖梁", "3-7-34", "盖梁", 3.031),    # 样例 0/0，按墩盖梁系数
    ("盖梁", "3-7-33", "盖梁", 3.031),      # 700.161/231
    ("墩", "3-7-29", "墩柱", 4.295),        # 12126.9325/2823.5（含柱系梁，样例整体口径）
    ("箱梁", "3-7-38", "现浇箱梁", 5.387),  # 51725.974/9602
    ("梯道", "3-7-41", "梯道", 5.387),      # 样例无梯道，用户指定默认 5.387
    ("搭板", "3-7-42", "搭板", 3.209),      # 346.572/108
    ("护栏", "3-7-47", "护栏", 4.810),      # 3261.18/678
    ("挡", "3-7-49", "挡墙", 1.61),         # 用户指定默认 1.61（挡土墙）
]
# 页面输入框顺序（键去重）
TPL_RATIOS = []
for _kw, _code, _key, _def in _TEMPLATE_MAP:
    if not any(k == _key for k, _ in TPL_RATIOS):
        TPL_RATIOS.append((_key, _def))


def _template_for(item) -> dict | None:
    """混凝土条目对应模板定额（模板量=混凝土量×系数，系数页面可改）。
    桩基（灌注桩）、桥面铺装、沥青混凝土不套模板。"""
    fx = _norm_text(item.get("fenxiang") or "")
    mat = _norm_text(item.get("mat") or "")
    if "桩基" in fx or "沥青" in mat:
        return None
    if not (re.search(r"C\d{2}", mat) or "混凝土" in mat):
        return None
    for kw, code, key, default in _TEMPLATE_MAP:
        if kw in fx:
            return {"code": code, "key": key, "default": default}
    return None


def _item_display(item):
    if item["kind"] == "桩基参数":
        return (f"{item['fenxiang']}（{item['count']:g} 根 φ{item['dia']:g}m "
                f"桩长 {item['depth']:g}m）")
    mat = item.get("mat") or ""
    return f"{item['fenxiang']} {mat}".strip()


def _classify(cands):
    """候选分类 → (status, picked, candidates, note)。"""
    cands = [c for c in cands if c["score"] >= MIN_CANDIDATE]
    if not cands:
        return "none", None, [], ""
    best = cands[0]
    if best["score"] >= 100:
        return "single", best, cands[:1], ""
    ok = [c for c in cands if c["unit_ok"]]
    best_ok = ok[0] if ok else None
    second = cands[1]["score"] if len(cands) > 1 else -999
    if best_ok is not None and best["score"] >= TH_SINGLE and best["score"] - second >= MARGIN:
        return "single", best_ok, cands[:MAX_CANDIDATES], ""
    if best["score"] >= TH_SINGLE and best["score"] - second < MARGIN:
        return "multi", None, cands[:MAX_CANDIDATES], "多个候选定额得分接近，请结合差异说明人工选择。"
    return "none", None, cands[:3], "得分最高的候选未达到可信阈值，需补充特征后重新匹配。"


def _checklist(item, feats=None):
    """缺失特征 → 补充信息清单（分项类型感知）。"""
    feats = feats or _extract_features(_item_display(item))
    t = feats["text"]
    fx = _norm_text(item["fenxiang"])
    mat = _norm_text(item.get("mat") or "")
    missing = []
    is_pile_param = item.get("kind") == "桩基参数"
    is_pile_material = "桩基础" == _map_group(item["fenxiang"]) and item.get("kind") == "材料"
    if not feats["comps"]:
        missing.append("构件类型（如 承台/墩柱/盖梁/护栏/支座…）")
    if (is_pile_param or ("桩" in fx and not is_pile_material)) and "dia" not in feats["specs"] and "dia_le" not in feats["specs"]:
        missing.append("桩径 φ（如 φ1200）")
    if (is_pile_param or ("桩" in fx and not is_pile_material)) and "depth_le" not in feats["specs"]:
        missing.append("桩长/钻孔深度 h（如 h≤40m）")
    if is_pile_param and not (feats["conds"] & set(SOIL)):
        missing.append("地层条件（如 砂土/黏土/砂砾/岩石）")
    if "混凝土" in feats["materials"] and not any(re.match(r"C\d{2}$", m) for m in feats["materials"]):
        missing.append("混凝土强度等级（如 C30）")
    if re.search(r"HPB|HRB|CRB|Q235|钢筋", mat + fx):
        missing.append("钢筋直径 φ10/φ14 以内或以外")
    if "预应力系" in feats["comps"]:
        missing.append("钢绞线束长/孔数")
    if "支座" in feats["comps"]:
        missing.append("支座类型（板式/四氟板式/盆式…）")
    if "护栏栏杆" in feats["comps"]:
        missing.append("栏杆材质类型（钢管/不锈钢/铸铁…）")
    if "排水" in feats["comps"]:
        if _unit_parts(item.get("mat_unit") or "")[0] == "kg":
            missing.append("泄水管长度（m，定额按 10m 计，kg 无法换算）")
        else:
            missing.append("泄水管材质（钢管/铸铁管/塑料管）")
    if item.get("kind") != "桩基参数" and not item.get("mat_unit") and not item.get("unit"):
        missing.append("单位（m/m2/m3/t/套）")
    return missing


def _match_text(item, entries):
    """全库文本打分（未映射/回退路径）。"""
    text = _item_display(item)
    feats = _extract_features(text)
    cands = []
    for e in entries:
        ef = _extract_features(e["name"])
        s, reasons, ok = _score(feats, ef, item.get("mat_unit") or item.get("unit"),
                                e["unit"], item_group=item.get("fenxiang", ""))
        cands.append({"entry": e, "score": s, "reasons": reasons, "unit_ok": ok})
    cands.sort(key=lambda c: -c["score"])
    return cands


def _find_candidates(item, entries, rules):
    """按材料语义在规则组（优先）与全库中找候选。

    返回 [{"entry": 定额条目副本, "boost": 规则加分, "boost_reason": str}, ...]
    """
    mat = _norm_text(item.get("mat") or "")
    fx = _norm_text(item["fenxiang"])
    group = _map_group(item["fenxiang"])
    rule_rows = rules.get(group) or []
    fam_kw = _GROUP_FAMILY_KW.get(group, "")
    pool = []

    def add_entry(e, boost=0, reason=""):
        pool.append({"entry": dict(e), "boost": boost, "boost_reason": reason})

    def add_rule_rows(class_token, boost_reason, boost_pred=None):
        """规则样例行：名称含类别词（混凝土/螺纹钢…）的给予规则加分。

        boost_pred(name) 可额外约束（如伸缩缝型号一致才加分）。
        """
        for rr in rule_rows:
            e = next((x for x in entries if x["code"] == rr["code"]), None)
            if e is None:
                e = {"code": rr["code"], "name": rr["name"], "unit": rr["unit"],
                     "price": None, "row": 0}
            nm = _norm_text(rr["name"])
            boost = 8 if class_token and class_token in nm and (boost_pred is None or boost_pred(nm)) else 0
            add_entry(e, boost, boost_reason if boost else "")

    def add_library(pred, boost=0, reason=""):
        for e in entries:
            if pred(e):
                add_entry(e, boost, reason)

    m_c = re.search(r"C\d{2}", mat)
    if "防水" in mat:
        # 防水混凝土/防水层 → 防水相关子目
        add_rule_rows("防水", f"样例规则指向（{group}）")
        add_library(lambda e: "防水" in e["name"], 0, "")
    elif m_c:
        # Cxx 混凝土材料 → 现浇混凝土定额（强度精确匹配优先）
        # 规则加分仅给浇筑类子目（灌注/现浇/预制/浇筑），模板/凿除等配套子目不抬分
        add_rule_rows("混凝土", f"样例规则指向（{group}）",
                      boost_pred=lambda nm: not re.search(r"模板|凿除|拆除|凿除桩顶", nm)
                      and re.search(r"灌注|现浇|预制|浇筑", nm) is not None)
        # 全库补强：同强度 + 工艺（现浇/预制/灌注）+ 同构件类别
        def pred(e, m=m_c):
            n = e["name"]
            return ("混凝土" in n and m.group() in n
                    and re.search(r"现浇|预制|灌注", n)
                    and (not fam_kw or re.search(fam_kw, n)))
        add_library(pred, 0, "")
    elif "沥青" in mat:
        add_rule_rows("沥青", f"样例规则指向（{group}）")
        add_library(lambda e: "沥青" in e["name"], 0, "")
    elif re.search(r"HPB|HRB|CRB|Q235|钢筋", mat):
        feats = _extract_features(mat)
        st = feats["specs"].get("steel_type", "")
        # 规则加分仅给钢种相符的样例行（HRB400 只抬 螺纹钢 行）
        add_rule_rows("钢筋", f"样例规则指向（{group}）",
                      boost_pred=lambda nm, st=st: bool(st) and st in nm)
        if st:
            add_library(lambda e, st=st: "钢筋" in e["name"] and st in e["name"], 0, "")
        else:
            add_library(lambda e: "钢筋" in e["name"]
                        and ("螺纹钢" in e["name"] or "3号钢" in e["name"]), 0, "")
        if group == "桩基础" and st in ("3号钢", "螺纹钢"):
            # 桩基钢筋按设计要求锚定：HPB300/Q235→3-2-6（3号钢 φ10以外）、
            # HRB400/CRB550→3-2-7（螺纹钢 φ14以外）
            want = "3-2-6" if st == "3号钢" else "3-2-7"
            dia_w = "φ10以外" if st == "3号钢" else "φ14以外"
            for p in pool:
                if str(p["entry"]["code"]).split("*")[0].strip() == want:
                    p["boost"] = 30
                    p["boost_reason"] = (f"桩基钢筋：{mat} → {want}"
                                         f"（{st} {dia_w}，设计要求）")
    elif "钢绞线" in mat:
        add_rule_rows("钢绞线", f"样例规则指向（{group}）")
        add_library(lambda e: "钢绞线" in e["name"], 0, "")
    elif "支座" in mat or "支座" in fx:
        add_rule_rows("支座", f"样例规则指向（{group}）")
        add_library(lambda e: "支座" in e["name"], 0, "")
    elif "伸缩" in mat or "伸缩" in fx:
        # 型号一致（KF-80 ↔ TS80）才给样例规则加分，避免型号不符的子目被抬分
        tn = _extract_features(mat)["specs"].get("type_no")
        add_rule_rows("伸缩缝", f"样例规则指向（{group}）",
                      boost_pred=lambda nm, tn=tn: tn is None
                      or _extract_features(nm)["specs"].get("type_no") == tn)
        add_library(lambda e: "伸缩缝" in e["name"]
                    and "沥青" not in e["name"] and "麻丝" not in e["name"], 0, "")
    elif group == "土方":
        # 挖方/填方：按语义锚定样例规则（挖方→开槽挖土、填方→填压土方）
        kw = ("挖" if "挖方" in fx or "挖方" in mat
              else "填" if "填方" in fx or "填方" in mat else None)
        for rr in rule_rows:
            e = next((x for x in entries if x["code"] == rr["code"]), None)
            if e is None:
                e = {"code": rr["code"], "name": rr["name"], "unit": rr["unit"],
                     "price": None, "row": 0}
            nm = _norm_text(rr["name"])
            boost = 25 if kw and kw in nm else 0
            add_entry(e, boost, f"样例规则指向（{group}）" if boost else "")
    else:
        add_rule_rows(None, "")
    # 去重（编号+名称+单位）
    seen, out = set(), []
    for p in pool:
        e = p["entry"]
        k = (e["code"], e["name"], e["unit"])
        if k not in seen:
            seen.add(k)
            out.append(p)
    return out


def _match_pile_params(item, entries, rules):
    """桩基参数（根数/桩径/桩长）→ 钻孔定额 + 样例组合。"""
    dia_mm = item["dia"] * 1000 if item["dia"] < 10 else item["dia"]
    depth = item["depth"]
    total_m = item["count"] * depth
    cands = []

    def _add_cand(e):
        ef = _extract_features(e["name"])
        specs = ef["specs"]
        if "钻孔" not in ef["methods"] or "灌注" in ef["methods"]:
            return
        if "dia_le" not in specs or "depth_le" not in specs:
            return
        if specs["dia_le"] < dia_mm or specs["depth_le"] < depth:
            return
        score, reasons = 10, [f"桩径 φ{dia_mm:g} ≤ φ≤{specs['dia_le']:g}"]
        if specs["dia_le"] == dia_mm:
            score += 8
            reasons[0] = f"桩径一致 φ≤{specs['dia_le']:g}"
        if specs["depth_le"] == depth:
            score += 6
            reasons.append(f"深度一致 h≤{specs['depth_le']:g}m")
        else:
            score += 3
            reasons.append(f"深度 h≤{depth:g}m 在定额 h≤{specs['depth_le']:g}m 范围内")
        if ef["conds"] & set(SOIL):
            reasons.append("定额适用地层：" + "、".join(sorted(ef["conds"] & set(SOIL))))
        score += min(len(_bigrams(_norm_text(item["fenxiang"])) & _bigrams(ef["text"])), 15)
        cands.append({"entry": e, "score": score, "reasons": reasons, "unit_ok": True})

    # 库中缺少小桩径钻孔变体（如 φ≤1800 的 3-1-32*0.87 换乘系数子目）时，
    # 从样例规则补入：单位借同号基础子目，单价按换乘系数折算
    extended = list(entries)
    for rr in rules.get("桩基础") or []:
        if "钻孔" not in _norm_text(rr["name"]):
            continue
        if any(str(x["code"]) == rr["code"] for x in entries):
            continue
        base = str(rr["code"]).split("*")[0].strip()
        be = next((x for x in entries if str(x["code"]) == base), None)
        price = (round(be["price"] * 0.87, 2)
                 if be and be["price"] is not None and "0.87" in rr["name"]
                 else be["price"] if be else None)
        extended.append({"code": rr["code"], "name": rr["name"],
                         "unit": be["unit"] if be else rr["unit"],
                         "price": price, "row": 0})
    for e in extended:
        _add_cand(e)
    if not cands:
        # 桩长超出库内该桩径所有深度档（如 h70 > 库内 φ1800 最大 h≤60）：
        # 取该桩径最大档位定额套用并标注，工程量仍按实际桩长计算
        best = None
        for e in extended:
            specs = _extract_features(e["name"])["specs"]
            if ("钻孔" not in _extract_features(e["name"])["methods"]
                    or "灌注" in _extract_features(e["name"])["methods"]):
                continue
            if specs.get("dia_le") == dia_mm and "depth_le" in specs:
                if best is None or specs["depth_le"] > best[0]:
                    best = (specs["depth_le"], e)
        if best:
            e = best[1]
            specs = _extract_features(e["name"])["specs"]
            cands.append({"entry": e, "score": 20, "reasons":
                          [f"桩径一致 φ≤{specs['dia_le']:g}",
                           f"桩长 {depth:g}m 超出库内 φ{dia_mm:g} 钻孔定额最大深度"
                           f" h≤{specs['depth_le']:g}m，按最大档套用（工程量按实际桩长）"],
                          "unit_ok": True})
    cands.sort(key=lambda c: -c["score"])
    # 钻孔工程量 = 根数 × 桩长（自然米）
    for c in cands:
        c["qty_natural"] = total_m
        c["qty_de"] = total_m / _unit_parts(c["entry"]["unit"])[1]
        c["amount"] = round(c["qty_de"] * c["entry"]["price"], 2) if c["entry"]["price"] is not None else None
    # 样例组合
    combo = []
    for rr in rules.get("桩基础") or []:
        nm = _norm_text(rr["name"])
        st, need = "需补充", ""
        if "钢护筒" in nm:
            m = re.search(r"φ(\d+)", nm)
            if m and float(m.group(1)) == dia_mm and "陆地" in rr["name"]:
                st = "已匹配"
                need = f"工程量=桩数×2m（每根桩 2m 护筒，{item['count']:g}根×2m）"
            else:
                st = "不适用"
                need = (f"桩径不同（定额 φ{m.group(1)}）" if m
                        else "水中埋设，陆地不套")
        elif "回旋钻机钻孔" in nm:
            st = "已匹配" if any(c["entry"]["code"] == rr["code"] for c in cands[:MAX_CANDIDATES]) else "不适用"
        elif "泥浆" in nm or "灌注桩混凝土" in nm:
            need = "工程量=桩基混凝土量（随桩基混凝土条目套用）"
        elif "凿除" in nm:
            st = "已匹配"
            need = "工程量=桩数×0.5m×截面积（分部内所有桩合计，套 3-10-44 一次）"
        elif "钢筋" in nm:
            need = "桩基钢筋用量及直径 φ10/φ14 内外"
        elif "运" in nm or "废料" in nm:
            st = "已匹配"
            need = ("凿桩头弃运：1-6-71=凿桩头量×2（建筑垃圾按虚方），"
                    "1-6-72=×2×(运距-1)，运距在页面设置")
        combo.append({"code": rr["code"], "name": rr["name"], "unit": rr["unit"],
                      "status": st, "need": need})
    # 陆地埋设钢护筒：每根桩 2m，工程量 = 桩数×2（定额单位 10m）
    liner = None
    for e in entries:
        nm = _norm_text(e["name"])
        if "陆地埋设钢护筒" in nm and f"φ{dia_mm:g}" in nm:
            liner = dict(e)
            liner["qty_natural"] = item["count"] * 2
            liner["qty_de"] = liner["qty_natural"] / _unit_parts(e["unit"])[1]
            liner["amount"] = (round(liner["qty_de"] * e["price"], 2)
                               if e["price"] is not None else None)
            break
    # 水中变体仅作提示（如实际水中施工改套）
    supplement = []
    for e in entries:
        if "水中埋设钢护筒" in e["name"] and f"φ{dia_mm:g}" in _norm_text(e["name"]):
            supplement.append(e)
    return cands, combo, supplement, liner


def _match_one(idx, item, entries, code_index, rules, all_items):
    status, matches, combo, supplement, checklist, notes = "none", [], [], [], [], []
    mud, cage = [], None
    feats = _extract_features(_item_display(item))
    if item["kind"] == "桩基参数":
        cands, combo, supplement, liner = _match_pile_params(item, entries, rules)
        status, picked, top, note = _classify(cands)
        matches = top
        checklist = _checklist(item)
        if note:
            notes.append(note)
        if not cands:
            notes.append(f"桩径 φ{item['dia']:g}/桩长 {item['depth']:g}m 超出库内回旋钻机钻孔定额范围，"
                         "无法匹配，请核对桩参数或补充地层条件。")
        else:
            notes.append(f"钻孔工程量 = Σ根数×桩长 = {item['count']:g}×{item['depth']:g} = "
                         f"{item['count'] * item['depth']:g}m；陆地埋设钢护筒按每根桩 2m 套用，"
                         f"工程量 = {item['count']:g}根×2m = {item['count'] * 2:g}m；"
                         "泥浆/灌注桩混凝土随桩基混凝土条目套用；"
                         "凿桩头按 0.5m/根计，分部内所有桩合计套 3-10-44 一次；"
                         "凿桩头弃运套 1-6-71/1-6-72（建筑垃圾按虚方×2，"
                         "1-6-72 次数=页面运距-1）。")
            if status in ("single", "multi") and supplement:
                for s in supplement:
                    notes.append(f"如水中施工改套：{s['code']} {s['name']}（{s['unit']}）")
        return {"idx": idx, "item": item, "status": status, "matches": matches,
                "rule_combo": combo, "checklist": checklist, "notes": notes,
                "pile_liner": liner}

    mat = _norm_text(item.get("mat") or "")
    if "面积" in mat:
        # 面积类列（桥梁面积等）为指标数据，非定额计量子目
        return {"idx": idx, "item": item, "status": "none", "matches": [],
                "rule_combo": [], "checklist": [],
                "notes": ["「面积」列为指标数据，非定额计量子目，不参与定额匹配。"]}
    group = _map_group(item["fenxiang"])
    # 混凝土条目随同套用相应模板定额（模板量=混凝土量×系数，系数默认参考表比例、页面可改）
    tpl = None
    tfi = _template_for(item)
    if tfi:
        e = next((x for x in entries
                  if str(x["code"]).split("*")[0].strip() == tfi["code"]), None)
        if e is not None:
            te = dict(e)
            te["tpl_key"] = tfi["key"]
            te["tpl_default"] = tfi["default"]
            te["conc_val"] = item.get("value")
            tpl = te
    if group and group in rules:
        pool = _find_candidates(item, entries, rules)
        cands = []
        for p in pool:
            e = p["entry"]
            ef = _extract_features(e["name"])
            s, reasons, ok = _score(feats, ef, item.get("mat_unit"), e["unit"],
                                    item_group=item.get("fenxiang", ""))
            if p["boost"]:
                s += p["boost"]
                reasons.insert(0, p["boost_reason"])
            cands.append({"entry": e, "score": s, "reasons": reasons, "unit_ok": ok})
        cands.sort(key=lambda c: -c["score"])
        if not cands:
            cands = _match_text(item, entries)
        status, picked, top, note = _classify(cands)
        matches = top
        if note:
            notes.append(note)
        # 桩基混凝土条目（top 为 3-1-60 灌注桩混凝土）：泥浆制作/运输随同套用
        is_pile_conc = (group == "桩基础" and re.search(r"C\d{2}", mat)
                        and matches
                        and str(matches[0]["entry"]["code"]).split("*")[0].strip() == "3-1-60")
        # 钻孔桩钢筋笼入孔 3-2-10：工程量=同部位 HPB300+HRB400 钢筋量之和（kg→t），
        # 由该部位 HPB300 条目携带套用（避免每行重复）
        cage = None
        if group == "桩基础" and mat == "HPB300":
            fx = _norm_text(item["fenxiang"])
            hpb_kg = hrb_kg = 0.0
            for o in all_items:
                if _norm_text(o.get("fenxiang") or "") != fx:
                    continue
                m = _norm_text(o.get("mat") or "")
                v = _to_float(o.get("value")) or 0
                if m == "HPB300":
                    hpb_kg += v
                elif m == "HRB400":
                    hrb_kg += v
            if hpb_kg or hrb_kg:
                e = next((x for x in entries if str(x["code"]) == "3-2-10"), None)
                if e is not None:
                    ce = dict(e)
                    ce["qty_natural"] = hpb_kg + hrb_kg
                    ce["qty_de"] = (hpb_kg + hrb_kg) / 1000 / _unit_parts(e["unit"])[1]
                    ce["amount"] = (round(ce["qty_de"] * e["price"], 2)
                                    if e["price"] is not None else None)
                    ce["cage_parts"] = (hpb_kg, hrb_kg)
                    cage = ce
        combo = []
        for rr in rules.get(group) or []:
            nm = _norm_text(rr["name"])
            st, need = "需补充", ""
            if any(c["entry"]["code"] == rr["code"] for c in cands[:MAX_CANDIDATES]):
                st = "已匹配" if cands and cands[0]["entry"]["code"] == rr["code"] else "候选"
            if "模板" in nm:
                if tpl and str(rr["code"]).split("*")[0].strip() == tpl["code"]:
                    st = "已匹配"
                    need = (f"模板量=混凝土量×系数（默认 {tpl['tpl_default']:g}，页面可改；"
                            f"定额单位 {tpl['unit']}）")
                else:
                    need = "模板面积（m2）"
            elif "泥浆" in nm:
                if is_pile_conc:
                    st = "已匹配"
                    need = "工程量=桩基混凝土量（随本条目套用，÷10 换算定额单位）"
                else:
                    need = "工程量=桩基混凝土量（随桩基混凝土条目套用）"
            elif "凿" in nm:
                if "梯道" in _norm_text(item["fenxiang"]):
                    st = "需补充"
                    need = "梯道下部输入表未提供桩数（无桩基类型块），凿桩头量需补充"
                else:
                    st = "已匹配"
                    need = "工程量=桩数×0.5m×截面积（分部内所有桩合计，套 3-10-44 一次）"
            elif "钢筋笼入孔" in nm:
                if group == "桩基础" and mat in ("HPB300", "HRB400"):
                    st = "已匹配" if cage else "候选"
                    need = "工程量=HPB300+HRB400 钢筋量之和（kg 按 t 换算）"
                else:
                    need = "钢筋笼重量（kg 按 t 换算）"
            elif "钢筋" in nm:
                if group == "桩基础":
                    need = ("桩基钢筋按设计要求套 φ10/φ14 以外子目（kg 按 t 换算）"
                            if "骨架" in nm else "钢筋笼重量（kg 按 t 换算）")
                else:
                    need = "钢筋用量及直径 φ10/φ14 内外"
            elif "钢绞线" in nm:
                need = "束长/孔数"
            elif "波纹管" in nm:
                need = "管道直径与长度"
            elif "运" in nm or "废料" in nm:
                if group == "桩基础":
                    if "梯道" in _norm_text(item["fenxiang"]):
                        st = "需补充"
                        need = "梯道下部无桩数信息，凿桩头弃运量需补充"
                    else:
                        st = "已匹配"
                        need = ("凿桩头弃运：1-6-71=凿桩头量×2（建筑垃圾按虚方），"
                                "1-6-72=×2×(运距-1)，运距在页面设置")
                else:
                    need = "弃运方量（按样例规则）"
            combo.append({"code": rr["code"], "name": rr["name"], "unit": rr["unit"],
                          "status": st, "need": need})
        checklist = _checklist(item)
        # 泥浆制作 3-1-58 / 泥浆运输 3-1-59：工程量=桩基混凝土量（定额单位 10m3）
        if is_pile_conc:
            val = item.get("value") or 0
            for code in ("3-1-58", "3-1-59"):
                e = next((x for x in entries if str(x["code"]) == code), None)
                if e is None:
                    continue
                me = dict(e)
                me["qty_natural"] = val
                me["qty_de"] = val / _unit_parts(e["unit"])[1]
                me["amount"] = (round(me["qty_de"] * e["price"], 2)
                                if e["price"] is not None else None)
                mud.append(me)
            notes.append("泥浆制作 3-1-58 / 泥浆运输 3-1-59 随灌注桩混凝土套用，"
                         f"工程量=桩基混凝土量 {_fmt_num(val)} m3。")
        if cage:
            hpb_kg, hrb_kg = cage["cage_parts"]
            notes.append("钻孔桩钢筋笼入孔 3-2-10 随本条目套用，工程量=HPB300+HRB400 钢筋量之和 "
                         f"{_fmt_num(hpb_kg)}+{_fmt_num(hrb_kg)} kg（÷1000 换 t）。")
    else:
        cands = _match_text(item, entries)
        status, picked, top, note = _classify(cands)
        matches = top
        if note:
            notes.append(note)
        checklist = _checklist(item)
        if group:
            notes.append(f"分项「{item['fenxiang']}」无套定额样例规则，已按全库文本匹配。")
        if tpl and not any(str(rr.get("code", "")).split("*")[0].strip() == tpl["code"]
                           for rr in combo):
            combo.append({"code": tpl["code"], "name": tpl["name"], "unit": tpl["unit"],
                          "status": "已匹配",
                          "need": (f"模板量=混凝土量×系数（默认 {tpl['tpl_default']:g}，"
                                   f"页面可改；定额单位 {tpl['unit']}）")})

    # 语义补充说明
    fx = _norm_text(item["fenxiang"])
    if "沥青" in mat and matches and _unit_parts(item.get("mat_unit") or "")[0] == "m3":
        notes.append("沥青混凝土定额按 t 计，清单为 m3，套用需按密度换算（请提供密度）。")
    if re.search(r"KF|TS", mat) and status == "single" and matches:
        if "TS" in matches[0]["entry"]["name"] and "KF" in mat:
            notes.append(f"清单为 KF 型定型产品，定额 {matches[0]['entry']['code']} 为 TS 型制作、"
                         "安装子目，伸缩量规格一致，型式不同请人工确认。")
    if status == "single" and matches:
        m_c = re.search(r"C\d{2}", mat)
        m_e = re.search(r"C\d{2}", matches[0]["entry"]["name"])
        if m_c and m_e and m_c.group() != m_e.group():
            notes.append(f"清单混凝土 {m_c.group()} 与定额 {m_e.group()} 强度不同，"
                         "已按构件/工艺匹配，套用请人工确认或换算。")
    if status == "none" and matches:
        notes.append("以下为库内较接近的定额子目（仅供参考，不得直接套用），"
                     "请按补充清单完善特征后重新匹配。")

    # 定额工程量与合价
    for c in matches:
        mult = _unit_parts(c["entry"]["unit"])[1]
        val = item.get("value") if item.get("value") is not None else 0
        ib = _unit_parts(item.get("mat_unit") or "")[0]
        eb = _unit_parts(c["entry"]["unit"])[0]
        qty_de = val / mult
        c["qty_natural"] = val
        if ib and eb and {ib, eb} == {"kg", "t"} and ib == "kg":
            qty_de = val / 1000 / mult
            c["qty_de"] = qty_de
            c["note_conv"] = f"已按 kg→t 换算（{val:g} kg = {qty_de:g} {c['entry']['unit']}）"
        else:
            c["qty_de"] = qty_de
        c["amount"] = round(qty_de * c["entry"]["price"], 2) if c["entry"]["price"] is not None else None
    if tpl:
        notes.append(f"模板 {tpl['code']} {tpl['name']} 随本条目套用，"
                     f"模板量=混凝土量×系数（默认 {tpl['tpl_default']:g}，页面可改）。")
    return {"idx": idx, "item": item, "status": status, "matches": matches,
            "rule_combo": combo, "checklist": checklist, "notes": notes,
            "mud": mud, "cage": cage, "tpl": tpl}


def match_quotas(items, lib) -> list[dict]:
    entries = lib["entries"]
    rules = lib.get("rules") or {}
    code_index = {}
    for e in entries:
        code_index.setdefault(e["code"], []).append(e)
    return [_match_one(i, it, entries, code_index, rules, items) for i, it in enumerate(items)]


# ---------------------------------------------------------------------------
# AI 说明（DeepSeek，仅解释）
# ---------------------------------------------------------------------------

_AI_SYSTEM_PROMPT = (
    "你是天津桥梁定额匹配的说明助手。你只能依据给定的候选定额文本撰写理由，"
    "禁止编造定额编号、单价或定额库中不存在的内容；不得改变候选列表，"
    "不得推荐候选之外的定额。输出一个 JSON 对象："
    '{"candidates":[{"index":1,"reason":"≤40字","diff":"≤40字"}],"overall":"≤60字"}。'
)

_client = None


def _get_client():
    global _client
    if _client is None:
        from openai import OpenAI

        from config import DEEPSEEK_API_KEY, DEEPSEEK_BASE_URL
        _client = OpenAI(api_key=DEEPSEEK_API_KEY, base_url=DEEPSEEK_BASE_URL)
    return _client


def _parse_json_loose(raw):
    raw = re.sub(r"```(?:json)?\s*(.*?)```", r"\1", raw, flags=re.S)
    s = raw[raw.find("{"):raw.rfind("}") + 1] if "{" in raw and "}" in raw else raw
    return json.loads(s)


def ai_explain(item_desc: str, candidates: list[dict]) -> dict | None:
    """多候选 AI 对比说明 → {"candidates":[...],"overall":...}；失败返回 None。"""
    try:
        from config import DEEPSEEK_MODEL
        cand_text = "\n".join(
            f"{i + 1}. {c['entry']['code']} {c['entry']['name']}（{c['entry']['unit']}，"
            f"单价 {c['entry']['price'] if c['entry']['price'] is not None else '—'}）"
            f" 规则命中：{'；'.join(c.get('reasons') or [])}"
            for i, c in enumerate(candidates))
        user = (f"## 清单条目\n{item_desc}\n"
                f"## 候选定额（规则检索结果）\n{cand_text}\n"
                "请输出 JSON 说明各候选的匹配理由与差异。")
        resp = _get_client().chat.completions.create(
            model=DEEPSEEK_MODEL,
            messages=[{"role": "system", "content": _AI_SYSTEM_PROMPT},
                      {"role": "user", "content": user}],
            temperature=0.2, max_tokens=800, timeout=20)
        raw = resp.choices[0].message.content or ""
        try:
            return _parse_json_loose(raw)
        except Exception:
            resp2 = _get_client().chat.completions.create(
                model=DEEPSEEK_MODEL,
                messages=[{"role": "system", "content": _AI_SYSTEM_PROMPT},
                          {"role": "assistant", "content": raw},
                          {"role": "user", "content": "你上次的输出格式不正确，请只输出一个 JSON 对象。"}],
                temperature=0.2, max_tokens=800, timeout=20)
            return _parse_json_loose(resp2.choices[0].message.content or "")
    except Exception:
        return None


# ---------------------------------------------------------------------------
# 结果导出
# ---------------------------------------------------------------------------

_STATUS_CN = {"single": "唯一匹配", "multi": "多候选", "none": "无匹配"}


def _fmt_num(v):
    """数值 → 无科学计数法的短字符串（Excel 公式/追溯标注用）。"""
    if v is None:
        return ""
    try:
        if float(v) == int(v):
            return str(int(v))
        s = f"{float(v):.6f}".rstrip("0").rstrip(".")
        return s or "0"
    except (TypeError, ValueError):
        return str(v)


def _grade_tag(item, entry_name):
    """混凝土强度等级（工程数量表材料列 C25/C30…）→ 定额名称后备注；
    定额名称已含同等级时不重复标注。"""
    m = re.search(r"C\d{2}", _norm_text(item.get("mat") or ""))
    if m and m.group() not in _norm_text(entry_name or ""):
        return f"{m.group()} 混凝土"
    return ""


def result_rows(results) -> list[dict]:
    rows = []
    for r in results:
        item = r["item"]
        if item.get("kind") == "桩基参数":
            mat_label = f"桩基参数 φ{item['dia']:g}/h{item['depth']:g}m×{item['count']:g}根"
        else:
            mat_label = item.get("mat", "")
        base = {"序号": r["idx"] + 1, "部位": item.get("sheet", ""),
                "分项名": item.get("fenxiang", ""),
                "材料": mat_label,
                "清单数量": (_fmt_num(item.get("value")) if item.get("value") is not None
                             else item.get("count", "")),
                "状态": _STATUS_CN.get(r["status"], r["status"])}
        # single：只输出入选定额一行（参考候选仅页面 expander 展示）
        cands = r["matches"][:1] if r["status"] == "single" else r["matches"]
        if cands:
            for c in cands:
                e = c["entry"]
                g = _grade_tag(item, e["name"])
                rows.append({**base,
                             "定额编号": e["code"],
                             "定额名称": e["name"] + (f"（{g}）" if g else ""),
                             "定额单位": e["unit"],
                             "定额工程量": c.get("qty_de", ""),
                             "单价": e["price"] if e["price"] is not None else "",
                             "合价": c.get("amount", ""),
                             "匹配理由": "；".join(c.get("reasons") or [])})
        else:
            ref = "；".join(f"{c['entry']['code']} {c['entry']['name']}"
                            for c in r.get("matches") or [])
            reason = "；".join(r.get("checklist") or [])
            if ref:
                reason += ("；库内较接近子目（仅供参考，不得直接套用）：" + ref) if reason else ref
            rows.append({**base, "定额编号": "", "定额名称": "", "定额单位": "",
                         "定额工程量": "", "单价": "", "合价": "", "匹配理由": reason})
    return rows


def build_result_excel(results, lib=None, waste_haul_km=5, tpl_ratios=None) -> bytes:
    """导出：
    Sheet1 分部分项工程项目预算计价表——仿表格1013 第二张表格式（两行表头、
    分部/子分部行、子分部内序号重排、单位/工程量/单价/合价/其中人工费）；
    只放最优候选（multi 在名称后标注候选数，无匹配条目不进表）；
    有计算过程的量写成 Excel 公式（工程量/单价换算/合价）。
    Sheet2 候选明细与理由——每候选一行（最优标★），含得分/理由/定额工程量换算。
    Sheet3 补充信息清单——无匹配条目 + 样例组合需补充项。
    lib 可选：传入定额库后，桩基础子分部末尾补 3-10-44 凿除桩顶
    （凿桩头按 0.5m/根计，分部内所有桩合计）+ 1-6-71/1-6-72 弃运
    （建筑垃圾按虚方 ×2，1-6-72 套 (waste_haul_km-1) 次）。
    tpl_ratios：混凝土→模板定额系数覆盖（{系数键: 系数}，缺省用参考表样例比例），
    模板量=混凝土量×系数，公式写入 F 列。
    """
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, Side

    wb = openpyxl.Workbook()
    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")

    # ---------------- Sheet1 计价表 ----------------
    ws = wb.active
    ws.title = "分部分项工程项目预算计价表"
    for col, w in (("A", 7), ("B", 11.5), ("C", 19.8), ("D", 7.2), ("E", 12.8),
                   ("F", 12.5), ("G", 7.3), ("H", 7.5), ("I", 16.7), ("J", 13.2)):
        ws.column_dimensions[col].width = w
    ws.merge_cells("A1:J1")
    ws["A1"] = "分部分项工程项目预算计价表"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].alignment = center
    ws["A2"] = "专业工程名称:"
    ws["H2"] = "金额单位：元"
    for k, v in (("A3", "序号"), ("B3", "编号"), ("C3", "项 目 名 称"),
                 ("E3", "单位"), ("F3", "工程量"), ("G3", "金  额")):
        ws[k] = v
    ws["G4"], ws["I4"], ws["J4"] = "单价", "合价", "其中：人工费"
    for rng in ("A3:A4", "B3:B4", "C3:D4", "E3:E4", "F3:F4", "G3:H3"):
        ws.merge_cells(rng)
    for cell in ("A3", "B3", "C3", "E3", "F3", "G3", "G4", "I4", "J4"):
        ws[cell].alignment = center

    def style_row(row, end="J"):
        for col in "ABCDEFGHIJ"[: "ABCDEFGHIJ".index(end) + 1]:
            ws[f"{col}{row}"].border = border

    # 基础工程（土石方→桩基础→承台）；下部结构只留 桥台/台盖梁/桥墩/墩盖梁
    _BASE_SUBS = {"土方", "桩基础", "承台"}
    _LOWER_SUBS = {"桥台", "台盖梁", "桥墩", "墩盖梁"}
    # 子分部顺序对齐参考表 sheet2 + 基础工程新分部（土方→桩基础→承台→桥台→台盖梁→桥墩→墩盖梁；
    # 上部：箱梁→小箱梁→板梁→桥面铺装→搭板→护栏→支座→伸缩缝→排水）
    _SUB_ORDER = ["土方", "桩基础", "承台", "桥台", "台盖梁", "桥墩", "墩盖梁",
                  "预应力砼现浇箱梁", "预应力混凝土小箱梁", "预应力板梁",
                  "桥面铺装", "搭板", "防撞护栏", "支座", "伸缩缝", "排水"]

    def _div_of(it):
        sub = _map_group(it.get("fenxiang", "")) or ""
        if _norm_text(it.get("fenxiang", "")).startswith("梯道下部"):
            return "梯道下部"
        if sub in _BASE_SUBS:
            return "基础工程"
        if sub in _LOWER_SUBS:
            return "下部结构"
        return "上部结构"

    # 凿桩头 3-10-44：按 0.5m/根 计，分部内所有桩合计（体积=桩数×0.5×π/4×φ²）
    head_entry = None
    if lib:
        head_entry = next((e for e in lib.get("entries") or []
                           if str(e["code"]) == "3-10-44"), None)
    # 凿出的桩头属建筑垃圾：弃运套 1-6-71（运距1km）+ 1-6-72（每增运1km）。
    # 建筑垃圾（石/混凝土类）弃运量按虚方 ×2；1-6-72 套 (运距-1) 次（运距 _HAUL_KM km）
    waste_entries = []
    if lib:
        for code, nm_kw in (("1-6-71", "运(弃)土、废料 运距1km"),
                            ("1-6-72", "运(弃)土、废料 每增运1km")):
            e = next((x for x in lib.get("entries") or []
                      if str(x["code"]) == code), None)
            if e is None:
                rr = next((x for x in (lib.get("rules") or {}).get("桩基础") or []
                           if str(x["code"]) == code), None)
                e = {"code": code, "name": rr["name"] if rr else nm_kw,
                     "unit": rr["unit"] if rr else "m3", "price": None, "row": 0}
            waste_entries.append(e)
    # 模板系数（模板量=混凝土量×系数）：默认参考表样例比例，页面可改
    ratios = dict(TPL_RATIOS)
    if tpl_ratios:
        ratios.update({k: float(v) for k, v in tpl_ratios.items() if v is not None})
    head_by_div = {}
    last_pile_by_div = {}
    if head_entry is not None:
        for r in results:
            it = r["item"]
            if it.get("kind") != "桩基参数":
                continue
            div = _div_of(it)
            h = head_by_div.setdefault(div, {"parts": [], "total": 0.0})
            v = it["count"] * 0.5 * 3.14159265 / 4 * (it["dia"] ** 2)
            h["parts"].append((it["count"], it["dia"]))
            h["total"] += v
            last_pile_by_div[div] = r["idx"]

    ordered = []  # [(div, sub, r)]：大分部在前、分部内按参考表子分部顺序聚组
    for r in results:
        if r["status"] == "none" or not r["matches"]:
            continue  # 无匹配条目不进计价表（见 补充信息清单）
        it = r["item"]
        sub = _map_group(it.get("fenxiang", "")) or (it.get("fenxiang", "") or "其他")
        div = _div_of(it)
        ordered.append((div, sub, r))
    ordered.sort(key=lambda x: ({"基础工程": 0, "下部结构": 1, "上部结构": 2,
                                 "梯道下部": 3}[x[0]],
                                _SUB_ORDER.index(x[1]) if x[1] in _SUB_ORDER else 99))

    def _write_head(div):
        """桩基础子分部收尾：写 3-10-44 凿除桩顶合计行 + 1-6-71/1-6-72 弃运行
        （每分部一次；建筑垃圾弃运按虚方 ×2，1-6-72 套 (运距-1) 次）。"""
        nonlocal row
        h = head_by_div.get(div)
        if head_entry is None or not h or h.get("done"):
            return
        h["done"] = True
        cnt_terms = "+".join(f"{c:g}×φ{d:g}" for c, d in h["parts"])
        vol_terms = "+".join(f"{c:g}*0.5*3.14159265/4*{d:g}^2" for c, d in h["parts"])
        hm = _unit_parts(head_entry["unit"])[1]
        ws[f"B{row}"] = head_entry["code"]
        ws[f"C{row}"] = (head_entry["name"]
                         + f"〔凿桩头按 0.5m/根计，本分部桩合计：{cnt_terms}〕")
        ws.merge_cells(f"C{row}:D{row}")
        ws[f"E{row}"] = head_entry["unit"]
        ws[f"F{row}"] = f"=({vol_terms})/{hm:g}"
        if head_entry["price"] is not None:
            ws[f"G{row}"] = round(head_entry["price"], 4)
            ws[f"I{row}"] = f"=F{row}*G{row}"
        style_row(row)
        row += 1
        # 凿出的桩头属建筑垃圾：1-6-71 弃运1km、1-6-72 每增运1km（弃运量×2 虚方）
        for k, we in enumerate(waste_entries):
            if k == 1 and waste_haul_km <= 1:
                continue  # 运距≤1km 不套每增运1km
            wm = _unit_parts(we["unit"])[1]
            if k == 0:
                wf = f"=({vol_terms})*2"
                note = (f"〔凿桩头弃运×2（建筑垃圾按虚方）；运距 {waste_haul_km:g}km"
                        + ("；单价库中无此子目，需补充" if we["price"] is None else "")
                        + "〕")
            else:
                wf = f"=({vol_terms})*2*{waste_haul_km - 1:g}"
                note = (f"〔每增运1km，套 {waste_haul_km - 1:g} 次"
                        + ("；单价库中无此子目，需补充" if we["price"] is None else "")
                        + "〕")
            if wm != 1:
                wf += f"/{wm:g}"
            ws[f"B{row}"] = we["code"]
            ws[f"C{row}"] = we["name"] + note
            ws.merge_cells(f"C{row}:D{row}")
            ws[f"E{row}"] = we["unit"]
            ws[f"F{row}"] = wf
            if we["price"] is not None:
                ws[f"G{row}"] = round(we["price"], 4)
                ws[f"I{row}"] = f"=F{row}*G{row}"
            style_row(row)
            row += 1

    row = 5
    last_div = None
    last_sub = None
    for div, sub, r in ordered:
        it = r["item"]
        if last_div != div:
            if last_div is not None:
                _write_head(last_div)
            ws[f"C{row}"] = div
            style_row(row)
            row += 1
            last_div = div
            last_sub = None
        if last_sub != sub:
            if last_sub == "桩基础":
                _write_head(div)  # 桩基础子分部收尾：凿除桩顶合计行
            ws[f"C{row}"] = sub
            style_row(row)
            row += 1
            last_sub = sub
        c = r["matches"][0]
        e = c["entry"]
        mult = _unit_parts(e["unit"])[1]
        ws[f"B{row}"] = e["code"]
        name = e["name"]
        tags = []
        if r["status"] == "multi" and len(r["matches"]) > 1:
            tags.append(f"另有{len(r['matches']) - 1}个候选")
        # 混凝土强度等级（工程数量表材料列）：定额名称未含该等级时备注在名称后
        g = _grade_tag(it, e["name"])
        if g:
            tags.append(g)
        # 清单原值追溯标注
        if it.get("kind") == "桩基参数":
            tags.append(f"清单 {_fmt_num(it['count'])}根×φ{_fmt_num(it['dia'])}"
                        f"×h{_fmt_num(it['depth'])}m")
        elif it.get("value") is not None:
            tags.append(f"清单 {_fmt_num(it['value'])} {_norm_text(it.get('mat_unit') or '')}".rstrip())
        if e.get("price") is None:
            tags.append("单价库中无此子目，需补充")
        if tags:
            name += "〔" + "；".join(tags) + "〕"
        ws[f"C{row}"] = name
        ws.merge_cells(f"C{row}:D{row}")
        ws[f"E{row}"] = e["unit"]  # 单位与定额单位一致
        # 工程量（定额单位口径）：有计算过程的写成公式
        if it.get("kind") == "桩基参数":
            ws[f"F{row}"] = f"={_fmt_num(it['count'])}*{_fmt_num(it['depth'])}/{mult:g}"
        elif c.get("note_conv"):  # kg→t
            ws[f"F{row}"] = (f"={_fmt_num(it['value'])}/1000/{mult:g}" if mult != 1
                             else f"={_fmt_num(it['value'])}/1000")
        elif mult != 1 and it.get("value") is not None:
            ws[f"F{row}"] = f"={_fmt_num(it['value'])}/{mult:g}"
        else:
            v = it.get("value")
            ws[f"F{row}"] = round(v, 4) if v is not None else 0
        # 单价 = 定额原单价
        price = e["price"]
        if price is not None:
            ws[f"G{row}"] = round(price, 4)
            ws[f"I{row}"] = f"=F{row}*G{row}"
        style_row(row)
        row += 1
        # 桩基混凝土条目：泥浆制作/运输随同套用（工程量=桩基混凝土量）
        for me in r.get("mud") or []:
            mmult = _unit_parts(me["unit"])[1]
            ws[f"B{row}"] = me["code"]
            ws[f"C{row}"] = (me["name"]
                             + f"〔工程量=桩基混凝土 {_fmt_num(me['qty_natural'])} m3〕")
            ws.merge_cells(f"C{row}:D{row}")
            ws[f"E{row}"] = me["unit"]
            ws[f"F{row}"] = f"={_fmt_num(me['qty_natural'])}/{mmult:g}"
            if me["price"] is not None:
                ws[f"G{row}"] = round(me["price"], 4)
                ws[f"I{row}"] = f"=F{row}*G{row}"
            style_row(row)
            row += 1
        # 桩基钢筋条目（HPB300 行携带）：钻孔桩钢筋笼入孔 3-2-10，
        # 工程量=同部位 HPB300+HRB400 钢筋量之和（kg→t）
        cage = r.get("cage")
        if cage:
            cmult = _unit_parts(cage["unit"])[1]
            hpb_kg, hrb_kg = cage.get("cage_parts", (0, 0))
            ws[f"B{row}"] = cage["code"]
            ws[f"C{row}"] = (cage["name"]
                             + f"〔工程量=HPB300+HRB400 钢筋量 "
                               f"{_fmt_num(hpb_kg)}+{_fmt_num(hrb_kg)} kg〕")
            ws.merge_cells(f"C{row}:D{row}")
            ws[f"E{row}"] = cage["unit"]
            ws[f"F{row}"] = (f"=({_fmt_num(hpb_kg)}+{_fmt_num(hrb_kg)})/1000/{cmult:g}"
                             if cmult != 1
                             else f"=({_fmt_num(hpb_kg)}+{_fmt_num(hrb_kg)})/1000")
            if cage["price"] is not None:
                ws[f"G{row}"] = round(cage["price"], 4)
                ws[f"I{row}"] = f"=F{row}*G{row}"
            style_row(row)
            row += 1
        # 混凝土条目：相应模板定额随同套用（模板量=混凝土量×系数，默认参考表比例、页面可改）
        tpl = r.get("tpl")
        if tpl and tpl.get("conc_val") is not None:
            tmult = _unit_parts(tpl["unit"])[1]
            ratio = ratios.get(tpl.get("tpl_key"), tpl.get("tpl_default", 0))
            ratio_changed = abs(ratio - tpl.get("tpl_default", 0)) > 1e-9
            ws[f"B{row}"] = tpl["code"]
            ws[f"C{row}"] = (tpl["name"]
                             + f"〔模板系数 {ratio:g}"
                             + (f"（默认 {tpl['tpl_default']:g}）" if ratio_changed else "")
                             + "；模板量=混凝土量×系数〕")
            ws.merge_cells(f"C{row}:D{row}")
            ws[f"E{row}"] = tpl["unit"]
            ws[f"F{row}"] = f"={_fmt_num(tpl['conc_val'])}*{ratio:g}/{tmult:g}"
            if tpl["price"] is not None:
                ws[f"G{row}"] = round(tpl["price"], 4)
                ws[f"I{row}"] = f"=F{row}*G{row}"
            style_row(row)
            row += 1
        # 桩基参数：钻孔行后再套一行陆地埋设钢护筒（每根桩 2m，工程量=桩数×2）
        liner = r.get("pile_liner")
        if it.get("kind") == "桩基参数" and liner:
            lmult = _unit_parts(liner["unit"])[1]
            ws[f"B{row}"] = liner["code"]
            ws[f"C{row}"] = (liner["name"]
                             + f"〔清单 {_fmt_num(it['count'])}根×2m 护筒〕")
            ws.merge_cells(f"C{row}:D{row}")
            ws[f"E{row}"] = liner["unit"]
            ws[f"F{row}"] = f"={_fmt_num(it['count'])}*2/{lmult:g}"
            if liner["price"] is not None:
                ws[f"G{row}"] = round(liner["price"], 4)
                ws[f"I{row}"] = f"=F{row}*G{row}"
            style_row(row)
            row += 1
    # 最后一个分部若以桩基础收尾：补凿除桩顶合计行
    if last_div is not None and last_sub == "桩基础":
        _write_head(last_div)
    # 序号：子分部内从 1 重排（B 列有编号的是明细行）
    seq = 0
    for rr in range(5, row):
        if ws[f"B{rr}"].value:
            seq += 1
            ws[f"A{rr}"] = seq
        else:
            seq = 0

    # ---------------- Sheet2 候选明细与理由 ----------------
    ws2 = wb.create_sheet("候选明细与理由")
    headers2 = ["序号", "部位", "分项名", "材料", "清单数量", "状态", "候选",
                "定额编号", "定额名称", "定额单位", "定额工程量", "定额单价",
                "合价", "得分", "匹配理由"]
    ws2.append(headers2)
    for col, w in (("A", 6), ("B", 12), ("C", 14), ("D", 22), ("E", 10),
                   ("F", 9), ("G", 8), ("H", 11), ("I", 40), ("J", 9),
                   ("K", 12), ("L", 10), ("M", 12), ("N", 6), ("O", 50)):
        ws2.column_dimensions[col].width = w
    r2 = 2
    for r in results:
        it = r["item"]
        if it.get("kind") == "桩基参数":
            mat_label = f"桩基参数 φ{it['dia']:g}/h{it['depth']:g}m×{it['count']:g}根"
        else:
            mat_label = it.get("mat", "")
        qty_label = (_fmt_num(it.get("value")) if it.get("value") is not None
                     else _fmt_num(it.get("count")) if it.get("count") is not None else "")
        base_row = [r["idx"] + 1, it.get("sheet", ""), it.get("fenxiang", ""),
                    mat_label, qty_label,
                    _STATUS_CN.get(r["status"], r["status"])]
        if r["matches"]:
            for j, c in enumerate(r["matches"]):
                e = c["entry"]
                mult = _unit_parts(e["unit"])[1]
                ws2.append(base_row + ["★最优" if j == 0 else "",
                                       e["code"], e["name"], e["unit"],
                                       None, e["price"], None,
                                       c.get("score", ""),
                                       "；".join(c.get("reasons") or [])])
                # 定额工程量（扩大单位）：有计算过程的写成公式
                if it.get("kind") == "桩基参数":
                    ws2[f"K{r2}"] = f"={_fmt_num(it['count'])}*{_fmt_num(it['depth'])}/{mult:g}"
                elif c.get("note_conv"):
                    ws2[f"K{r2}"] = (f"={_fmt_num(it['value'])}/1000/{mult:g}" if mult != 1
                                     else f"={_fmt_num(it['value'])}/1000")
                else:
                    ws2[f"K{r2}"] = (f"={_fmt_num(it['value'])}/{mult:g}" if mult != 1
                                     else round(it["value"], 4) if it.get("value") is not None else 0)
                if e["price"] is not None:
                    ws2[f"M{r2}"] = f"=K{r2}*L{r2}"
                r2 += 1
            liner = r.get("pile_liner")
            if it.get("kind") == "桩基参数" and liner:
                lmult = _unit_parts(liner["unit"])[1]
                ws2.append(base_row + ["★套用",
                                       liner["code"], liner["name"], liner["unit"],
                                       None, liner["price"], None, "",
                                       "陆地埋设钢护筒：每根桩 2m，工程量=桩数×2m"])
                ws2[f"K{r2}"] = f"={_fmt_num(it['count'])}*2/{lmult:g}"
                if liner["price"] is not None:
                    ws2[f"M{r2}"] = f"=K{r2}*L{r2}"
                r2 += 1
            if (head_entry is not None and it.get("kind") == "桩基参数"
                    and r["idx"] == last_pile_by_div.get(_div_of(it))):
                h = head_by_div.get(_div_of(it))
                if h and h.get("parts"):
                    cnt_terms = "+".join(f"{c:g}×φ{d:g}" for c, d in h["parts"])
                    vol_terms = "+".join(f"{c:g}*0.5*3.14159265/4*{d:g}^2"
                                         for c, d in h["parts"])
                    hm = _unit_parts(head_entry["unit"])[1]
                    ws2.append(base_row + ["★套用",
                                           head_entry["code"], head_entry["name"],
                                           head_entry["unit"], None, head_entry["price"],
                                           None, "",
                                           "凿桩头按 0.5m/根计，本分部所有桩合计套用一次"
                                           f"（{cnt_terms}）"])
                    ws2[f"K{r2}"] = f"=({vol_terms})/{hm:g}"
                    if head_entry["price"] is not None:
                        ws2[f"M{r2}"] = f"=K{r2}*L{r2}"
                    r2 += 1
                    # 凿出的桩头属建筑垃圾：弃运 1-6-71 ×2、1-6-72 ×2×(运距-1)
                    for k, we in enumerate(waste_entries):
                        if k == 1 and waste_haul_km <= 1:
                            continue
                        wm = _unit_parts(we["unit"])[1]
                        if k == 0:
                            wf = f"=({vol_terms})*2"
                            reason = (f"凿桩头弃运：×2（建筑垃圾按虚方），运距 {waste_haul_km:g}km"
                                      + ("；单价库中无此子目，需补充"
                                         if we["price"] is None else ""))
                        else:
                            wf = f"=({vol_terms})*2*{waste_haul_km - 1:g}"
                            reason = (f"每增运1km：套 {waste_haul_km - 1:g} 次"
                                      f"（运距 {waste_haul_km:g}km）"
                                      + ("；单价库中无此子目，需补充"
                                         if we["price"] is None else ""))
                        if wm != 1:
                            wf += f"/{wm:g}"
                        ws2.append(base_row + ["★套用",
                                               we["code"], we["name"], we["unit"],
                                               None, we["price"], None, "", reason])
                        ws2[f"K{r2}"] = wf
                        if we["price"] is not None:
                            ws2[f"M{r2}"] = f"=K{r2}*L{r2}"
                        r2 += 1
            for me in r.get("mud") or []:
                mmult = _unit_parts(me["unit"])[1]
                ws2.append(base_row + ["★套用",
                                       me["code"], me["name"], me["unit"],
                                       None, me["price"], None, "",
                                       "随灌注桩混凝土套用：工程量=桩基混凝土量（÷10 换算）"])
                ws2[f"K{r2}"] = f"={_fmt_num(me['qty_natural'])}/{mmult:g}"
                if me["price"] is not None:
                    ws2[f"M{r2}"] = f"=K{r2}*L{r2}"
                r2 += 1
            cage = r.get("cage")
            if cage:
                cmult = _unit_parts(cage["unit"])[1]
                hpb_kg, hrb_kg = cage.get("cage_parts", (0, 0))
                ws2.append(base_row + ["★套用",
                                       cage["code"], cage["name"], cage["unit"],
                                       None, cage["price"], None, "",
                                       "钢筋笼入孔随桩基钢筋套用：工程量=HPB300+HRB400 钢筋量之和（÷1000 换 t）"])
                ws2[f"K{r2}"] = (f"=({_fmt_num(hpb_kg)}+{_fmt_num(hrb_kg)})/1000/{cmult:g}"
                                 if cmult != 1
                                 else f"=({_fmt_num(hpb_kg)}+{_fmt_num(hrb_kg)})/1000")
                if cage["price"] is not None:
                    ws2[f"M{r2}"] = f"=K{r2}*L{r2}"
                r2 += 1
            tpl = r.get("tpl")
            if tpl and tpl.get("conc_val") is not None:
                tmult = _unit_parts(tpl["unit"])[1]
                ratio = ratios.get(tpl.get("tpl_key"), tpl.get("tpl_default", 0))
                ratio_changed = abs(ratio - tpl.get("tpl_default", 0)) > 1e-9
                ws2.append(base_row + ["★套用",
                                       tpl["code"], tpl["name"], tpl["unit"],
                                       None, tpl["price"], None, "",
                                       "模板随混凝土套用：模板量=混凝土量×系数 "
                                       f"{ratio:g}"
                                       + (f"（默认 {tpl['tpl_default']:g}，页面可改）"
                                          if ratio_changed else "（页面可改）")])
                ws2[f"K{r2}"] = f"={_fmt_num(tpl['conc_val'])}*{ratio:g}/{tmult:g}"
                if tpl["price"] is not None:
                    ws2[f"M{r2}"] = f"=K{r2}*L{r2}"
                r2 += 1
        else:
            ref = "；".join(f"{c['entry']['code']} {c['entry']['name']}"
                            for c in r.get("matches") or [])
            reason = "；".join(r.get("checklist") or [])
            if ref:
                reason += ("；库内较接近子目（仅供参考，不得直接套用）：" + ref) if reason else ref
            ws2.append(base_row + ["", "", "", "", "", "", "", "", reason])
            r2 += 1

    # ---------------- Sheet3 补充信息清单 ----------------
    ws3 = wb.create_sheet("补充信息清单")
    ws3.append(["序号", "部位", "分项名", "材料", "缺失维度/补充说明"])
    for col, w in (("A", 6), ("B", 12), ("C", 14), ("D", 24), ("E", 60)):
        ws3.column_dimensions[col].width = w
    for r in results:
        if r["status"] == "none":
            ws3.append([r["idx"] + 1, r["item"].get("sheet", ""),
                        r["item"].get("fenxiang", ""), r["item"].get("mat", ""),
                        "；".join(r.get("checklist") or [])])
        for rr in r.get("rule_combo") or []:
            if rr.get("status") == "需补充" and rr.get("need"):
                ws3.append([r["idx"] + 1, r["item"].get("sheet", ""),
                            r["item"].get("fenxiang", ""), rr["code"] + " " + rr["name"],
                            rr["need"]])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
