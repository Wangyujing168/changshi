# -*- coding: utf-8 -*-
"""概算成果压缩包生成（总概算表 Excel + 编制说明 Word + 成套概算 PDF）。

流程：
1. python-docx 打开内置 docx 模板（由用户提供的 .doc 一次性转换而来）：
   - 天津/其他省份：templates/deliverable_天津.docx（封面3页+目录+编制说明）
   - 河北/雄安：天津封面+目录 与 templates/deliverable_雄安.docx 的编制说明合并
2. 按项目信息与二类费计算结果填充：封面名称/编号/日期、工程概况、编制依据
   （省份相关）、三、工程建设其他费用（按实际计算费种动态生成，只写费用+计取
   标准不写金额，单倍行距）、预备费率、资金来源、工程投资数字。
3. Word COM：重新分页 → 查「编制说明」起止页 → 更新目录页码 → 另存带真实
   页码的 Word 文件（交付②）→ 导出 PDF。
4. Excel COM：概算表主表（总概算表/费用汇总/项目汇总表）复制到新工作簿、
   横向缩放一页宽 → 导出 PDF。
5. pypdf 按序合并：docx PDF + Excel PDF + 广联达源 PDF（交付③）。
6. zip 打包：①总概算表 Excel（原 xlsx 全 sheet）②编制说明 Word ③成套 PDF。

依赖：python-docx / pypdf（纯 Python）+ 运行时 Word/Excel COM（用户机器）。
"""
import copy
import os
import re
import shutil
import tempfile
import zipfile
from io import BytesIO

from docx import Document
from docx.enum.text import (WD_ALIGN_PARAGRAPH, WD_LINE_SPACING,
                            WD_TAB_ALIGNMENT, WD_TAB_LEADER)
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.text.paragraph import Paragraph

_TEMPLATE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates")
_TJ_TEMPLATE = os.path.join(_TEMPLATE_DIR, "deliverable_天津.docx")
_XA_TEMPLATE = os.path.join(_TEMPLATE_DIR, "deliverable_雄安.docx")

# 河北/雄安 项目用雄安编制说明结构，其余省份用天津结构
XA_REGIONS = ("河北", "雄安")

# ---- 三、工程建设其他费用：模板语句（key=引擎费种名，值为完整句子含标点） ----
FEE_SENTENCE_TJ = {
    "场地准备费及临时设施费": "场地准备及建设单位临时设施费：按照第一部分工程费的0.5%计取；",
    "建设管理费": "建设单位管理费：按财政部财建[2016]504号计取；",
    "监理费": "工程监理费：参考国家发改委发改价格[2007]670号文，下浮20%计取；",
    "可行性研究费": "项目前期工作费：按合同价格计取；",
    "勘察费": "勘察费：按合同价格计取；",
    "工程设计费": "设计费：参照计价格[2002]10号文，下浮20%计取；",
    "施工图审查费": "施工图审查费：按勘察设计费的3.2%计取；",
    "招标代理费": "招标代理服务费：根据计价格[2002]1980号文计取；",
    "造价咨询费": "工程造价咨询费：参照津价房地[2008]136号计取；",
    "环境影响咨询费": "环境影响评价费：按计价格〔2002〕125号文计取；",
    "工程保险费": "工程保险费：按第一部分工程费的0.3%计算计取；",
    "水土保持补偿费": "水土保持补偿费：按津发改价综[2020]351号文计取；",
    "水土保持费": "水土保持咨询费：参考保监[2005]22号，按暂估计取；",
}

FEE_SENTENCE_XA = {
    "建设管理费": "项目建设管理费：根据财建〔2016〕504号文《基本建设财务管理规定》计取；",
    "监理费": "工程建设监理费：参考冀建市研[2016]6号文，执行发改价格[2015]299号文按市场价计算；",
    "环境影响咨询费": "环境影响评价报告编制费：按合同价计入；",
    "水土保持费": "水土保持方案编制费：按合同价计入；",
    "勘察费": "勘察费：按建安费×1%计入；",
    "工程设计费": "设计费：按合同价计入；",
    "造价咨询费": "建设工程造价咨询服务费：参考冀建市研〔2017〕2号文《河北省工程造价咨询"
                  "服务收费管理暂行办法》，执行发改价格[2015]299号文按市场价计算；",
    "场地准备费及临时设施费": "场地准备及临时设施费：按建安费的0.5%计列。",
    "工程保险费": "工程保险费：根据建标[2011]1号的规定，以建安费的0.3%计列。",
    "招标代理费": "招标代理服务费：按照国家计委计价格〔2002〕1980号文及发改价格[2011]534号文计列；",
    "施工图审查费": "施工图审查费：按勘察设计费的6.5%计列；",
    "水土保持补偿费": "水土保持补偿费：参考冀价行费[2017]173号文计列；",
}


def _is_xa(region: str) -> bool:
    r = str(region or "")
    return any(x in r for x in XA_REGIONS)


def fee_sentence(region: str, name: str, label: str, yiju: str) -> str:
    """费种 → 编制说明语句（模板语句优先，缺省按引擎依据拼）。"""
    table = FEE_SENTENCE_XA if _is_xa(region) else FEE_SENTENCE_TJ
    if name in table:
        return table[name]
    yiju = (yiju or "").replace("\n", " ").strip()
    if not yiju or yiju in ("请查阅相关政策文件", "用户自定义金额"):
        yiju = "合同价" if name == "自定义" else "相关规定"
    if yiju.startswith(("按", "参照", "依据", "根据", "以", "参考")):
        return f"{label}：{yiju}计取；"
    return f"{label}：按{yiju}计取；"


# ---- python-docx 工具 ----

def _paras(doc, pred):
    return [p for p in doc.paragraphs if pred(p.text.strip())]


def _set_para_text(p, text):
    """整段替换文本（保留首 run 字体格式）。"""
    runs = p.runs
    if not runs:
        p.add_run(text)
        return
    runs[0].text = text
    for r in runs[1:]:
        r.text = ""


def _delete_para(p):
    p._element.getparent().remove(p._element)


def _fill_between(doc, start_pred, end_pred, lines, style_name, single_spacing=False):
    """删除两个锚定标题之间的段落，并在 end 段前按 style 插入 lines。

    single_spacing=True 时对新段显式写单倍行距+段前段后0。
    注意：保持文档网格对齐（不写 snapToGrid=0）——模板 docGrid
    linePitch≈31pt，一/二节正文都在该网格上，三节对齐网格才能与
    前面版式一致（关掉网格反而会比一/二节紧、显得不协调）。
    """
    ps = doc.paragraphs
    start_i = next(i for i, p in enumerate(ps) if start_pred(p.text.strip()))
    end_i = next(i for i, p in enumerate(ps) if i > start_i and end_pred(p.text.strip()))
    for p in ps[start_i + 1:end_i]:
        _delete_para(p)
    anchor = next(p for p in doc.paragraphs if end_pred(p.text.strip()))
    if not lines:
        lines = ["本项目无其他计取的工程建设其他费用。"]
    for line in lines:
        p = anchor.insert_paragraph_before(line, style=style_name)
        # 样式自带首行缩进两字（江建权/jwang0ZW 均为 w:firstLineChars="200"）：
        # 动态生成的内容一律顶格写，与二节正文一致。注意必须同时清零
        # firstLine 与 firstLineChars——Word 中 firstLineChars 优先级高于
        # firstLine，只清 firstLine 仍会继承样式的两字缩进
        ind = p._p.get_or_add_pPr().get_or_add_ind()
        ind.set(qn("w:firstLine"), "0")
        ind.set(qn("w:firstLineChars"), "0")
        if single_spacing:
            pf = p.paragraph_format
            pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)


def _heading_para(doc, heading_text):
    """定位标题段；其后第一个非空段落（用于工程概况等单段替换）。"""
    ps = doc.paragraphs
    hi = next(i for i, p in enumerate(ps) if p.text.strip() == heading_text)
    for p in ps[hi + 1:]:
        if p.text.strip():
            return p
    return None


# ---- 模板填充 ----

# 封面第3页签署栏（自上而下 6 项）
SIGN_LABELS = ["审定人", "审核人", "校核人", "设计项目负责人", "概算专业负责人", "编制人"]

# 封面项目名称：28pt 黑体，版心宽 = 21 - 3.5 - 2.0 = 15.5cm ≈ 439pt，
# 28pt 全角字每字 28pt → 每行最多 15 个全角字宽（半角按 0.5 计）
_COVER_NAME_MAX_W = 15.0


def _char_w(c):
    """字符显示宽度（全角 1.0 / 半角 0.5）。"""
    return 1.0 if ord(c) > 0x2E7F else 0.5


def _disp_len(s):
    return sum(_char_w(c) for c in s)


def _split_project_name(name, max_w=_COVER_NAME_MAX_W):
    """项目名称自动换行：行1 尽量长（≥行2），每行不超过 max_w 个全角字宽。

    规则：总宽超过一行时，从行首累计越过一半即切行（偶数总宽时行1 比行2 长 1）；
    若两行仍放不下，行1/行2 各取满一行，剩余放行3。
    """
    name = re.sub(r"\s+", "", (name or "").strip())
    if not name:
        return []
    total = _disp_len(name)
    if total <= max_w:
        return [name]
    target = (total + 1) / 2
    acc, i = 0.0, 0
    while i < len(name):
        w = _char_w(name[i])
        if acc + w > max_w + 1e-9:
            break
        acc += w
        i += 1
        if acc >= target - 1e-9:
            break
    if i <= 0:
        i = 1
    rest = name[i:]
    if _disp_len(rest) > max_w + 1e-9:
        acc2, j = 0.0, i
        while j < len(name) and acc2 + _char_w(name[j]) <= max_w + 1e-9:
            acc2 += _char_w(name[j])
            j += 1
        return [name[:i], name[i:j], name[j:]]
    return [name[:i], rest]


def _distribute_label(label, target):
    """标签两端对齐：在字符间均分全角空格，补到 target 个全角字宽。"""
    extra = int(target) - len(label)
    if extra <= 0 or len(label) < 2:
        return label
    base, rem = divmod(extra, len(label) - 1)
    parts = []
    for i, ch in enumerate(label):
        parts.append(ch)
        if i < len(label) - 1:
            parts.append("　" * (base + (1 if i < rem else 0)))
    return "".join(parts)


def _add_sign_run(p, text, underline, font_pt):
    """签署栏 run：eastAsia 字体（与模板一致），可带下划线。"""
    r = p.add_run(text)
    r.font.size = Pt(font_pt)
    r.font.underline = underline
    rpr = r._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.insert(0, rf)
    rf.set(qn("w:hint"), "eastAsia")
    return r


# 签署栏几何（两页共用）：下划线 12 个全角空格；版心宽 = 21cm − 左3.5cm − 右2cm
# = 15.5cm ≈ 439.4pt；块宽 =（标签宽 + 冒号 1em + 12em）× 字号 → 左缩进 = (439.4−块宽)/2，
# 左右边距相等 → 字和横线整块在版心居中（与上方名称/概算书/编号的居中对齐一致）
_SIGN_LINE_CHARS = 12
_SIGN_TEXT_W_PT = 15.5 * 72 / 2.54
# 第3页：标签两端对齐到 7 全角字 + 冒号 + 横线，15pt → 块宽 20em×15 = 300pt
_SIGN3_INDENT_PT = (_SIGN_TEXT_W_PT - (7 + 1 + _SIGN_LINE_CHARS) * 15) / 2  # ≈ 69.7pt
# 第2页：标签两端对齐到 6 全角字 + 冒号 + 横线，16pt → 块宽 19em×16 = 304pt
_P2_LABEL_W = 6
_P2_FONT_PT = 16
_P2_INDENT_PT = (_SIGN_TEXT_W_PT - (_P2_LABEL_W + 1 + _SIGN_LINE_CHARS) * _P2_FONT_PT) / 2  # ≈ 67.7pt


def _rewrite_sign_para(p, text, font_pt, indent_pt):
    """签署栏段落重排：固定左缩进使整块（标签+下划线）在版心居中。

    不用 jc=center：Word 段落居中不计行尾空格（整块会偏），且 pPr 元素
    顺序放错时 Word 会整体忽略 jc。改用左对齐 + 固定左缩进：全角字符宽度
    精确（1em），各行同缩进同线长 → 左端（标签同宽）右端（线尾）全齐、
    横线等长、整块居中。下划线为全角空格 U+3000（=1em 精确）。
    """
    ppr = p._p.get_or_add_pPr()
    for tag in ("w:ind", "w:tabs", "w:jc"):
        for el in ppr.findall(qn(tag)):
            ppr.remove(el)
    pf = p.paragraph_format  # python-docx API 按 schema 顺序插入，不会再被 Word 忽略
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.left_indent = Pt(indent_pt)
    for child in list(p._p):
        if child.tag != qn("w:pPr"):
            p._p.remove(child)
    _add_sign_run(p, text, underline=False, font_pt=font_pt)
    _add_sign_run(p, "　" * _SIGN_LINE_CHARS, underline=True, font_pt=font_pt)


def _set_cover_name(doc, lines):
    """封面三页各一组项目名称段：行1→第一段、行2→第二段、行3+插在第二段后。"""
    firsts = _paras(doc, lambda t: t == "创新道（富兴路—智兴路）")
    seconds = _paras(doc, lambda t: t == "道路及配套管线工程")
    for p1, p2 in zip(firsts, seconds):
        _set_para_text(p1, lines[0] if lines else "　")
        _set_para_text(p2, lines[1] if len(lines) > 1 else "　")
        for extra in lines[2:]:
            new_el = copy.deepcopy(p2._element)
            p2._element.addnext(new_el)
            _set_para_text(Paragraph(new_el, p2._parent), extra)


# 目录条目：标题左对齐，页码经右对齐点前导制表位落到版心右缘
_TOC_ENTRIES = [("编制说明", 1), ("总概算表", 8), ("单位工程概算表", 11)]


def _toc_entry_para(p, label, page):
    """目录条目段：清缩进/旧制表位，加右对齐点前导制表位，写「标题\\t页码」。"""
    ppr = p._p.get_or_add_pPr()
    for tag in ("w:ind", "w:tabs"):
        for el in ppr.findall(qn(tag)):
            ppr.remove(el)
    p.paragraph_format.tab_stops.add_tab_stop(
        Pt(_SIGN_TEXT_W_PT), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    _set_para_text(p, f"{label}\t{page}")


def _rebuild_toc(doc):
    """目录页重建：每行一条目（标题左对齐 + 页码右对齐点前导），页码为占位值，
    由 Word COM 目录更新（_word_docx_to_pdf）在导出时写入真实页码。"""
    toc_title = next((p for p in doc.paragraphs if p.text.strip() == "目    录"), None)
    if toc_title is None:
        return
    ps = doc.paragraphs
    i = next(idx for idx, p in enumerate(ps) if p._p is toc_title._p)
    entries = [p for p in ps[i + 1:i + 5]
               if p.text.strip().startswith(("编制说明", "总概算表"))]
    if len(entries) < 2:
        return
    first, second = entries[0], entries[1]
    _toc_entry_para(first, *_TOC_ENTRIES[0])
    _toc_entry_para(second, *_TOC_ENTRIES[1])
    if len(entries) < 3:  # 第三条目插在第二条目后
        third_el = copy.deepcopy(second._element)
        second._element.addnext(third_el)
        _toc_entry_para(Paragraph(third_el, second._parent), *_TOC_ENTRIES[2])


def _fill_cover(doc, info):
    """封面三页：项目名称（自动换行，行1 略长）、签署栏两端对齐、工程编号、日期。"""
    raw_name = str(info.get("name") or "").strip()
    if not raw_name:  # 兼容旧字段：两行拼接
        raw_name = "".join((info.get("name_lines") or [])).strip()
    name_lines = _split_project_name(raw_name)
    if name_lines:
        _set_cover_name(doc, name_lines)

    # 第2页签署区：总院总经理/总院总工程师/院长/院总工程师（模板一行三职称，拆为每行一项）
    _p2_labels = ["总院总经理", "总院总工程师", "院长", "院总工程师"]
    _p2_first = next((p for p in doc.paragraphs
                      if p.text.strip().startswith("总院总经理：")), None)
    _p2_last = next((p for p in doc.paragraphs
                     if p.text.strip().startswith("院总工程师：")), None)
    if _p2_first is not None and _p2_last is not None:
        _inserted = []
        _anchor = _p2_first
        for _lbl in _p2_labels[1:3]:  # 复制首段格式，插入第 2、3 项
            _new_el = copy.deepcopy(_p2_first._element)
            _anchor._element.addnext(_new_el)
            _new_p = Paragraph(_new_el, _p2_first._parent)
            _inserted.append(_new_p)
            _anchor = _new_p
        for _p, _lbl in zip([_p2_first] + _inserted + [_p2_last], _p2_labels):
            _rewrite_sign_para(_p, _distribute_label(_lbl, _P2_LABEL_W) + "：",
                               _P2_FONT_PT, _P2_INDENT_PT)
    # 第3页签署栏：标签两端对齐 + 等长下划线 + 整体居中
    sign_w = max(_disp_len(l) for l in SIGN_LABELS)
    for p in doc.paragraphs:
        t = p.text.strip()
        lbl = next((l for l in SIGN_LABELS if t == l + "："), None)
        if lbl is not None:
            _rewrite_sign_para(p, _distribute_label(lbl, sign_w) + "：",
                               15, _SIGN3_INDENT_PT)

    if info.get("proj_no"):
        for p in _paras(doc, lambda t: t.startswith("工程编号：")):
            _set_para_text(p, f"工程编号： {info['proj_no']}")
    if info.get("date"):
        for p in _paras(doc, lambda t: t in ("2026年1月",)):
            _set_para_text(p, info["date"])

    # 目录页：标题左对齐、页码右对齐（点前导）
    _rebuild_toc(doc)


def _fill_tianjin(doc, region, info, fee_lines, totals):
    _fill_cover(doc, info)

    if info.get("overview"):
        p = _heading_para(doc, "一、工程概况")
        if p is not None:
            _set_para_text(p, info["overview"])

    # 二、编制依据：天津保留模板；其他省份只留通用两条 + 占位提示
    if region != "天津" and not _is_xa(region):
        _fill_between(doc,
                      lambda t: t == "二、编制依据",
                      lambda t: t == "三、工程建设其他费用有关说明",
                      ["本阶段工程的设计文本及图纸。",
                       "2011年中华人民共和国住房和城乡建设部关于印发《市政工程设计概算"
                       "编制办法》的通知（建标[2011]1号）。",
                       f"（{region}编制依据文本待补充，提供后由系统自动填入。）"],
                      "江建权")

    # 三、工程建设其他费用：按实际计算费种动态生成（只写费用+计取标准，不写金额），
    # 逐条编号（与雄安版式一致）
    lines = [f"{i}、{fee_sentence(region, f['name'], f['label'], f.get('yiju', ''))}"
             for i, f in enumerate(fee_lines, 1)]
    _fill_between(doc,
                  lambda t: t == "三、工程建设其他费用有关说明",
                  lambda t: t == "四、其他有关说明",
                  lines, "江建权", single_spacing=True)

    # 四、预备费率 + 资金来源
    pct = float(totals.get("yubei_rate") or 5.0)
    for p in _paras(doc, lambda t: t.startswith("预备费按第一、第二部分费用之和的")):
        _set_para_text(p, f"预备费按第一、第二部分费用之和的{pct:g}%计取。")
    if info.get("fund_source"):
        for p in _paras(doc, lambda t: t.startswith("资金来源")):
            _set_para_text(p, f"资金来源为{info['fund_source']}。")

    # 五、工程投资
    zd = float(totals.get("zhengchai") or 0)
    inv = (f"本项目总投资{float(totals.get('total') or 0):.2f}万元，"
           f"其中第一部分工程费用{float(totals.get('part1') or 0):.2f}万元，"
           f"建设工程其他费用{float(totals.get('fees') or 0):.2f}万元，"
           f"预备费{float(totals.get('yubei') or 0):.2f}万元"
           + (f"，征地拆迁费{zd:.2f}万元" if zd > 0 else "") + "。")
    for p in _paras(doc, lambda t: t.startswith("本项目总投资")):
        _set_para_text(p, inv)


def _fill_xiong_an(doc, region, info, fee_lines, totals):
    _fill_cover(doc, info)

    if info.get("overview"):
        p = _heading_para(doc, "编制范围")
        if p is not None:
            _set_para_text(p, info["overview"])
            # 建设内容段落与模板项目绑定，用户填了概况即删除避免陈旧内容
            for p2 in _paras(doc, lambda t: t.startswith("本项目共设置")):
                _delete_para(p2)

    # 三、工程建设其他费用取费标准（编号列表）
    lines = [f"{i}、{fee_sentence(region, f['name'], f['label'], f.get('yiju', ''))}"
             for i, f in enumerate(fee_lines, 1)]
    _fill_between(doc,
                  lambda t: t == "三、工程建设其他费用取费标准",
                  lambda t: t == "四、预备费",
                  lines, "jwang0ZW", single_spacing=True)

    # 四、预备费率
    pct = float(totals.get("yubei_rate") or 5.0)
    for p in _paras(doc, lambda t: "乘以费率" in t and "计列" in t):
        _set_para_text(p, re.sub(r"费率\d+(\.\d+)?%", f"费率{pct:g}%", p.text))

    # 五、征地拆迁费：本项目无征地拆迁费时整节删除
    zd = float(totals.get("zhengchai") or 0)
    if zd <= 0:
        for p in _paras(doc, lambda t: t == "五、征地拆迁费"):
            _delete_para(p)
        for p in _paras(doc, lambda t: t.startswith("1、永久征地按照")):
            _delete_para(p)

    # 七、资金筹措
    if info.get("fund_source"):
        for p in _paras(doc, lambda t: t.startswith("本项目建设资金为")):
            _set_para_text(p, f"本项目建设资金为{info['fund_source']}。")

    # 八、工程总投资
    inv = (f"本项目总投资{float(totals.get('total') or 0):.2f}万元，"
           f"其中第一部分工程费用为{float(totals.get('part1') or 0):.2f}万元，"
           f"第二部分其他费用为{float(totals.get('fees') or 0):.2f}万元，"
           f"预备费为{float(totals.get('yubei') or 0):.2f}万元"
           + (f"，征地拆迁费为{zd:.2f}万元" if zd > 0 else "") + "。")
    for p in _paras(doc, lambda t: t.startswith("本项目总投资")):
        _set_para_text(p, inv)


def _merge_xa_into_tj(tj_doc, xa_doc):
    """雄安成果 = 天津封面+目录 + 雄安编制说明（合并两份 docx 的 body）。"""
    ps = tj_doc.paragraphs
    start = next(i for i, p in enumerate(ps)
                 if p.text.replace(" ", "") == "编制说明" and p.style.name.startswith("Heading"))
    for p in ps[start:]:
        _delete_para(p)
    # 样式补齐（雄安样式的段落复制过来后仍需样式定义）
    tj_styles = tj_doc.styles.element
    existing = {s.get(qn("w:styleId")) for s in tj_styles.findall(qn("w:style"))}
    for s in xa_doc.styles.element.findall(qn("w:style")):
        if s.get(qn("w:styleId")) not in existing:
            tj_styles.append(copy.deepcopy(s))
    # 追加雄安全部 body 段落（置于末尾 sectPr 之前）
    body = tj_doc.element.body
    body_sect_pr = body.find(qn("w:sectPr"))
    for child in list(xa_doc.element.body):
        if child.tag == qn("w:sectPr"):
            continue
        body_sect_pr.addprevious(copy.deepcopy(child))


# ---- COM：docx → PDF（含目录页码更新）、xlsx → PDF ----

def _com_init():
    """当前线程 COM 初始化（Streamlit 脚本线程未初始化 COM，直接 Dispatch 会报
    -2147221008 尚未调用 CoInitialize）。返回是否真正需要配对 CoUninitialize。"""
    import pythoncom
    try:
        pythoncom.CoInitialize()
        return True
    except Exception:
        return False  # 线程已是 MTA 等场景：交给 COM 自身调度，不配对反初始化


def _com_deinit(ok):
    if ok:
        import pythoncom
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

def _com_para_adjusted_page(doc, pred):
    """段落末尾在 Word 中显示的页码（考虑分节重新编号）。"""
    for i in range(1, doc.Paragraphs.Count + 1):
        para = doc.Paragraphs.Item(i)
        if pred(para.Range.Text.rstrip("\r\x07")):
            rng = para.Range
            rng.Collapse(0)  # 折叠到段尾
            return rng.Information(1)  # wdActiveEndAdjustedPageNumber
    return None


def _word_docx_to_pdf(docx_path, pdf_path, excel_pages, docx_out=None):
    import win32com.client as w
    com_ok = _com_init()
    app = w.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        doc = app.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
        doc.Repaginate()
        start_page = _com_para_adjusted_page(
            doc, lambda t: t.replace(" ", "") == "编制说明" and not t.startswith("目录"))
        end_page = _com_para_adjusted_page(doc, lambda t: t.startswith("本项目总投资"))
        # 更新目录页码：编制说明→start / 总概算表→end+1 / 单位工程概算表→end+1+excel页
        # （目录条目已重建为「标题\t页码」+ 右对齐点前导制表位）
        if start_page and end_page:
            updates = [("编制说明", str(start_page)),
                       ("总概算表", str(end_page + 1)),
                       ("单位工程概算表", str(end_page + 1 + excel_pages))]
            for i in range(1, doc.Paragraphs.Count + 1):
                para = doc.Paragraphs.Item(i)
                t = para.Range.Text.rstrip("\r\x07")
                for label, num in updates:
                    if t.startswith(label) and "\t" in t:
                        # 关键：Range.End 回退一位，不替换段落标记；
                        # 否则 Word 会重建段落并把 pPr 重置为样式默认（制表位丢失/变居中）
                        rng = para.Range
                        rng.End = rng.End - 1
                        rng.Text = re.sub(r"(\d+)\s*$", num, t)
            doc.Repaginate()
        if docx_out:
            # 另存为带真实目录页码的 Word 文件（交付用）
            doc.SaveAs(os.path.abspath(docx_out))
        doc.ExportAsFixedFormat(os.path.abspath(pdf_path), 17)  # wdExportFormatPDF
        doc.Close(False)
    finally:
        app.Quit()
        _com_deinit(com_ok)


def _excel_main_sheet_to_pdf(xlsx_path, pdf_path):
    import win32com.client as w
    com_ok = _com_init()
    xl = w.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(os.path.abspath(xlsx_path))
        main = None
        for nm in ("总概算表", "费用汇总", "项目汇总表"):
            for i in range(1, wb.Worksheets.Count + 1):
                if wb.Worksheets.Item(i).Name == nm:
                    main = wb.Worksheets.Item(i)
                    break
            if main is not None:
                break
        if main is None:
            main = wb.Worksheets.Item(1)
        nb = xl.Workbooks.Add()
        main.Copy(Before=nb.Worksheets.Item(1))
        for i in range(nb.Worksheets.Count, 0, -1):
            if nb.Worksheets.Item(i).Name != main.Name:
                nb.Worksheets.Item(i).Delete()
        ws = nb.Worksheets.Item(1)
        # 打印区域（参考用户成品 PDF「2024.2.1总概算表」：只打印到 K 备注列，
        # L/M「说明」列（自动链接/不可修改/可选）与 N「总投资计算标记」列是
        # 内部机制，不进打印；模板自带 A1:K19 正是此意图）：
        # 有模板打印区域 → 保留其列范围、行数扩展到内容末行（引擎插行后
        # 总投资等行曾落到区域外被截断）；
        # 无（legacy 费用汇总等 openpyxl 新建表）→ 用整个 UsedRange。
        ur = ws.UsedRange
        last_row = ur.Row + ur.Rows.Count - 1
        _pa = (ws.PageSetup.PrintArea or "").strip()
        _m = re.match(
            r"^(?:'[^']*'!)?\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)$", _pa)
        if _m:
            _c1, _r1 = _m.group(1), int(_m.group(2))
            _c2, _r2 = _m.group(3), int(_m.group(4))
            _area = f"${_c1}${_r1}:${_c2}{max(last_row, _r2)}"
        else:
            _area = ws.Range(
                ws.Cells(1, 1),
                ws.Cells(last_row,
                         ur.Column + ur.Columns.Count - 1)).Address
        ws.PageSetup.PrintArea = _area
        ps = ws.PageSetup
        ps.Orientation = 2  # 横向
        ps.Zoom = False
        ps.FitToPagesWide = 1
        ps.FitToPagesTall = False
        nb.ExportAsFixedFormat(0, os.path.abspath(pdf_path))  # xlTypePDF
        nb.Close(False)
        wb.Close(False)
    finally:
        xl.Quit()
        _com_deinit(com_ok)


# ---- 主入口 ----

def _colnum(c: str) -> int:
    n = 0
    for ch in c:
        n = n * 26 + (ord(ch.upper()) - 64)
    return n


def _read_totals_from_path(xlsx_path, fallback=None) -> dict:
    """Excel COM CalculateFull 后按行标签读回 {part1, fees, yubei, total(, zhengchai)}。

    天津：总概算表 B 列（第一部分工程费/第二部分其它建设费用/预 备 费/
    工程总投资）→ G 列；legacy：费用汇总 E 列（第一部分工程费用/
    第二部分工程建设其他费/征地拆迁费/预备费/工程总投资）→ J 列。
    标签匹配前移除全部空白（「预 备 费」「其它/其他」等模板写法差异）。
    用于分开导出：每单项概算表的二类费/总投资按该单项自算，编制说明
    总投资句须与概算表数值一致 → 直接从该表读回。
    """
    import win32com.client as w
    com_ok = _com_init()
    xl = w.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(os.path.abspath(xlsx_path))
        try:
            xl.CalculateFull()
        except Exception:
            pass
        ws = None
        for nm in ("总概算表", "费用汇总", "项目汇总表"):
            for i in range(1, wb.Worksheets.Count + 1):
                if wb.Worksheets.Item(i).Name == nm:
                    ws = wb.Worksheets.Item(i)
                    break
            if ws is not None:
                break
        if ws is None:
            ws = wb.Worksheets.Item(1)
        if ws.Name == "总概算表":
            _keys = (("第一部分工程费", "part1"),
                     ("第二部分其它建设费用", "fees"),
                     ("预备费", "yubei"), ("工程总投资", "total"))
            _lcol, _vcol = _colnum("B"), _colnum("G")
        else:
            _keys = (("第一部分工程费用", "part1"),
                     ("第二部分工程建设其他费", "fees"),
                     ("征地拆迁费", "zhengchai"),
                     ("预备费", "yubei"), ("工程总投资", "total"))
            _lcol, _vcol = _colnum("E"), _colnum("J")
        ur = ws.UsedRange
        _last = ur.Row + ur.Rows.Count - 1
        out = {"part1": 0.0, "fees": 0.0, "yubei": 0.0, "total": 0.0}
        for r in range(2, _last + 1):
            lab = "".join(str(ws.Cells(r, _lcol).Value or "").split())
            for _k, _key in _keys:
                if lab == _k:
                    try:
                        out[_key] = float(ws.Cells(r, _vcol).Value or 0)
                    except (TypeError, ValueError):
                        pass
        wb.Close(False)
        return {**out, **(fallback or {})}
    finally:
        xl.Quit()
        _com_deinit(com_ok)


def read_totals_from_xlsx(excel_bytes, fallback=None) -> dict:
    """字节版包装：写入临时文件后读回（供 app 校验/预读用）。"""
    tmp = tempfile.mkdtemp(prefix="rdtot_")
    try:
        path = os.path.join(tmp, "概算表.xlsx")
        with open(path, "wb") as f:
            f.write(excel_bytes)
        return _read_totals_from_path(path, fallback)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def build_deliverable(*, region, info, fee_lines, totals, excel_bytes,
                      source_pdf_bytes=None) -> bytes:
    """生成概算成果压缩包（zip），内含三件：
    ① {项目名}-总概算表.xlsx（原样交付用户概算表 Excel，含全部 sheet）
    ② {项目名}-编制说明.docx（封面+目录+编制说明，目录页码为真实分页值）
    ③ {项目名}-概算成果.pdf（封面+编制说明 → 概算表 → 广联达源 PDF 合并）

    region: 省份（河北/雄安→雄安编制说明，其余→天津结构）
    info: {name:项目名称(单行,封面自动换行), proj_no, date, overview, fund_source}
    fee_lines: [{"name": 引擎费种名, "label": 显示名, "yiju": 依据}]
    totals: {part1, fees, yubei, total, zhengchai, yubei_rate(百分数)}；可只传
        {zhengchai, yubei_rate} 部分键——其余（part1/fees/yubei/total）用 Excel
        COM CalculateFull 从 excel_bytes 的概算表主表读回（分开导出时每单项的
        二类费/总投资按该单项自算，编制说明总投资句与概算表数值保持一致）。
    excel_bytes: 概算表 xlsx 字节；source_pdf_bytes: 广联达源 PDF 字节（可缺）。
    """
    if not excel_bytes:
        raise ValueError("缺少概算表 Excel 数据")
    base = re.sub(r'[\\/:*?"<>|\r\n\t]', "",
                  str(info.get("name") or "").strip()) or "概算成果"
    tmp = tempfile.mkdtemp(prefix="deliv_")
    try:
        docx_path = os.path.join(tmp, "成果.docx")
        docx_pdf = os.path.join(tmp, "封面编制说明.pdf")
        xlsx_path = os.path.join(tmp, "概算表.xlsx")
        excel_pdf = os.path.join(tmp, "概算表.pdf")
        src_pdf = os.path.join(tmp, "广联达源.pdf")
        final_pdf = os.path.join(tmp, f"{base}-概算成果.pdf")
        zip_path = os.path.join(tmp, f"{base}-概算成果.zip")

        with open(xlsx_path, "wb") as f:
            f.write(excel_bytes)
        if source_pdf_bytes:
            with open(src_pdf, "wb") as f:
                f.write(source_pdf_bytes)
        if not totals or not totals.get("total"):
            # 缺总投资 → 从概算表主表读回该表自己的数值（分开模式：每单项各自口径）
            totals = {**(_read_totals_from_path(xlsx_path) or {}),
                      **(totals or {})}

        if _is_xa(region):
            tj = Document(_TJ_TEMPLATE)
            xa = Document(_XA_TEMPLATE)
            _merge_xa_into_tj(tj, xa)
            _fill_xiong_an(tj, region, info, fee_lines, totals)
            tj.save(docx_path)
        else:
            doc = Document(_TJ_TEMPLATE)
            _fill_tianjin(doc, region, info, fee_lines, totals)
            doc.save(docx_path)

        # Excel 主表 → PDF，并取得页数（目录单位工程概算表页码 = 编制说明末页+1+该页数）
        _excel_main_sheet_to_pdf(xlsx_path, excel_pdf)
        from pypdf import PdfReader
        excel_pages = max(1, len(PdfReader(excel_pdf).pages))

        _word_docx_to_pdf(docx_path, docx_pdf, excel_pages,
                          docx_out=os.path.join(tmp, f"{base}-编制说明.docx"))

        # 合并：封面+编制说明 → Excel 概算表 → 广联达源 PDF
        from pypdf import PdfWriter
        writer = PdfWriter()
        for path in (docx_pdf, excel_pdf, src_pdf):
            if not os.path.exists(path) or os.path.getsize(path) == 0:
                continue
            for page in PdfReader(path).pages:
                writer.add_page(page)
        with open(final_pdf, "wb") as f:
            writer.write(f)

        # 打包：①总概算表 Excel ②编制说明 Word ③成套概算 PDF
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.write(xlsx_path, f"{base}-总概算表.xlsx")
            zf.write(os.path.join(tmp, f"{base}-编制说明.docx"),
                     f"{base}-编制说明.docx")
            zf.write(final_pdf, f"{base}-概算成果.pdf")
        with open(zip_path, "rb") as f:
            return f.read()
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
