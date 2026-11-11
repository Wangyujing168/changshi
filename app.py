"""
工程造价智能助手 - Streamlit 网页界面（Phase 1: 问答功能）
"""
import re
import streamlit as st
import streamlit.components.v1 as components
from rag_engine import get_engine
from fee_engine import (
    detect_and_calculate, is_hebei_region,
    ALL_PROVINCES, DEFAULT_REGION,
    FEE_PARAM_SPECS,
    FEE_CATALOG, CHECKLIST_STAGES,
    fee_calc_intent, build_checklist_meta, advance_checklist,
    detect_new_task_intent,
    current_stage, param_cards_for, settle_checklist,
    linear_rate_options, _build_shencha_rate_options, fee_catalog_for,
    calc_discount_scenarios,
    _get_coef_config_simple,
    _preset_fee_numbers,
    build_excel_summary, TJ_TEMPLATE_XLSX,
    compute_selected_fees,
)
from ai_planner import (
    PlannerError, available_keys, build_ctx_summary, question_done, run_planner,
    sanitize_plan_questions,
)
from config import CASE_BASE_DIR

# ===== 清单式任务流（checklist task flow）双轨开关 =====
# _TASK_FLOW_START：接管「缺金额/缺费种」场景（旧路径会丢的场景）
# _TASK_FLOW_CASCADE：金额齐全的 cascade 也进清单机（B3 已翻转，面板变死代码）
_TASK_FLOW_START = True
_TASK_FLOW_CASCADE = True
# AI 对话式提问（概算助手）：默认开启；用户可手动切换传统清单模式（会话内记住）
st.session_state.setdefault("ai_guide_pref", True)

_TASK_STAGE_LABELS = {
    "fees": "费种确认",
    "amounts": "基础金额",
    "custom": "自定义费用/合同价",
    "params": "费率与系数",
    "discount": "折扣",
}

# 环评/可研服务类型（政策文件固定 4 项；造价咨询服务按省份在 FEE_CATALOG conf 中）
_HP_SERVICES = ["编制报告书", "编制报告表", "评估报告书", "评估报告表"]
_KY_SERVICES = ["编制项目建议书", "编制可研报告", "评估项目建议书", "评估可研报告"]

# 水土保持补偿费计征类型（迁移自 FEE_PARAM_SPECS["水土保持补偿费"]）
_SBUBAO_TYPES = [("一般项目（按面积）", "general"),
                 ("开采矿产", "mining_oil_gas"),
                 ("取土/挖沙", "mining_other")]

# ===== 政策依据可点击链接 =====

_POLICY_URLS: dict[str, str] = {
    # 河北省
    "冀建市研[2017]2号": "https://www.baidu.com/s?wd=冀建市研[2017]2号+工程造价咨询服务收费管理暂行办法",
    "冀价行费[2018]57号": "https://www.baidu.com/s?wd=冀价行费[2018]57号+施工图审查费",
    "发改价格〔2011〕534号": "https://www.baidu.com/s?wd=发改价格〔2011〕534号+降低部分建设项目收费标准",
    # 天津市
    "津价房地[2008]136号": "https://www.baidu.com/s?wd=津价房地[2008]136号+建设工程造价咨询服务",
    "津价管[2011]46号": "https://www.baidu.com/s?wd=津价管[2011]46号+施工图设计文件审查",
    # 国家层面
    "计价格[2002]10号": "https://www.baidu.com/s?wd=计价格[2002]10号+工程勘察设计收费管理规定",
    "发改价格[2007]670号": "https://www.baidu.com/s?wd=发改价格[2007]670号+建设工程监理收费",
    "计价格[2002]125号": "https://www.baidu.com/s?wd=计价格[2002]125号+环境影响咨询收费",
    "计价格[1999]1283号": "https://www.baidu.com/s?wd=计价格[1999]1283号+建设项目前期工作咨询费",
    "计价格[2002]1980号": "https://www.baidu.com/s?wd=计价格[2002]1980号+招标代理服务收费",
    "建市[2007]86号": "https://www.baidu.com/s?wd=建市[2007]86号+工程设计资质标准",
}


def show_policy_badge(policy_id: str):
    """Display a clickable policy badge in the Streamlit UI.

    Renders a styled box with a link to search for the policy document.
    Falls back to st.info() if the policy ID is unknown.
    """
    url = _POLICY_URLS.get(policy_id)
    if url:
        st.markdown(
            f'<div style="background-color:#d4e6f1;padding:10px 14px;'
            f'border-radius:4px;text-align:center;border:1px solid #aed6f1;">'
            f'<a href="{url}" target="_blank" rel="noopener"'
            f' style="color:#0d47a1;text-decoration:none;font-weight:bold;font-size:14px;">'
            f'📋 {policy_id}</a></div>',
            unsafe_allow_html=True,
        )
    else:
        st.info(policy_id)


def _round2(val: float) -> str:
    """Round to 2 decimal places using round-half-up (matches Excel ROUND).

    Python's built-in round() and format() use banker's rounding (round half
    to even), which can differ from Excel at the exact .005 boundary by 0.01.
    """
    import math
    return f"{math.floor(val * 100 + 0.5) / 100:.2f}"


def _basis_with_links(basis_text: str) -> str:
    """Insert clickable HTML links into a 依据 text for panel captions."""
    result = basis_text
    for pid, url in _POLICY_URLS.items():
        if pid in result:
            result = result.replace(
                pid,
                f'<a href="{url}" target="_blank" rel="noopener"'
                f' style="text-decoration:none;color:#0d47a1;">{pid}</a>',
            )
    return result




# ===== 近期计算记录 =====

import json as _json

_RECENTS_FILE = "recent_calculations.json"


def _load_recents() -> list[dict]:
    """从文件加载近期计算记录。"""
    try:
        with open(_RECENTS_FILE, "r", encoding="utf-8") as f:
            return _json.load(f)
    except (FileNotFoundError, _json.JSONDecodeError):
        return []


def _save_recents(recents: list[dict]):
    """将近期计算记录持久化到文件。"""
    try:
        with open(_RECENTS_FILE, "w", encoding="utf-8") as f:
            _json.dump(recents, f, ensure_ascii=False, indent=2)
    except Exception:
        pass  # 写入失败不影响计算


def _record_calculation(query: str, fee_result: dict | None):
    """将一次成功计算追加到近期计算列表。"""
    if not query or not fee_result:
        return
    recents = list(st.session_state.get("recent_calculations", []))
    # 去重：与上一条相同则不重复记录
    if recents and recents[-1].get("query") == query:
        return

    from datetime import datetime
    mode_map = {
        "__cascade__": "二类费计算",
        "__iteration__": "迭代计算",
        "__comparison__": "方案比选",
    }
    mode = mode_map.get(fee_result.get("fee_type", ""), "单费种")

    # 构建摘要
    summary_parts = []
    fee_name = fee_result.get("费种", "")
    if fee_name:
        summary_parts.append(fee_name)
    val = fee_result.get("结果(万元)")
    if val is not None:
        summary_parts.append(f"{val:.2f}万")
    elif fee_result.get("结果汇总"):
        total = fee_result["结果汇总"].get("项目总投资(万元)")
        if total is not None:
            summary_parts.append(f"总投资 {total:.2f}万")

    recents.append({
        "query": query,
        "timestamp": datetime.now().strftime("%m-%d %H:%M"),
        "summary": " · ".join(summary_parts) or query[:30],
        "mode": mode,
    })
    # 最多保留 20 条
    if len(recents) > 20:
        recents = recents[-20:]
    st.session_state.recent_calculations = recents
    _save_recents(recents)


# ===== 页面设置 =====
st.set_page_config(
    page_title="造价智能助手",
    page_icon="🌿",
    layout="wide",
)

# ===== 辅助渲染函数 =====




def _build_cascade_excel(ctx: dict) -> bytes:
    """根据计算结果生成「活公式」Excel 文件，返回 bytes 供下载。

    四 sheet 构建（费用汇总/输入参数/计算规则/计算过程）在
    fee_engine.build_excel_summary：全部费用为活公式，修改黄色输入格
    （建安费/设备费/折扣/费率/档位表）后自动重算；循环引用依赖 Excel
    迭代计算（文件已内置 calcPr 设置）。
    """
    return build_excel_summary(ctx)


# ===== 功能页通用标题行（含「刷新本功能」按钮） =====

def _page_header(title: str, refresh_key: str, refresh_keys,
                 refresh_prefixes=None, keep_keys=None):
    """功能页标题行 + 「🔄 刷新本功能」按钮。

    点击后清除本功能的全部 session 状态（上传文件、解析结果、计算结果、
    参数输入、动态子键），立即重跑——无需刷新整个网页（F5 不清
    session_state，这正是本按钮存在的意义）。导航开关（*_mode/active_case）
    与全局偏好（selected_region 等，keep_keys 白名单）不受影响。
    """
    c1, c2 = st.columns([0.78, 0.22], vertical_alignment="center")
    with c1:
        st.title(title)
    with c2:
        if st.button("🔄 刷新本功能", key=refresh_key, use_container_width=True,
                     help="清除本功能的全部上传、参数与计算结果，重新开始"):
            keep = {k: st.session_state[k]
                    for k in (keep_keys or []) if k in st.session_state}
            for k in refresh_keys:
                st.session_state.pop(k, None)
            for prefix in refresh_prefixes or []:
                for k in list(st.session_state.keys()):
                    if k.startswith(prefix):
                        st.session_state.pop(k, None)
            st.session_state.update(keep)
            st.rerun()


# ===== 广联达 PDF 导入页 =====

def _render_pdf_import_page():
    """上传广联达 PDF → 解析「单位工程造价汇总表」/「工程费用汇总表」→ 编辑分组 → 存专业清单。

    清单保存到 st.session_state["zhuan_ye_list"]，Excel 导出的第一部分
    工程费用按 组（一）→ 专业 → 节/细目 列示（见 fee_engine 导出区块）。
    """
    import pandas as pd
    from pdf_import import parse_pdf, build_groups

    _page_header("📄 广联达报表导入", "pdf_refresh",
                 ["pdf_up", "pdf_up_id", "excel_up", "excel_up_id", "pdf_rows",
                  "pdf_editor", "pdf_split_mode", "zhuan_ye_list", "pdf_bytes",
                  "src_pdf_up", "src_pdf_up_id", "src_pdf_files"],
                 refresh_prefixes=["excel_len_", "excel_cross_",
                                   "excel_item_editor_"])
    st.caption("只解析每个专业的「单位工程造价汇总表」（河北等地区）或"
               "「工程费用汇总表」（河南等地区）：提取 工程名称 + 工程造价"
               "（PDF 费用金额单位为元，已换算为万元，原值保留）。"
               "同一工程名称出现多批时自动按（一）（二）…分组，可在下方编辑。"
               "也可直接上传广联达 Excel 造价分析表（单项造价分析/项目造价分析），"
               "除工程造价外还会识别 工程规模/单位造价 指标（可多选，新上传会替换当前表格）。"
               "用 Excel 导入时，如需在概算成果末尾附上广联达源 PDF，"
               "请在下方「📎 附加广联达源 PDF」上传。")

    up = st.file_uploader("上传广联达导出的 PDF 报表", type=["pdf"], key="pdf_up")
    if up is not None and st.session_state.get("pdf_up_id") != up.file_id:
        with st.spinner("正在解析 PDF…"):
            try:
                leaves = parse_pdf(up.getvalue())
            except Exception as e:
                st.error(f"PDF 解析失败：{e}")
                st.stop()
        if not leaves:
            st.error("未识别到「单位工程造价汇总表」或「工程费用汇总表」，请确认是广联达软件导出的报表。")
            st.stop()
        st.session_state.pdf_up_id = up.file_id
        st.session_state.pop("excel_up_id", None)
        # 保存源 PDF 字节：概算成果导出（封面+编制说明+概算表+广联达PDF 合并）时使用
        st.session_state.pdf_bytes = up.getvalue()
        # 新 PDF 即新的源文件：清除此前单独附加的源 PDF（避免成果误附旧文件）
        st.session_state.pop("src_pdf_files", None)
        st.session_state.pop("src_pdf_up_id", None)
        # 新 PDF：清除上一次的导出方式选择（radio 按当前解析结果重新显示）
        st.session_state.pop("pdf_split_mode", None)
        st.session_state.pdf_rows = [
            {"组号": l["group_no"], "组名": "", "专业": l["zhuanye"], "节": l["jie"],
             "细目": l["ximu"], "造价(万元)": l["amount_wan"],
             "设备费(万元)": l["device_wan"], "电气专业": l["electrical"],
             "单位": "", "工程规模": None, "单位造价(元)": None}
            for l in leaves
        ]
        st.rerun()

    ex_up = st.file_uploader(
        "或上传广联达 Excel 造价分析表（单项造价分析 / 项目造价分析，可多选）",
        type=["xls", "xlsx"], key="excel_up", accept_multiple_files=True)
    if ex_up and st.session_state.get("excel_up_id") != [f.file_id for f in ex_up]:
        from excel_import import parse_analysis
        leaves = []
        _goff = 1
        for f in ex_up:
            with st.spinner(f"正在解析「{f.name}」…"):
                try:
                    _part = parse_analysis(f.getvalue(), group_offset=_goff,
                                           filename=f.name)
                except Exception as e:
                    st.error(f"「{f.name}」解析失败：{e}")
                    st.stop()
            leaves += _part
            _goff = max((l["组号"] for l in _part), default=0) + 1
        if not leaves:
            st.error("未识别到「序号/名称/金额」表头，请确认是造价分析类报表。")
            st.stop()
        st.session_state.excel_up_id = [f.file_id for f in ex_up]
        st.session_state.pop("pdf_up_id", None)
        st.session_state.pop("pdf_bytes", None)
        st.session_state.pop("pdf_split_mode", None)
        st.session_state.pdf_rows = leaves
        st.rerun()

    # 附加源 PDF（仅存档、不解析）：Excel 导入时专业清单来自表格，概算成果
    # 末尾的广联达 PDF 需要在这里单独上传；上传解析 PDF 时无需（本身即源）。
    # 可多选：分开导出（每单项一份概算表）时按上传顺序对应各单项
    src_up = st.file_uploader(
        "📎 附加广联达源 PDF（可选；仅附在「概算成果」末尾，不参与解析——"
        "用 Excel 导入专业清单时在此补传，可一次多选："
        "分开导出时按上传顺序对应各单项，只传一份则每份成果都附这一份）",
        type=["pdf"], key="src_pdf_up", accept_multiple_files=True)
    if src_up and st.session_state.get("src_pdf_up_id") != [f.file_id for f in src_up]:
        st.session_state.src_pdf_up_id = [f.file_id for f in src_up]
        st.session_state.src_pdf_files = [(f.name, f.getvalue()) for f in src_up]
        st.rerun()
    if st.session_state.get("src_pdf_files"):
        _sf = st.session_state["src_pdf_files"]
        # 上传控件离开页面后状态会被 Streamlit 清理（widget state 丢失），
        # 这里以 session 里的附件为准，不清除；控件为空时提示可重新上传/刷新清除
        _hint = ("（上传控件当前为空：重新上传会替换附件；"
                 "点「🔄 刷新本功能」清除）" if not src_up else "")
        if len(_sf) == 1:
            st.caption(f"✅ 已附加源 PDF：{_sf[0][0]}（将附在概算成果末尾）{_hint}")
        else:
            st.caption(
                f"✅ 已附加 {len(_sf)} 份源 PDF：{'、'.join(n for n, _b in _sf)}。"
                "分开导出（每单项一份概算表）时按上传顺序对应各单项；"
                "合并为同一概算表时全部附在成果末尾。" + _hint)
    elif not st.session_state.get("pdf_bytes"):
        st.caption("⚠ 未附加广联达源 PDF。通过 Excel 导入时请在此上传源 PDF，"
                   "概算成果末尾才能附上广联达报表。")

    rows = st.session_state.get("pdf_rows")
    if not rows:
        st.info("请先上传 PDF。示例：工作区 参考2.PDF（河北雄安项目，577 页）。")
        return

    st.markdown("### 解析结果（可直接编辑）")
    st.caption("组号=批次分组（同一次上传中同名工程第 2 次出现起算新组）；"
               "电气专业（照明/智能交通/电气）造价放合计列＝安装＋设备，其余放建筑列。")
    df = pd.DataFrame(rows)
    edited = st.data_editor(
        df,
        key="pdf_editor",
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "组号": st.column_config.NumberColumn("组号", min_value=1, max_value=20,
                                                step=1),
            "组名": st.column_config.TextColumn("组名（工程或费用名称，自行填写）"),
            "专业": st.column_config.TextColumn("专业（单项工程）"),
            "节": st.column_config.TextColumn("节（单位工程）"),
            "细目": st.column_config.TextColumn("细目"),
            "造价(万元)": st.column_config.NumberColumn("造价(万元)", format="%.4f"),
            "设备费(万元)": st.column_config.NumberColumn("设备费(万元)", format="%.4f"),
            "电气专业": st.column_config.CheckboxColumn("电气专业"),
            "单位": st.column_config.TextColumn("单位"),
            "工程规模": st.column_config.NumberColumn("工程规模", format="%.2f"),
            "单位造价(元)": st.column_config.NumberColumn("单位造价(元)", format="%.2f"),
        },
    )

    _cn = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]
    # 编辑器中文列名 → build_groups 英文键
    groups = build_groups([
        {"group_no": r.get("组号"), "group_name": r.get("组名"),
         "zhuanye": r.get("专业"), "jie": r.get("节"), "ximu": r.get("细目"),
         "amount_wan": r.get("造价(万元)"), "device_wan": r.get("设备费(万元)"),
         "electrical": r.get("电气专业"), "unit": r.get("单位"),
         "qty": r.get("工程规模"), "unit_val": r.get("单位造价(元)")}
        for r in edited.to_dict("records")
    ])
    tot = 0.0
    dev_tot = 0.0
    st.markdown("### 汇总预览")
    for gi, g in enumerate(groups, 1):
        gsum = sum(x["amount"] for it in g["items"] for x in it["leaves"])
        gdev = sum(x["device"] for it in g["items"] for x in it["leaves"])
        tot += gsum
        dev_tot += gdev
        title = f"（{_cn[gi - 1] if gi <= len(_cn) else gi}）{g['name'] or '未命名组'} — {gsum:,.2f} 万元"
        with st.expander(title, expanded=True):
            for it in g["items"]:
                s = sum(x["amount"] for x in it["leaves"])
                st.markdown(f"- {it['zhuanye']}：**{s:,.2f} 万元**"
                            f"（{len(it['leaves'])} 行{'，电气专业' if it['electrical'] else ''}）")
    st.markdown(f"**第一部分 工程费用合计：{tot:,.2f} 万元"
                f"（其中设备费 {dev_tot:,.2f} 万元）**")

    if len(groups) > 1:
        _split_opts = ["同一概算表拆项", "分开做概算表（每单项一份）"]
        _mode = st.radio(
            "概算表导出方式（多项工程）",
            _split_opts,
            index=(_split_opts.index(st.session_state["pdf_split_mode"])
                   if st.session_state.get("pdf_split_mode") in _split_opts else 0),
            horizontal=True,
            help="同一概算表拆项：所有单项列在同一张概算表中（按单项 1/1.1/2/2.1 拆项列示）。"
                 "分开做概算表：每个单项工程单独一份概算表文件（二类费/预备费/总投资"
                 "按该单项自己的工程费计算）；征地拆迁费/自定义费用按单项逐一提问、"
                 "分别写入自己那份概算表；概算成果（编制说明+PDF）也按每单项一套分别生成。",
        )
        # 镜像到普通 session_state 键：radio 只在 PDF 导入页渲染，跳转到结果页后
        # Streamlit 会清理未渲染 widget 的状态（带 key 的 widget 状态会丢），
        # 普通键不会被清理，导出时从这里读
        st.session_state["pdf_split_mode"] = _mode

    # ── 技术经济指标：大项（专业行）指标默认口径 + 提问 + 可直接编辑 ──
    from excel_import import item_default_tech, tech_from_text, tech_to_text
    _any_leaf_tech = any("tech" in lf for g in groups for it in g["items"]
                         for lf in it["leaves"])
    _lens: dict = {}
    _cross: dict = {}
    if _any_leaf_tech:
        st.markdown("### 技术经济指标（导出到概算表 单位/数量/单位价值 列）")
        st.caption("Excel 细项行的 工程规模/单位造价 直接作为细目指标；大项（专业行）"
                   "默认口径：道路=m2 路面面积（行车道/车行道+非机动车道+人行道）"
                   "+m 道路长度、照明=m 道路长度、信号灯/电子警察/智能交通=路口"
                   "（智能交通未填路口数时默认 m=道路长度），"
                   "其余按专业名推断单位同单位求和、单位价值=金额÷数量。"
                   "下方可直接编辑。")
        for gi, g in enumerate(groups, 1):
            _zys = [str(it.get("zhuanye") or "") for it in g["items"]]
            _need_len = any(("道路" in z) or ("照明" in z) for z in _zys)
            _need_cross = any(k in z for z in _zys
                              for k in ("信号灯", "电子警察", "智能交通"))
            if not (_need_len or _need_cross):
                continue
            _t1, _t2 = st.columns(2)
            with _t1:
                if _need_len:
                    _lens[gi] = st.number_input(
                        f"（{_cn[gi - 1] if gi <= len(_cn) else gi}）"
                        f"{g['name'] or '未命名组'} 道路长度(m)",
                        min_value=0.0, step=1.0, key=f"excel_len_{gi}") or 0.0
            with _t2:
                if _need_cross:
                    _cross[gi] = st.number_input(
                        f"（{_cn[gi - 1] if gi <= len(_cn) else gi}）"
                        f"{g['name'] or '未命名组'} 路口数(个)",
                        min_value=0.0, step=1.0, key=f"excel_cross_{gi}") or 0.0
        _item_rows = []
        for gi, g in enumerate(groups, 1):
            for it in g["items"]:
                _u, _q, _v = tech_to_text(item_default_tech(
                    it["zhuanye"], it["leaves"], _lens.get(gi, 0.0),
                    _cross.get(gi, 0.0)))
                _item_rows.append({"组号": gi, "组名": g["name"] or "",
                                   "专业": it["zhuanye"], "单位": _u,
                                   "数量": _q, "单位价值(元)": _v})
        st.markdown("**大项技术经济指标（可直接编辑）**")
        _item_edited = st.data_editor(
            pd.DataFrame(_item_rows),
            key=f"excel_item_editor_{len(_item_rows)}",
            num_rows="fixed",
            use_container_width=True,
            hide_index=True,
            column_config={
                "组号": st.column_config.NumberColumn("组号", disabled=True),
                "组名": st.column_config.TextColumn("组名", disabled=True),
                "专业": st.column_config.TextColumn("专业", disabled=True),
                "单位": st.column_config.TextColumn("单位（多组用「、」分隔）"),
                "数量": st.column_config.TextColumn("数量"),
                "单位价值(元)": st.column_config.TextColumn("单位价值(元)"),
            },
        )
    else:
        _item_edited = pd.DataFrame()
    _item_recs = [dict(r) for r in _item_edited.to_dict("records")]

    def _attach_item_tech(gs, recs):
        """把大项指标编辑器内容按 (组,专业) 顺序写回 items[].tech。"""
        out = []
        _idx = 0
        for g in gs:
            _items = []
            for it in g["items"]:
                _it = dict(it)
                if _idx < len(recs):
                    _t = tech_from_text(recs[_idx].get("单位"),
                                        recs[_idx].get("数量"),
                                        recs[_idx].get("单位价值(元)"))
                    if _t:
                        _it["tech"] = _t
                    else:
                        _it.pop("tech", None)
                _idx += 1
                _items.append(_it)
            out.append({"name": g["name"], "items": _items})
        return out

    groups_out = _attach_item_tech(groups, _item_recs)

    if len(groups) > 1 and st.session_state.get("pdf_split_mode") != \
            "分开做概算表（每单项一份）":
        st.warning("⚠ 当前导出方式为「同一概算表拆项」：所有单项放在**同一张概算表**，"
                   "征地拆迁费/自定义费用按项目整体提问一次。"
                   "如需**每个单项单独一份概算表**（征地拆迁费/自定义费用按单项逐一填写），"
                   "请在上方「概算表导出方式」选择「**分开做概算表（每单项一份）**」。")

    c1, c2 = st.columns(2)
    with c1:
        if st.button("💾 保存专业清单（用于 Excel 导出）", use_container_width=True,
                     type="primary"):
            st.session_state.zhuan_ye_list = groups_out
            st.success("已保存。返回主页面发起二类费计算，导出 Excel 的第一部分"
                       "将按 组→专业→细目 列示，各层自动求和，并带技术经济指标。")
    with c2:
        if st.button("🚀 按合计发起二类费计算", use_container_width=True):
            st.session_state.zhuan_ye_list = groups_out
            st.session_state.pdf_import_mode = False
            # tot 已含设备费（电气专业造价 = 安装 + 设备）→ 建安费须扣除设备费，
            # 否则第一部分工程费 = 建安 + 设备 会把设备费双计（面板与 Excel 合计不符）
            st.session_state.current_query = (
                f"建安费{round(tot - dev_tot, 2)}万，设备费{round(dev_tot, 2)}万，"
                f"帮我算全部费用")
            st.rerun()

    if st.session_state.get("zhuan_ye_list"):
        st.caption("当前已保存的专业清单将在下次导出 Excel 时生效。")


def _render_preprocess_page():
    """表格预处理：① 数量总表 .xls → 解析映射 → 预览 → 下载填好的表格（模板已内置）；
    ② 植物表 .xlsx → 把 胸（地）径/高度/冠幅 拼接进名称列 → 预览 → 下载改写后的表。

    解析/填充逻辑分别在 quantity_import.py 与 plant_spec.py（纯逻辑模块）。
    """
    import pandas as pd

    from plant_spec import process_plant_table, to_bytes as plant_to_bytes
    from quantity_import import (QuantityImportError, fill_template,
                                 load_builtin_template, mapping_rows,
                                 parse_xls)

    _page_header("📋 表格预处理", "pre_refresh",
                 ["pre_xls_up", "pre_src_id", "pre_mapping", "pre_filled",
                  "pre_xls_err", "pre_plant_up", "pre_plant_src_id",
                  "pre_plant_changes", "pre_plant_file", "pre_plant_err"])

    # ---------- ① 数量总表 → 填充模板 ----------
    st.markdown("### ① 桥梁工程量表 → 填充模板")
    st.caption("上传「主要工程数量总表.xls」，自动把桥梁上、下部结构各部位工程量填入"
               "内置表格模板的信息输入表对应单元格（计价表与混凝土方量/凿桩头"
               "为公式自动联动）。数量总表里模板没有的部位会自动新增分项。")

    up_xls = st.file_uploader("上传数量总表 (.xls)", type=["xls"],
                              key="pre_xls_up")
    if up_xls is None:
        st.info("请上传数量总表 .xls，解析后可在下方核对映射结果并下载。")
    elif st.session_state.get("pre_src_id") != up_xls.file_id:
        with st.spinner("正在解析并填充…"):
            try:
                m = parse_xls(up_xls.getvalue())
                filled = fill_template(load_builtin_template(), m)
                st.session_state.pre_xls_err = None
            except QuantityImportError as e:
                m = None
                st.session_state.pre_xls_err = f"解析失败：{e}"
        st.session_state.pre_mapping = m
        st.session_state.pre_filled = filled
        st.session_state.pre_src_id = up_xls.file_id
        st.rerun()
    else:
        if st.session_state.get("pre_xls_err"):
            st.error(st.session_state.pre_xls_err)
        else:
            m = st.session_state.pre_mapping
            st.markdown("#### 映射结果核对")
            rows = mapping_rows(m)
            st.dataframe(pd.DataFrame(rows, columns=["单元格", "项目", "值"]),
                         use_container_width=True, hide_index=True)
            for w in m.get("warnings") or []:
                st.warning(w)

            st.download_button(
                "⬇️ 下载填好的表格.xlsx", data=st.session_state.pre_filled,
                file_name="表格_已导入.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)

    st.divider()

    # ---------- ② 植物表 → 规格写入名称 ----------
    st.markdown("### ② 植物表 → 规格写入名称")
    st.caption("上传植物表 .xlsx，把 胸（地）径/高度/冠幅 拼接进名称列，如"
               "「法桐 胸径14-16cm 高度10m 冠幅3.5-4m」。乔木写胸径、灌木写地径，"
               "拿不准的保留「胸（地）径」；缺哪项跳过哪项，三项全空的保留原名。")

    up_plant = st.file_uploader("上传植物表 (.xlsx)", type=["xlsx"],
                                key="pre_plant_up")
    if up_plant is None:
        st.info("请上传植物表 .xlsx，改写后的新表可在下方下载。")
    elif st.session_state.get("pre_plant_src_id") != up_plant.file_id:
        with st.spinner("正在改写名称列…"):
            try:
                changes, wb = process_plant_table(up_plant)
                st.session_state.pre_plant_err = None
            except Exception as e:
                changes = []
                st.session_state.pre_plant_err = f"解析失败：{e}"
        st.session_state.pre_plant_changes = changes
        st.session_state.pre_plant_file = plant_to_bytes(wb) if changes else None
        st.session_state.pre_plant_src_id = up_plant.file_id
        st.rerun()
    else:
        if st.session_state.get("pre_plant_err"):
            st.error(st.session_state.pre_plant_err)
        else:
            changes = st.session_state.pre_plant_changes
            st.markdown(f"#### 名称改写预览（共 {len(changes)} 行）")
            if changes:
                st.dataframe(pd.DataFrame(changes, columns=["行号", "原名称", "新名称"]),
                             use_container_width=True, hide_index=True)
                st.download_button(
                    "⬇️ 下载改写后的植物表.xlsx", data=st.session_state.pre_plant_file,
                    file_name="植物表_已改写.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)
            else:
                st.info("未找到可改写的苗木行（需 B 列名称、C 列表头含「径」的区块）。")


# ===== 桥梁定额匹配：上传定额库 + 海泰式数量总表 → 自动匹配天津桥梁定额 =====
_STATUS_EMOJI = {"single": "✅", "multi": "🤔", "none": "❌"}


def _render_bridge_match_page():
    import pandas as pd
    from bridge_match import (BridgeMatchError, parse_quota_library,
                              parse_rule_template, load_builtin_rule_template,
                              parse_quantity_excel, parse_text_lines,
                              match_quotas, build_result_excel, result_rows,
                              ai_explain, _item_display, _grade_tag, TPL_RATIOS)

    _page_header("🌉 市政桥梁定额匹配", "bm_refresh",
                 ["bm_lib_up", "bm_lib", "bm_lib_id", "bm_lib_err",
                  "bm_rules_up", "bm_rules", "bm_rules_id", "bm_rules_err",
                  "bm_qty_up", "bm_qty_id", "bm_qty_err", "bm_items",
                  "bm_text_area", "bm_text_btn", "bm_text_warns",
                  "bm_results", "bm_results_key", "bm_lib_merged",
                  "bm_rules_src", "bm_ai_cache", "bm_span_len", "bm_soil",
                  "bm_bridge_width", "bm_earth_reuse", "bm_soil_haul_km",
                  "bm_waste_haul_km", "bm_run"],
                 refresh_prefixes=["bm_tpl_", "bm_ai_btn_"])
    st.caption("上传导定额库 + 套定额模板（单位工程）+ 海泰式数量总表，"
               "按构件/工艺/材料/规格自动匹配市政定额子目（适用于广联达软件）；"
               "严格以定额库文字为准，不编造子目；多候选全部列出并附差异说明；"
               "特征缺失时给出需补充的信息清单。")

    # ---------- ① 定额库 ----------
    st.markdown("### ① 定额库")
    up_lib = st.file_uploader("上传定额库 Excel（导定额.xlsx 分页计价表，含定额编号/名称/单位/单价）",
                              type=["xlsx"], key="bm_lib_up",
                              help="新版导定额库为单表 38 页分页计价表；"
                                   "旧版天津桥梁.xlsx（定额库+规则样例双表）也兼容。")
    if up_lib is not None and st.session_state.get("bm_lib_id") != up_lib.file_id:
        try:
            lib = parse_quota_library(up_lib.getvalue())
            st.session_state.bm_lib = lib
            st.session_state.bm_lib_id = up_lib.file_id
            st.session_state.bm_lib_err = None
        except BridgeMatchError as e:
            st.session_state.bm_lib_err = str(e)
            st.session_state.bm_lib = None
            st.session_state.bm_lib_id = None
        except Exception as e:  # 解析兜底
            st.session_state.bm_lib_err = f"定额库解析失败：{e}"
            st.session_state.bm_lib = None
            st.session_state.bm_lib_id = None
    if st.session_state.get("bm_lib_err"):
        st.error(st.session_state.bm_lib_err)
    elif st.session_state.get("bm_lib"):
        lib = st.session_state.bm_lib
        dup = lib.get("duplicates") or {}
        n_mat = len(lib.get("materials", []))
        st.success(f"已加载 {len(lib['entries'])} 条定额"
                   + (f" + {n_mat} 条材料行" if n_mat else "")
                   + (f" + {len(lib.get('rules', {}))} 组套定额规则样例"
                      if lib.get("rules") else "")
                   + (f"（⚠️ 库中重复编号：{'、'.join(dup)}）" if dup else ""))
        with st.expander("定额库预览（前 10 条）"):
            st.dataframe(pd.DataFrame(
                [{"编号": e["code"], "名称": e["name"], "单位": e["unit"],
                  "单价": e["price"]} for e in lib["entries"][:10]]),
                use_container_width=True, hide_index=True)

    st.divider()

    # ---------- ② 套定额模板（内置单位工程.xlsx，可上传自定义覆盖） ----------
    st.markdown("### ② 套定额模板")
    try:
        builtin_rules = load_builtin_rule_template()
    except BridgeMatchError as e:
        builtin_rules = None
        st.error(str(e))
    up_rules = st.file_uploader(
        "上传自定义套定额模板（可选）——默认使用内置 单位工程.xlsx 模板规则，上传后以自定义模板为准",
        type=["xlsx"], key="bm_rules_up",
        help="内置模板：quota_base/单位工程.xlsx（分部→分项→子组→明细行）；无需上传即可匹配。")
    if up_rules is None and st.session_state.get("bm_rules_id"):
        # 移除上传的自定义模板后回到内置规则
        st.session_state.bm_rules = None
        st.session_state.bm_rules_id = None
    # 旧版解析结果（无 divs 分部结构）在代码升级后按同 file_id 残留于会话——
    # 检测到缺 divs 时用 uploader 缓存的原始字节重新解析
    if up_rules is not None and (
            st.session_state.get("bm_rules_id") != up_rules.file_id
            or "divs" not in (st.session_state.get("bm_rules") or {})):
        try:
            rules_data = parse_rule_template(up_rules.getvalue())
            st.session_state.bm_rules = rules_data
            st.session_state.bm_rules_id = up_rules.file_id
            st.session_state.bm_rules_err = None
        except BridgeMatchError as e:
            st.session_state.bm_rules_err = str(e)
            st.session_state.bm_rules = None
            st.session_state.bm_rules_id = None
        except Exception as e:
            st.session_state.bm_rules_err = f"模板解析失败：{e}"
            st.session_state.bm_rules = None
            st.session_state.bm_rules_id = None
    if st.session_state.get("bm_rules_err"):
        st.error(st.session_state.bm_rules_err)
    elif st.session_state.get("bm_rules"):
        rules_data = st.session_state.bm_rules
        n_rows = sum(len(v) for v in rules_data["rules"].values())
        st.success(f"已加载自定义套定额模板：{len(rules_data['rules'])} 组分项规则"
                   f"（共 {n_rows} 条样例）——将覆盖内置模板")
        with st.expander("模板规则预览（按分项分组）"):
            st.dataframe(pd.DataFrame(
                [{"分项组": g, "样例条数": len(rs)}
                 for g, rs in rules_data["rules"].items()]),
                use_container_width=True, hide_index=True)
    elif builtin_rules is not None:
        n_rows = sum(len(v) for v in builtin_rules["rules"].values())
        st.success(f"内置套定额模板：单位工程.xlsx（{len(builtin_rules['rules'])} 组分项规则，"
                   f"共 {n_rows} 条样例）——无需上传，匹配时自动套用")
        with st.expander("模板规则预览（按分项分组）"):
            st.dataframe(pd.DataFrame(
                [{"分项组": g, "样例条数": len(rs)}
                 for g, rs in builtin_rules["rules"].items()]),
                use_container_width=True, hide_index=True)

    st.divider()

    # ---------- ③ 工程量 ----------
    st.markdown("### ③ 工程量清单")
    tab_up, tab_text = st.tabs(["上传数量总表（替换）", "粘贴文本（追加）"])
    with tab_up:
        up_qty = st.file_uploader(
            "上传海泰式数量总表（分项材料表：分项名/适用位置/备注/数量/材料列）",
            type=["xls", "xlsx"], key="bm_qty_up",
            help="旧版 .xls 与新版 .xlsx 均支持；上传后替换当前清单。")
        if up_qty is not None and st.session_state.get("bm_qty_id") != up_qty.file_id:
            try:
                items = parse_quantity_excel(up_qty.getvalue())
                st.session_state.bm_items = items
                st.session_state.bm_qty_id = up_qty.file_id
                st.session_state.bm_qty_err = None
                st.session_state.bm_text_warns = None
            except BridgeMatchError as e:
                st.session_state.bm_qty_err = str(e)
            except Exception as e:
                st.session_state.bm_qty_err = f"数量表解析失败：{e}"
        if st.session_state.get("bm_qty_err"):
            st.error(st.session_state.bm_qty_err)
    with tab_text:
        st.text_area("每行一条：名称 数量 单位（如：钻孔灌注桩 φ1000 h≤40m 砂土 300 m）",
                     key="bm_text_area", height=120,
                     help="解析失败的行会跳过并提示；追加到当前清单末尾。")
        if st.button("追加文本条目", key="bm_text_btn", use_container_width=True):
            txt = st.session_state.bm_text_area.strip()
            if not txt:
                st.warning("请先粘贴文本。")
            else:
                items, warns = parse_text_lines(txt)
                if not items and warns:
                    st.session_state.bm_text_warns = warns
                else:
                    cur = st.session_state.get("bm_items") or []
                    cur.extend(items)
                    st.session_state.bm_items = cur
                    st.session_state.bm_text_warns = warns
                    st.session_state.bm_qty_id = None
        if st.session_state.get("bm_text_warns"):
            st.warning("部分行解析失败已跳过：" + "；".join(st.session_state.bm_text_warns))
    if st.session_state.get("bm_items"):
        items = st.session_state.bm_items
        n_pile = sum(1 for it in items if it.get("kind") == "桩基参数")
        st.success(f"已加载 {len(items)} 条清单条目（其中桩基参数块 {n_pile} 个）")
        with st.expander("条目预览（前 20 条）"):
            rows = []
            for it in items[:20]:
                if it.get("kind") == "桩基参数":
                    rows.append({"部位": it["sheet"], "分项名": it["fenxiang"],
                                 "材料": f"桩基参数 φ{it['dia']:g}/h{it['depth']:g}m",
                                 "单位": "根", "数量": it["count"]})
                else:
                    rows.append({"部位": it["sheet"], "分项名": it["fenxiang"],
                                 "材料": it.get("mat"), "单位": it.get("mat_unit"),
                                 "数量": it.get("value")})
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    st.divider()

    # ---------- ④ 匹配 ----------
    st.markdown("### ④ 匹配")
    has_tpl = bool(st.session_state.get("bm_rules"))
    # 内置模板规则兜底：② 未上传时也按内置 单位工程.xlsx 规则套用
    ready = bool(st.session_state.get("bm_lib")) and bool(st.session_state.get("bm_items"))
    if not ready:
        st.info("请先完成 ① 定额库 与 ③ 工程量清单 的上传（② 模板可选，默认内置）。")
    elif not has_tpl and builtin_rules is not None:
        st.info("当前将按**内置模板规则**（单位工程.xlsx）套用；"
                "如需自定义，上传 ② 模板覆盖后重新匹配。")
    # 钢绞线束长：按一跨桥长计，默认 30m（→ 40m 以内档）；量表有锚具→后张法（孔数=锚具），无锚具→先张法
    # 钻孔地层：按所选地层优先匹配钻孔子目；砂土/黏土 分开可选，唯一匹配对应变体
    # （2026-10-27 用户口径：如 φ≤1000 h≤40m 砂土 / 黏土 两定额靠选地层唯一命中）
    # 桥梁宽度：沥青摊铺宽度档取 ≥W 的最小档，默认 8.5 → 8.5m 以内档（2026-10-27 用户口径）
    c_span, c_soil, c_width = st.columns([1, 1, 1])
    with c_span:
        span_len = float(st.number_input(
            "🌉 钢绞线束长/一跨桥长（m）", min_value=1.0, max_value=200.0, value=30.0,
            step=1.0, key="bm_span_len",
            help="钢绞线束长按一跨桥长计，默认 30m → 套 40m 以内档（≤20m 套 20m 以内档）。"
                 "量表有锚具按后张法（孔数与锚具一致），无锚具按先张法。"))
    with c_soil:
        soil = st.selectbox(
            "🪨 钻孔地层条件", ["砂土/黏土", "砂土", "黏土", "砂砾", "岩石"],
            key="bm_soil",
            help="定额名称含所选地层者优先，含所选以外地层的变体不进候选；"
                 "默认 砂土/黏土（砂土、黏土变体并列）；单独选 砂土 或 黏土 可唯一命中对应变体"
                 "（如 冲击式钻机钻孔 φ≤1000 h≤40m 砂土 / 黏土）。")
    with c_width:
        bridge_width = float(st.number_input(
            "🌉 桥梁宽度（m）", min_value=1.0, max_value=100.0, value=8.5,
            step=0.5, key="bm_bridge_width",
            help="沥青混凝土摊铺宽度档按桥梁宽度取 ≥宽度 的最小档（默认 8.5 → 摊铺宽度8.5m以内）；"
                 "量表含 改性/SBS 套改性沥青混凝土，否则套普通沥青混凝土。"))
    if st.button("🔍 开始匹配", key="bm_run", use_container_width=True, disabled=not ready):
        with st.spinner("正在匹配定额..."):
            lib = dict(st.session_state.bm_lib)
            # 规则优先级：② 上传模板 > 内置模板（单位工程.xlsx）> 库自带规则样例
            rules_data = st.session_state.get("bm_rules") or builtin_rules
            if rules_data:
                lib["rules"] = rules_data["rules"]
                if rules_data.get("divs"):
                    lib["tpl_divs"] = rules_data["divs"]
            results = match_quotas(st.session_state.bm_items, lib,
                                   span_len=span_len, soil=soil,
                                   bridge_width=bridge_width)
            st.session_state.bm_results = results
            st.session_state.bm_lib_merged = lib
            st.session_state.bm_rules_src = (
                "单位工程.xlsx 模板（上传）" if st.session_state.get("bm_rules")
                else "内置单位工程.xlsx 模板" if rules_data
                else "定额库自带规则样例" if lib.get("rules")
                else "全库文本匹配（无模板）")
            st.session_state.bm_results_key = (
                f"{st.session_state.bm_lib_id}|{st.session_state.get('bm_rules_id') or 'builtin'}|"
                f"{st.session_state.bm_qty_id}|{len(st.session_state.bm_items)}|"
                f"{'v2' if (rules_data or {}).get('divs') else 'v1'}|span={span_len:g}"
                f"|soil={soil}|width={bridge_width:g}")
            st.session_state.bm_ai_cache = {}

    results = st.session_state.get("bm_results")
    if results:
        items = st.session_state.bm_items
        cur_key = (f"{st.session_state.bm_lib_id}|{st.session_state.get('bm_rules_id') or 'builtin'}|"
                   f"{st.session_state.bm_qty_id}|{len(items)}|"
                   f"{'v2' if ((st.session_state.get('bm_rules') or builtin_rules) or {}).get('divs') else 'v1'}"
                   f"|span={st.session_state.get('bm_span_len', 30):g}"
                   f"|soil={st.session_state.get('bm_soil', '砂土/黏土')}"
                   f"|width={st.session_state.get('bm_bridge_width', 8.5):g}")
        if st.session_state.get("bm_results_key") != cur_key:
            st.error("⚠️ 清单、定额库或模板已变化，以下为**上一次**匹配结果（旧规则），"
                     "请重新点击「开始匹配」后查看。")

        # ---------- ⑤ 结果 ----------
        st.markdown(f"### ⑤ 匹配结果（规则来源：{st.session_state.get('bm_rules_src', '—')}）")
        n_single = sum(1 for r in results if r["status"] == "single")
        n_multi = sum(1 for r in results if r["status"] == "multi")
        n_none = sum(1 for r in results if r["status"] == "none")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("总条目", len(results))
        c2.metric("唯一匹配", n_single)
        c3.metric("多候选", n_multi)
        c4.metric("无匹配", n_none)

        st.dataframe(pd.DataFrame(result_rows(results)),
                     use_container_width=True, hide_index=True, height=420)
        # 土方规则（用户 2026-10-26 口径）：填方=压实方、挖方=自然方（清单量原样）；
        # 土方不利用 → 买土 1-6-20=填方×1.22（自然方）；余土外运=挖方×1.32（松方）
        earth_reuse_opt = st.selectbox(
            "⛰️ 土方利用情况",
            ["不利用（需外购土）", "利用（不外购）"],
            key="bm_earth_reuse",
            help="土方不利用时套买土定额 1-6-20，买土量=填方（压实方）×1.22（自然方），"
                 "余土外运=挖方（自然方）×1.32（松方）；"
                 "土方利用时不外购土，余土外运=(挖方−填方×1.22)×1.32（松方）。")
        earth_reuse = earth_reuse_opt == "利用（不外购）"
        # 弃运运距：土方与建筑垃圾分开设置，1-6-72（每增运1km）套 (运距-1) 次
        c_soil, c_waste = st.columns(2)
        with c_soil:
            soil_km = int(st.number_input(
                "🚛 土方弃运运距（km）", min_value=1, max_value=100, value=5,
                step=1, key="bm_soil_haul_km",
                help="余土外运 1-6-72（每增运1km）套 (运距-1) 次"))
        with c_waste:
            waste_km = int(st.number_input(
                "🚛 建筑垃圾弃运运距（km）", min_value=1, max_value=100, value=5,
                step=1, key="bm_waste_haul_km",
                help="凿桩头等建筑垃圾弃运 1-6-72 套 (运距-1) 次；弃运量按虚方×2"))
        st.caption("土方：挖方=自然方、填方=压实方（清单量原样套用）。"
                   "土方不利用 → 买土 1-6-20=填方×1.22（自然方），"
                   "余土外运 1-6-71/72=挖方×1.32（松方）；"
                   "土方利用 → 不外购土，余土外运=(挖方−填方×1.22)×1.32。"
                   "凿桩头等建筑垃圾弃运按虚方 ×2、运距按上方建筑垃圾运距。")
        # 模板系数：模板量=混凝土量×系数，默认参考表样例比例，可自行修改
        st.markdown("**📐 模板系数**（模板量=混凝土量×系数，默认按参考表样例比例，可自行修改）")
        tpl_ratios = {}
        tcols = st.columns(4)
        for i, (k, d) in enumerate(TPL_RATIOS):
            with tcols[i % 4]:
                tpl_ratios[k] = float(st.number_input(
                    f"{k}模板系数", min_value=0.0, max_value=100.0, value=float(d),
                    step=0.001, format="%.3f", key=f"bm_tpl_{k}"))
        st.download_button(
            "⬇️ 下载匹配结果 Excel",
            data=build_result_excel(results, st.session_state.get("bm_lib_merged"),
                                    waste_haul_km=waste_km,
                                    soil_haul_km=soil_km,
                                    earth_reuse=earth_reuse,
                                    tpl_ratios=tpl_ratios),
            file_name="桥梁定额匹配结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)

        st.markdown("#### 逐条明细")
        for r in results:
            it = r["item"]
            mat = it.get("mat") or (f"桩基参数 φ{it['dia']:g}/h{it['depth']:g}m"
                                    if it.get("kind") == "桩基参数" else "")
            title = (f"{_STATUS_EMOJI.get(r['status'], '')} #{r['idx'] + 1} "
                     f"[{it['sheet']}] {it['fenxiang']} {mat} "
                     f"— {'唯一匹配' if r['status'] == 'single' else '多候选' if r['status'] == 'multi' else '无匹配'}")
            with st.expander(title):
                if r["matches"]:
                    for c in r["matches"]:
                        e = c["entry"]
                        extra = ""
                        if c.get("note_conv"):
                            extra = f"（{c['note_conv']}）"
                        elif c.get("qty_de") is not None:
                            extra = f"→ 定额工程量 {c['qty_de']:g} {e['unit']}"
                        g = _grade_tag(it, e["name"])
                        st.markdown(
                            f"- **{e['code']} {e['name']}**"
                            f"{'（' + g + '）' if g else ''}"
                            f"（{e['unit']}，"
                            f"单价 {e['price'] if e['price'] is not None else '—'}）"
                            f"　得分 {c['score']}{extra}")
                        if c.get("reasons"):
                            st.caption("　理由：" + "；".join(c["reasons"]))
                for note in r.get("notes") or []:
                    st.info(note)
                if r.get("checklist"):
                    st.warning("**需补充信息**：" + "；".join(r["checklist"]))
                if r.get("rule_combo"):
                    st.markdown("**该分部套定额样例组合（逐条状态）**")
                    st.dataframe(pd.DataFrame([
                        {"编号": x["code"], "名称": x["name"], "单位": x["unit"],
                         "状态": x["status"], "补充说明": x.get("need", "")}
                        for x in r["rule_combo"]]),
                        use_container_width=True, hide_index=True)
                if r["status"] == "multi" and len(r["matches"]) >= 2:
                    ai_key = (f"{r['idx']}|"
                              f"{'-'.join(sorted(c['entry']['code'] for c in r['matches']))}")
                    ai_cache = st.session_state.get("bm_ai_cache") or {}
                    if st.button("🤖 AI 生成说明", key=f"bm_ai_btn_{ai_key}"):
                        with st.spinner("AI 正在对比候选定额..."):
                            expl = ai_explain(_item_display(it), r["matches"])
                            if expl:
                                ai_cache[ai_key] = expl
                            else:
                                ai_cache[ai_key] = {"error": True}
                            st.session_state.bm_ai_cache = ai_cache
                    if ai_key in ai_cache:
                        expl = ai_cache[ai_key]
                        if expl.get("error"):
                            st.warning("AI 说明生成失败，请参考上方规则命中理由。")
                        else:
                            for cand in expl.get("candidates") or []:
                                i = int(cand.get("index", 0)) - 1
                                if 0 <= i < len(r["matches"]):
                                    e = r["matches"][i]["entry"]
                                    st.markdown(f"- **{e['code']} {e['name']}**")
                                    st.caption(f"　理由：{cand.get('reason', '')}")
                                    st.caption(f"　差异：{cand.get('diff', '')}")
                            if expl.get("overall"):
                                st.info("综合建议：" + expl["overall"])


# ===== 公路桥梁定额匹配：公路定额库 + 秦滨式数量总表 → 模板列结构套定额 =====
def _render_road_bridge_page():
    import pandas as pd
    from bridge_match import BridgeMatchError, parse_quantity_excel
    from bridge_match_road import (DEFAULT_PARAMS, build_road_excel,
                                   default_road_library, match_road_quotas,
                                   parse_road_library)

    _page_header("🛣 公路桥梁定额匹配", "rb_refresh",
                 ["rb_lib_up", "rb_lib", "rb_lib_id", "rb_lib_err",
                  "rb_qty_up", "rb_qty_bytes", "rb_qty_id", "rb_qty_err",
                  "rb_items", "rb_result",
                  "rb_hun_km", "rb_gang_km", "rb_tu_km", "rb_long_km",
                  "rb_liq_km", "rb_nijiang_dj", "rb_shatu_pct", "rb_yashi",
                  "rb_ganghu_t", "rb_gjx_shu", "rb_qiaochang", "rb_weiyan_t"])
    st.caption("上传公路定额库 + 秦滨式数量总表，按公路工程量定额自动套用，"
               "输出模板列结构（项/目/节/细目代号+费率编号+定额代号+名称+单位+"
               "工程量1+工程量2+定额调整情况+说明）；"
               "钢筋类型抽换/实际运距/实际桩径等换算集中在「定额调整情况」与「说明」两列。")

    # ---------- ① 定额库 ----------
    st.markdown("### ① 定额库")
    up_lib = st.file_uploader(
        "上传公路定额库 .xls/.xlsx（可选）",
        type=["xls", "xlsx"], key="rb_lib_up",
        help="不传时按参考模板内置口径套用；上传后同编号以库为准，"
             "库缺口的固定子目仍按模板口径兜底。")
    lib = None
    if up_lib is not None:
        if st.session_state.get("rb_lib_id") != up_lib.file_id:
            try:
                lib = parse_road_library(up_lib.getvalue())
                st.session_state.rb_lib = lib
                st.session_state.rb_lib_id = up_lib.file_id
                st.session_state.rb_lib_err = None
            except Exception as e:
                st.session_state.rb_lib_err = f"定额库解析失败：{e}"
                st.session_state.rb_lib = None
                st.session_state.rb_lib_id = None
        lib = st.session_state.get("rb_lib")
    if st.session_state.get("rb_lib_err"):
        st.error(st.session_state.rb_lib_err)
    elif lib:
        st.success(f"已加载 {lib['n']} 条公路定额（同编号以库内名称为准）")
    else:
        lib = default_road_library()  # 内置公路概算定额库（quota_base/公路定额+预算.xls）
        st.info(f"未上传定额库：已使用内置公路概算定额库（{lib['n']} 条），"
                f"上传后同编号以库为准，库缺口子目按模板口径兜底。")

    st.divider()

    # ---------- ② 数量表 ----------
    st.markdown("### ② 工程量数量表")
    up_qty = st.file_uploader(
        "上传秦滨式数量总表（总表含桩基块/支座块/伸缩缝文本块）",
        type=["xls", "xlsx"], key="rb_qty_up")
    items = None
    if up_qty is not None:
        if st.session_state.get("rb_qty_id") != up_qty.file_id:
            try:
                items = parse_quantity_excel(up_qty.getvalue())
                st.session_state.rb_items = items
                st.session_state.rb_qty_bytes = up_qty.getvalue()
                st.session_state.rb_qty_id = up_qty.file_id
                st.session_state.rb_qty_err = None
            except Exception as e:
                st.session_state.rb_qty_err = f"数量表解析失败：{e}"
                st.session_state.rb_items = None
                st.session_state.rb_qty_id = None
        items = st.session_state.get("rb_items")
    if st.session_state.get("rb_qty_err"):
        st.error(st.session_state.rb_qty_err)
    elif items:
        st.success(f"已解析 {len(items)} 条清单项")

    st.divider()

    # ---------- ③ 参数 ----------
    st.markdown("### ③ 换算与参数（默认按参考模板，可修改）")
    pc = st.columns(3)
    with pc[0]:
        hun_km = float(st.number_input("🚛 混凝土搅拌车运距（km）", min_value=1.0,
                                       max_value=100.0, value=float(DEFAULT_PARAMS["hun_km"]),
                                       step=1.0, key="rb_hun_km"))
        gang_km = float(st.number_input("🚚 钢材运输运距（km）", min_value=1.0,
                                        max_value=100.0, value=float(DEFAULT_PARAMS["gang_km"]),
                                        step=1.0, key="rb_gang_km"))
        tu_km = float(st.number_input("⛰ 土方外运运距（km）", min_value=1.0,
                                      max_value=100.0, value=float(DEFAULT_PARAMS["tu_km"]),
                                      step=1.0, key="rb_tu_km"))
        long_km = float(st.number_input("🚛 钢筋笼运输运距（km）", min_value=1.0,
                                        max_value=100.0, value=float(DEFAULT_PARAMS["long_km"]),
                                        step=1.0, key="rb_long_km"))
        liq_km = float(st.number_input("🚚 沥青混合料运距（km）", min_value=1.0,
                                       max_value=100.0, value=float(DEFAULT_PARAMS["liq_km"]),
                                       step=1.0, key="rb_liq_km"))
    with pc[1]:
        nijiang_dj = float(st.number_input("💰 泥浆处置费单价（元/m3）", min_value=0.0,
                                           value=float(DEFAULT_PARAMS["nijiang_dj"]),
                                           step=1.0, key="rb_nijiang_dj"))
        shatu_pct = float(st.number_input("🏖 钻孔砂土比例（%）", min_value=0.0,
                                          max_value=100.0, value=float(DEFAULT_PARAMS["shatu_pct"]),
                                          step=1.0, key="rb_shatu_pct"))
        yashi = float(st.number_input("📐 填方压实换算系数", min_value=1.0, max_value=2.0,
                                      value=float(DEFAULT_PARAMS["yashi"]), step=0.01,
                                      key="rb_yashi",
                                      help="余土外运 = 挖方 − 填方×系数（天然方）"))
        ganghu_t = float(st.number_input(
            "⚙ 钢护筒总量（t，0=按定额说明(16)表内插×每根2m估算）",
            min_value=0.0, value=0.0, step=0.001, format="%.3f",
            key="rb_ganghu_t",
            help="设计已给定钢护筒总质量时填设计值；填 0 则按桩径查表内插（φ1.0→267/"
                 "φ1.5→568/φ1.8→778.6kg/m）×每根2m×根数自动估算"))
        gjx_shu = float(st.number_input("🎯 钢绞线束长档（m）", min_value=10.0,
                                        max_value=100.0, value=float(DEFAULT_PARAMS["gjx_shu"]),
                                        step=10.0, key="rb_gjx_shu"))
    with pc[2]:
        qiaochang = float(st.number_input("🌉 桥梁长度（m，护栏节统计量）", min_value=0.0,
                                          value=0.0, step=0.01, key="rb_qiaochang"))
        weiyan_t = float(st.number_input("🛟 钢围堰量（t，0=从数量表自动识别）", min_value=0.0,
                                         value=0.0, step=0.1, key="rb_weiyan_t"))

    st.divider()

    # ---------- ④ 匹配 ----------
    if items is None:
        st.info("请先上传数量表，然后点击「开始匹配」。")
        return
    if st.button("⚙️ 开始匹配", type="primary", use_container_width=True):
        with st.spinner("正在按公路定额匹配..."):
            try:
                params = {"hun_km": hun_km, "gang_km": gang_km, "tu_km": tu_km,
                          "long_km": long_km, "liq_km": liq_km,
                          "nijiang_dj": nijiang_dj, "shatu_pct": shatu_pct,
                          "yashi": yashi, "ganghu_t": ganghu_t,
                          "gjx_shu": gjx_shu, "qiaochang": qiaochang,
                          "weiyan_t": weiyan_t}
                res = match_road_quotas(items, lib,
                                        st.session_state.get("rb_qty_bytes"),
                                        params)
                st.session_state.rb_result = res
            except BridgeMatchError as e:
                st.error(str(e))
            except Exception as e:
                st.error(f"匹配失败：{e}")

    res = st.session_state.get("rb_result")
    if not res:
        return

    if res.get("warnings"):
        st.warning("**请核对以下口径**：\n\n" + "\n\n".join(
            f"- {w}" for w in res["warnings"]))
    if res.get("skipped"):
        with st.expander("未套定额的清单项（按参考模板口径）"):
            for s in res["skipped"]:
                st.caption("· " + s)

    n_q = sum(1 for r in res["rows"] if r["t"] == "quota")
    st.success(f"匹配完成：共 {len(res['rows'])} 行（{n_q} 条定额子目）——"
               f"工程名称：{res['title']}")
    with st.expander("结果预览（模板列结构）"):
        st.dataframe(pd.DataFrame([
            {"项": r["d"] if r["t"] == "xiang" else "",
             "目": r["d"] if r["t"] == "mu" else "",
             "节": r["d"] if r["t"] == "jie" else "",
             "细目": r["d"] if r["t"] == "ximu" else "",
             "费率编号": r["rate"], "定额代号": r["code"], "名称": r["name"],
             "单位": r["unit"], "工程量1": r["q1"], "工程量2": r["q2"],
             "定额调整情况": r["adj"], "说明": r["note"]}
            for r in res["rows"]]),
            use_container_width=True, hide_index=True, height=420)
    out_name = f"{res['title']}_公路桥梁定额匹配.xlsx" if res.get("title") \
        else "公路桥梁定额匹配结果.xlsx"
    st.download_button(
        "⬇️ 下载匹配结果 Excel",
        data=build_road_excel(res),
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True)


# ===== 正反项目案例：子项注册表（侧边栏点击 → 案例页）=====
CASE_ITEMS = [
    {"key": "saimiyi", "short": "塞米伊环城公路（国外项目）",
     "file": "案例_塞米伊环城公路.md"},
    {"key": "yuxinqiao", "short": "于新桥拆除重建",
     "file": "案例_于新桥拆除重建.md"},
    {"key": "guogaowang", "short": "国高网高速概算",
     "file": "案例_国高网高速概算.md"},
    {"key": "dongcheng", "short": "湛江东盛大道全过程咨询",
     "file": "案例_东盛大道全过程咨询.md"},
]

# 正面案例模板节序（2026-11 起按《建设项目造价专业正反案例模板》重设计）：
# 基本信息/案例摘要/项目与背景/面临的问题或挑战/措施与做法/实施成效/
# 经验总结 突出展示；推广应用建议/主要依据/附件清单/关键指标 折叠查看。
_CASE_MAIN_SECTIONS = [
    ("基本信息", None),
    ("案例摘要", "📝 案例摘要"),
    ("项目与背景", "🏗️ 项目与背景"),
    ("面临的问题或挑战", "🧩 面临的问题或挑战"),
    ("措施与做法", "🛠️ 措施与做法"),
    ("实施成效", "✅ 实施成效"),
    ("经验总结", "💡 经验总结"),
]


def _parse_case_md(path) -> dict:
    """解析案例 md：按 ## 分节，返回 {'title': 主标题, 小节名: 内容}。"""
    result = {}
    cur, buf = "title", []
    for line in path.read_text(encoding="utf-8").splitlines():
        if line.startswith("## "):
            result[cur] = "\n".join(buf).strip()
            cur, buf = line[3:].strip(), []
        else:
            buf.append(line)
    result[cur] = "\n".join(buf).strip()
    if result.get("title", "").startswith("# "):
        result["title"] = result["title"][2:].strip()
    return result


def _case_type_of(sections: dict) -> str:
    """从基本信息表提取案例类型：'正面' / '反面' / ''（未标注）。"""
    for line in (sections.get("基本信息") or "").splitlines():
        if "案例类型" not in line:
            continue
        if "正面" in line:
            return "正面"
        if "反面" in line:
            return "反面"
    return ""


def _case_type_badge(sections: dict) -> str:
    """案例类型小标题徽标 HTML（正面绿色/反面红色），未标注时返回空串。"""
    _t = _case_type_of(sections)
    if _t == "正面":
        return ('<span style="background:#e8f5e9;color:#2e7d32;'
                'padding:2px 10px;border-radius:12px;font-size:0.85rem;'
                'font-weight:600;">📗 正面案例</span>')
    if _t == "反面":
        return ('<span style="background:#fdecea;color:#c62828;'
                'padding:2px 10px;border-radius:12px;font-size:0.85rem;'
                'font-weight:600;">📕 反面案例</span>')
    return ""


def _render_case_page(case_key: str):
    """正反项目案例页：按正反案例模板展示——基本信息表 + 案例摘要 + 项目与背景 +
    面临的问题或挑战 + 措施与做法 + 实施成效 + 经验总结 突出展示；
    推广应用建议/主要依据/附件清单/关键指标 折叠；
    底部围绕本案例提问（本案例优先检索，查不到回退全库）。"""
    item = next((c for c in CASE_ITEMS if c["key"] == case_key), None)
    if item is None:
        st.session_state.active_case = None
        st.rerun()

    if st.button("← 返回功能导航", use_container_width=True):
        st.session_state.active_case = None
        st.rerun()

    st.title("🏗️ 正反项目案例")
    st.header(item["short"])

    sections = _parse_case_md(CASE_BASE_DIR / item["file"])
    badge = _case_type_badge(sections)
    if badge:
        st.markdown(badge, unsafe_allow_html=True)
    if sections.get("title"):
        st.caption(sections["title"])

    main_names = {n for n, _ in _CASE_MAIN_SECTIONS}
    for name, header in _CASE_MAIN_SECTIONS:
        content = sections.get(name, "")
        if not content:
            continue
        if header:
            st.markdown(f"### {header}")
        st.markdown(content)

    rest = [(t, c) for t, c in sections.items() if t not in main_names and t != "title"]
    if rest:
        with st.expander("📖 查看推广应用建议与依据材料"):
            for t, c in rest:
                st.markdown(f"#### {t}")
                st.markdown(c)

    st.divider()
    st.markdown("### 💬 围绕本案例提问")
    st.caption("优先在本案例培训资料中检索，查不到时回退全库（政策/定额/其他案例）。")

    if st.session_state.get("case_history_owner") != case_key:
        st.session_state.case_history = []
        st.session_state.case_history_owner = case_key

    for msg in st.session_state.case_history:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    engine = get_engine()
    question = st.chat_input(f"关于「{item['short']}」有什么想问的？")
    if question:
        st.session_state.case_history.append({"role": "user", "content": question})
        with st.chat_message("user"):
            st.markdown(question)
        with st.chat_message("assistant"):
            with st.spinner("正在检索本案例资料并生成回答..."):
                history = [
                    {"role": m["role"], "content": m["content"]}
                    for m in st.session_state.case_history[:-1]
                ]
                answer = engine.chat(
                    question, history,
                    kb_chunks=engine.case_kb_chunks(item["file"]),
                    kb_label=f"项目案例「{item['short']}」",
                )
            st.markdown(answer)
        st.session_state.case_history.append({"role": "assistant", "content": answer})





def _render_engine_card(fee_result):
    """渲染单个费种的引擎计算结果卡片"""
    params = fee_result.get("参数", {})
    result_val = fee_result.get("结果(万元)") or fee_result.get("结果(元)")
    unit = "万元" if "结果(万元)" in fee_result else "元"

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(label="费种", value=fee_result.get("费种", ""))
    with col2:
        st.metric(label=f"金额（{unit}）", value=f"{result_val}")
    with col3:
        basis_text = _basis_with_links(fee_result.get('依据', ''))
        st.markdown(f"<small>依据：{basis_text}</small>", unsafe_allow_html=True)

    with st.expander("查看计算明细", expanded=True):
        if params:
            st.markdown("**输入参数**：")
            for k, v in params.items():
                st.markdown(f"- {k}：**{v}**")
        basic = fee_result.get("基本设计收费(万元)")
        if basic is not None:
            st.markdown(f"**基本设计收费**：**{basic} 万元**")
        other = fee_result.get("其他设计收费明细")
        if other:
            st.markdown("**其他设计收费**：")
            for od in other:
                st.markdown(f"- {od['项目']}：**{od['费用(万元)']} 万元**")
        items = fee_result.get("分项明细")
        if items and len(items) > 1:
            st.markdown("**分项明细**：")
            for item in items:
                st.markdown(
                    f"- {item.get('类别', '')}：基数 {item.get('基数(万元)', '')} 万元 "
                    f"→ **{item.get('费用(元)', '')} 元**"
                )
        steps = fee_result.get("计算步骤")
        if steps:
            # 区分三种步骤格式：
            # A. 统一格式（有"代入"键）→ 步骤：公式；代入 代入值 → 结果
            # B. 粗略估算（有"步骤"键但无"代入"）→ 步骤：公式 → 结果
            # C. 分档累进（有"区间"键但无"步骤"）→ 区间：金额 × 费率 = 费用
            if steps[0].get("代入"):
                st.markdown("**计算过程**：")
                for i, s in enumerate(steps, 1):
                    step_name = s.get("步骤", "")
                    formula = s.get("公式", "")
                    substitution = s.get("代入", "")
                    result_step = s.get("结果", "")
                    if substitution:
                        st.markdown(
                            f"**{i}. {step_name}**：{formula}；"
                            f"代入 {substitution} → **{result_step}**"
                        )
                    else:
                        st.markdown(
                            f"**{i}. {step_name}**：{formula} → **{result_step}**"
                        )
            elif steps[0].get("步骤"):
                st.markdown("**计算过程**：")
                for i, s in enumerate(steps, 1):
                    formula = s.get("公式", "")
                    result_step = s.get("结果", "")
                    st.markdown(
                        f"**{i}. {s.get('步骤', '')}**：{formula} → **{result_step}**"
                    )
            else:
                st.markdown("**分档计算**：")
                for s in steps:
                    rate = s.get("费率(%)", s.get("费率(‰)", ""))
                    unit = "%" if "费率(%)" in s else "‰"
                    st.markdown(
                        f"- {s.get('区间', '')}：{s.get('金额(万元)', '')}万元 "
                        f"× {rate}{unit} = **{s.get('费用(万元)', '')}万元**"
                    )
        rate_detail = fee_result.get("费率明细")
        if rate_detail:
            st.markdown("**费率档位**：")
            final_val = fee_result.get("结果(万元)")
            for rd in rate_detail:
                rv = rd.get("费用(万元)")
                mark = ""
                if (final_val is not None and rv is not None
                        and isinstance(rv, (int, float))
                        and abs(float(rv) - float(final_val)) < 0.005):
                    mark = " ✔（最终采用）"
                st.markdown(f"- {rd.get('费率', '')} → {rv} 万元{mark}")
        if "分摊" in fee_result:
            st.caption(fee_result["分摊"])
        adjustment = fee_result.get("计费额调整")
        if adjustment and adjustment.get("触发调整"):
            st.info(adjustment.get("说明", ""))






# ===== 多费种迭代计算渲染函数 =====



def _render_iteration_result(result):
    """渲染模式2：迭代收敛结果。"""
    import pandas as pd

    st.markdown("## 迭代计算（总投资收敛）")

    params = result["输入参数"]
    col1, col2 = st.columns(2)
    col1.metric("建安工程费", f"{params['建安工程费(万元)']} 万元")
    col2.metric("设备购置费", f"{params['设备购置费(万元)']} 万元")

    st.info(
        "**迭代原理**：建设管理费、建设项目前期工作咨询费、环境影响咨询费依赖总投资；"
        "总投资又包含这些二类费本身。通过反复迭代使总投资收敛到稳定值。"
    )

    # 收敛过程表
    history = result["迭代过程"]
    steps_data = []
    for h in history:
        steps_data.append({
            "迭代": h["迭代次数"],
            "总投资(万元)": round(h["总投资(万元)"], 2),
            "二类费合计(万元)": round(h["二类费合计(万元)"], 2),
            "变化(万元)": round(h.get("变化(万元)", 0), 4),
        })

    st.markdown("### 收敛过程")
    st.dataframe(pd.DataFrame(steps_data), use_container_width=True, hide_index=True)

    final = result["收敛结果"]
    converged = result["已收敛"]

    # 额外费用提示
    extra_fees = result.get("额外费用", [])
    extra_note = ""
    if extra_fees:
        extra_total = sum(e["金额(万元)"] for e in extra_fees)
        extra_names = "、".join(f"{e['名称']} {e['金额(万元)']}万" for e in extra_fees)
        extra_note = f"\n\n含用户指定额外费用：{extra_names}（已计入合计）"

    if converged:
        yubei_val = final.get("预备费(万元)", 0)
        proj_total = final.get("项目总投资(万元)", final["总投资(万元)"])
        st.success(
            f"✅ 经过 **{result['迭代次数']}** 次迭代已收敛 "
            f"（阈值 {result['收敛阈值(万元)']} 万元）。\n\n"
            f"静态总投资：**{final['总投资(万元)']:.2f} 万元**，"
            f"二类费合计：**{final['二类费合计(万元)']:.2f} 万元**\n\n"
            f"预备费：**{yubei_val:.2f} 万元**（(一类费+二类费)×5%），"
            f"项目总投资：**{proj_total:.2f} 万元**"
            f"{extra_note}"
        )
    else:
        st.warning(
            f"⚠️ 经过 {result['迭代次数']} 次迭代未完全收敛"
        )

    # 最终明细
    st.markdown("### 收敛后各项费用")
    fees = final["各项费用"]
    for fee_key, val in sorted(fees.items()):
        st.markdown(f"- **{fee_key}**：{val:.2f} 万元")
    # 预备费单独显示
    yubei_final_val = final.get("预备费(万元)")
    if yubei_final_val is not None and yubei_final_val > 0:
        st.markdown(f"- **预备费**：{yubei_final_val:.2f} 万元")
    proj_final = final.get("项目总投资(万元)")
    if proj_final is not None:
        st.markdown(f"\n**项目总投资（含预备费）：{proj_final:.2f} 万元**")

    with st.expander("查看每轮迭代详细数据"):
        for h in history:
            st.markdown(f"#### 第 {h['迭代次数']} 轮")
            fees = h["各项费用"]
            for fee_key, val in sorted(fees.items()):
                st.markdown(f"- {fee_key}：{val:.2f} 万元")
            st.caption(f"总投资：{h['总投资(万元)']:.2f} 万元 ｜ 变化：{h['变化(万元)']:.2f} 万元")

    # 响应文本（含完整明细，确保 rerun 后不丢失）
    yb_val = final.get("预备费(万元)", 0)
    proj_total = final.get("项目总投资(万元)", final["总投资(万元)"])
    yb_text = f"\n预备费：**{yb_val:.2f} 万元**（(一类费+二类费)×5%）" if yb_val > 0 else ""
    fee_lines = []
    fees_detail = final["各项费用"]
    for fee_key, val in sorted(fees_detail.items()):
        fee_lines.append(f"- **{fee_key}**：{val:.2f} 万元")
    if yb_val > 0:
        fee_lines.append(f"- **预备费**：{yb_val:.2f} 万元")
    return (
        f"## 迭代计算结果\n\n"
        f"经过 **{result['迭代次数']}** 次迭代，静态总投资收敛至 "
        f"**{final['总投资(万元)']:.2f} 万元**，"
        f"二类费合计 **{final['二类费合计(万元)']:.2f} 万元**。"
        f"{yb_text}\n\n"
        f"### 收敛后各项费用\n\n" + "\n".join(fee_lines) + "\n\n"
        f"项目总投资（含预备费）：**{proj_total:.2f} 万元**。"
    )


def _render_comparison_result(result):
    """渲染模式3：多方案比选结果。"""
    import pandas as pd

    st.markdown("## 多方案比选 / 敏感性分析")

    sweep = result["扫描参数"]
    st.info(
        f"**扫描参数**：{sweep['参数描述']}，"
        f"共 {len(sweep['值列表'])} 个方案："
        f"{', '.join(str(v) + sweep.get('单位', '') for v in sweep['值列表'])}"
    )

    # 对比表
    st.markdown("### 费用对比表（单位：万元）")
    comparison_rows = result["对比表"]
    df = pd.DataFrame(comparison_rows)
    st.dataframe(df, use_container_width=True, hide_index=True)

    # 柱状图
    st.markdown("### 二类费合计对比")
    chart_data = {}
    for s in result["方案列表"]:
        chart_data[s["方案名称"]] = s["二类费合计(万元)"]
    st.bar_chart(chart_data)

    # 方案明细
    with st.expander("查看每个方案的各项费用明细"):
        for s in result["方案列表"]:
            st.markdown(f"#### {s['方案名称']}")
            fees = s["各项费用"]
            for fee_key, val in sorted(fees.items()):
                st.markdown(f"- **{fee_key}**：{val:.2f} 万元")
            st.metric("二类费合计", f"{s['二类费合计(万元)']:.2f} 万元")
            st.metric("总投资", f"{s['总投资(万元)']:.2f} 万元")

    # 构建持久化响应文本（含完整对比表，确保 rerun 后不丢失）
    # Markdown 表格
    col_keys = list(comparison_rows[0].keys()) if comparison_rows else []
    md_table = ""
    if col_keys:
        md_table += "| " + " | ".join(str(c) for c in col_keys) + " |\n"
        md_table += "|" + "|".join(":--" for _ in col_keys) + "|\n"
        for row in comparison_rows:
            md_table += "| " + " | ".join(str(row.get(c, "")) for c in col_keys) + " |\n"

    # 方案明细文本
    detail_lines = []
    for s in result["方案列表"]:
        detail_lines.append(f"**{s['方案名称']}**：二类费合计 {s['二类费合计(万元)']:.2f} 万，"
                           f"总投资 {s['总投资(万元)']:.2f} 万")

    return (
        f"## 多方案比选结果\n\n"
        f"扫描参数：{sweep['参数描述']}，共 {len(sweep['值列表'])} 个方案。\n\n"
        f"### 费用对比表（单位：万元）\n\n"
        f"{md_table}\n\n"
        f"### 各方案汇总\n\n" + "\n".join(f"- {l}" for l in detail_lines)
    )


# ===== 清单式任务流：pending_task 渲染与触发（阶段 B） =====

def _render_task_area():
    """清单式任务流渲染：ask 阶段提问气泡 / done 阶段结果重绘。

    每次 rerun 调用（位置仿 pending_comparison），状态存于 st.session_state.pending_task。
    """
    task = st.session_state.get("pending_task")
    if not task:
        return
    # region 切换随时刷新进 ctx（费率档位/河北服务清单据此重建）
    task["region"] = st.session_state.get("selected_region")
    if task.get("phase") == "ask":
        with st.chat_message("assistant"):
            _ask_bubble(task)
    elif task.get("phase") == "done":
        _render_result_payload(task)


def _cancel_task(ctx: dict | None = None):
    """结束本次任务：清状态 + 告别消息。"""
    st.session_state.pop("pending_task", None)
    st.session_state.messages.append({
        "role": "assistant",
        "content": "已结束本次计算。如需重新计算，请直接输入新的问题。",
    })
    st.rerun()


def _split_item_mode() -> bool:
    """是否处于「分开做概算表（每单项一份）」模式（多单项）。"""
    return (st.session_state.get("pdf_split_mode") == "分开做概算表（每单项一份）"
            and len(st.session_state.get("zhuan_ye_list") or []) > 1)


def _item_gl(grp: dict, idx: int) -> str:
    """单项标签（与结果页导出文件名/概算成果命名一致）：组名或「单项N」。"""
    _cn = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十"]
    return (str((grp or {}).get("name") or "").strip()
            or f"单项{_cn[idx - 1] if idx <= len(_cn) else idx}")


def _step_back(ctx: dict, from_stage: str | None = None) -> None:
    """⬅ 上一步：回退到上一阶段重新修改（已填数据保留，各卡以原值为默认）。

    撤销目标阶段及之后阶段的「完成」标记，并置 ctx["_reask"] 强制重问
    目标阶段；目标阶段重新确认完成后（确认回调清除 _reask）恢复常规推进。
    """
    stages = CHECKLIST_STAGES
    stage = from_stage if from_stage is not None else current_stage(ctx)
    idx = stages.index(stage) if stage in stages else len(stages)
    if idx <= 0:
        return  # 已在第一步，无上一步
    target_idx = idx - 1
    ctx["_reask"] = stages[target_idx]
    ctx["_subflow"] = None
    # 撤销目标及之后阶段的完成标记（金额等已填数据保留作默认值）
    if target_idx <= 0:
        ctx["fees_confirmed"] = False
    if target_idx <= 2:
        ctx["custom_answered"] = False
        # 征地拆迁费提问重新确认（已填金额保留；已有识别费用时自动跳过提问）
        ctx.pop("zd_asked", None)
        ctx.pop("zd_exists", None)
        ctx.pop("zd_placement", None)
        ctx.pop("zd_refine_asked", None)
        # 分开导出按单项问：重问各单项征地拆迁/自定义费用
        # （item_custom_fees 数据保留作默认值；item_fees_base 重新快照）
        ctx.pop("item_fees_done", None)
        ctx.pop("_item_stage", None)
        ctx.pop("item_fees_base", None)
    if target_idx <= 3:
        ctx["params_done"] = False
        ctx["param_idx"] = 0
    if target_idx <= 4:
        ctx["discount_answered"] = False
    ctx["qno"] = ctx.get("qno", 1) + 1


def _task_progress_header(ctx: dict, stage: str | None):
    """提问气泡顶部：进度徽标 + 已识别信息摘要。"""
    stages = CHECKLIST_STAGES
    idx = stages.index(stage) if stage in stages else len(stages)
    st.markdown(
        f"### 📋 二类费计算 — 信息确认\n\n"
        f"**第 {idx + 1} 步/共 {len(stages) + 1} 步**："
        f"{_TASK_STAGE_LABELS.get(stage, '开始计算')}"
    )
    st.progress((idx + 1) / (len(stages) + 2))
    if ctx.get("original_mode") in ("iteration", "comparison"):
        st.info("原请求为**迭代/比选**模式但缺少金额信息："
                "先按清单补齐信息，随后按标准级联流程计算（含内部迭代收敛）。")

    fees = ctx.get("fees") or []
    if fees:
        fees_txt = "、".join(FEE_CATALOG.get(f, {}).get("label", f) for f in fees)
        st.markdown(f"**已选费种**：{fees_txt}")
    am = ctx.get("amounts", {})
    known_bits = []
    if am.get("jianan") is not None:
        known_bits.append(f"建安费 {am['jianan']} 万")
    if am.get("shebei_known") and am.get("shebei"):
        known_bits.append(f"设备费 {am['shebei']} 万")
    if am.get("total_investment") is not None:
        known_bits.append(f"项目总投资 {am['total_investment']} 万")
    if known_bits:
        st.markdown(f"**已识别金额**：{'，'.join(known_bits)}")
    elif am.get("jianan") is not None and not am.get("shebei_known"):
        st.caption("💡 设备费未提及（按 0 计算）。可直接回复「设备费 XXX 万」或「无设备费」。")
    st.markdown("---")


def _ask_bubble(ctx: dict):
    """ask 阶段总调度：按 current_stage 分发子渲染；全部完成 → 结算确认。"""
    if (not ctx.get("item_fees_done") and _split_item_mode()
            and ctx.get("fees_confirmed")
            and current_stage(ctx) in ("custom", "params", "discount")):
        # 分开导出（每单项一份概算表）：征地拆迁费/自定义费用按单项逐个问
        # （AI 与传统模式统一走此流程；合同价仍项目级，流程内提问）
        _ask_item_fees(ctx)
        return
    if ctx.get("ai_guided") and not ctx.get("_ai_fallback"):
        _ask_ai_guide(ctx)
        return
    _msg = ctx.pop("_ai_fallback_msg", None)
    if _msg:
        st.warning(_msg)
    stage = current_stage(ctx)
    _task_progress_header(ctx, stage)

    if stage == "fees":
        _ask_fees(ctx)
    elif stage == "amounts":
        _ask_amounts(ctx)
    elif stage == "custom":
        _ask_custom(ctx)
    elif stage == "params":
        _ask_params(ctx)
    elif stage == "discount":
        _ask_discount(ctx)
    else:
        _finalize_confirm(ctx)

    st.markdown("---")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("⬅ 上一步", key=f"ck_back_{stage}_{ctx.get('qno', 1)}",
                     use_container_width=True, disabled=stage == "fees"):
            _step_back(ctx, from_stage=stage)
            st.rerun()
    with c2:
        if st.button("🗑 结束本次计算", key=f"ck_cancel_{stage}_{ctx.get('qno', 1)}",
                     use_container_width=True):
            _cancel_task(ctx)


def _ask_fees(ctx: dict):
    qno = ctx.get("qno", 1)
    if ctx.get("preset_all") and not ctx.get("fees_confirmed"):
        # 未指明费种：默认全量计算，先确认是否有不需要的费种
        preset = _preset_fee_numbers()
        excluded = ctx.get("fee_exclusions") or set()
        st.markdown(
            f"未指明具体费种，将**默认计算全部 {len(preset)} 项二类费**"
            f"（不含交易服务费）。\n\n**是否有不需要计算的费用？**"
            f"点击序号即可排除，或直接回复序号。")
        st.caption("预备费默认计入（稍后单独确认是否考虑），不在以上可排除清单中。")
        cols = st.columns(3)
        for n, name, label in preset:
            is_ex = name in excluded
            with cols[(n - 1) % 3]:
                if st.button(
                    f"{'🚫 ' if is_ex else ''}{n}. {label}",
                    key=f"ck_fees_ex_{name}_{qno}",
                    use_container_width=True,
                    type="primary" if is_ex else "secondary",
                ):
                    if is_ex:
                        excluded.discard(name)
                        if name not in ctx["fees"]:
                            ctx["fees"].append(name)
                    else:
                        excluded.add(name)
                        if name in ctx["fees"]:
                            ctx["fees"].remove(name)
                    ctx["fee_exclusions"] = excluded
                    ctx["qno"] = qno + 1
                    st.rerun()
        remain = len(ctx["fees"])
        if st.button(f"✅ 确认，按剩余 {remain} 项计算",
                     key=f"ck_fees_allok_{qno}",
                     use_container_width=True, type="primary"):
            ctx["fees_confirmed"] = True
            ctx.pop("_reask", None)
            ctx["qno"] = qno + 1
            st.rerun()
        st.caption(
            "也可以直接回复：**「不需要 3、5」** / **「3、5」**"
            "（排除第 3、5 项）/ **「除了 3-5，其他都要」**。"
            "全部都要算的话直接点上方确认即可。")
        return
    st.markdown("请问需要计算哪些**二类费**？（点击选择，可多选）")
    defs = fee_catalog_for(ctx.get("region"), exclude_reserve=True)
    selected = set(ctx.get("fees") or [])
    cols = st.columns(3)
    for i, fd in enumerate(defs):
        is_sel = fd["name"] in selected
        with cols[i % 3]:
            if st.button(
                f"{'✅ ' if is_sel else ''}{fd['label']}",
                key=f"ck_fees_{i}_{qno}",
                use_container_width=True,
                type="primary" if is_sel else "secondary",
            ):
                if is_sel:
                    selected.discard(fd["name"])
                else:
                    selected.add(fd["name"])
                ctx["fees"] = [d["name"] for d in defs if d["name"] in selected]
                ctx["qno"] = qno + 1
                st.rerun()
    col1, col2 = st.columns(2)
    with col1:
        if st.button("📦 全部二类费（不含交易服务费）", key=f"ck_fees_all_{qno}",
                     use_container_width=True):
            ctx["fees"] = [f for f in FEE_CATALOG if f not in ("交易服务费", "预备费")]
            ctx["preset_all"] = True
            ctx["qno"] = qno + 1
            st.rerun()
    with col2:
        if st.button("✅ 确认所选费种", key=f"ck_fees_ok_{qno}", use_container_width=True,
                     disabled=not selected):
            ctx["fees_confirmed"] = True
            ctx.pop("_reask", None)
            ctx["qno"] = qno + 1
            st.rerun()


def _ask_amounts(ctx: dict):
    qno = ctx.get("qno", 1)
    am = ctx["amounts"]
    if am.get("total_investment") is not None:
        st.info(
            f"已识别**项目总投资 {am['total_investment']} 万元**。"
            f"总投资及派生费用由程序内部迭代计算，这里只需提供建安工程费。"
        )
    st.markdown("**建安工程费**是多少？（万元）")
    c1, c2 = st.columns(2)
    with c1:
        jianan_in = st.number_input(
            "建安工程费（万元）", min_value=0.0,
            value=float(am.get("jianan") or 0.0), step=100.0,
            key=f"ck_amounts_jianan_{qno}")
    with c2:
        shebei_in = st.number_input(
            "设备购置费（万元，没有可留空）", min_value=0.0,
            value=float(am.get("shebei") or 0.0), step=50.0,
            key=f"ck_amounts_shebei_{qno}")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("✅ 确认金额", key=f"ck_amounts_ok_{qno}",
                     use_container_width=True,
                     disabled=jianan_in is None or jianan_in <= 0):
            am["jianan"] = float(jianan_in)
            if shebei_in and shebei_in > 0:
                am["shebei"] = float(shebei_in)
                am["shebei_known"] = True
            else:
                am.pop("shebei", None)
                am.pop("shebei_known", None)
            ctx.pop("_reask", None)
            ctx["qno"] = qno + 1
            st.rerun()
    with col2:
        if st.button("🚫 无设备费", key=f"ck_amounts_nosb_{qno}",
                     use_container_width=True):
            am["shebei"] = 0.0
            am["shebei_known"] = True
            st.rerun()


def _ask_zhengchai(ctx: dict, qno: int) -> bool:
    """计算前主动询问征地拆迁费：是否存在 → 列示方式（单独列项/计入其他费）
    → 是否需要细化 → 进入对应填写子流程。

    Returns True 表示本回合正在提问征地拆迁费（调用方应立即返回，不渲染后续问题）。
    查询已自带征地拆迁费时跳过提问，默认单独列项（确认页可取消）。
    """
    if _zd_amount(ctx.get("custom_fees")) > 0:
        ctx["zd_asked"] = True
        ctx["zd_exists"] = True
        ctx.setdefault("zhengchai_separate", True)
        return False
    # AI 模式降级兜底：问答记录里已答过征地拆迁 → 一律不再重复提问
    if not ctx.get("zd_asked"):
        for _rd in ctx.get("_ai_rounds") or []:
            for _qa in _rd.get("qa") or []:
                if re.search(r"征地|拆迁", str(_qa.get("text") or "")) \
                        and str(_qa.get("answer") or "").strip():
                    ctx["zd_asked"] = True
                    ctx["zd_exists"] = not re.search(
                        r"没有|无|不涉及", str(_qa["answer"]))
                    break
            if ctx.get("zd_asked"):
                break
    if not ctx.get("zd_asked"):
        st.markdown(
            "**🏗️ 是否有征地拆迁费？**"
            "（如土地征用、房屋拆迁、青苗补偿等，通常不随其他费用打折）")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("✅ 有征地拆迁费", key=f"ck_zd_yes_{qno}",
                         use_container_width=True):
                ctx["zd_asked"] = True
                ctx["zd_exists"] = True
                ctx["qno"] = qno + 1
                st.rerun()
        with c2:
            if st.button("🚫 没有", key=f"ck_zd_no_{qno}",
                         use_container_width=True):
                ctx["zd_asked"] = True
                ctx["zd_exists"] = False
                ctx["qno"] = qno + 1
                st.rerun()
        return True
    if ctx.get("zd_exists") and ctx.get("zd_placement") is None:
        st.markdown("**🏗️ 征地拆迁费如何列示？**")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("📑 单独列项（与工程费、工程建设其他费并列展示）",
                         key=f"ck_zd_pl_sep_{qno}", use_container_width=True):
                ctx["zd_placement"] = "separate"
                ctx["zhengchai_separate"] = True
                ctx["qno"] = qno + 1
                st.rerun()
        with c2:
            if st.button("📦 计入工程建设其他费（作为二类费组成部分）",
                         key=f"ck_zd_pl_in_{qno}", use_container_width=True):
                ctx["zd_placement"] = "inside"
                ctx["zhengchai_separate"] = False
                ctx["qno"] = qno + 1
                st.rerun()
        return True
    if ctx.get("zd_exists") and not ctx.get("zd_refine_asked"):
        st.markdown(
            "**🏗️ 征地拆迁费是否需要细化？**"
            "（细化后逐项列示，如 房屋征收 / 青苗补偿 / 管线迁改）")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("📋 需要细化（逐项填写）", key=f"ck_zd_refine_yes_{qno}",
                         use_container_width=True):
                ctx["zd_refine_asked"] = True
                ctx["_subflow"] = "zd_detail"
                ctx["qno"] = qno + 1
                st.rerun()
        with c2:
            if st.button("💰 不需要，填一笔总额", key=f"ck_zd_refine_no_{qno}",
                         use_container_width=True):
                ctx["zd_refine_asked"] = True
                ctx["_subflow"] = "zd_single"
                ctx["qno"] = qno + 1
                st.rerun()
        return True
    return False


def _zd_detail_loop(ctx: dict):
    """征地拆迁费细化填写：多行表格，逐项自动加「征地拆迁费·」前缀。

    前缀保证各细项仍被识别为征地拆迁费（单独列项归组、建管费基数扣除、
    预备费基数剔除均按 征地|拆迁 正则匹配）。
    """
    import pandas as pd
    qno = ctx.get("qno", 1)
    st.markdown("**📋 征地拆迁费明细填写**")
    st.caption("每行一项：填「细项名称」与「金额（万元）」，留空行忽略；"
               "名称自动加前缀「征地拆迁费·」。")
    rows = st.session_state.get(f"zd_rows_{qno}")
    if not rows:
        rows = [{"细项名称": "", "金额(万元)": 0.0} for _ in range(3)]
        st.session_state[f"zd_rows_{qno}"] = rows
    edited = st.data_editor(
        pd.DataFrame(rows),
        key=f"ck_zd_editor_{qno}",
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "细项名称": st.column_config.TextColumn("细项名称", width="large"),
            "金额(万元)": st.column_config.NumberColumn(
                "金额(万元)", min_value=0.0, format="%.2f"),
        },
    )
    c1, c2 = st.columns(2)
    with c1:
        if st.button("✅ 添加明细", key=f"ck_zd_det_ok_{qno}",
                     use_container_width=True):
            added = 0
            for r in edited.to_dict("records"):
                nm = str(r.get("细项名称") or "").strip()
                try:
                    amt = float(r.get("金额(万元)") or 0)
                except (TypeError, ValueError):
                    amt = 0.0
                if nm and amt > 0:
                    ctx.setdefault("custom_fees", []).append(
                        {"名称": f"征地拆迁费·{nm}",
                         "金额(万元)": round(amt, 4)})
                    added += 1
            if added:
                ctx["_subflow"] = None
                ctx["qno"] = qno + 1
                st.rerun()
            else:
                st.warning("请至少填写一项：细项名称 + 金额（万元）> 0。")
    with c2:
        if st.button("↩ 返回", key=f"ck_zd_det_back_{qno}",
                     use_container_width=True):
            ctx["_subflow"] = None
            ctx["qno"] = qno + 1
            st.rerun()


def _zd_single_loop(ctx: dict):
    qno = ctx.get("qno", 1)
    st.markdown("**💰 征地拆迁费（一笔总额）**")
    amount = st.number_input("金额（万元）", min_value=0.0, value=0.0, step=10.0,
                             key=f"ck_zd_single_{qno}")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("✅ 添加", key=f"ck_zd_single_ok_{qno}",
                     use_container_width=True, disabled=amount <= 0):
            ctx.setdefault("custom_fees", []).append(
                {"名称": "征地拆迁费", "金额(万元)": round(float(amount), 4)})
            ctx["_subflow"] = None
            ctx["qno"] = qno + 1
            st.rerun()
    with c2:
        if st.button("↩ 返回", key=f"ck_zd_single_back_{qno}",
                     use_container_width=True):
            ctx["_subflow"] = None
            ctx["qno"] = qno + 1
            st.rerun()


def _ask_custom(ctx: dict):
    qno = ctx.get("qno", 1)
    cf = ctx.get("custom_fees") or []
    co = ctx.get("contract_overrides") or {}
    if cf:
        st.markdown("**已添加自定义费用**：")
        for e in cf:
            st.markdown(f"- {e.get('名称')}：{e.get('金额(万元)')} 万元")
    if co:
        st.markdown("**已设置合同价/合同费率**：")
        for fn, ov in co.items():
            if ov.get("type") == "rate":
                desc = f"合同费率 {ov.get('rate')}% × {ov.get('base', '工程费')}"
            else:
                desc = f"一口价 {ov.get('amount_wan')} 万元"
            st.markdown(f"- {FEE_CATALOG.get(fn, {}).get('label', fn)}：{desc}")

    sub = ctx.get("_subflow")
    if sub == "zd_detail":
        _zd_detail_loop(ctx)
        return
    if sub == "zd_single":
        _zd_single_loop(ctx)
        return
    if _ask_zhengchai(ctx, qno):
        return
    if sub == "custom_loop":
        _custom_fee_loop(ctx)
        return
    if sub == "contract_loop":
        _contract_loop(ctx)
        return

    st.markdown("是否需要添加**自定义费用**或**合同价**？")
    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("🚫 不需要", key=f"ck_custom_no_{qno}", use_container_width=True):
            ctx["custom_answered"] = True
            ctx.pop("_reask", None)
            ctx["qno"] = qno + 1
            st.rerun()
    with c2:
        if st.button("➕ 添加自定义费用", key=f"ck_custom_add_{qno}",
                     use_container_width=True):
            ctx["_subflow"] = "custom_loop"
            ctx["qno"] = qno + 1
            st.rerun()
    with c3:
        if st.button("📝 添加合同价/合同费率", key=f"ck_custom_ct_{qno}",
                     use_container_width=True):
            ctx["_subflow"] = "contract_loop"
            ctx["qno"] = qno + 1
            st.rerun()
    if cf or co:
        if st.button("✅ 确认，无需更多", key=f"ck_custom_done_{qno}",
                     use_container_width=True):
            ctx["custom_answered"] = True
            ctx.pop("_reask", None)
            ctx["qno"] = qno + 1
            st.rerun()


def _item_flow_footer(ctx: dict, stage: str):
    """分开导出按单项问流程的底部操作条（上一步/结束本次计算）。"""
    st.markdown("---")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("⬅ 上一步", key=f"ck_item_back_{stage}_{ctx.get('qno', 1)}",
                     use_container_width=True):
            _step_back(ctx, from_stage="custom")
            st.rerun()
    with c2:
        if st.button("🗑 结束本次计算", key=f"ck_item_cancel_{stage}_{ctx.get('qno', 1)}",
                     use_container_width=True):
            _cancel_task(ctx)


def _item_fees_finish(ctx: dict, labels: list[str]):
    """整体 custom_fees = 项目级既有费用 + 各单项条目（名称前加「单项标签·」）；
    置位完成标志（含 AI 模式门控键），结算与「同一概算表拆项」导出按整体口径。"""
    # 项目级既有费用在完成时按 _item 标记重算（流程期间自由文本新增的
    # 无 _item 条目也会并入，不会被重建覆盖）
    base = [dict(e) for e in (ctx.get("custom_fees") or []) if not e.get("_item")]
    ctx["item_fees_base"] = base
    overall = list(base)
    for lb in labels:
        for e in (ctx.get("item_custom_fees") or {}).get(lb) or []:
            nm = str(e.get("名称") or "")
            overall.append({"名称": f"{lb}·{nm}",
                            "金额(万元)": e.get("金额(万元)"), "_item": lb})
    ctx["custom_fees"] = overall
    ctx["zd_asked"] = True
    ctx["zd_exists"] = _zd_amount(overall) > 0
    # 逐项细化在按单项流程中已由每个单项展开块内的「📋 逐项细化」覆盖，
    # 置位 zd_refine_asked 防止 AI 问题库把项目级「是否逐项细化」再问一遍
    ctx["zd_refine_asked"] = True
    if ctx.get("zd_exists") and ctx.get("zd_placement") is None:
        ctx["zd_placement"] = "separate"  # 兜底（正常流程 placement 屏已问）
        ctx["zhengchai_separate"] = True
    ctx["custom_answered"] = True
    ctx["item_fees_done"] = True
    ctx.pop("_item_stage", None)
    ctx.pop("_reask", None)
    ctx["_ai_dirty"] = True  # AI 模式重规划：跳过项目级征地拆迁/自定义提问
    ctx["qno"] = ctx.get("qno", 1) + 1


def _ask_item_fees(ctx: dict):
    """分开导出（每单项一份概算表）模式：征地拆迁费 / 自定义费用按单项逐个询问。

    阶段推进 ctx["_item_stage"]：zd → zd_placement（任一单项有征地拆迁时）
    → custom → contract → 完成（item_fees_done=True）。合同价仍项目级。

    数据存放：
      ctx["item_custom_fees"] = {单项标签: [{名称, 金额(万元)}, ...]}
        （名称不带单项前缀；征地拆迁细项带「征地拆迁费·」前缀供引擎识别，
        每份单项概算表按该单项自己的条目计算）；
      ctx["item_fees_base"] = 流程开始前已有的项目级自定义费用（查询自带等，
        名称无 _item 标记的条目），每份单项概算表也会包含；
      ctx["custom_fees"] = 整体口径 = base + 各单项条目（名称加「单项标签·」，
        带 _item 标记），供整体结算与「同一概算表拆项」导出使用。
    """
    qno = ctx.get("qno", 1)
    zy_list = st.session_state.get("zhuan_ye_list") or []
    labels = [_item_gl(g, i) for i, g in enumerate(zy_list, 1)]
    item_cf = ctx.setdefault("item_custom_fees", {})
    if "item_fees_base" not in ctx:
        # 首次进入快照项目级既有费用；回退重问时按 _item 标记还原快照
        ctx["item_fees_base"] = [dict(e) for e in (ctx.get("custom_fees") or [])
                                 if not e.get("_item")]
    stage = ctx.get("_item_stage") or "zd"

    if ctx.get("_subflow") == "contract_loop":
        _contract_loop(ctx)
        return

    _stage_lbl = {"zd": "各单项征地拆迁费", "zd_placement": "征地拆迁费列示方式",
                  "custom": "各单项自定义费用", "contract": "合同价/合同费率"}[stage]
    st.markdown(f"### 📋 分开导出 — 各单项工程费用确认（{_stage_lbl}）")
    st.caption("「分开做概算表」模式：征地拆迁费、自定义费用每个单项各填一次，"
               "分别写入自己那份概算表；费种选择/费率/折扣/合同价仍按项目整体。")

    # ── 第 1 屏：各单项征地拆迁费 ──
    if stage == "zd":
        base_zd = _zd_amount(ctx.get("item_fees_base"))
        st.markdown("**🏗️ 各单项工程的征地拆迁费**"
                    "（土地征用、房屋拆迁、青苗补偿等，通常不随其他费用打折）")
        if base_zd > 0:
            st.info(f"查询中已识别**项目整体**征地拆迁费 {base_zd} 万元，"
                    "已计入每份概算表；如需按单项拆分，请在下方各单项填写。")
        st.caption("每个单项填一笔总额即可；需要逐项列示（如 房屋征收/青苗补偿）时"
                   "展开「📋 逐项细化」填明细，明细有内容时优先于总额。")
        import pandas as pd
        edits = {}
        for i, lb in enumerate(labels, 1):
            prev = item_cf.get(lb) or []
            prev_zd = [e for e in prev
                       if _ZD_RE.search(str(e.get("名称") or ""))]
            with st.expander(f"{lb} — 征地拆迁费", expanded=False):
                has = st.radio("是否有征地拆迁费？", ["没有", "有"], horizontal=True,
                               index=1 if prev_zd else 0,
                               key=f"ck_item_zd_has_{qno}_{i}")
                detail_rows: list = []
                total_val = 0.0
                if has == "有":
                    prev_total = next((float(e.get("金额(万元)") or 0)
                                       for e in prev_zd
                                       if str(e.get("名称")) == "征地拆迁费"), 0.0)
                    total_val = st.number_input(
                        "征地拆迁费总额（万元）", min_value=0.0,
                        value=float(prev_total), step=10.0,
                        key=f"ck_item_zd_amt_{qno}_{i}")
                    with st.expander("📋 逐项细化（可选）"):
                        rows = ([{"细项名称": str(e.get("名称"))[len("征地拆迁费·"):],
                                  "金额(万元)": float(e.get("金额(万元)") or 0)}
                                 for e in prev_zd
                                 if str(e.get("名称")) != "征地拆迁费"]
                                or [{"细项名称": "", "金额(万元)": 0.0}
                                    for _ in range(2)])
                        ed = st.data_editor(
                            pd.DataFrame(rows),
                            key=f"ck_item_zd_edit_{qno}_{i}",
                            num_rows="dynamic", use_container_width=True,
                            hide_index=True,
                            column_config={
                                "细项名称": st.column_config.TextColumn(
                                    "细项名称", width="large"),
                                "金额(万元)": st.column_config.NumberColumn(
                                    "金额(万元)", min_value=0.0, format="%.2f"),
                            })
                        detail_rows = ed.to_dict("records")
                edits[i] = {"lb": lb, "has": has == "有",
                            "amt": total_val, "rows": detail_rows}
        if st.button("✅ 确认各单项征地拆迁费", key=f"ck_item_zd_ok_{qno}",
                     use_container_width=True, type="primary"):
            ok = True
            for i, lb in enumerate(labels, 1):
                ed = edits[i]
                if not ed["has"]:
                    item_cf[lb] = [e for e in item_cf.get(lb) or []
                                   if not _ZD_RE.search(str(e.get("名称") or ""))]
                    continue
                rows = []
                for r in ed["rows"]:
                    nm = str(r.get("细项名称") or "").strip()
                    try:
                        amt = float(r.get("金额(万元)") or 0)
                    except (TypeError, ValueError):
                        amt = 0.0
                    if nm and amt > 0:
                        rows.append({"名称": f"征地拆迁费·{nm}",
                                     "金额(万元)": round(amt, 4)})
                if not rows and ed["amt"] > 0:
                    rows = [{"名称": "征地拆迁费",
                             "金额(万元)": round(float(ed["amt"]), 4)}]
                if not rows:
                    st.warning(f"「{lb}」选择了「有征地拆迁费」，请填写总额或明细。")
                    ok = False
                    continue
                item_cf[lb] = ([e for e in item_cf.get(lb) or []
                                if not _ZD_RE.search(str(e.get("名称") or ""))]
                               + rows)
            if ok:
                ctx["zd_asked"] = True
                ctx["zd_exists"] = any(
                    _ZD_RE.search(str(e.get("名称") or ""))
                    for lb in labels for e in (item_cf.get(lb) or []))
                ctx["_item_stage"] = ("zd_placement" if ctx["zd_exists"] else "custom")
                ctx["qno"] = qno + 1
                st.rerun()
        _item_flow_footer(ctx, stage)
        return

    # ── 第 2 屏：征地拆迁费列示方式（项目级，各单项统一口径） ──
    if stage == "zd_placement":
        st.markdown("**🏗️ 征地拆迁费如何列示？**（各单项的征地拆迁费按同一方式列示）")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("📑 单独列项（与工程费、工程建设其他费并列展示）",
                         key=f"ck_item_zd_pl_sep_{qno}", use_container_width=True):
                ctx["zd_placement"] = "separate"
                ctx["zhengchai_separate"] = True
                ctx["_item_stage"] = "custom"
                ctx["qno"] = qno + 1
                st.rerun()
        with c2:
            if st.button("📦 计入工程建设其他费（作为二类费组成部分）",
                         key=f"ck_item_zd_pl_in_{qno}", use_container_width=True):
                ctx["zd_placement"] = "inside"
                ctx["zhengchai_separate"] = False
                ctx["_item_stage"] = "custom"
                ctx["qno"] = qno + 1
                st.rerun()
        _item_flow_footer(ctx, stage)
        return

    # ── 第 3 屏：各单项自定义费用 ──
    if stage == "custom":
        st.markdown("**➕ 各单项工程的自定义费用**"
                    "（合同价以外的零散费用，如管线切改费）")
        st.caption("每个单项分别填写；没有的单项留空即可。"
                   "查询中已识别的项目级费用会自动计入每份概算表。")
        import pandas as pd
        edits = {}
        for i, lb in enumerate(labels, 1):
            prev = [e for e in item_cf.get(lb) or []
                    if not _ZD_RE.search(str(e.get("名称") or ""))]
            rows = ([{"费用名称": str(e.get("名称")),
                      "金额(万元)": float(e.get("金额(万元)") or 0)} for e in prev]
                    or [{"费用名称": "", "金额(万元)": 0.0} for _ in range(2)])
            with st.expander(f"{lb} — 自定义费用", expanded=bool(prev)):
                ed = st.data_editor(
                    pd.DataFrame(rows), key=f"ck_item_cf_{qno}_{i}",
                    num_rows="dynamic", use_container_width=True, hide_index=True,
                    column_config={
                        "费用名称": st.column_config.TextColumn(
                            "费用名称", width="large", help="如：管线切改费"),
                        "金额(万元)": st.column_config.NumberColumn(
                            "金额(万元)", min_value=0.0, format="%.2f"),
                    })
                edits[i] = (lb, ed.to_dict("records"))
        c1, c2, c3 = st.columns(3)
        with c1:
            if st.button("🚫 各单项都没有自定义费用", key=f"ck_item_cf_none_{qno}",
                         use_container_width=True):
                for _i, (lb, _rows) in edits.items():
                    item_cf[lb] = [e for e in item_cf.get(lb) or []
                                   if _ZD_RE.search(str(e.get("名称") or ""))]
                ctx["_item_stage"] = "contract"
                ctx["qno"] = qno + 1
                st.rerun()
        with c2:
            if st.button("✅ 确认，下一步", key=f"ck_item_cf_ok_{qno}",
                         use_container_width=True, type="primary"):
                ok = True
                for _i, (lb, recs) in edits.items():
                    rows = []
                    for r in recs:
                        nm = str(r.get("费用名称") or "").strip()
                        try:
                            amt = float(r.get("金额(万元)") or 0)
                        except (TypeError, ValueError):
                            amt = 0.0
                        if nm and amt > 0:
                            rows.append({"名称": nm, "金额(万元)": round(amt, 4)})
                        elif nm or amt > 0:
                            st.warning(f"「{lb}」有未填完整的自定义费用行"
                                       "（需 名称 + 金额>0），请补全或留空。")
                            ok = False
                            break
                    if not ok:
                        break
                    item_cf[lb] = ([e for e in item_cf.get(lb) or []
                                    if _ZD_RE.search(str(e.get("名称") or ""))]
                                   + rows)
                if ok:
                    ctx["_item_stage"] = "contract"
                    ctx["qno"] = qno + 1
                    st.rerun()
        with c3:
            if st.button("↩ 修改征地拆迁费", key=f"ck_item_cf_back_{qno}",
                         use_container_width=True):
                ctx["_item_stage"] = "zd"
                ctx["qno"] = qno + 1
                st.rerun()
        _item_flow_footer(ctx, stage)
        return

    # ── 第 4 屏：合同价/合同费率（项目级，各单项共用） ──
    co = ctx.get("contract_overrides") or {}
    if co:
        st.markdown("**已设置合同价/合同费率**：")
        for fn, ov in co.items():
            if ov.get("type") == "rate":
                desc = f"合同费率 {ov.get('rate')}% × {ov.get('base', '工程费')}"
            else:
                desc = f"一口价 {ov.get('amount_wan')} 万元"
            st.markdown(f"- {FEE_CATALOG.get(fn, {}).get('label', fn)}：{desc}")
    st.markdown("**📝 合同价/合同费率**（项目级，各单项共用）")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("📝 添加合同价/合同费率", key=f"ck_item_ct_{qno}",
                     use_container_width=True):
            ctx["_subflow"] = "contract_loop"
            ctx["qno"] = qno + 1
            st.rerun()
    with c2:
        if st.button("✅ 不需要，完成各单项费用确认", key=f"ck_item_ct_done_{qno}",
                     use_container_width=True, type="primary"):
            _item_fees_finish(ctx, labels)
            st.rerun()
    _item_flow_footer(ctx, stage)


def _custom_fee_loop(ctx: dict):
    qno = ctx.get("qno", 1)
    st.markdown("**➕ 添加自定义费用**（合同价外的零散费用，如管线切改费）")
    name = st.text_input("费用名称", key=f"ck_cf_name_{qno}", placeholder="如：管线切改费")
    amount = st.number_input("金额（万元）", min_value=0.0, value=0.0, step=10.0,
                             key=f"ck_cf_amt_{qno}")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("✅ 添加", key=f"ck_cf_ok_{qno}", use_container_width=True,
                     disabled=not name or amount <= 0):
            ctx.setdefault("custom_fees", []).append(
                {"名称": name, "金额(万元)": round(float(amount), 4)})
            ctx["_subflow"] = None
            ctx["qno"] = qno + 1
            st.rerun()
    with c2:
        if st.button("↩ 返回", key=f"ck_cf_back_{qno}", use_container_width=True):
            ctx["_subflow"] = None
            ctx["qno"] = qno + 1
            st.rerun()


def _contract_loop(ctx: dict):
    qno = ctx.get("qno", 1)
    st.markdown("**📝 添加合同价/合同费率**")
    fee_options = [(f, FEE_CATALOG.get(f, {}).get("label", f)) for f in FEE_CATALOG]
    fee_name = st.selectbox("费种", fee_options, format_func=lambda x: x[1],
                            key=f"ck_ct_fee_{qno}")
    ctype = st.radio("覆盖类型", ["合同费率", "一口价（固定金额）"],
                     key=f"ck_ct_type_{qno}", horizontal=True)
    ok = False
    ov = None
    if ctype == "合同费率":
        base_label = st.selectbox("计费基数",
                                  ["工程费", "建安费", "项目总投资", "自定义金额", "选定费种"],
                                  key=f"ck_ct_base_{qno}")
        rate = st.number_input("费率（%）", min_value=0.0, value=1.0, step=0.1,
                               key=f"ck_ct_rate_{qno}")
        base_custom = None
        base_fees = None
        if base_label == "自定义金额":
            base_custom = st.number_input("基数金额（万元）", min_value=0.0, value=0.0,
                                          step=100.0, key=f"ck_ct_baseamt_{qno}")
        elif base_label == "选定费种":
            base_fees = st.multiselect("基数费种（可多选）",
                                       [f for f, _ in fee_options],
                                       key=f"ck_ct_basefees_{qno}")
        if rate > 0:
            ov = {"type": "rate", "rate": float(rate), "base": base_label}
            if base_custom is not None:
                ov["base_custom"] = float(base_custom)
            if base_fees:
                ov["base_fees"] = list(base_fees)
            ok = True
    else:
        amount_wan = st.number_input("一口价金额（万元）", min_value=0.0, value=0.0,
                                     step=10.0, key=f"ck_ct_amt_{qno}")
        if amount_wan > 0:
            ov = {"type": "price", "amount_wan": float(amount_wan)}
            ok = True
    c1, c2 = st.columns(2)
    with c1:
        if st.button("✅ 添加", key=f"ck_ct_ok_{qno}", use_container_width=True,
                     disabled=not ok):
            fn = fee_name[0]
            ctx.setdefault("contract_overrides", {})[fn] = ov
            if fn not in ctx["fees"]:
                ctx["fees"].append(fn)
            ctx["_subflow"] = None
            ctx["qno"] = qno + 1
            st.rerun()
    with c2:
        if st.button("↩ 返回", key=f"ck_ct_back_{qno}", use_container_width=True):
            ctx["_subflow"] = None
            ctx["qno"] = qno + 1
            st.rerun()


def _params_advance(ctx: dict, pidx: int, n_cards: int):
    """推进参数卡（不 bump qno，由调用方负责）。"""
    if pidx + 1 < n_cards:
        ctx["param_idx"] = pidx + 1
    else:
        ctx["params_done"] = True
        ctx.pop("_reask", None)


def _clear_card_overrides(ctx: dict, card: dict):
    """「本卡用默认值」：清除该卡已收集的覆盖值。"""
    fn = card["fee"]
    kind = card["kind"]
    if kind == "coef":
        ctx.get("coef_overrides", {}).pop(fn, None)
        ctx.pop("sheji_extra_separate", None)  # 附加项默认=不单独计算
    elif kind == "rate":
        ctx.get("rate_overrides", {}).pop(fn, None)
    elif kind == "service":
        ctx.get("service_selections", {}).pop(fn, None)
    elif kind == "party":
        ctx["jiaoyi_party"] = None
    elif kind == "shencha":
        ctx.get("rate_overrides", {}).pop("施工图审查费", None)
        ctx.get("spec_overrides", {}).pop("施工图审查费", None)
    elif kind == "yubei_include":
        ctx["yubei_considered"] = True  # 默认：需要考虑
    elif kind == "yubei_rate":
        ctx["yubei_rate"] = 5.0
    elif kind == "shuibao_comp":
        ctx.get("spec_overrides", {}).pop("水土保持补偿费", None)
    elif kind == "spec":
        ctx.get("spec_params", {}).pop(fn, None)


def _ask_params(ctx: dict):
    qno = ctx.get("qno", 1)
    cards = param_cards_for(ctx)
    if not cards:
        ctx["params_done"] = True
        st.rerun()
    pidx = min(ctx.get("param_idx", 0), len(cards) - 1)
    if ctx.get("_ai_fallback") and ctx.pop("_params_seek", False):
        # AI 模式降级而来：仅首次顺延定位到第一张未填卡；之后逐张显示
        # （已填卡保留原值、确认即过），编号连续不跳号
        while pidx < len(cards) and question_done(ctx, _card_qkey(cards[pidx])):
            pidx += 1
        if pidx >= len(cards):
            ctx["params_done"] = True
            st.rerun()
        ctx["param_idx"] = pidx
    card = cards[pidx]
    label = FEE_CATALOG.get(card["fee"], {}).get("label", card["fee"])
    if question_done(ctx, _card_qkey(card)):
        st.markdown(f"### ⚙️ 参数设置（{pidx + 1}/{len(cards)}）· 已设置")
    else:
        st.markdown(f"### ⚙️ 参数设置（{pidx + 1}/{len(cards)}）")
    st.caption(f"当前费种：{label}")

    kind = card["kind"]
    if kind == "coef":
        _ask_coef_card(ctx, card, qno)
    elif kind == "rate":
        _ask_linear_rate_card(ctx, card, qno)
    elif kind == "shencha":
        _ask_shencha_card(ctx, card, qno)
    elif kind == "service":
        _ask_service_card(ctx, card, qno)
    elif kind == "party":
        _ask_party_card(ctx, card, qno)
    elif kind == "yubei_include":
        _ask_yubei_include_card(ctx, card, qno)
    elif kind == "yubei_rate":
        _ask_yubei_rate_card(ctx, card, qno)
    elif kind == "shuibao_comp":
        _ask_shuibao_comp_card(ctx, card, qno)
    elif kind == "spec":
        _ask_spec_card(ctx, card, qno)

    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("✅ 确认并继续", key=f"ck_params_ok_{qno}", use_container_width=True):
            _commit_card_widgets(ctx, card, qno)
            ctx["qno"] = qno + 1
            _params_advance(ctx, pidx, len(param_cards_for(ctx)))
            st.rerun()
    with col2:
        if st.button("⏭ 本卡用默认值", key=f"ck_params_skip_{qno}", use_container_width=True):
            _clear_card_overrides(ctx, card)
            ctx["qno"] = qno + 1
            _params_advance(ctx, pidx, len(param_cards_for(ctx)))
            st.rerun()
    with col3:
        if st.button("⚡ 剩余全部用默认值", key=f"ck_params_alldef_{qno}",
                     use_container_width=True):
            ctx["params_done"] = True
            ctx.pop("_reask", None)
            ctx["qno"] = qno + 1
            st.rerun()
    if pidx > 0:
        if st.button("⬅ 上一项参数（返回修改上一个费种的设置）",
                     key=f"ck_params_prev_{qno}", use_container_width=True):
            # 回退目标 = 上一个费种：跳过同费种相邻卡（如环评 service→coef），
            # 与按钮文案「返回修改上一个费种的设置」一致；已填卡保留原值显示
            target = pidx - 1
            while target > 0 and cards[target]["fee"] == card["fee"]:
                target -= 1
            ctx["param_idx"] = target
            ctx["qno"] = qno + 1
            st.rerun()


def _ask_coef_card(ctx: dict, card: dict, qno: int):
    """系数卡：展示 FEE_CATALOG 中该费种的系数配置（按钮 + 自定义输入）。"""
    fee_name = card["fee"]
    query_text = " ".join(ctx.get("query_history") or [ctx.get("query", "")])
    config = _get_coef_config_simple(fee_name, query_text) or {}
    coefs = config.get("coefs", [])
    fee_overrides = ctx.setdefault("coef_overrides", {}).setdefault(fee_name, {})
    for ci, coef_def in enumerate(coefs):
        param_name = coef_def["param_name"]
        key = coef_def["key"]
        current_val = fee_overrides.get(param_name, coef_def.get("current", 1.0))
        options = coef_def.get("options", [])
        st.markdown(f"**{key}**")
        if coef_def.get("description"):
            st.caption(coef_def["description"])
        if options:
            cols = st.columns(min(len(options), 4))
            for oi, (opt_label, opt_val) in enumerate(options):
                try:
                    is_active = abs(float(opt_val) - float(current_val)) < 0.005
                except (TypeError, ValueError):
                    is_active = str(opt_val) == str(current_val)
                with cols[oi % 4]:
                    if st.button(
                        f"{opt_label}（{opt_val}）",
                        key=f"ck_coef_{fee_name}_{param_name}_{oi}_{qno}",
                        use_container_width=True,
                        type="primary" if is_active else "secondary",
                    ):
                        try:
                            fee_overrides[param_name] = float(opt_val)
                        except (TypeError, ValueError):
                            fee_overrides[param_name] = opt_val
                        ctx["qno"] = qno + 1
                        st.rerun()
        try:
            default_custom = float(current_val)
        except (TypeError, ValueError):
            default_custom = 1.0
        custom_val = st.number_input(
            f"✏️ 自定义 {key}", min_value=0.10, max_value=5.00,
            value=default_custom, step=0.05, format="%.2f",
            key=f"ck_coef_{fee_name}_{param_name}_custom_{qno}")
        if abs(custom_val - default_custom) > 0.001:
            fee_overrides[param_name] = custom_val

    # 设计费附加项单独计算（所有省份提问；天津模板两行按勾选计取/置零）
    # 渲染不写 ctx，确认时由 _commit_coef_extras 提交（防 AI 模式提问一闪而过）
    if fee_name == "工程设计费":
        st.markdown("---")
        st.markdown("**设计费附加项**（计价格[2002]10号 1.0.16，默认不单独计算）")
        st.caption("施工图预算编制费=设计费×10%；竣工图编制费=设计费×8%。"
                   "不勾选则不计取。")
        st.checkbox("单独计算施工图预算编制费（设计费×10%）",
                    value=bool((ctx.get("sheji_extra_separate") or {})
                               .get("施工图预算编制费")),
                    key=f"ck_sjext_ys_{qno}")
        st.checkbox("单独计算竣工图编制费（设计费×8%）",
                    value=bool((ctx.get("sheji_extra_separate") or {})
                               .get("竣工图编制费")),
                    key=f"ck_sjext_jg_{qno}")


def _ask_linear_rate_card(ctx: dict, card: dict, qno: int):
    """费率卡：档位按钮「X% → Y万元」（数据源 = calc 函数费率明细）+ 自定义费率。"""
    fee_name = card["fee"]
    am = ctx.get("amounts", {})
    opts = linear_rate_options(
        fee_name, am.get("jianan") or 0.0, am.get("shebei") or 0.0,
        ctx.get("project_type") or "通用")
    if not opts or not opts.get("rate_options"):
        st.info("费率档位数据暂不可用，请直接输入自定义费率。")
        rate_opts = []
        default_rate = None
    else:
        rate_opts = opts["rate_options"]
        default_rate = opts.get("default_rate")
    cur = ctx.get("rate_overrides", {}).get(fee_name)

    if rate_opts:
        st.markdown("**选择费率档位**（档位与金额联动，金额随建安费变化）：")
        cols = st.columns(min(len(rate_opts), 3))
        for oi, o in enumerate(rate_opts):
            is_active = (cur == o["rate"]) or (cur is None and o["rate"] == default_rate)
            with cols[oi % 3]:
                if st.button(
                    o["label"],
                    key=f"ck_rate_{fee_name}_{oi}_{qno}",
                    use_container_width=True,
                    type="primary" if is_active else "secondary",
                ):
                    ctx.setdefault("rate_overrides", {})[fee_name] = o["rate"]
                    ctx["qno"] = qno + 1
                    _params_advance(ctx, ctx.get("param_idx", 0),
                                    len(param_cards_for(ctx)))
                    st.rerun()
    c1, c2 = st.columns(2)
    with c1:
        custom_rate = st.number_input(
            "✏️ 自定义费率（%）", min_value=0.0, value=0.0, step=0.1, format="%.1f",
            key=f"ck_rate_{fee_name}_cust_{qno}")
    with c2:
        st.write("")
        if st.button("使用自定义费率", key=f"ck_rate_{fee_name}_custok_{qno}",
                     use_container_width=True, disabled=custom_rate <= 0):
            ctx.setdefault("rate_overrides", {})[fee_name] = f"{custom_rate}%"
            ctx["qno"] = qno + 1
            _params_advance(ctx, ctx.get("param_idx", 0), len(param_cards_for(ctx)))
            st.rerun()


def _ask_shencha_card(ctx: dict, card: dict, qno: int):
    """施工图审查费卡：河北统一 6.5% / 津价管46号按类型+规模档位。"""
    sopts = _build_shencha_rate_options(ctx.get("region"))
    cur = ctx.get("rate_overrides", {}).get("施工图审查费")
    labels = [o["label"] for o in sopts["rate_options"]]
    keys = [o["key"] for o in sopts["rate_options"]]
    idx = keys.index(cur) if cur in keys else keys.index(sopts["default_key"])
    choice = st.radio("选择适用档位", labels, index=idx,
                      key=f"ck_shencha_{qno}", label_visibility="collapsed")
    chosen = sopts["rate_options"][labels.index(choice)]
    # 不在渲染时写 ctx：渲染即写默认值会让 question_done 提前成立（AI 模式
    # 该提问一闪而过，用户看不到）。档位/面积仅在用户点「✅ 确认」时由
    # _commit_shencha_card 写入。
    if chosen.get("billing") == "area":
        st.number_input("建筑面积（m²）", min_value=0.0, value=0.0, step=1000.0,
                        key=f"ck_shencha_area_{qno}")


def _commit_shencha_card(ctx: dict, card: dict, qno: int):
    """施工图审查卡确认：把当前档位/面积写入 ctx（唯一写入点）。"""
    sopts = _build_shencha_rate_options(ctx.get("region"))
    labels = [o["label"] for o in sopts["rate_options"]]
    keys = [o["key"] for o in sopts["rate_options"]]
    choice = st.session_state.get(f"ck_shencha_{qno}")
    if choice not in labels:
        choice = labels[keys.index(sopts["default_key"])]
    chosen = sopts["rate_options"][labels.index(choice)]
    ctx.setdefault("rate_overrides", {})["施工图审查费"] = chosen["key"]
    if chosen.get("billing") == "area":
        area = st.session_state.get(f"ck_shencha_area_{qno}") or 0.0
        if float(area) > 0:
            ctx.setdefault("spec_overrides", {}).setdefault(
                "施工图审查费", {})["area_m2"] = float(area)
        else:
            ctx.setdefault("spec_overrides", {}).pop("施工图审查费", None)


def _ask_service_card(ctx: dict, card: dict, qno: int):
    """服务类型多选卡（环评/可研/造价咨询）。"""
    fee_name = card["fee"]
    conf = FEE_CATALOG.get(fee_name, {}).get("conf", {})
    if fee_name == "造价咨询费":
        hebei = is_hebei_region(ctx.get("region"))
        svc_list = conf.get("services_hebei") if hebei else conf.get("services_tianjin")
        default = conf.get("default_hebei") if hebei else conf.get("default_tianjin")
    elif fee_name == "环境影响咨询费":
        svc_list = [{"name": s, "label": s} for s in _HP_SERVICES]
        default = ["编制报告表", "评估报告表"]
    else:  # 可行性研究费
        svc_list = [{"name": s, "label": s} for s in _KY_SERVICES]
        default = ["编制可研报告"]
    names = [s["name"] for s in svc_list]
    labels = [s["label"] for s in svc_list]
    cur = ctx.get("service_selections", {}).get(fee_name)
    picked = st.multiselect(
        f"选择需要计算的服务类型（{FEE_CATALOG[fee_name]['label']}）",
        names, default=list(cur) if cur else list(default),
        key=f"ck_svc_{fee_name}_{qno}",
        format_func=lambda n: labels[names.index(n)])
    # 不在渲染时写 ctx：渲染即写默认值会让 question_done 提前成立（AI 模式
    # 该提问一闪而过，用户看不到）。选择值仅在用户点「✅ 确认」时由
    # _commit_service_card 写入。
    if not picked:
        if fee_name == "造价咨询费":
            st.warning(f"未选择任何服务类型，确认后将按默认服务"
                       f"（{'、'.join(default)}）计算。")
        else:
            st.warning("未选择任何服务类型，该费种将计为 0 元。")


def _commit_service_card(ctx: dict, card: dict, qno: int):
    """服务多选卡确认：把当前勾选写入 ctx（唯一写入点）。"""
    fee_name = card["fee"]
    picked = st.session_state.get(f"ck_svc_{fee_name}_{qno}")
    if picked is None:
        return
    ctx.setdefault("service_selections", {})[fee_name] = list(picked)


def _ask_party_card(ctx: dict, card: dict, qno: int):
    """交易服务费计费方卡（点选即推进）。"""
    st.markdown("**交易服务费计费方**：")
    cur = ctx.get("jiaoyi_party")
    cols = st.columns(3)
    for ci, (plabel, pval) in enumerate(
            [("招标方", "招标方"), ("中标方", "中标方"), ("双方合计", None)]):
        with cols[ci]:
            is_active = (cur == pval) or (pval is None and cur is None)
            if st.button(plabel, key=f"ck_party_{ci}_{qno}", use_container_width=True,
                         type="primary" if is_active else "secondary"):
                ctx["jiaoyi_party"] = pval
                ctx["qno"] = qno + 1
                _params_advance(ctx, ctx.get("param_idx", 0), len(param_cards_for(ctx)))
                st.rerun()


def _ask_yubei_include_card(ctx: dict, card: dict, qno: int):
    """预备费是否考虑（默认要考虑；不考虑则跳过费率卡，结算时预备费=0）。"""
    st.markdown("**是否要考虑预备费？**\n\n"
                "预备费 =（第一部分工程费＋工程建设其他费）× 预备费率，"
                "默认计入工程总投资。")
    yc = ctx.get("yubei_considered")
    if yc is True:
        st.info("已确认：需要考虑预备费（可修改）。")
    elif yc is False:
        st.info("已确认：不考虑预备费（可修改）。")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("✅ 需要考虑预备费", key=f"ck_ybi_yes_{qno}",
                     use_container_width=True, type="primary"):
            ctx["yubei_considered"] = True
            ctx["qno"] = qno + 1
            _params_advance(ctx, ctx.get("param_idx", 0), len(param_cards_for(ctx)))
            st.rerun()
    with c2:
        if st.button("🚫 不考虑预备费", key=f"ck_ybi_no_{qno}",
                     use_container_width=True):
            ctx["yubei_considered"] = False
            ctx["qno"] = qno + 1
            _params_advance(ctx, ctx.get("param_idx", 0), len(param_cards_for(ctx)))
            st.rerun()


def _ask_yubei_rate_card(ctx: dict, card: dict, qno: int):
    """预备费率卡（默认 5%）。"""
    cur = ctx.get("yubei_rate", 5.0)
    v = st.number_input("预备费率（%）", min_value=0.0, value=float(cur), step=0.5,
                        format="%.1f", key=f"ck_yubei_{qno}")
    if st.button("✅ 使用该费率", key=f"ck_yubei_ok_{qno}", use_container_width=True):
        ctx["yubei_rate"] = float(v)
        ctx["qno"] = qno + 1
        _params_advance(ctx, ctx.get("param_idx", 0), len(param_cards_for(ctx)))
        st.rerun()


def _ask_shuibao_comp_card(ctx: dict, card: dict, qno: int):
    """水土保持补偿费物理参数卡。"""
    sbp = dict(ctx.get("spec_overrides", {}).get("水土保持补偿费", {}) or {})
    sbp.setdefault("calc_type", "general")
    sbp.setdefault("land_input", 0.0)
    sbp.setdefault("land_unit", "m²")
    ctype = st.radio("计征类型", [t[0] for t in _SBUBAO_TYPES], horizontal=True,
                     index=[t[1] for t in _SBUBAO_TYPES].index(sbp.get("calc_type", "general")),
                     key=f"ck_sb_type_{qno}")
    sbp["calc_type"] = dict((t[0], t[1]) for t in _SBUBAO_TYPES)[ctype]
    c1, c2 = st.columns(2)
    with c1:
        land = st.number_input("占地面积", min_value=0.0,
                               value=float(sbp.get("land_input", 0.0) or 0.0),
                               step=1.0, key=f"ck_sb_land_{qno}")
    with c2:
        unit = st.selectbox("单位", ["m²", "亩", "公顷"],
                            index=["m²", "亩", "公顷"].index(sbp.get("land_unit", "m²")),
                            key=f"ck_sb_unit_{qno}")
    sbp["land_input"] = float(land)
    sbp["land_unit"] = unit
    c3, c4 = st.columns(2)
    with c3:
        wells = st.number_input("井数（口）", min_value=0,
                                value=int(sbp.get("well_cnt", 0) or 0),
                                step=1, key=f"ck_sb_wells_{qno}")
        sbp["well_cnt"] = int(wells)
    with c4:
        addw = st.number_input("增加井数（口）", min_value=0,
                               value=int(sbp.get("add_wells", 0) or 0),
                               step=1, key=f"ck_sb_addwells_{qno}")
        sbp["add_wells"] = int(addw)
    if sbp.get("calc_type") in ("mining_oil_gas", "mining_other"):
        c5, c6, c7 = st.columns(3)
        with c5:
            st.number_input("开采量（m³）", min_value=0.0,
                            value=float(sbp.get("extract_vol", 0.0) or 0.0),
                            step=1000.0, key=f"ck_sb_ev_{qno}")
        with c6:
            st.number_input("材料量（m³）", min_value=0.0,
                            value=float(sbp.get("material_vol", 0.0) or 0.0),
                            step=1000.0, key=f"ck_sb_mv_{qno}")
        with c7:
            st.number_input("弃方量（m³）", min_value=0.0,
                            value=float(sbp.get("waste_vol", 0.0) or 0.0),
                            step=1000.0, key=f"ck_sb_wv_{qno}")
    # 不在渲染时写 ctx：渲染即写默认值会让 question_done 提前成立（AI 模式
    # 该提问一闪而过，用户看不到）。参数仅在用户点「✅ 确认」时由
    # _commit_shuibao_comp_card 写入。


def _commit_shuibao_comp_card(ctx: dict, card: dict, qno: int):
    """水土保持补偿费卡确认：从控件值组装参数写入 ctx（唯一写入点）。"""
    type_label = st.session_state.get(f"ck_sb_type_{qno}")
    type_map = dict((t[0], t[1]) for t in _SBUBAO_TYPES)
    calc_type = type_map.get(type_label, "general")
    sbp = {
        "calc_type": calc_type,
        "land_input": float(st.session_state.get(f"ck_sb_land_{qno}") or 0.0),
        "land_unit": st.session_state.get(f"ck_sb_unit_{qno}") or "m²",
        "well_cnt": int(st.session_state.get(f"ck_sb_wells_{qno}") or 0),
        "add_wells": int(st.session_state.get(f"ck_sb_addwells_{qno}") or 0),
    }
    if calc_type in ("mining_oil_gas", "mining_other"):
        sbp["extract_vol"] = float(st.session_state.get(f"ck_sb_ev_{qno}") or 0.0)
        sbp["material_vol"] = float(st.session_state.get(f"ck_sb_mv_{qno}") or 0.0)
        sbp["waste_vol"] = float(st.session_state.get(f"ck_sb_wv_{qno}") or 0.0)
    ctx.setdefault("spec_overrides", {})["水土保持补偿费"] = sbp


def _ask_spec_card(ctx: dict, card: dict, qno: int):
    """通用 spec 参数卡（FEE_PARAM_SPECS 驱动，如水土保持费）。"""
    fn = card["fee"]
    spec = FEE_PARAM_SPECS.get(fn)
    params = dict(ctx.get("spec_params", {}).get(fn, {}))
    for ps in spec["params"]:
        if ps.get("type") == "select":
            opts = ps.get("options", [])
            labels = [o[0] for o in opts]
            vals = [o[1] for o in opts]
            cur = params.get(ps["key"], ps.get("default"))
            idx = vals.index(cur) if cur in vals else 0
            v = st.selectbox(ps["label"], labels, index=idx,
                             key=f"ck_spec_{fn}_{ps['key']}_{qno}")
            params[ps["key"]] = vals[labels.index(v)]
        else:
            default_v = params.get(ps["key"], ps.get("default", 0.0))
            step = 0.01 if ps.get("unit") == "亿元" else 100.0
            st.number_input(f"{ps['label']}（{ps.get('unit', '')}）",
                            min_value=0.0, value=float(default_v or 0.0),
                            step=step, key=f"ck_spec_{fn}_{ps['key']}_{qno}")
    # 不在渲染时写 ctx：渲染即写默认值会让 question_done 提前成立（AI 模式
    # 该提问一闪而过，用户看不到）。参数仅在用户点「✅ 确认」时由
    # _commit_spec_card 写入。


def _commit_spec_card(ctx: dict, card: dict, qno: int):
    """通用 spec 参数卡确认：从控件值组装参数写入 ctx（唯一写入点）。"""
    fn = card["fee"]
    spec = FEE_PARAM_SPECS.get(fn)
    if spec is None:
        return
    params = dict(ctx.get("spec_params", {}).get(fn, {}))
    for ps in spec["params"]:
        raw = st.session_state.get(f"ck_spec_{fn}_{ps['key']}_{qno}")
        if raw is None:
            continue
        if ps.get("type") == "select":
            opts = ps.get("options", [])
            labels = [o[0] for o in opts]
            vals = [o[1] for o in opts]
            params[ps["key"]] = vals[labels.index(raw)] if raw in labels else ps.get("default")
        else:
            params[ps["key"]] = float(raw)
    ctx.setdefault("spec_params", {})[fn] = params


def _commit_coef_extras(ctx: dict, card: dict, qno: int):
    """系数卡确认时提交「设计费附加项单独计算」勾选（所有省份渲染）。"""
    if card["fee"] != "工程设计费":
        return
    _cb1 = st.session_state.get(f"ck_sjext_ys_{qno}")
    _cb2 = st.session_state.get(f"ck_sjext_jg_{qno}")
    if _cb1 is None and _cb2 is None:
        return  # 未渲染该卡（如旧 qno 残留）→ 不动
    if _cb1 or _cb2:
        ctx["sheji_extra_separate"] = {
            "施工图预算编制费": bool(_cb1),
            "竣工图编制费": bool(_cb2),
        }
    else:
        ctx.pop("sheji_extra_separate", None)  # 都不勾 = 不计算（默认）


def _commit_card_widgets(ctx: dict, card: dict, qno: int):
    """确认按钮统一提交：把卡内控件当前值写入 ctx（各卡唯一写入点）。

    服务/施工图审查/水保补偿/spec/coef 卡渲染时不再写 ctx（防渲染即写
    默认值导致 AI 模式提问一闪而过），只在用户点确认时在这里提交。
    """
    kind = card["kind"]
    if kind == "service":
        _commit_service_card(ctx, card, qno)
    elif kind == "shencha":
        _commit_shencha_card(ctx, card, qno)
    elif kind == "shuibao_comp":
        _commit_shuibao_comp_card(ctx, card, qno)
    elif kind == "spec":
        _commit_spec_card(ctx, card, qno)
    elif kind == "coef":
        _commit_coef_extras(ctx, card, qno)


def _ask_discount(ctx: dict):
    qno = ctx.get("qno", 1)
    if ctx.get("_subflow") == "discount_loop":
        _discount_loop(ctx)
        return
    discs = ctx.get("discounts") or {}
    if discs:
        st.markdown("**已设折扣**：")
        for fn, d in discs.items():
            st.markdown(f"- {FEE_CATALOG.get(fn, {}).get('label', fn)}：打 {d:g} 折")
    st.markdown("是否有费用需要**打折**？")
    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("🚫 不需要", key=f"ck_disc_no_{qno}", use_container_width=True):
            ctx["discount_answered"] = True
            ctx.pop("_reask", None)
            ctx["qno"] = qno + 1
            st.rerun()
    with c2:
        if st.button("💸 指定费用打折", key=f"ck_disc_add_{qno}", use_container_width=True):
            ctx["_subflow"] = "discount_loop"
            ctx["qno"] = qno + 1
            st.rerun()
    with c3:
        if st.button("📊 查看折扣情景对比", key=f"ck_disc_scen_{qno}",
                     use_container_width=True):
            ctx["discount_scenario"] = True
            ctx["discount_answered"] = True
            ctx.pop("_reask", None)
            ctx["qno"] = qno + 1
            st.rerun()


def _discount_loop(ctx: dict):
    qno = ctx.get("qno", 1)
    st.markdown("**💸 指定费用打折**")
    discs = ctx.get("discounts") or {}
    if discs:
        st.markdown("**已设折扣**：")
        for fn, d in discs.items():
            st.markdown(f"- {FEE_CATALOG.get(fn, {}).get('label', fn)}：打 {d:g} 折")
    fee_names = [f for f in ctx.get("fees") or []
                 if FEE_CATALOG.get(f, {}).get("tier") is not None and f != "预备费"]
    fee_options = [(f, FEE_CATALOG.get(f, {}).get("label", f)) for f in fee_names]
    if not fee_options:
        st.info("当前费种无适用折扣项。")
        ctx["_subflow"] = None
        st.rerun()

    def _lbl(x):
        lab = x[1]
        if x[0] in discs:
            lab += f"（已设 {discs[x[0]]:g} 折）"
        return lab

    # 默认选中第一个尚未设折扣的费种，方便连续添加
    default_idx = next((i for i, (f, _) in enumerate(fee_options)
                        if f not in discs), 0)
    fn = st.selectbox("费种（可连续添加多个）", fee_options, format_func=_lbl,
                      index=default_idx, key=f"ck_dl_fee_{qno}")
    coef = st.number_input("折扣系数（0.8 = 打八折）", min_value=0.01, max_value=2.0,
                           value=0.8, step=0.05, format="%.2f", key=f"ck_dl_coef_{qno}")
    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("✅ 添加/更新", key=f"ck_dl_ok_{qno}", use_container_width=True):
            ctx.setdefault("discounts", {})[fn[0]] = float(coef)
            ctx["qno"] = qno + 1
            st.rerun()  # 留在本流程，可继续添加下一个费种
    with c2:
        if st.button("🏁 完成", key=f"ck_dl_done_{qno}", use_container_width=True):
            ctx["discount_answered"] = True
            ctx.pop("_reask", None)
            ctx["_subflow"] = None
            ctx["qno"] = qno + 1
            st.rerun()
    with c3:
        if st.button("↩ 返回", key=f"ck_dl_back_{qno}", use_container_width=True):
            ctx["_subflow"] = None
            ctx["qno"] = qno + 1
            st.rerun()


# ══════════════ AI 对话式提问（概算助手）══════════════
# AI 主导的信息确认：run_planner 动态决定要问什么，控件作答写入 ctx（与固定清单同源），
# 异常或手动切换时降级传统 5 阶段清单（数据无损）。
_AI_SUBFLOWS = ("zd_detail", "zd_single", "custom_loop", "contract_loop", "discount_loop")

# 语义问题键的规范选项（AI 只定问题措辞，动作由固定选项保证可靠）
_AI_CANON_OPTIONS: dict[str, list[tuple[str, object]]] = {
    "zd_exists": [("✅ 有", True), ("🚫 没有", False)],
    "zd_placement": [("📑 单独列项（与工程费、工程建设其他费并列）", "separate"),
                     ("📦 计入工程建设其他费（作为二类费组成部分）", "inside")],
    "zd_refine": [("📋 需要细化（逐项填写）", "detail"),
                  ("💰 不需要，填一笔总额", "single")],
    "custom_need": [("🚫 不需要", "none"), ("➕ 添加自定义费用", "custom"),
                    ("📝 有合同价/合同费率", "contract")],
    "discount_need": [("🚫 不需要", "none"), ("💸 指定费用打折", "discount"),
                      ("📊 查看折扣情景对比", "scenario")],
    "card:party": [("招标方", "招标方"), ("中标方", "中标方"), ("双方合计", None)],
    "card:yubei_include": [("✅ 需要考虑预备费", True), ("🚫 不考虑预备费", False)],
}

_CARD_RENDERERS = {
    "coef": _ask_coef_card, "rate": _ask_linear_rate_card, "shencha": _ask_shencha_card,
    "service": _ask_service_card, "shuibao_comp": _ask_shuibao_comp_card,
    "spec": _ask_spec_card,
}

_SUBFLOW_RENDERERS = {
    "zd_detail": _zd_detail_loop, "zd_single": _zd_single_loop,
    "custom_loop": _custom_fee_loop, "contract_loop": _contract_loop,
    "discount_loop": _discount_loop,
}


def _ai_parse_card_key(key: str) -> dict:
    parts = key.split(":", 2)
    kind = parts[1]
    if kind in ("coef", "rate", "service", "spec"):
        return {"kind": kind, "fee": parts[2]}
    fee_map = {"party": "交易服务费", "yubei_rate": "预备费",
               "yubei_include": "预备费", "shuibao_comp": "水土保持补偿费",
               "shencha": "施工图审查费"}
    return {"kind": kind, "fee": fee_map[kind]}


def _card_qkey(card: dict) -> str:
    """参数卡 → 问题键（与 ai_planner._card_key 同格式，供 question_done 判定）。"""
    kind = card["kind"]
    if kind in ("coef", "rate", "service", "spec"):
        return f"card:{kind}:{card['fee']}"
    return f"card:{kind}"


def _ai_record_answer(ctx: dict, text: str, answer: str):
    rounds = ctx.setdefault("_ai_rounds", [])
    if not rounds:
        rounds.append({"think": "", "qa": []})
    rounds[-1]["qa"].append({"text": text, "answer": answer})


def _ai_remaining(plan: dict, ctx: dict) -> list[dict]:
    answered = ctx.get("_ai_answered") or []
    # force：回看已填参数卡（⬅ 返回上一参数设置构造）→ 绕过 question_done
    # 的「已填不重问」过滤，但仍受 answered 控制（确认后正常消失）
    return [q for q in plan.get("questions") or []
            if q.get("key") not in answered
            and (q.get("force") or not question_done(ctx, q.get("key")))]


_BARE_YES = ("有", "是的", "是", "对", "对的", "涉及", "要", "需要", "有的",
             "有啊", "有呀", "有吧", "嗯")
_BARE_NO = ("没有", "无", "不涉及", "不用", "不需要", "否", "不是", "不对",
            "没有了", "不需要了")


def _ai_map_bare_answer(ctx: dict, text: str):
    """当前可见问题唯一且为二选一键时，把「有/没有/单独/计入…」裸答映射为按钮回答。

    用户常打字回答而不是点按钮；裸答若不登记，同一问题会被下一轮重复提问
    （或降级传统清单后从头重问）。只在唯一可见问题、短文本时映射，避免误伤。
    """
    if not ctx.get("ai_guided") or ctx.get("_ai_fallback"):
        return None
    visible = _ai_remaining(ctx.get("_ai_plan") or {}, ctx)
    if len(visible) != 1:
        return None
    q = visible[0]
    key = q.get("key") or ""
    t = re.sub(r"[。.！!？?，,\s]", "", (text or "").strip())
    if not t or len(t) > 8:
        return None
    qtxt = q.get("text") or ""
    if key == "zd_exists":
        if t in _BARE_YES:
            return key, True, qtxt
        if t in _BARE_NO:
            return key, False, qtxt
    elif key == "zd_placement":
        if "单独" in t or t in ("并列", "单列"):
            return key, "separate", qtxt
        if "计入" in t or "合算" in t or t == "二类费":
            return key, "inside", qtxt
    elif key == "zd_refine":
        if t in ("需要", "要", "细化", "逐项", "填", "填写"):
            return key, "detail", qtxt
        if t in ("不用", "不需要", "总额", "一笔", "填总额"):
            return key, "single", qtxt
    elif key in ("custom_need", "discount_need"):
        if t in _BARE_NO:
            return key, "none", qtxt
    return None


def _ai_mark_done(ctx: dict, key: str, text: str, answer: str):
    if key.startswith("card:"):
        # 用户明确关闭的参数卡（✅已设置/⏭用默认值）→ 后续轮次不再重问
        done_cards = ctx.setdefault("_ai_card_done", [])
        if key not in done_cards:
            done_cards.append(key)
    answered = ctx.setdefault("_ai_answered", [])
    if key not in answered:
        answered.append(key)
    _ai_record_answer(ctx, text, answer)
    ctx["qno"] = ctx.get("qno", 1) + 1
    if not _ai_remaining(ctx.get("_ai_plan") or {}, ctx):
        ctx["_ai_dirty"] = True
    st.rerun()


def _ai_write(ctx: dict, key: str, value):
    """语义键的 ctx 写入（镜像固定流程写入路径）。"""
    if key == "zd_exists":
        ctx["zd_asked"] = True
        ctx["zd_exists"] = bool(value)
    elif key == "zd_placement":
        ctx["zd_placement"] = value
        ctx["zhengchai_separate"] = (value == "separate")
    elif key == "zd_refine":
        ctx["zd_refine_asked"] = True
        ctx["_subflow"] = "zd_detail" if value == "detail" else "zd_single"
    elif key == "custom_need":
        if value == "none":
            ctx["custom_answered"] = True
            ctx.pop("_reask", None)
        elif value == "custom":
            ctx["_subflow"] = "custom_loop"
        else:
            ctx["_subflow"] = "contract_loop"
    elif key == "discount_need":
        if value == "none":
            ctx["discount_answered"] = True
            ctx.pop("_reask", None)
        elif value == "discount":
            ctx["_subflow"] = "discount_loop"
        else:
            ctx["discount_scenario"] = True
            ctx["discount_answered"] = True
    elif key == "card:party":
        ctx["jiaoyi_party"] = value  # None = 双方合计（也是有效答案）
        ctx["_ai_party_done"] = True
    elif key == "card:yubei_include":
        ctx["yubei_considered"] = bool(value)
    else:
        ctx.setdefault("ai_notes", []).append(f"{value}")
        # 保险：按钮值本身含征地/拆迁字样（LLM 自造按钮选项、free 键问题）
        # → 直接登记状态，杜绝同一问题被重复提问
        _v = str(value or "")
        if re.search(r"征地|拆迁", _v):
            ctx["zd_asked"] = True
            ctx["zd_exists"] = not re.search(r"没有|无|不涉及|不需要", _v)


# ── 控件渲染 ──

def _ai_render_buttons(ctx: dict, q: dict, qno: int):
    opts = _AI_CANON_OPTIONS.get(q["key"])
    pairs = opts if opts else [(o, o) for o in q.get("options") or []]
    cols = st.columns(min(len(pairs), 3))
    for oi, (label, val) in enumerate(pairs):
        with cols[oi % 3]:
            if st.button(label, key=f"ai_btn_{q['key']}_{oi}_{qno}",
                         use_container_width=True):
                _ai_write(ctx, q["key"], val)
                _ai_mark_done(ctx, q["key"], q["text"], label)


def _ai_render_number(ctx: dict, q: dict, qno: int):
    key = q["key"]
    if key == "amounts_jianan":
        am = ctx.setdefault("amounts", {})
        v = st.number_input("建安工程费（万元）", min_value=0.0,
                            value=float(am.get("jianan") or 0.0), step=100.0,
                            key=f"ai_num_jianan_{qno}")
        if st.button("✅ 确认", key=f"ai_num_ok_jianan_{qno}",
                     use_container_width=True):
            if v and v > 0:
                am["jianan"] = float(v)
                ctx.pop("_reask", None)
                _ai_mark_done(ctx, key, q["text"], f"建安费 {v} 万元")
            else:
                st.warning("请输入大于 0 的金额")
    elif key == "amounts_shebei":
        am = ctx.setdefault("amounts", {})
        v = st.number_input("设备购置费（万元）", min_value=0.0,
                            value=float(am.get("shebei") or 0.0), step=50.0,
                            key=f"ai_num_shebei_{qno}")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("✅ 确认", key=f"ai_num_ok_shebei_{qno}",
                         use_container_width=True):
                am["shebei"] = float(v)
                am["shebei_known"] = True
                _ai_mark_done(ctx, key, q["text"], f"设备费 {v} 万元")
        with c2:
            if st.button("🚫 无设备费", key=f"ai_num_nosb_{qno}",
                         use_container_width=True):
                am["shebei"] = 0.0
                am["shebei_known"] = True
                _ai_mark_done(ctx, key, q["text"], "无设备费")
    elif key == "card:yubei_rate":
        v = st.number_input("预备费率（%）", min_value=0.0,
                            value=float(ctx.get("yubei_rate", 5.0)), step=0.5,
                            format="%.1f", key=f"ai_num_yubei_{qno}")
        if st.button("✅ 使用该费率", key=f"ai_num_ok_yubei_{qno}",
                     use_container_width=True):
            ctx["yubei_rate"] = float(v)
            ctx["_ai_yubei_done"] = True
            _ai_mark_done(ctx, key, q["text"], f"预备费率 {v}%")


def _ai_render_text(ctx: dict, q: dict, qno: int):
    txt = st.text_input("（自由回答）", key=f"ai_txt_free_{qno}",
                        placeholder="在此输入，或直接使用下方聊天输入框")
    if st.button("✅ 发送", key=f"ai_txt_ok_free_{qno}",
                 use_container_width=True, disabled=not txt):
        # 问题本身是征地拆迁问句（LLM 误用自由文本提问）→ 裸答直接登记状态
        _qt = str(q.get("text") or "")
        if re.search(r"征地|拆迁", _qt):
            _t = re.sub(r"[。.！!？?，,\s]", "", txt.strip())
            if _t in _BARE_YES:
                _ai_write(ctx, "zd_exists", True)
                _ai_mark_done(ctx, "zd_exists", _qt, txt)
            elif _t in _BARE_NO:
                _ai_write(ctx, "zd_exists", False)
                _ai_mark_done(ctx, "zd_exists", _qt, txt)
        # 复用自由文本解析；未识别的内容存为补充说明供 AI 参考
        res = advance_checklist(ctx, txt, region=ctx.get("region"))
        if not res.get("recognized"):
            ctx.setdefault("ai_notes", []).append(txt)
        ctx["_ai_dirty"] = True
        _ai_mark_done(ctx, q["key"], q["text"], txt)


def _ai_card_answer_text(ctx: dict, card: dict) -> str:
    kind, fee = card["kind"], card["fee"]
    label = FEE_CATALOG.get(fee, {}).get("label", fee)
    try:
        if kind in ("rate", "shencha"):
            v = ctx.get("rate_overrides", {}).get("施工图审查费" if kind == "shencha" else fee)
            return f"{label}：{v}"
        if kind == "coef":
            _ans = f"{label}：{ctx.get('coef_overrides', {}).get(fee)}"
            _sep = ctx.get("sheji_extra_separate") or {}
            _sep_txt = "、".join(_nm for _nm, _on in _sep.items() if _on)
            if _sep_txt:
                _ans += f"；单独计算：{_sep_txt}"
            return _ans
        if kind == "service":
            sv = ctx.get("service_selections", {}).get(fee) or ["默认"]
            return f"{label}：{'、'.join(sv)}"
        if kind == "shuibao_comp":
            return f"{label}：{ctx.get('spec_overrides', {}).get('水土保持补偿费')}"
        if kind == "spec":
            return f"{label}：{ctx.get('spec_params', {}).get(fee)}"
    except Exception:
        pass
    return f"{label}已设置"


_AI_KIND_LABEL = {
    "coef": "系数设置", "rate": "费率档位", "shencha": "施工图审查档位",
    "service": "服务类型", "party": "计费方", "yubei_include": "预备费是否考虑",
    "yubei_rate": "预备费率", "shuibao_comp": "水土保持补偿费参数",
    "spec": "参数设置",
}


def _ai_card_idx(ctx: dict, card: dict):
    """卡片在 param_cards_for 中的位置（找不到返回 None）。"""
    return next((i for i, c in enumerate(param_cards_for(ctx))
                 if c["kind"] == card["kind"] and c["fee"] == card["fee"]), None)


def _ai_goto_card(ctx: dict, card: dict, toast_msg: str):
    """构造「回看」plan 强制渲染指定参数卡（已填值保留作默认，可确认可修改）。"""
    label = FEE_CATALOG.get(card["fee"], {}).get("label", card["fee"])
    kind_label = _AI_KIND_LABEL.get(card["kind"], card["kind"])
    plan = {"think": f"返回查看{label}参数",
            "questions": [{"type": "buttons", "key": _card_qkey(card), "force": True,
                           "text": f"{label}—{kind_label}"}]}
    ctx["_ai_plan"] = plan
    ctx["_ai_plan_region"] = ctx.get("region")
    ctx["_ai_answered"] = []
    ctx["_ai_dirty"] = False
    ctx.setdefault("_ai_rounds", []).append({"think": plan["think"], "qa": []})
    ctx["qno"] = ctx.get("qno", 1) + 1
    st.toast(toast_msg)
    st.rerun()


def _ai_back_prev_fee(ctx: dict, card: dict):
    """AI 卡界面返回按钮：

    - 同费种上一张卡 / 上一费种的第一张卡 → 机械回退（不清已填参数，
      「回看」plan 强制渲染，卡自带返回按钮可连续逐级回退）；
    - 已是第一个费种的第一张卡 → 「返回上一步」：回退到参数阶段之前，
      重新确认征地拆迁费、自定义费用等基础信息（已填值保留作默认）。
    """
    cards = param_cards_for(ctx)
    cur_idx = _ai_card_idx(ctx, card)
    if cur_idx is None:
        return
    if cur_idx == 0:
        # 第一个费种的第一张卡：无更早参数设置 → 返回上一步重确认基础信息
        _step_back(ctx, from_stage="params")
        ctx["_ai_dirty"] = True
        st.toast("已返回上一步：可重新确认征地拆迁费、自定义费用等基础信息"
                 "（已填内容保留，可直接确认或修改）。")
        st.rerun()
        return
    prev = cards[cur_idx - 1]
    if prev["fee"] != card["fee"]:
        # 跨费种回退：跳到上一个费种的第一张卡
        prev = next(c for c in cards if c["fee"] == prev["fee"])
    label = FEE_CATALOG.get(prev["fee"], {}).get("label", prev["fee"])
    _ai_goto_card(ctx, prev, f"已返回{label}参数设置（已填选项保留，可直接确认或修改）。")


def _ai_render_card_question(ctx: dict, q: dict, qno: int):
    card = _ai_parse_card_key(q["key"])
    renderer = _CARD_RENDERERS.get(card["kind"])
    # 第一个费种的第一张卡：无更早参数设置 → 按钮改为「返回上一步」
    _back_label = ("⬅ 返回上一步" if _ai_card_idx(ctx, card) == 0
                   else "⬅ 返回上一参数设置")
    if renderer is None:  # party 走按钮控件；未知 kind 兜底
        _ai_render_buttons(ctx, q, qno)
        if st.button(_back_label, key=f"ai_card_back_{q['key']}_{qno}",
                     use_container_width=True):
            _ai_back_prev_fee(ctx, card)
        return
    renderer(ctx, card, qno)
    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("✅ 已设置，继续", key=f"ai_card_ok_{q['key']}_{qno}",
                     use_container_width=True):
            _commit_card_widgets(ctx, card, qno)
            _ai_mark_done(ctx, q["key"], q["text"], _ai_card_answer_text(ctx, card))
    with c2:
        if st.button("⏭ 本卡用默认值", key=f"ai_card_def_{q['key']}_{qno}",
                     use_container_width=True):
            _clear_card_overrides(ctx, card)
            _ai_mark_done(ctx, q["key"], q["text"], "用默认值")
    with c3:
        if st.button(_back_label, key=f"ai_card_back_{q['key']}_{qno}",
                     use_container_width=True):
            _ai_back_prev_fee(ctx, card)


def _ai_render_fees_question(ctx: dict, q: dict, qno: int):
    """费种确认/选择：镜像 _ask_fees 两分支，但完成动作走 _ai_mark_done。"""
    if q["key"] == "fees_preset_confirm":
        preset = _preset_fee_numbers()
        excluded = ctx.get("fee_exclusions") or set()
        st.caption("点击序号即可排除，或直接回复序号（如「不需要 3、5」）。")
        cols = st.columns(3)
        for n, name, label in preset:
            is_ex = name in excluded
            with cols[(n - 1) % 3]:
                if st.button(f"{'🚫 ' if is_ex else ''}{n}. {label}",
                             key=f"ai_fees_ex_{name}_{qno}",
                             use_container_width=True,
                             type="primary" if is_ex else "secondary"):
                    if is_ex:
                        excluded.discard(name)
                        if name not in ctx["fees"]:
                            ctx["fees"].append(name)
                    else:
                        excluded.add(name)
                        if name in ctx["fees"]:
                            ctx["fees"].remove(name)
                    ctx["fee_exclusions"] = excluded
                    ctx["qno"] = qno + 1
                    st.rerun()
        remain = len(ctx["fees"])
        if st.button(f"✅ 确认，按剩余 {remain} 项计算",
                     key=f"ai_fees_allok_{qno}", use_container_width=True,
                     type="primary"):
            ctx["fees_confirmed"] = True
            ctx.pop("_reask", None)
            _ai_mark_done(ctx, q["key"], q["text"], f"按剩余 {remain} 项计算")
        return
    # fees_select
    defs = fee_catalog_for(ctx.get("region"), exclude_reserve=True)
    selected = set(ctx.get("fees") or [])
    cols = st.columns(3)
    for i, fd in enumerate(defs):
        is_sel = fd["name"] in selected
        with cols[i % 3]:
            if st.button(f"{'✅ ' if is_sel else ''}{fd['label']}",
                         key=f"ai_fees_{i}_{qno}", use_container_width=True,
                         type="primary" if is_sel else "secondary"):
                if is_sel:
                    selected.discard(fd["name"])
                else:
                    selected.add(fd["name"])
                ctx["fees"] = [d["name"] for d in defs if d["name"] in selected]
                ctx["qno"] = qno + 1
                st.rerun()
    col1, col2 = st.columns(2)
    with col1:
        if st.button("📦 全部二类费（不含交易服务费）", key=f"ai_fees_all_{qno}",
                     use_container_width=True):
            ctx["fees"] = [f for f in FEE_CATALOG if f not in ("交易服务费", "预备费")]
            ctx["preset_all"] = True
            ctx["qno"] = qno + 1
            st.rerun()
    with col2:
        if st.button("✅ 确认所选费种", key=f"ai_fees_ok_{qno}",
                     use_container_width=True, disabled=not selected):
            ctx["fees_confirmed"] = True
            ctx.pop("_reask", None)
            _ai_mark_done(ctx, q["key"], q["text"],
                          "、".join(FEE_CATALOG.get(f, {}).get("label", f)
                                    for f in ctx["fees"]))


def _ai_render_question(ctx: dict, q: dict, qno: int):
    key = q.get("key")
    if key in ("fees_preset_confirm", "fees_select"):
        _ai_render_fees_question(ctx, q, qno)
    elif key in ("amounts_jianan", "amounts_shebei", "card:yubei_rate"):
        _ai_render_number(ctx, q, qno)
    elif key in _AI_CANON_OPTIONS:
        _ai_render_buttons(ctx, q, qno)
    elif key and key.startswith("card:"):
        _ai_render_card_question(ctx, q, qno)
    else:  # free（或 AI 自行给出的 buttons/text）
        if q.get("type") == "buttons":
            _ai_render_buttons(ctx, q, qno)
        else:
            _ai_render_text(ctx, q, qno)


# ── 气泡渲染 ──

def _ai_header_history(ctx: dict):
    st.markdown("### 🧠 概算助手 — 信息确认")
    st.caption("AI 将根据已提供的信息动态确认计算所需内容；点击选项、填写数字或直接自由输入均可。")
    fees = ctx.get("fees") or []
    if fees and ctx.get("fees_confirmed"):
        # 未确认前费种问题卡已展示列表，头部不重复列
        st.markdown("**已选费种**：" + "、".join(
            FEE_CATALOG.get(f, {}).get("label", f) for f in fees))
    am = ctx.get("amounts") or {}
    bits = []
    if am.get("jianan") is not None:
        bits.append(f"建安费 {am['jianan']} 万")
    if am.get("shebei_known"):
        bits.append(f"设备费 {am.get('shebei', 0)} 万")
    if bits:
        st.markdown("**已识别金额**：" + "，".join(bits))
    rounds = ctx.get("_ai_rounds") or []
    if len(rounds) > 1:
        with st.expander("📜 查看此前的问答记录"):
            for rd in rounds[:-1]:
                if rd.get("think"):
                    st.markdown(f"💭 *{rd['think']}*")
                for qa in rd.get("qa") or []:
                    st.markdown(f"- 问：{qa.get('text')}\n  - 答：{qa.get('answer')}")
    if rounds:
        for qa in rounds[-1].get("qa") or []:
            st.caption(f"✅ {qa.get('text')} → {qa.get('answer')}")
    st.markdown("---")


def _ai_footer(ctx: dict):
    st.markdown("---")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("🔧 切换传统清单模式", key=f"ai_switch_{ctx.get('qno', 1)}",
                     use_container_width=True):
            ctx["ai_guided"] = False
            st.session_state["ai_guide_pref"] = False
            ctx["param_idx"] = 0
            for k in ("_ai_plan", "_ai_answered", "_ai_dirty", "_ai_plan_region"):
                ctx.pop(k, None)
            st.rerun()
    with c2:
        if st.button("🗑 结束本次计算", key=f"ai_cancel_{ctx.get('qno', 1)}",
                     use_container_width=True):
            _cancel_task(ctx)


def _ai_fallback(ctx: dict, msg: str):
    ctx["_ai_fallback"] = True
    ctx["_ai_fallback_msg"] = msg
    # AI 模式卡内按钮曾调用 _params_advance 污染过 param_idx/params_done，
    # 这里复位：传统清单从参数卡起点重走；_params_seek 让 _ask_params
    # 仅首次顺延到第一张未填卡，之后逐张显示（编号连续不跳号）
    ctx["param_idx"] = 0
    ctx["_params_seek"] = True
    ctx["params_done"] = False
    st.rerun()


def _ask_ai_guide(ctx: dict):
    """AI 对话式提问主流程：规划 → 提问 → 作答 → 重规划，直至信息齐全进最终确认。"""
    qno = ctx.get("qno", 1)
    sub = ctx.get("_subflow")
    last_sub = ctx.get("_ai_last_subflow")
    if last_sub in _AI_SUBFLOWS and sub is None:
        # 刚从子流程返回：有实际数据才补记完成标志，避免 AI 重复提问
        # （固定流程里这些标志由阶段级「确认，无需更多」按钮置位，AI 模式没有该按钮）
        if last_sub == "custom_loop" and ctx.get("custom_fees"):
            ctx["custom_answered"] = True
        elif last_sub == "contract_loop" and ctx.get("contract_overrides"):
            ctx["custom_answered"] = True
        elif last_sub == "discount_loop":
            # 折扣子流程返回（含「↩ 返回」）即视为已作答：子流程内已给过
            # 添加机会，结果（discounts 有无）就是答案，避免 AI 重复提问
            ctx["discount_answered"] = True
        ctx["_ai_dirty"] = True
    ctx["_ai_last_subflow"] = sub
    if sub in _AI_SUBFLOWS:
        _ai_header_history(ctx)
        _SUBFLOW_RENDERERS[sub](ctx)
        _ai_footer(ctx)
        return

    plan = ctx.get("_ai_plan")
    if (ctx.get("_ai_dirty") or plan is None
            or ctx.get("_ai_plan_region") != ctx.get("region")):
        avail = available_keys(ctx)
        try:
            with st.spinner("🧠 正在思考接下来需要确认的信息…"):
                plan = run_planner(build_ctx_summary(ctx, avail), avail)
        except PlannerError:
            _ai_fallback(ctx, "AI 引导暂时不可用（接口异常），已切换为传统清单模式。")
        # 问题净化：free 参数打包问题 → 第一个未填参数卡（卡片式提问）；
        # free 键基础问题（含折扣/打折/下浮）→ 恢复规范按钮键，杜绝重复提问
        # （逻辑在 ai_planner.sanitize_plan_questions，纯函数可直接单测）
        _avail2 = available_keys(ctx)
        plan["questions"] = sanitize_plan_questions(
            plan.get("questions") or [], dict(_avail2))
        ctx["_ai_plan"] = plan
        ctx["_ai_plan_region"] = ctx.get("region")
        if plan.get("questions"):
            ctx.pop("_ai_empty_plans", None)  # 正常出题 → 空计划计数复位
        ctx["_ai_answered"] = []
        ctx["_ai_dirty"] = False
        ctx.setdefault("_ai_rounds", []).append({"think": plan.get("think") or "", "qa": []})

    # params 阶段完成标志与参数卡实际填写状态同步：AI 模式卡内控件直接写
    # ctx（不依赖 param_idx 推进），避免 current_stage 误判「params 未完成」
    # 而降级传统清单重问参数卡
    ctx["params_done"] = all(
        question_done(ctx, _card_qkey(c)) for c in param_cards_for(ctx))

    # 卡内控件直接写入 ctx 的问题（未走「✅已设置/⏭默认值」按钮）→ 补记
    # 完成标志与问答记录，避免同一费种参数在重规划或降级时被重复提问
    for q in plan.get("questions") or []:
        key = q.get("key") or ""
        if q.get("force"):
            continue  # 回看卡：强制渲染已填卡，不补记（补记会被 answered 过滤掉）
        if key not in (ctx.get("_ai_answered") or []) and question_done(ctx, key):
            ctx.setdefault("_ai_answered", []).append(key)
            if key.startswith("card:"):
                done_cards = ctx.setdefault("_ai_card_done", [])
                if key not in done_cards:
                    done_cards.append(key)
                _ai_record_answer(ctx, q.get("text"),
                                  _ai_card_answer_text(ctx, _ai_parse_card_key(key)))
            else:
                _ai_record_answer(ctx, q.get("text"), "已通过页面设置")

    if not plan.get("questions"):
        if current_stage(ctx) is None:
            _ai_header_history(ctx)
            _finalize_confirm(ctx)
            _ai_footer(ctx)
            return
        # 计划为空但仍有缺失：可能是 LLM 本轮问题全被丢弃（违规提前问参数），
        # 重规划重试；连续三轮为空才降级传统清单（避免重放已答问题）
        if ctx.get("_ai_empty_plans", 0) < 2:
            ctx["_ai_empty_plans"] = ctx.get("_ai_empty_plans", 0) + 1
            ctx["_ai_dirty"] = True
            st.rerun()
        _ai_fallback(ctx, "AI 判断信息已齐全，但程序校验发现仍有缺失，已切换为传统清单模式。")

    visible = _ai_remaining(plan, ctx)
    if not visible:
        # 本轮问题已全部通过卡内控件答完：重新规划（可用问题库已排除
        # 已填参数卡），不再降级传统清单——传统模式会从 param_idx 重放
        # 参数卡，导致已答费种被重复提问
        ctx["_ai_dirty"] = True
        st.rerun()

    _ai_header_history(ctx)
    if plan.get("think"):
        st.markdown(f"💭 *{plan['think']}*")
    for q in visible:
        st.markdown(f"**{q.get('text')}**")
        _ai_render_question(ctx, q, qno)
    _ai_footer(ctx)


def _finalize_confirm(ctx: dict):
    qno = ctx.get("qno", 1)
    st.markdown("**✅ 信息已齐全**，点击开始计算。")
    fees_txt = "、".join(FEE_CATALOG.get(f, {}).get("label", f)
                         for f in (ctx.get("fees") or []))
    am = ctx.get("amounts", {})
    st.markdown(f"- 费种：{fees_txt}")
    st.markdown(f"- 建安费 **{am.get('jianan', 0)}** 万 + 设备费 **{am.get('shebei', 0)}** 万")
    if am.get("total_investment"):
        st.markdown(f"- 项目总投资（已识别）：{am['total_investment']} 万")
    for fn, ov in (ctx.get("contract_overrides") or {}).items():
        if ov.get("type") == "rate":
            st.markdown(f"- {FEE_CATALOG.get(fn, {}).get('label', fn)}："
                        f"合同费率 {ov['rate']}% × {ov.get('base', '工程费')}")
        else:
            st.markdown(f"- {FEE_CATALOG.get(fn, {}).get('label', fn)}："
                        f"一口价 {ov.get('amount_wan')} 万")
    for fn, d in (ctx.get("discounts") or {}).items():
        st.markdown(f"- {FEE_CATALOG.get(fn, {}).get('label', fn)}：打 {d:g} 折")
    for e in (ctx.get("custom_fees") or []):
        st.markdown(f"- 自定义：{e.get('名称')} {e.get('金额(万元)')} 万")

    # 征地拆迁费单独列项：计算前确认口径（结算与 Excel 导出据此渲染）
    _zc_sep = st.checkbox(
        "🏗️ 征地拆迁费单独列项（与工程费、工程建设其他费并列；建管费基数仍扣除）",
        value=bool(ctx.get("zhengchai_separate",
                           st.session_state.get("zhengchai_separate", False))),
        key=f"ck_zc_sep_input_{qno}",
    )
    ctx["zhengchai_separate"] = _zc_sep
    st.session_state["zhengchai_separate"] = _zc_sep

    col1, col2 = st.columns(2)
    with col1:
        if st.button("🚀 开始计算", key=f"ck_final_go_{qno}",
                     use_container_width=True, type="primary"):
            _finalize_and_render(ctx)
    with col2:
        if st.button("🗑 结束本次计算", key=f"ck_final_cancel_{qno}",
                     use_container_width=True):
            _cancel_task(ctx)


# ── 征地拆迁费单独列项：面板口径对齐 Excel 导出 ──
# 识别规则与 fee_engine 导出区块、建管费基数扣除（_match_custom_fee_deductions）同源
_ZD_RE = re.compile(r"征地|拆迁|建设用地|土地征用|用地费")


def _zd_amount(custom_fees) -> float:
    """自定义费用中征地拆迁类金额合计（万元）。"""
    return round(sum(
        float(e.get("amount_wan") or e.get("金额(万元)") or 0)
        for e in (custom_fees or [])
        if _ZD_RE.search(str(e.get("名称") or e.get("name") or ""))
    ), 4)


def _apply_zhengchai_panel(payload: dict, ctx: dict) -> bool:
    """单独列项时调整面板汇总口径，与 Excel 导出五个部分一致：

    工程建设其他费（原二类费合计）= 二类费合计 − 征地拆迁；
    预备费 = (第一部分工程费 + 工程建设其他费) × 费率（基数不含征地拆迁）；
    总投资 = 工程费 + 征地拆迁 + 工程建设其他费 + 预备费。
    引擎各费种数值不动（建管/可研基数仍含征地拆迁，与导出迭代层口径一致）。
    """
    if not ctx.get("zhengchai_separate"):
        return False
    zd = _zd_amount(payload.get("custom_fees"))
    if zd <= 0:
        return False
    preview = payload.get("preview") or {}
    p1 = payload.get("total_part1") or 0.0
    rate = 0.0 if ctx.get("yubei_considered") is False \
        else float(ctx.get("yubei_rate") or 5.0)
    fee_excl = round((preview.get("fee_total_with_custom") or 0.0) - zd, 4)
    yb = preview.get("yubei_total") or 0.0
    ctr_yb = (payload.get("contract_overrides") or {}).get("预备费") or {}
    if ctr_yb.get("type") != "price":  # 合同一口价保持不变
        yb = round((p1 + fee_excl) * rate / 100.0, 4)
    preview["zhengchai_total"] = zd
    preview["fee_total_with_custom"] = fee_excl
    preview["yubei_total"] = yb
    preview["project_total_with_custom"] = round(p1 + fee_excl + zd + yb, 4)
    numerical = preview.get("numerical")
    if numerical:
        numerical["预备费(万元)"] = yb
    yb_rr = (preview.get("raw") or {}).get("原始结果") or {}
    if isinstance(yb_rr, dict):
        yb_rr = yb_rr.get("预备费")
    if isinstance(yb_rr, dict):
        _src = yb_rr.get("预备费率来源") or "默认"
        yb_rr["计算公式"] = (f"（{p1} + {round(fee_excl, 4)}）× "
                             f"{rate}%（{_src}）")
        yb_rr["结果(万元)"] = yb
        for _s in yb_rr.get("计算步骤") or []:
            if _s.get("步骤") == "计算":
                _s["公式"] = yb_rr["计算公式"] + "，不含征地拆迁费"
                _s["结果"] = f"{yb} 万元"
                break
    return True


def _apply_zhengchai_scenarios(result: dict, ctx: dict) -> None:
    """折扣情景对比：单独列项时各方案 二类费合计/预备费/总投资 同口径剔除征地拆迁。"""
    zd = _zd_amount(ctx.get("custom_fees"))
    if zd <= 0:
        return
    am = ctx.get("amounts") or {}
    p1 = round(float(am.get("jianan") or 0) + float(am.get("shebei") or 0), 4)
    rate = 0.0 if ctx.get("yubei_considered") is False \
        else float(ctx.get("yubei_rate") or 5.0)
    ctr_yb = (ctx.get("contract_overrides") or {}).get("预备费") or {}
    is_price = ctr_yb.get("type") == "price"
    for s in result.get("方案列表") or []:
        fee_excl = round((s.get("二类费合计(万元)") or 0.0) - zd, 4)
        d = s.get("各项费用") or {}
        yb = d.get("预备费(万元)", 0.0) if is_price \
            else round((p1 + fee_excl) * rate / 100.0, 4)
        s["二类费合计(万元)"] = fee_excl
        s["总投资(万元)"] = round(p1 + fee_excl + zd + yb, 4)
        s["项目总投资(万元)"] = s["总投资(万元)"]
        if d is not None:
            d["预备费(万元)"] = yb
    for row in result.get("对比表") or []:
        fname = row.get("费用名称")
        for i, s in enumerate(result.get("方案列表") or [], 1):
            k = f"方案{i}"
            if fname == "预备费(万元)":
                row[k] = (s.get("各项费用") or {}).get("预备费(万元)")
            elif fname == "二类费合计(万元)":
                row[k] = s["二类费合计(万元)"]
            elif fname in ("总投资(万元)", "项目总投资(万元)"):
                row[k] = s["项目总投资(万元)"]


def _gen_discount_scenarios(ctx: dict):
    """折扣情景对比（单独列项时同步剔除征地拆迁口径）。"""
    result = calc_discount_scenarios(ctx)
    if ctx.get("zhengchai_separate"):
        _apply_zhengchai_scenarios(result, ctx)
    return result


def _finalize_and_render(ctx: dict):
    """结算并进入 done 阶段（结果持久化在 pending_task["payload"]）。"""
    with st.spinner("正在计算..."):
        result = settle_checklist(ctx)
    if result["kind"] == "cascade":
        payload = result["payload"]
        # 征地拆迁费单独列项：面板汇总口径对齐 Excel 导出（剔除征地拆迁）
        _apply_zhengchai_panel(payload, ctx)
        ctx["payload"] = payload
        ctx["phase"] = "done"
        _record_calculation(ctx.get("query", ""), {
            "fee_type": "__cascade__",
            "结果汇总": {"项目总投资(万元)": payload["preview"]["project_total_with_custom"]},
        })
        # 折扣情景：结算后自动生成对比（缓存防重复计算）
        if ctx.get("discount_scenario"):
            ctx["_discount_scenarios"] = _gen_discount_scenarios(ctx)
        _zd_disp = payload["preview"].get("zhengchai_total")
        if _zd_disp:
            summary = (
                f"## 二类费计算结果\n\n"
                f"工程费 **{payload['total_part1']:.2f} 万元**，"
                f"征地拆迁费 **{_zd_disp:.2f} 万元**（单独列项），"
                f"工程建设其他费 **{payload['preview']['fee_total_with_custom']:.2f} 万元**，"
                f"预备费 **{payload['preview']['yubei_total']:.2f} 万元**，"
                f"工程总投资 **{payload['preview']['project_total_with_custom']:.2f} 万元**。\n\n"
                f"详细结果见下方面板。"
            )
        else:
            summary = (
                f"## 二类费计算结果\n\n"
                f"二类费合计 **{payload['preview']['fee_total_with_custom']:.2f} 万元**，"
                f"项目总投资 **{payload['preview']['project_total_with_custom']:.2f} 万元**。\n\n"
                f"详细结果见下方面板。"
            )
    else:
        ctx["phase"] = "done"
        ctx["simple_results"] = result["results"]
        lines = []
        for fn, r in (result["results"] or {}).items():
            if r:
                val = r.get("结果(万元)", r.get("结果(元)", ""))
                lines.append(f"- **{FEE_CATALOG.get(fn, {}).get('label', fn)}**：{val}")
            else:
                lines.append(f"- **{fn}**：计算失败（参数不足）")
        summary = "## 计算结果\n\n" + "\n".join(lines)
    st.session_state.messages.append({"role": "assistant", "content": summary})
    st.rerun()


def _render_result_payload(ctx: dict):
    """done 阶段：汇总指标 + 各费种分步计算 + Excel/折扣对比按钮。"""
    if st.button("⬅ 返回修改设置", key=f"ck_result_back_{ctx.get('qno', 1)}"):
        # 回到最后一个步骤（折扣）重新确认；可继续「上一步」回退到更早阶段
        ctx["phase"] = "ask"
        ctx["_reask"] = "discount"
        ctx["discount_answered"] = False
        ctx["qno"] = ctx.get("qno", 1) + 1
        st.rerun()
    payload = ctx.get("payload")
    if payload:
        preview = payload["preview"]
        _zd_disp = preview.get("zhengchai_total")
        st.markdown("## 二类费计算结果（程序精确计算）")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("第一部分工程费", f"{payload['total_part1']:.2f} 万元")
        c2.metric("工程建设其他费" if _zd_disp else "二类费合计",
                  f"{preview['fee_total_with_custom']:.2f} 万元")
        c3.metric("预备费", f"{preview['yubei_total']:.2f} 万元")
        c4.metric("项目总投资", f"{preview['project_total_with_custom']:.2f} 万元")
        if _zd_disp:
            st.info(
                f"🏗️ 征地拆迁费 **{_zd_disp:,.2f} 万元**已单独列项，"
                f"与工程费、工程建设其他费并列；预备费基数不含征地拆迁费。"
            )

        numerical = preview.get("numerical", {})
        rows = []
        for fd in payload.get("fee_defs", []):
            if fd["name"] not in payload.get("selected_fees", set()):
                continue
            val = numerical.get(f"{fd['name']}(万元)")
            if val is None:
                continue
            if val <= 0 and fd["name"] != "预备费":
                continue
            rows.append({"费种": fd["label"], "金额(万元)": round(val, 2)})
        if rows:
            import pandas as pd
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

        custom_fees = payload.get("custom_fees") or []
        if custom_fees:
            st.markdown("**自定义费用**：")
            for e in custom_fees:
                st.markdown(f"- {e.get('名称') or e.get('name')}："
                            f"{e.get('金额(万元)') or e.get('amount_wan')} 万元")

        # 设计费附加项单独计算（非天津勾选才出现）
        _sep_lines = []
        for _nm, _lab in (("施工图预算编制费", "施工图预算编制费（设计费×10%）"),
                          ("竣工图编制费", "竣工图编制费（设计费×8%）")):
            _v = numerical.get(f"{_nm}(万元)")
            if _v:
                _sep_lines.append(f"- {_lab}：{_v} 万元")
        if _sep_lines:
            st.markdown("**设计费附加项（单独计算）**：")
            st.markdown("\n".join(_sep_lines))

        # 各费种计算过程（复用 _render_engine_card 统一步骤格式，零改动）
        raw = preview.get("raw", {}).get("原始结果", {})
        with st.expander("查看各费种详细计算步骤", expanded=False):
            for fd in payload.get("fee_defs", []):
                fn = fd["name"]
                if fn not in payload.get("selected_fees", set()):
                    continue
                detail = raw.get(fn)
                if detail:
                    st.markdown(f"#### {fd['label']}")
                    _render_engine_card(detail)
            sb = raw.get("水土保持补偿费")
            if sb:
                st.markdown("#### 水土保持补偿费")
                _render_engine_card(sb)
            _sep_rr = raw.get("设计费附加项")
            if _sep_rr:
                st.markdown("#### 设计费附加项（单独计算）")
                _render_engine_card(_sep_rr)

        # 折扣情景对比（结算后按需生成，缓存防重复计算）
        if ctx.get("_discount_scenarios"):
            _render_comparison_result(ctx["_discount_scenarios"])

        # Excel 概算表：下载按钮与概算成果 PDF 共用同一份字节（先统一生成一次）
        _is_tj = str(ctx.get("region") or DEFAULT_REGION) == "天津"
        _zy_list = st.session_state.get("zhuan_ye_list") or []
        _per_item = (st.session_state.get("pdf_split_mode")
                     == "分开做概算表（每单项一份）" and len(_zy_list) > 1)
        _region = str(ctx.get("region") or DEFAULT_REGION)
        _excel_files = []   # [(标签, bytes, 文件名)]；分开模式 = N 单项各一份
        excel_err = None
        try:
            _base = {
                **payload,
                "region": ctx.get("region"),
                "yubei_rate": ctx.get("yubei_rate"),
                "yubei_considered": ctx.get("yubei_considered"),
                "jiaoyi_party": ctx.get("jiaoyi_party"),
                "spec_overrides": ctx.get("spec_overrides"),
                "discounts": ctx.get("discounts"),
                "zhuan_ye": _zy_list,
                "zhengchai_separate": ctx.get("zhengchai_separate", False),
                "sheji_extra_separate": ctx.get("sheji_extra_separate"),
                "split_mode": "one_sheet",
            }
            if _per_item:
                for _i, _g in enumerate(_zy_list, 1):
                    _gl = _item_gl(_g, _i)
                    # 每单项一份完整概算表：第一部分=该单项，二类费/预备费/总投资
                    # 按该单项自己的工程费计算。交易服务费/可研费/环评费在模板
                    # 里是写死的数字（不是公式），必须用该单项工程费重跑引擎
                    # 得到单项自算值，否则每个文件都会抄同一个整体值
                    _p1 = sum(float(x.get("amount") or 0)
                              for it in (_g.get("items") or [])
                              for x in (it.get("leaves") or []))
                    _dev = sum(float(x.get("device") or 0)
                               for it in (_g.get("items") or [])
                               for x in (it.get("leaves") or []))
                    # 按单项问的征地拆迁/自定义费用：该单项自己的条目
                    # + 项目级既有费用；未走按单项流程时回退整体口径
                    _cf_map = ctx.get("item_custom_fees") or {}
                    if _cf_map:
                        _cf_i = ((ctx.get("item_fees_base") or [])
                                 + (_cf_map.get(_gl) or []))
                    else:
                        _cf_i = payload.get("custom_fees") or None
                    _payload_i = compute_selected_fees(
                        jianan=round(_p1 - _dev, 4), shebei=round(_dev, 4),
                        project_type=payload.get("project_type") or "通用",
                        query=str(payload.get("query") or ""),
                        region=_region,
                        selected_fees=payload.get("selected_fees") or set(),
                        coef_overrides=payload.get("coef_overrides") or None,
                        rate_overrides=payload.get("rate_overrides") or None,
                        service_selections=payload.get("service_selections") or None,
                        spec_overrides=payload.get("spec_overrides") or None,
                        contract_overrides=payload.get("contract_overrides") or None,
                        fee_discounts=payload.get("fee_discounts") or None,
                        custom_fees=_cf_i,
                        yubei_rate=(0.0 if ctx.get("yubei_considered") is False
                                    else float(ctx.get("yubei_rate") or 5.0)),
                        jiaoyi_party=ctx.get("jiaoyi_party"),
                        # 不用整体覆盖值：让引擎按该单项自算收敛总投资
                        total_investment_override=None,
                        sheji_extra_separate=ctx.get("sheji_extra_separate") or None,
                    )
                    # 单独列项口径同步（预备费基数不含征地拆迁等）
                    _apply_zhengchai_panel(_payload_i, ctx)
                    _bytes = _build_cascade_excel({
                        **_base, **_payload_i,
                        "zhuan_ye": [_g], "split_mode": "one_sheet"})
                    _excel_files.append(
                        (_gl, _bytes, f"{_region}-{_gl}-概算表.xlsx"))
            else:
                _bytes = _build_cascade_excel(_base)
                _excel_files.append(
                    ("概算表", _bytes, f"{_region}-总概算表.xlsx"))
        except Exception as e:
            excel_err = str(e)
        col1, col2 = st.columns(2)
        with col1:
            if excel_err:
                st.error(f"Excel 生成失败：{excel_err}")
            elif len(_excel_files) == 1:
                _lbl, _bytes, _fname = _excel_files[0]
                st.download_button(
                    "📥 导出 Excel 汇总表",
                    data=_bytes,
                    file_name=_fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"ck_excel_{ctx.get('qno', 1)}",
                    use_container_width=True,
                )
                if _is_tj and not TJ_TEMPLATE_XLSX.exists():
                    st.caption("⚠ 未找到天津模板文件，导出将回退为通用布局。")
            else:
                st.markdown("**📥 导出 Excel 概算表（每单项一份）**")
                for _i, (_lbl, _bytes, _fname) in enumerate(_excel_files):
                    st.download_button(
                        f"📥 {_lbl}概算表",
                        data=_bytes,
                        file_name=_fname,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"ck_excel_{ctx.get('qno', 1)}_{_i}",
                        use_container_width=True,
                    )
        with col2:
            if st.button("📊 折扣方案对比", key=f"ck_scen_btn_{ctx.get('qno', 1)}",
                         use_container_width=True):
                with st.spinner("正在生成折扣情景..."):
                    ctx["_discount_scenarios"] = _gen_discount_scenarios(ctx)
                st.rerun()

        # ── 📦 概算成果压缩包：总概算表 Excel + 编制说明 Word + 成套概算 PDF ──
        with st.expander("📦 生成概算成果（总概算表Excel + 编制说明Word + 成套PDF 压缩包）",
                         expanded=bool(st.session_state.get("deliverable_pdf"))):
            _deliv = st.session_state.get("deliverable_pdf")
            if _deliv:
                _dl = _deliv if isinstance(_deliv, list) else [_deliv]
                for _i, _d in enumerate(_dl):
                    st.download_button(
                        "📥 下载概算成果压缩包"
                        + (f"（{_d['label']}）" if len(_dl) > 1 else ""),
                        data=_d["data"], file_name=_d["name"], mime="application/zip",
                        key=f"deliv_dl_{ctx.get('qno', 1)}_{_i}",
                        use_container_width=True)
                st.caption(f"已生成 {len(_dl)} 套成果（每套内含总概算表Excel、"
                           f"编制说明Word、成套概算PDF 三个文件；"
                           f"修改表单后再次点击即可重新生成）")
            if st.session_state.get("src_pdf_files") or st.session_state.get("pdf_bytes"):
                st.caption("✅ 已检测到本次会话的广联达源 PDF，将附在成果末尾。")
            else:
                st.warning("⚠ 未检测到本次会话的广联达源 PDF（PDF 导入页上传解析时自动保存；"
                           "Excel 导入时在该页「📎 附加广联达源 PDF」上传），"
                           "成果将只含封面、编制说明与概算表。")
            _zy_list = st.session_state.get("zhuan_ye_list") or []
            _def_name = (_zy_list[0].get("name") or "").strip() if _zy_list else ""
            import datetime as _dt
            _per_item_infos: list[dict] = []
            _today_str = _dt.datetime.now().strftime("%Y年%m月")
            with st.form(f"deliv_form_{ctx.get('qno', 1)}"):
                if _per_item:
                    # 分开导出：每个单项一套成果 → 信息按单项分别填写
                    _lbls = [_item_gl(g, i) for i, g in enumerate(_zy_list, 1)]
                    # 首组名若与某单项标签相同，说明文件里没有真正的主项目名
                    # → 默认项目名称直接用单项名（避免「12#水系桥-14#水系桥」）
                    _has_main = bool(_def_name) and _def_name not in _lbls
                    for _ix, _lb in enumerate(_lbls):
                        with st.expander(f"📦 第 {_ix + 1} 项成果信息 — {_lb}",
                                         expanded=(_ix == 0)):
                            _n = st.text_input(
                                "项目名称",
                                value=(f"{_def_name}-{_lb}" if _has_main
                                       else _lb),
                                key=f"deliv_name_{_ix}",
                                help="该单项概算成果封面上的工程名称"
                                     "（封面将按排版自动换行，无需手动分行）。")
                            _no = st.text_input("工程编号", value="",
                                                key=f"deliv_no_{_ix}")
                            _dt2 = st.text_input("日期", value=_today_str,
                                                 key=f"deliv_date_{_ix}")
                            _ov = st.text_area(
                                "工程概况", value="", height=120,
                                key=f"deliv_ov_{_ix}",
                                placeholder="如：本工程为东西走向城市次干路，路线全长约400米，"
                                            "配套建设排水、照明、绿化、交通工程等。")
                            _fu = st.text_input(
                                "资金来源（可空，留空保留模板原文）", value="",
                                key=f"deliv_fund_{_ix}")
                            _per_item_infos.append({
                                "name": (_n or "").strip(),
                                "proj_no": (_no or "").strip(),
                                "date": (_dt2 or "").strip(),
                                "overview": (_ov or "").strip(),
                                "fund_source": (_fu or "").strip()})
                    _d_submit = st.form_submit_button(
                        f"📦 生成概算成果压缩包（{len(_lbls)} 套，每单项一套）",
                        use_container_width=True)
                else:
                    _d_name = st.text_input(
                        "项目名称", value=_def_name,
                        help="封面将按排版自动换行（第一行略长），无需手动分行。")
                    _d_no = st.text_input("工程编号", value="")
                    _d_date = st.text_input("日期", value=_today_str)
                    _d_ov = st.text_area("工程概况", value="", height=120,
                                         placeholder="如：本工程为东西走向城市次干路，路线全长约400米，"
                                                     "配套建设排水、照明、绿化、交通工程等。")
                    _d_fund = st.text_input("资金来源（可空，留空保留模板原文）", value="")
                    _d_submit = st.form_submit_button("📦 生成概算成果压缩包",
                                                      use_container_width=True)
            if _d_submit:
                if excel_err:
                    st.error(f"概算表生成失败，无法生成成果：{excel_err}")
                else:
                    # 编制说明费用段：选中且金额>0 的费种（不含预备费/交易服务费）
                    # + 自定义费用（不含单独列项的征地拆迁）+ 设计费附加项
                    _raw_rr = preview.get("raw", {}).get("原始结果", {})
                    _fl = []
                    for fd in payload.get("fee_defs", []):
                        fn = fd["name"]
                        if fn in ("预备费", "交易服务费"):
                            continue
                        if fn not in payload.get("selected_fees", set()):
                            continue
                        _v = numerical.get(f"{fn}(万元)")
                        if not _v or _v <= 0:
                            continue
                        _yj = ""
                        _rr = _raw_rr.get(fn)
                        if isinstance(_rr, dict):
                            _yj = str(_rr.get("依据") or "")
                        _fl.append({"name": fn, "label": fd["label"], "yiju": _yj})
                    for e in payload.get("custom_fees") or []:
                        _nm = str(e.get("名称") or e.get("name") or "").strip()
                        if _nm and not _ZD_RE.search(_nm):
                            _fl.append({"name": "自定义", "label": _nm, "yiju": ""})
                    for _nm, _yj in (("施工图预算编制费", "设计费的10%"),
                                     ("竣工图编制费", "设计费的8%")):
                        if numerical.get(f"{_nm}(万元)"):
                            _fl.append({"name": _nm, "label": _nm, "yiju": _yj})
                    _tot = {
                        "part1": payload["total_part1"],
                        "fees": preview["fee_total_with_custom"],
                        "yubei": preview["yubei_total"],
                        "total": preview["project_total_with_custom"],
                        "zhengchai": preview.get("zhengchai_total") or 0,
                        "yubei_rate": ctx.get("yubei_rate") or 5.0,
                    }
                    if _per_item:
                        _lbls2 = [_item_gl(g, i) for i, g in enumerate(_zy_list, 1)]
                        _bad_lbls = [_lbls2[i] for i, d in enumerate(_per_item_infos)
                                     if not d.get("name")]
                        if _bad_lbls:
                            st.error("请先填写项目名称：" + "、".join(_bad_lbls))
                            _infos = None
                        else:
                            _infos = _per_item_infos
                    else:
                        _info = {"name": (_d_name or "").strip(),
                                 "proj_no": (_d_no or "").strip(),
                                 "date": (_d_date or "").strip(),
                                 "overview": (_d_ov or "").strip(),
                                 "fund_source": (_d_fund or "").strip()}
                        if not _info["name"]:
                            st.error("请先填写项目名称。")
                            _infos = None
                        else:
                            _infos = [_info]
                    if _infos is not None:
                        # 分开模式：每个单项一套成果，信息按单项分别填写
                        # （概算表=该单项工作簿，编制说明投资段/概算表PDF 随之按单项；
                        # 费用段只列费种与计取标准、不含金额，跨单项共用无误）。
                        # totals 只传 zhengchai（按单项问时=该单项自己的
                        # 征地拆迁费）与预备费率，part1/fees/yubei/total 由
                        # build_deliverable 从各单项工作簿读回（= 各单项分别计算口径）。
                        if _per_item:
                            _targets = [(_gl, _infos[_ii]["name"], _bytes)
                                        for _ii, (_gl, _bytes, _fn)
                                        in enumerate(_excel_files)]
                        else:
                            _targets = [(None, _infos[0]["name"],
                                         _excel_files[0][1])]
                        # 源 PDF 与单项匹配（分开导出时）：
                        #  0 份 → 不附；1 份 → 每份成果都附这一份；
                        #  份数==单项数 → 按上传顺序对应（第1份→单项一…）；
                        #  份数>1 且对不上 → 合并后附到每份成果（提示说明）
                        _srcs = st.session_state.get("src_pdf_files") or []
                        if not _srcs and st.session_state.get("pdf_bytes"):
                            _srcs = [("广联达源.pdf", st.session_state["pdf_bytes"])]
                        if len(_srcs) > 1 and len(_srcs) != len(_targets):
                            st.warning(
                                f"源 PDF 共 {len(_srcs)} 份，与成果套数 {len(_targets)} "
                                "不一致：每份成果将附上合并后的全部源 PDF。"
                                "如需按单项分别对应，请按单项顺序上传相同份数的 PDF。")
                            try:
                                import io as _io
                                from pypdf import PdfWriter
                                _w = PdfWriter()
                                for _n2, _b2 in _srcs:
                                    _w.append(_io.BytesIO(_b2))
                                _buf = _io.BytesIO()
                                _w.write(_buf)
                                _srcs = [("广联达源（合并）.pdf", _buf.getvalue())]
                            except Exception as _me:
                                st.error(f"源 PDF 合并失败：{_me}")
                                _srcs = []
                        _out = []
                        with st.spinner("正在生成概算成果压缩包"
                                        "（Word/Excel 转换，每套约 30~60 秒）…"):
                            try:
                                from deliverable_pdf import build_deliverable
                                for _ii, (_gl, _nm, _bytes) in enumerate(_targets):
                                    if _per_item:
                                        # 征地拆迁费按单项：该单项自己的条目
                                        # + 项目级既有费用；未走按单项流程回退整体
                                        _cf_map = ctx.get("item_custom_fees") or {}
                                        if _cf_map:
                                            _zd_i = _zd_amount(
                                                (ctx.get("item_fees_base") or [])
                                                + (_cf_map.get(_gl) or []))
                                        else:
                                            _zd_i = _tot.get("zhengchai") or 0
                                        _tot_i = {
                                            "zhengchai": _zd_i,
                                            "yubei_rate": _tot.get("yubei_rate") or 5.0,
                                        }
                                    else:
                                        _tot_i = _tot
                                    _src_i = None
                                    if _srcs:
                                        _src_i = (_srcs[_ii][1]
                                                  if len(_srcs) == len(_targets)
                                                  else _srcs[0][1])
                                    _zip = build_deliverable(
                                        region=_region,
                                        info=_infos[_ii],
                                        fee_lines=_fl, totals=_tot_i,
                                        excel_bytes=_bytes,
                                        source_pdf_bytes=_src_i)
                                    _out.append({"label": _nm,
                                                 "name": f"{_nm}-概算成果.zip",
                                                 "data": _zip})
                            except Exception as e:
                                st.error(f"成果生成失败：{e}")
                                _out = []
                        if _out:
                            st.session_state.deliverable_pdf = _out
                            st.rerun()
    else:
        # 纯独立费种直算结果
        st.markdown("## 计算结果（程序精确计算）")
        for fn, r in (ctx.get("simple_results") or {}).items():
            if r:
                st.markdown(f"#### {FEE_CATALOG.get(fn, {}).get('label', fn)}")
                _render_engine_card(r)
            else:
                st.warning(f"**{fn}**：计算失败（参数不足）")


def _try_start_fee_task(query: str, *, force_task: bool = False) -> bool:
    """尝试启动清单式任务流。返回 True 表示已接管（调用方立即 rerun）。

    路由规则：
    - fee_calc_intent False（查规则等）→ 落 LLM，返回 False；
      force_task=True（任务进行中「新项目」重启）时跳过意图检查，
      即使新文本只有金额没有费种，也从清单第一步重新收集
    - cascade/iteration/comparison 且金额齐 → 旧直算路径（B3 翻转后进清单机）
    - 其余 → build_checklist_meta → pending_task，返回 True
    """
    if not fee_calc_intent(query) and not force_task:
        return False
    ctx = build_checklist_meta(
        query, region=st.session_state.get("selected_region"),
        force=_TASK_FLOW_CASCADE or force_task,
    )
    if ctx is None:
        return False
    if not _TASK_FLOW_START and not _TASK_FLOW_CASCADE:
        return False
    if st.session_state.get("ai_guide_pref", True):
        ctx["ai_guided"] = True
        ctx["_ai_dirty"] = True
        ctx.setdefault("_ai_rounds", [])
        ctx.setdefault("ai_notes", [])
    st.session_state.pending_task = ctx
    return True


# 初始化近期计算列表（从文件恢复，刷新不丢失，必须在 sidebar 之前）
if "recent_calculations" not in st.session_state:
    st.session_state.recent_calculations = _load_recents()

# ===== 侧边栏 =====
with st.sidebar:
    st.title("🏗️ 造价智能助手")
    st.divider()

    # ── 省份选择 ──
    if "selected_region" not in st.session_state:
        st.session_state.selected_region = DEFAULT_REGION

    selected_region = st.selectbox(
        "📍 所在省份",
        ALL_PROVINCES,
        index=ALL_PROVINCES.index(st.session_state.selected_region)
              if st.session_state.selected_region in ALL_PROVINCES else 1,
        key="region_selector",
    )
    st.session_state.selected_region = selected_region
    st.caption("有特殊政策的省份将自动应用对应费率，其余暂按默认（天津）计算。")

    st.divider()

    st.markdown("### 功能导航")
    # 「智能问答」勾选框 = 主界面选中态（单选互斥中的一项）：
    # 其他功能都未开启时勾选；点它 = 回主界面（退出所有整页模式）并保持勾选
    _main_on = not (st.session_state.get("pdf_import_mode", False)
                    or st.session_state.get("preprocess_mode", False)
                    or st.session_state.get("bridge_match_mode", False)
                    or st.session_state.get("road_bridge_match_mode", False)
                    or st.session_state.get("active_case"))

    def _chat_nav_clicked():
        # 在回调里（控件尚未实例化，改自身 key 合法）：退出所有整页模式，
        # 并保持自身勾选——点选后本运行直接渲染主界面，无需 rerun
        st.session_state.pdf_import_mode = False
        st.session_state.preprocess_mode = False
        st.session_state.bridge_match_mode = False
        st.session_state.road_bridge_match_mode = False
        st.session_state.active_case = None
        st.session_state.chat_mode_cb = True

    if "chat_mode_cb" in st.session_state:
        # 每次运行在实例化前把勾选状态与主界面态同步
        # （勾其他功能后的 rerun 会让它自动取消勾选）
        st.session_state.chat_mode_cb = _main_on
        st.checkbox("💬 智能问答（已上线）", key="chat_mode_cb",
                    on_change=_chat_nav_clicked)
    else:
        # 首次运行：主界面即智能问答，默认勾选
        st.checkbox("💬 智能问答（已上线）", value=True, key="chat_mode_cb",
                    on_change=_chat_nav_clicked)
    # 子项勾选框样式：用 <style> 元素兼作范围标记——此标记之后出现的 checkbox
    # （PDF 导入/表格预处理/桥梁定额匹配）才套子项样式；「智能问答」在其之前，
    # 保持主条目默认样式（正常字号、正常颜色、无竖线缩进）
    st.markdown(
        "<style>"
        '[data-testid="stSidebar"] [data-testid="stElementContainer"]:has(style)'
        ' ~ [data-testid="stElementContainer"] [data-testid="stCheckbox"] {'
        "  border-left: 2px solid rgba(49,51,63,0.18);"
        "  padding-left: 0.9rem;"
        "  margin: 0.12rem 0;"
        "}"
        '[data-testid="stSidebar"] [data-testid="stElementContainer"]:has(style)'
        ' ~ [data-testid="stElementContainer"] [data-testid="stCheckbox"] label > div {'
        "  margin-left: 0.4rem !important;"
        "  font-size: 1rem !important;"
        "  color: rgba(49,51,63,0.62) !important;"
        "}"
        # 「智能问答」勾选框：紧邻本 <style> 标记的容器，恢复主条目字号/行高/颜色，
        # 与「概算助手」等导航标题一致。字号必须打到最内层 p/stWidgetLabel 上——
        # Streamlit 1.58 勾选框标签文本 p 自带 emotion 字号，父级继承会被覆盖
        '[data-testid="stSidebar"] [data-testid="stElementContainer"]:has('
        '[data-testid="stCheckbox"]):has(+ [data-testid="stElementContainer"]:has(style))'
        ' [data-testid="stCheckbox"] label > div,'
        '[data-testid="stSidebar"] [data-testid="stElementContainer"]:has('
        '[data-testid="stCheckbox"]):has(+ [data-testid="stElementContainer"]:has(style))'
        ' [data-testid="stCheckbox"] label [data-testid="stWidgetLabel"],'
        '[data-testid="stSidebar"] [data-testid="stElementContainer"]:has('
        '[data-testid="stCheckbox"]):has(+ [data-testid="stElementContainer"]:has(style))'
        ' [data-testid="stCheckbox"] label p {'
        "  font-size: 1rem !important;"
        "  line-height: 1.6 !important;"
        "  color: rgb(49,51,63) !important;"
        "}"
        "</style>",
        unsafe_allow_html=True)
    # 子项说明：小字号灰色、无项目符号、轻微缩进，与主条目区分层级
    st.markdown(
        '<div style="font-size:0.875rem; line-height:1.9; color:rgba(49,51,63,0.62);'
        ' padding-left:0.9rem; border-left:2px solid rgba(49,51,63,0.18);'
        ' margin:-1rem 0 0.55rem 0;">'
        '查询苗木品种指标<br/>'
        '查询二类费政策文件<br/>'
        '查询定额说明'
        '</div>',
        unsafe_allow_html=True)
    # 渲染完成后做三件事（不依赖 CSS 选择器，直接改元素内联样式；放在说明块之后，
    # 避免零高度 iframe 挤在勾选框与说明块之间撑大间距）：
    # 1) 把「概算助手」等导航标题的实际字号/行高/颜色复制到「智能问答」勾选框文字上；
    # 2) 等整页渲染完（案例按钮在本 iframe 之后才出现，轮询等待），量出
    #    「智能问答→子项说明」与「正反项目案例→案例类型小标题」两个实际间距，
    #    把差值补偿到小标题容器上，使两者完全一致；
    # 3) 量出相邻案例按钮的实际行距，设到子项说明块 line-height 上，
    #    使说明块三行子项之间的间距与案例按钮子项间距一致
    components.html(
        "<script>"
        "(function _fit(_n){"
        "const _d = parent.document;"
        "const _li = _d.querySelector('[data-testid=\"stSidebar\"] li');"
        "const _p = _d.querySelector('[data-testid=\"stSidebar\"] [data-testid=\"stCheckbox\"] label p');"
        "const _desc = _d.querySelector('[data-testid=\"stSidebar\"] [data-testid=\"stMarkdownContainer\"] div[style*=\"0.875rem\"]');"
        "const _caseLi = Array.from(_d.querySelectorAll('[data-testid=\"stSidebar\"] li')).find(function(x){return x.textContent.indexOf('正反项目案例')>=0;});"
        "const _head = _d.querySelector('[data-testid=\"stSidebar\"] [data-testid=\"stMarkdownContainer\"] div[style*=\"--case-typetag\"]');"
        "const _btn = _d.querySelector('[data-testid=\"stSidebar\"] [data-testid^=\"stBaseButton-primary\"]');"
        "if (!(_li && _p && _desc && _caseLi && _head && _btn)) { if (_n < 50) setTimeout(function(){_fit(_n+1);}, 100); return; }"
        "const _cs = getComputedStyle(_li);"
        "for (const _k of ['fontSize','lineHeight','color','fontFamily','fontWeight']) { _p.style[_k] = _cs[_k]; }"
        "const _r = function(e){ return e.getBoundingClientRect(); };"
        "const _btns = _d.querySelectorAll('[data-testid=\"stSidebar\"] [data-testid^=\"stBaseButton-primary\"]');"
        "if (_btns.length >= 2) {"
        "  const _bgap = _r(_btns[1].querySelector('p')).top - _r(_btns[0].querySelector('p')).top;"
        "  _desc.style.lineHeight = _bgap + 'px';"
        "}"
        "const _gapA = _r(_desc).top - _r(_p).bottom;"
        "const _gapB = _r(_head).top - _r(_caseLi).bottom;"
        "const _c = _head.closest('[data-testid=\"stElementContainer\"]');"
        "_c.style.marginTop = (parseFloat(getComputedStyle(_c).marginTop) + (_gapA - _gapB)) + 'px';"
        "})(0);"
        "</script>",
        height=0)
    st.markdown("- 概算助手（已上线）")
    # 三个功能勾选框带 key，勾选状态由模式标志唯一决定（_mode_cb 只作显示镜像，
    # 每次实例化前从模式标志同步）。互斥在 on_change 回调里做。
    # 注意不能用 keyless 勾选框：其自身状态跨运行残留，会把回调复位的模式
    # 又写回 True（曾导致「从功能页回不到智能问答」——点智能问答后 PDF 页纹丝不动）
    def _pdf_nav_clicked():
        _on = st.session_state.pdf_mode_cb          # 交互后的新值
        st.session_state.pdf_import_mode = _on
        if _on:
            st.session_state.preprocess_mode = False
            st.session_state.bridge_match_mode = False
            st.session_state.road_bridge_match_mode = False
            st.session_state.active_case = None

    def _pre_nav_clicked():
        _on = st.session_state.pre_mode_cb
        st.session_state.preprocess_mode = _on
        if _on:
            st.session_state.pdf_import_mode = False
            st.session_state.bridge_match_mode = False
            st.session_state.road_bridge_match_mode = False
            st.session_state.active_case = None

    def _bm_nav_clicked():
        _on = st.session_state.bm_mode_cb
        st.session_state.bridge_match_mode = _on
        if _on:
            st.session_state.pdf_import_mode = False
            st.session_state.preprocess_mode = False
            st.session_state.road_bridge_match_mode = False
            st.session_state.active_case = None

    def _rbm_nav_clicked():
        _on = st.session_state.rbm_mode_cb
        st.session_state.road_bridge_match_mode = _on
        if _on:
            st.session_state.pdf_import_mode = False
            st.session_state.preprocess_mode = False
            st.session_state.bridge_match_mode = False
            st.session_state.active_case = None

    if "pdf_mode_cb" in st.session_state:
        st.session_state.pdf_mode_cb = st.session_state.get("pdf_import_mode", False)
    if "pre_mode_cb" in st.session_state:
        st.session_state.pre_mode_cb = st.session_state.get("preprocess_mode", False)
    if "bm_mode_cb" in st.session_state:
        st.session_state.bm_mode_cb = st.session_state.get("bridge_match_mode", False)
    if "rbm_mode_cb" in st.session_state:
        st.session_state.rbm_mode_cb = st.session_state.get("road_bridge_match_mode", False)
    st.checkbox("📄 广联达报表导入", key="pdf_mode_cb",
                on_change=_pdf_nav_clicked)
    st.markdown("- 表格预处理（已上线）")
    st.checkbox("📋 表格预处理", key="pre_mode_cb",
                on_change=_pre_nav_clicked)
    st.markdown("- 桥梁定额匹配（已上线）")
    st.checkbox("🌉 市政桥梁定额匹配", key="bm_mode_cb",
                on_change=_bm_nav_clicked)
    st.checkbox("🛣 公路桥梁定额匹配", key="rbm_mode_cb",
                on_change=_rbm_nav_clicked)
    st.markdown("- 指标对比分析（开发中）")
    st.markdown("- 正反项目案例（已上线）")
    # 侧边栏按钮文字左对齐（1.58 版按钮标签有内层 div/span flex 居中包装，需一并覆盖）
    st.markdown(
        "<style>"
        '[data-testid="stSidebar"] [data-testid="stButton"] button,'
        '[data-testid="stSidebar"] [data-testid="stButton"] button > div,'
        '[data-testid="stSidebar"] [data-testid="stButton"] button > div > span {'
        "  justify-content: flex-start;"
        "}"
        '[data-testid="stSidebar"] [data-testid="stButton"] button p {'
        "  text-align: left;"
        "  width: 100%;"
        "}"
        # 案例子项按钮：仿智能问答子项格式（小灰字+左竖线），保留可点击
        '[data-testid="stSidebar"] [data-testid="stButton"] '
        '[data-testid^="stBaseButton-primary"],'
        '[data-testid="stSidebar"] [data-testid="stButton"] button[kind="primary"] {'
        "  background: transparent !important;"
        "  color: rgba(49,51,63,0.62) !important;"
        "  border: none !important;"
        "  border-left: 2px solid rgba(49,51,63,0.18) !important;"
        "  border-radius: 0 !important;"
        "  box-shadow: none !important;"
        "  font-size: 1rem !important;"
        "  line-height: 1.9 !important;"
        "  min-height: 0 !important;"
        "  height: auto !important;"
        "  padding: 0 0 0 0.9rem !important;"
        "  justify-content: flex-start;"
        "  text-align: left;"
        "}"
        '[data-testid="stSidebar"] [data-testid="stButton"] '
        '[data-testid^="stBaseButton-primary"]:hover,'
        '[data-testid="stSidebar"] [data-testid="stButton"] button[kind="primary"]:hover {'
        "  color: rgb(255,75,75) !important;"
        "  background: rgba(49,51,63,0.05) !important;"
        "  border-left: 2px solid rgb(255,75,75) !important;"
        "}"
        '[data-testid="stSidebar"] [data-testid="stButton"] '
        '[data-testid^="stBaseButton-primary"] p {'
        "  text-align: left;"
        "  width: 100%;"
        "}"
        # 案例子项按钮容器：收紧纵向间距，贴近竖线分组效果
        '[data-testid="stSidebar"] [data-testid="stButton"]:has('
        '[data-testid^="stBaseButton-primary"]) {'
        "  margin: 0;"
        "}"
        # 案例按钮所在元素容器：抵消竖排块 16px 间距的一半，缩小项目名称之间的空隙
        '[data-testid="stSidebar"] [data-testid="stElementContainer"]:has('
        '[data-testid^="stBaseButton-primary"]) {'
        "  margin-top: -0.5rem;"
        "}"
        # 案例类型小标题（div[style*="--case-typetag"]）后的首个案例按钮：
        # 不再上移（覆写上方案例按钮容器 -0.5rem 规则），与小标题保持正常间距
        '[data-testid="stSidebar"] [data-testid="stElementContainer"]:has('
        'div[style*="--case-typetag"])'
        ' + [data-testid="stElementContainer"]:has('
        '[data-testid^="stBaseButton-primary"]) {'
        "  margin-top: 0;"
        "}"
        "</style>",
        unsafe_allow_html=True)
    def _case_nav_button(_case):
        if st.button(_case["short"], key=f"case_nav_{_case['key']}",
                     use_container_width=True, type="primary"):
            # 进入案例页时退出其他整页模式
            st.session_state.pdf_import_mode = False
            st.session_state.preprocess_mode = False
            st.session_state.bridge_match_mode = False
            st.session_state.active_case = _case["key"]
            st.rerun()

    # 案例子项按钮：按案例类型分组，组前加「正面案例/反面案例」小标题；
    # 类型从各案例 md 基本信息表读取（未标注类型的案例不分组、照旧排按钮）
    _case_groups, _case_ungrouped = {"正面": [], "反面": []}, []
    for _case in CASE_ITEMS:
        try:
            _t = _case_type_of(_parse_case_md(CASE_BASE_DIR / _case["file"]))
        except Exception:
            _t = ""
        if _t in _case_groups:
            _case_groups[_t].append(_case)
        else:
            _case_ungrouped.append(_case)
    for _group in ("正面", "反面"):
        if not _case_groups[_group]:
            continue
        _emoji = "📗" if _group == "正面" else "📕"
        st.markdown(
            f'<div style="--case-typetag:1;font-size:0.875rem;line-height:1.9;'
            f'color:rgba(49,51,63,0.62);padding-left:0.9rem;'
            f'border-left:2px solid rgba(49,51,63,0.18);'
            f'margin:-1rem 0 0.55rem 0;">{_emoji} {_group}案例</div>',
            unsafe_allow_html=True)
        for _case in _case_groups[_group]:
            _case_nav_button(_case)
    for _case in _case_ungrouped:
        _case_nav_button(_case)

    st.divider()

    st.markdown("### 数据状态")
    try:
        engine = get_engine()
        total = len(engine.chunks)
        cats = len(engine.data)
        kb_count = len(engine.knowledge_chunks)
        quota_count = len(engine.quota_chunks)
        case_count = len(engine.case_chunks)
        st.success(f"已加载 {cats} 个类别，共 {total} 条记录")
        st.caption(f"知识库：{kb_count} 个片段（政策文件 + 定额说明 {quota_count} 段"
                   f" + 项目案例 {case_count} 段）")
    except Exception as e:
        st.error(f"数据加载失败：{e}")

    st.divider()

    st.markdown("### 绿化指标查询")
    examples_green = [
        "白皮松高度3.5米的综合指标是多少？",
        "落叶乔木胸径14cm的有哪些品种？",
        "常绿乔木和落叶乔木的综合指标对比",
        "灌木球类中综合指标最低的是哪个？",
        "银杏的综合指标是多少？",
    ]
    for i, ex in enumerate(examples_green):
        if st.button(ex, use_container_width=True, key=f"green_{i}"):
            st.session_state.current_query = ex

    st.divider()

    st.markdown("### 二类费计算")
    examples_multi = [
        "建安费131万，设备费160万，桥梁工程，帮我算全部费用",
        "建安费8000万，工程总概算迭代计算",
        "建安费5000万，设备费3000万，方案比选",
    ]
    for i, ex in enumerate(examples_multi):
        if st.button(ex, use_container_width=True, key=f"multi_{i}"):
            st.session_state.current_query = ex

    # ── 近期计算 ──
    if st.session_state.get("recent_calculations"):
        st.divider()
        st.markdown("### 🕐 近期计算")
        recents = st.session_state.recent_calculations
        for i, rec in enumerate(reversed(recents)):
            label = f"{rec['timestamp']} · {rec['summary'][:28]}"
            if st.button(
                label,
                use_container_width=True,
                key=f"recent_{i}",
                help=rec.get("query", ""),
            ):
                st.session_state.current_query = rec["query"]
                st.rerun()

    st.divider()

    # 持久化调试面板
    if "debug_info" in st.session_state:
        st.markdown("### 🔍 引擎调试")
        d = st.session_state.debug_info
        if "error" in d:
            st.error(f"异常: {d['error']}")
        else:
            st.write(f"prompt: `{d.get('prompt','')}`")
            st.write(f"fee_type: `{d['fee_type']}`")
            st.write(f"has_amount: `{d['has_amount']}`")
            st.write(f"费种: {d.get('费种','')}")
            st.write(f"结果: {d.get('结果','')}")
        if st.button("清除调试", key="clear_debug"):
            del st.session_state.debug_info
            st.rerun()

    st.divider()
    st.caption("Powered by DeepSeek v4")

# ===== 广联达 PDF 导入模式：整页替换主界面 =====
if st.session_state.get("pdf_import_mode"):
    _render_pdf_import_page()
    st.stop()

# ===== 表格预处理模式：整页替换主界面 =====
if st.session_state.get("preprocess_mode"):
    _render_preprocess_page()
    st.stop()

# ===== 桥梁定额匹配模式：整页替换主界面 =====
if st.session_state.get("bridge_match_mode"):
    _render_bridge_match_page()
    st.stop()

# ===== 公路桥梁定额匹配模式：整页替换主界面 =====
if st.session_state.get("road_bridge_match_mode"):
    _render_road_bridge_page()
    st.stop()

# ===== 正反项目案例：点击侧边栏子项整页替换 =====
if st.session_state.get("active_case"):
    _render_case_page(st.session_state.active_case)
    st.stop()

# ===== 主界面 =====
_page_header("🏗️ 工程造价智能问答", "main_refresh",
             ["messages", "current_query", "pending_comparison",
              "pending_task", "zhengchai_separate", "deliverable_pdf"],
             refresh_prefixes=["ck_", "ai_", "deliv_"], keep_keys=["ai_guide_pref"])
st.caption("基于工程造价指标数据库，提供专业造价问答服务 | v2026-07-03")

# 初始化引擎
with st.spinner("正在加载数据库和 AI 模型..."):
    engine = get_engine()

# 初始化聊天历史
if "messages" not in st.session_state:
    st.session_state.messages = [
        {
            "role": "assistant",
            "content": (
                "你好！我是造价智能助手。\n\n"
                "我可以回答以下问题：\n"
                "- 🌱 查询具体苗木品种的综合指标、栽植费用\n"
                "- 📊 对比不同规格、不同品种的造价差异\n"
                "- 🏗️ 计算工程建设二类费（建设管理费、设计费、监理费等）\n"
                "- 📋 查询二类费政策文件和费率表\n\n"
                "请在下方输入你的问题。"
            ),
        }
    ]

# 显示聊天历史
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# ===== 方案比选结果持久化渲染（rerun 后重新绘制图表） =====
if "pending_comparison" in st.session_state:
    _render_comparison_result(st.session_state.pending_comparison)

# ===== 清单式任务流渲染（pending_task 气泡 + result_payload 重绘） =====
_render_task_area()

# ===== 输入框 =====
st.divider()
st.markdown("### 输入你的问题")

if "current_query" in st.session_state:
    prompt = st.session_state.current_query
    del st.session_state.current_query
else:
    prompt = st.chat_input("请输入你的造价问题，例如：白皮松高度3.5米多少钱？")

if prompt:
    # === 清单式任务流：进行中的任务 → 自由文本推进 ===
    pending_task = st.session_state.get("pending_task")
    _user_msg_appended = False
    _force_restart = False
    if pending_task and pending_task.get("phase") == "ask":
        _user_msg_appended = True
        st.session_state.messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)
        _restart_q = detect_new_task_intent(pending_task, prompt)
        if _restart_q:
            # 中途的新问题 → 结束旧任务，按新查询重新启动（裸「重新算」用原查询）
            # 新文本可能只有金额没有费种（如「帮我算一个新的项目，建安费8000万」），
            # 强制从清单第一步重新收集，不因 fee_calc_intent 为 False 落 LLM
            _force_restart = True
            _typed = prompt
            prompt = _restart_q
            st.session_state.pop("pending_task", None)
            st.session_state.messages.append({
                "role": "assistant",
                "content": ("🔄 好的，按原问题重新开始计算。" if _restart_q != _typed
                            else "🔄 检测到新的计算问题，已结束上一个任务，重新开始。"),
            })
            with st.chat_message("assistant"):
                st.markdown(st.session_state.messages[-1]["content"])
        else:
            # AI 模式裸答登记：当前唯一可见问题是二选一键且输入是「有/没有」等
            # 短答 → 直接按按钮路径登记，防止同一问题下一轮被重复提问
            _bare = _ai_map_bare_answer(pending_task, prompt)
            if _bare is not None:
                _b_key, _b_val, _b_qtxt = _bare
                _ai_write(pending_task, _b_key, _b_val)
                _ai_mark_done(pending_task, _b_key, _b_qtxt, prompt)
            res = advance_checklist(
                pending_task, prompt, region=st.session_state.get("selected_region"))
            st.session_state.pending_task = pending_task
            if res.get("recognized"):
                if pending_task.get("ai_guided"):
                    pending_task["_ai_dirty"] = True
                    pending_task.setdefault("_ai_rounds", [])
                    pending_task["_ai_rounds"].append(
                        {"think": "", "qa": [{"text": "（自由输入）", "answer": prompt}]})
                st.rerun()
            if res.get("fee_domain_token"):
                # 有费种/金额关键词但没解析出新信息 → 留在任务中提示
                st.session_state.messages.append({
                    "role": "assistant",
                    "content": "🤔 没有从这句话中识别出新的信息。"
                               "请回复当前步骤所需的内容，或点击「🗑 结束本次计算」重新提问。",
                })
                st.rerun()
            if pending_task.get("ai_guided"):
                # AI 模式：无关输入不杀任务，记为补充说明供 AI 参考并重新规划
                pending_task.setdefault("ai_notes", []).append(prompt)
                pending_task["_ai_dirty"] = True
                st.rerun()
            # 完全无关输入 → 结束任务，落普通问答（消息已记录，继续走下方流程）
            st.session_state.pop("pending_task", None)

    # 新提问时清除旧的待处理选择
    st.session_state.pop("pending_comparison", None)
    st.session_state.pop("pending_task", None)  # 新提问清掉旧任务/旧结果
    # 添加用户消息
    if not _user_msg_appended:
        st.session_state.messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)

    # === 清单式任务流接管（缺金额/缺费种等旧路径会丢的场景）===
    if _try_start_fee_task(prompt, force_task=_force_restart):
        st.rerun()

    # 生成回答
    with st.chat_message("assistant"):
        try:
            fee_result = detect_and_calculate(prompt, region=st.session_state.get("selected_region"))
        except Exception as e:
            import traceback
            st.code(traceback.format_exc())
            fee_result = None

        with st.spinner("正在检索数据并生成回答..."):
            # 二类费规则引擎
            # fee_result = detect_and_calculate(prompt)  # 已在上方调用

            if fee_result and fee_result.get("has_amount"):
                # 记录到近期计算（级联模式在确认时记录，此处跳过）
                if fee_result.get("mode") != "cascade":
                    _record_calculation(prompt, fee_result)

                # === 多费种迭代模式路由 ===
                mode = fee_result.get("mode")
                if mode == "iteration":
                    response = _render_iteration_result(fee_result)
                elif mode == "comparison":
                    response = _render_comparison_result(fee_result)
                    # 存储结果以便 rerun 后重新渲染交互式图表
                    st.session_state.pending_comparison = fee_result

                if mode is None:
                    # === 兜底直出：面板与对话机已退役，任何遗留费种直接渲染 ===
                    st.markdown("### 计算结果（程序精确计算）")
                    _render_engine_card(fee_result)
                    fee_name = fee_result.get("费种", "")
                    result_val = fee_result.get("结果(万元)") or fee_result.get("结果(元)")
                    unit = "万元" if "结果(万元)" in fee_result else "元"
                    response = (
                        f"## {fee_name}\n\n"
                        f"**结果**：{result_val} {unit}\n\n"
                        f"{fee_result.get('说明', '')}"
                    )
            elif fee_result and not fee_result.get("has_amount"):
                # === 无金额参考模式 ===
                is_sheji = (fee_result.get("fee_type") == "工程设计费")

                if is_sheji:
                    import pandas as pd

                    st.markdown("### 附表一：工程设计收费基价表")
                    st.caption("依据：《工程勘察设计收费管理规定》（计价格[2002]10号）单位：万元")
                    rate_table = fee_result.get("费率表", [])
                    if rate_table:
                        header = rate_table[0]
                        rows = rate_table[1:]
                        df_rate = pd.DataFrame(rows, columns=header)
                        st.table(df_rate)
                        st.caption("注：计费额 > 2000000 万元的，以计费额乘以 1.6% 的收费率计算收费基价。")

                    st.markdown("---")
                    st.markdown("### 附表二：工程设计收费专业调整系数表")
                    sheji_table_data = [
                        ("1、矿山采选工程", "黑色、黄金、化学、非金属及其他矿采选工程", "1.1"),
                        ("", "采煤工程，有色、铀矿采选工程", "1.2"),
                        ("", "选煤及其他煤炭工程", "1.3"),
                        ("2、加工冶炼工程", "各类冷加工工程", "1"),
                        ("", "船舶水工工程", "1.1"),
                        ("", "各类冶炼、热加工、压力加工工程", "1.2"),
                        ("", "核加工工程", "1.3"),
                        ("3、石油化工工程", "石油、化工、石化、化纤、医药工程", "1.2"),
                        ("", "核化工工程", "1.6"),
                        ("4、水利电力工程", "风力发电、其他水利工程", "0.8"),
                        ("", "火电工程", "1"),
                        ("", "核电常规岛、水电、水库、送变电工程", "1.2"),
                        ("", "核能工程", "1.6"),
                        ("5、交通运输工程", "机场场道工程", "0.8"),
                        ("", "公路、城市道路工程", "0.9"),
                        ("", "机场空管和助航灯光、轻轨工程", "1"),
                        ("", "水运、地铁、桥梁、隧道工程", "1.1"),
                        ("", "索道工程", "1.3"),
                        ("6、建筑市政工程", "邮政工艺工程", "0.8"),
                        ("", "建筑、市政、电信工程", "1"),
                        ("", "人防、园林绿化、广电工艺工程", "1.1"),
                        ("7、农业林业工程", "农业工程", "0.9"),
                        ("", "林业工程", "0.8"),
                    ]
                    df_coef = pd.DataFrame(sheji_table_data, columns=["工程类型", "具体专业", "专业调整系数"])
                    st.table(df_coef)
                    st.info(
                        "工程设计费调整系数共三个：\n\n"
                        "1. **专业调整系数**（上表 附表二）\n"
                        "2. **工程复杂程度调整系数**：I级 0.85 / II级 1.0 / III级 1.15\n"
                        "3. **附加调整系数**：多个时合并计算（相加 − 个数 + 1）\n\n"
                        "计算公式：基本设计收费 = 收费基价（附表一） × 专业调整系数（附表二） × 复杂程度系数 × 附加调整系数\n\n"
                        "⚠️ 工程设计费**不含**高程调整系数（高程系数仅用于监理费 发改价格[2007]670号）"
                    )
                    st.divider()
                    st.caption("以下为 AI 补充说明（仅作解释性描述，数字以上表为准）：")

                is_kancha = (fee_result.get("fee_type") == "勘察费")

                if is_kancha:
                    # 勘察费（无金额时）：展示计算方法和费率说明
                    import pandas as pd
                    st.markdown("### 工程勘察费 — 计算方法")
                    st.info(
                        "**工程勘察费**依据《工程勘察设计收费管理规定》（计价格[2002]10号）"
                        "的 **工程勘察收费标准** 部分计算。\n\n"
                        "⚠️ **与工程设计费的重要区别**：工程勘察费按**实物工作量**定额计费，"
                        "不是按投资额比例。\n\n"
                        "**精确计算公式**：\n"
                        "- 工程勘察收费 = 工程勘察收费基准价 × (1 ± 20%)\n"
                        "- 工程勘察收费基准价 = 实物工作收费 + 技术工作收费\n"
                        "- 实物工作收费 = 收费基价 × 实物工作量 × 附加调整系数\n"
                        "- 技术工作收费 = 实物工作收费 × 技术工作收费比例\n\n"
                        "**粗略估算方法**（《市政工程设计概算编制办法》，中国计划出版社）：\n"
                        "- 通用项目：第一部分工程费 × **0.8%~1.1%**\n"
                        "- 建筑项目：第一部分工程费 × **0.3%~0.5%**\n\n"
                        "💡 提供建安费和设备费金额，程序可按上述百分比法粗略估算。"
                        "精确计算请提供勘察类型和实物工作量。"
                    )

                    jianan_detected = fee_result.get("检测到建安费(万元)")
                    shebei_detected = fee_result.get("检测到设备费(万元)")
                    amt_detected = fee_result.get("检测到金额(万元)")
                    if jianan_detected is not None:
                        st.metric("建安工程费", f"{jianan_detected} 万元")
                        if shebei_detected is not None:
                            st.metric("设备购置费", f"{shebei_detected} 万元")
                    elif amt_detected is not None:
                        st.metric("检测到金额", f"{amt_detected} 万元")

                    st.divider()
                    st.markdown("**需明确的参数（精确计算）**：")
                    st.markdown(
                        "1. 勘察类型（工程测量/岩土工程勘察/水文地质勘察/工程物探等 16 大类）\n"
                        "2. 实物工作量（钻孔深度、测量面积/比例尺、取样数量等）\n"
                        "3. 复杂程度等级（简单/中等/复杂）\n"
                        "4. 附加调整系数（气温/高程/带状/水域等）"
                    )
                    st.caption("详细费率表见知识库《计价格[2002]10号》工程勘察收费标准章节。")

                    # 构建响应文本
                    response = (
                        f"工程勘察费依据《工程勘察设计收费管理规定》（计价格[2002]10号）"
                        f"的工程勘察收费标准计算。\n\n"
                        f"⚠️ 与工程设计费不同，勘察费按**实物工作量**定额计费"
                        f"（如钻探米数、测量面积等），不是按投资额比例。\n\n"
                        f"**粗略估算**（《市政工程设计概算编制办法》）："
                        f"通用 0.8%~1.1%，建筑 0.3%~0.5%。\n"
                        f"**精确计算**需提供勘察类型（16大类）、实物工作量、复杂程度、附加调整系数。\n\n"
                        f"详细费率表见知识库《计价格[2002]10号》工程勘察收费标准章节。"
                    )

                else:
                    history = [
                        {"role": m["role"], "content": m["content"]}
                        for m in st.session_state.messages[:-1]
                        if m["role"] in ("user", "assistant")
                    ]
                    response = engine.chat(prompt, history)
                    st.markdown(response)

                with st.expander("查看计费依据"):
                    st.markdown(f"**{fee_result.get('费种', '')}**")
                    st.markdown(f"<small>依据：{_basis_with_links(fee_result.get('依据', ''))}</small>", unsafe_allow_html=True)
                    st.caption(f"计费方式：{fee_result.get('计费方式', '')}")
                    rate_table = fee_result.get("费率表", [])
                    if rate_table:
                        import pandas as pd
                        header = rate_table[0]
                        rows = rate_table[1:]
                        df = pd.DataFrame(rows, columns=header)
                        st.table(df)
                    auto_coefs = fee_result.get("auto_detected_coefs", {})
                    if auto_coefs:
                        st.markdown("**引擎自动检测的系数**：")
                        for k, v in auto_coefs.items():
                            st.markdown(f"- {k}：**{v}**")
                    steps = fee_result.get("计算步骤")
                    if steps:
                        st.markdown("**分档计算明细**：")
                        for s in steps:
                            st.markdown(
                                f"- {s.get('区间', '')}：{s.get('金额(万元)', '')}万元 "
                                f"× {s.get('费率(%)', '')}% = **{s.get('费用(万元)', '')}万元**"
                            )
                    adjustment = fee_result.get("计费额调整")
                    if adjustment and adjustment.get("触发调整"):
                        st.info(adjustment.get("说明", ""))
                    if "分摊" in fee_result:
                        st.caption(fee_result["分摊"])

                    results = engine.search(prompt, top_k=5)
                    if results:
                        st.divider()
                        st.markdown("**数据库匹配**：")
                        for r in results:
                            u = r.get("unit", "元/株")
                            st.markdown(
                                f"- [{r['category']}] {r['name']}（{r['spec']}）"
                                f" -> 综合指标 **{r['comprehensive']}**{u}"
                            )

            else:
                # === 正常 LLM 对话 ===
                history = [
                    {"role": m["role"], "content": m["content"]}
                    for m in st.session_state.messages[:-1]
                    if m["role"] in ("user", "assistant")
                ]
                response = engine.chat(prompt, history)
                st.markdown(response)

                with st.expander("查看检索到的数据"):
                    results = engine.search(prompt, top_k=5)
                    if results:
                        for r in results:
                            u = r.get("unit", "元/株")
                            st.markdown(
                                f"- [{r['category']}] {r['name']}（{r['spec']}）"
                                f" -> 综合指标 **{r['comprehensive']}**{u}"
                            )
                    else:
                        st.caption("未匹配到相关数据")

    st.session_state.messages.append({"role": "assistant", "content": response})
    st.rerun()