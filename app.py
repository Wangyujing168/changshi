"""
工程造价智能助手 - Streamlit 网页界面（Phase 1: 问答功能）
"""
import re
import streamlit as st
from rag_engine import get_engine
from fee_engine import (
    detect_and_calculate, is_hebei_region,
    ALL_PROVINCES, DEFAULT_REGION,
    FEE_PARAM_SPECS,
    FEE_CATALOG, CHECKLIST_STAGES,
    fee_calc_intent, build_checklist_meta, advance_checklist,
    current_stage, param_cards_for, settle_checklist,
    linear_rate_options, _build_shencha_rate_options, fee_catalog_for,
    calc_discount_scenarios,
    _get_coef_config_simple,
)

# ===== 清单式任务流（checklist task flow）双轨开关 =====
# _TASK_FLOW_START：接管「缺金额/缺费种」场景（旧路径会丢的场景）
# _TASK_FLOW_CASCADE：金额齐全的 cascade 也进清单机（B3 已翻转，面板变死代码）
_TASK_FLOW_START = True
_TASK_FLOW_CASCADE = True

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
    """根据级联计算结果生成 Excel 文件，返回 bytes 供下载。"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill, numbers
    from openpyxl.utils import get_column_letter
    import datetime

    wb = Workbook()
    ws = wb.active
    ws.title = "费用汇总"

    # ── 样式定义 ──
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"))
    header_font = Font(name="微软雅黑", bold=True, size=11)
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font_w = Font(name="微软雅黑", bold=True, size=11, color="FFFFFF")
    title_font = Font(name="微软雅黑", bold=True, size=14)
    subtotal_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
    normal_font = Font(name="微软雅黑", size=10)
    bold_font = Font(name="微软雅黑", bold=True, size=10)
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")
    money_fmt = '#,##0.00'

    preview = ctx.get("preview", {})
    numerical = preview.get("numerical", {}) if preview else {}
    fee_defs = ctx.get("fee_defs", [])
    selected = ctx.get("selected_fees", set())
    custom_fees = ctx.get("custom_fees", [])
    fee_discounts = ctx.get("fee_discounts", {})

    # ── 列宽预设 ──
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 18
    ws.column_dimensions["F"].width = 28

    row = 1
    # ── 标题 ──
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
    ws.cell(row=row, column=1, value="建设项目二类费计算汇总表").font = title_font
    ws.cell(row=row, column=1).alignment = center_align
    row += 1

    # ── 项目基本信息 ──
    info_data = [
        ("建安工程费", f"{ctx.get('jianan', 0):.2f} 万元"),
        ("设备购置费", f"{ctx.get('shebei', 0):.2f} 万元"),
        ("第一部分工程费", f"{ctx.get('total_part1', 0):.2f} 万元"),
        ("项目类型", ctx.get("project_type", "")),
        ("计算日期", datetime.date.today().isoformat()),
    ]
    for label, val in info_data:
        ws.cell(row=row, column=1, value=label).font = bold_font
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
        ws.cell(row=row, column=2, value=val).font = normal_font
        row += 1
    row += 1

    # ── 表头 ──
    headers = ["序号", "费用名称", "费用（万元）", "打折后（万元）", "备注", "依据"]
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col_idx, value=h)
        cell.font = header_font_w
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = thin_border
    row += 1

    # ── 按层级输出各费种 ──
    seq = 0
    raw_total = 0.0
    discounted_total = 0.0
    tier_names = {0: "第一部分工程费相关", 1: "勘察设计费相关", 2: "总投资相关"}

    for tier in [0, 1, 2]:
        tier_fees = sorted(
            [fd for fd in fee_defs if fd["tier"] == tier and fd["name"] in selected],
            key=lambda fd: fd["name"])
        if not tier_fees:
            continue
        # 层级小标题
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        tier_cell = ws.cell(row=row, column=1, value=tier_names.get(tier, f"Tier {tier}"))
        tier_cell.font = Font(name="微软雅黑", bold=True, size=10, color="4472C4")
        tier_cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
        row += 1

        for fd in tier_fees:
            fn = fd["name"]
            val = numerical.get(f"{fn}(万元)")
            if val is None or val <= 0:
                continue
            seq += 1
            disc = fee_discounts.get(fn, 1.0)
            # 数值已含折扣（T0 由引擎应用、非 T0 结算侧补乘），避免二次打折
            disc_val = val
            raw_total += val
            discounted_total += disc_val

            # 备注
            notes = []
            if fn in ctx.get("coef_overrides", {}):
                for k, v in ctx["coef_overrides"][fn].items():
                    if abs(v - 1.0) > 0.005:
                        notes.append(f"{k}={v}")
            if fn in ctx.get("rate_overrides", {}):
                notes.append(f"费率={ctx['rate_overrides'][fn]}")
            if fn in ctx.get("service_selections", {}):
                svcs = ctx["service_selections"][fn]
                notes.append(f"{'、'.join(svcs)}")
            if abs(disc - 1.0) >= 0.005:
                notes.append(f"打折={disc:.2f}")
            note_str = "；".join(notes) if notes else ""
            display_val = disc_val if abs(disc - 1.0) >= 0.005 else val

            ws.cell(row=row, column=1, value=seq).font = normal_font
            ws.cell(row=row, column=1).alignment = center_align
            ws.cell(row=row, column=2, value=fd["label"]).font = normal_font
            ws.cell(row=row, column=3, value=val).font = normal_font
            ws.cell(row=row, column=3).number_format = money_fmt
            ws.cell(row=row, column=3).alignment = right_align
            ws.cell(row=row, column=4, value=display_val).font = normal_font
            ws.cell(row=row, column=4).number_format = money_fmt
            ws.cell(row=row, column=4).alignment = right_align
            ws.cell(row=row, column=5, value=note_str).font = Font(name="微软雅黑", size=9)
            ws.cell(row=row, column=6, value=fd.get("依据", "")).font = Font(name="微软雅黑", size=9)
            for c in range(1, 7):
                ws.cell(row=row, column=c).border = thin_border
            row += 1

    # ── 自定义费用 ──
    if custom_fees:
        for cf in custom_fees:
            seq += 1
            cf_amount = cf["amount_wan"]
            raw_total += cf_amount
            discounted_total += cf_amount
            ws.cell(row=row, column=1, value=seq).font = normal_font
            ws.cell(row=row, column=1).alignment = center_align
            ws.cell(row=row, column=2, value=f"【自定义】{cf['name']}").font = normal_font
            ws.cell(row=row, column=3, value=cf_amount).font = normal_font
            ws.cell(row=row, column=3).number_format = money_fmt
            ws.cell(row=row, column=3).alignment = right_align
            ws.cell(row=row, column=4, value=cf_amount).font = normal_font
            ws.cell(row=row, column=4).number_format = money_fmt
            ws.cell(row=row, column=4).alignment = right_align
            ws.cell(row=row, column=5, value="自定义费用，不打折").font = Font(name="微软雅黑", size=9)
            for c in range(1, 7):
                ws.cell(row=row, column=c).border = thin_border
            row += 1

    # ── 二类费合计 ──
    for c in range(1, 7):
        ws.cell(row=row, column=c).fill = subtotal_fill
        ws.cell(row=row, column=c).border = thin_border
    ws.cell(row=row, column=2, value="二类费合计").font = bold_font
    ws.cell(row=row, column=3, value=round(raw_total, 4)).font = bold_font
    ws.cell(row=row, column=3).number_format = money_fmt
    ws.cell(row=row, column=3).alignment = right_align
    display_disc_total = round(discounted_total, 4)
    ws.cell(row=row, column=4, value=display_disc_total).font = bold_font
    ws.cell(row=row, column=4).number_format = money_fmt
    ws.cell(row=row, column=4).alignment = right_align
    row += 1

    # ── 预备费 ──
    yb_val = preview.get("yubei_total", 0) if preview else 0
    if yb_val > 0:
        for c in range(1, 7):
            ws.cell(row=row, column=c).border = thin_border
        ws.cell(row=row, column=2, value="预备费（基本预备费）").font = bold_font
        ws.cell(row=row, column=3, value=round(yb_val, 4)).font = bold_font
        ws.cell(row=row, column=3).number_format = money_fmt
        ws.cell(row=row, column=3).alignment = right_align
        ws.cell(row=row, column=4, value=round(yb_val, 4)).font = bold_font
        ws.cell(row=row, column=4).number_format = money_fmt
        ws.cell(row=row, column=4).alignment = right_align
        row += 1

    # ── 项目总投资 ──
    project_total = preview.get("project_total_with_custom", 0) if preview else 0
    for c in range(1, 7):
        ws.cell(row=row, column=c).fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        ws.cell(row=row, column=c).border = thin_border
    ws.cell(row=row, column=2, value="项目总投资").font = Font(name="微软雅黑", bold=True, size=11)
    ws.cell(row=row, column=3, value=round(project_total, 4)).font = Font(name="微软雅黑", bold=True, size=11)
    ws.cell(row=row, column=3).number_format = money_fmt
    ws.cell(row=row, column=3).alignment = right_align

    # ── 保存到内存 ──
    import io
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()




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
                    st.markdown(
                        f"- {s.get('区间', '')}：{s.get('金额(万元)', '')}万元 "
                        f"× {s.get('费率(%)', '')}% = **{s.get('费用(万元)', '')}万元**"
                    )
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
    if st.button("🗑 结束本次计算", key=f"ck_cancel_{stage}_{ctx.get('qno', 1)}"):
        _cancel_task(ctx)


def _ask_fees(ctx: dict):
    qno = ctx.get("qno", 1)
    if ctx.get("preset_all") and not ctx.get("fees_confirmed"):
        # 未指明费种：默认全量计算，先确认是否有不需要的费种
        st.markdown(
            f"未指明具体费种，将**默认计算全部 {len(ctx.get('fees') or [])} 项二类费**"
            f"（不含交易服务费）。\n\n**是否有不需要计算的费用？**")
        if st.button("✅ 没有，全部计算", key=f"ck_fees_allok_{qno}",
                     use_container_width=True, type="primary"):
            ctx["fees_confirmed"] = True
            ctx["qno"] = qno + 1
            st.rerun()
        st.caption("如有不需要的费种，直接回复即可，"
                   "例如：**「不需要勘察费、不算监理费」**。")
        return
    st.markdown("请问需要计算哪些**二类费**？（点击选择，可多选）")
    defs = fee_catalog_for(ctx.get("region"))
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
            ctx["fees"] = [f for f in FEE_CATALOG if f != "交易服务费"]
            ctx["preset_all"] = True
            ctx["qno"] = qno + 1
            st.rerun()
    with col2:
        if st.button("✅ 确认所选费种", key=f"ck_fees_ok_{qno}", use_container_width=True,
                     disabled=not selected):
            ctx["fees_confirmed"] = True
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
            "建安工程费（万元）", min_value=0.0, value=0.0, step=100.0,
            key=f"ck_amounts_jianan_{qno}")
    with c2:
        shebei_in = st.number_input(
            "设备购置费（万元，没有可留空）", min_value=0.0, value=0.0, step=50.0,
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
            ctx["qno"] = qno + 1
            st.rerun()
    with col2:
        if st.button("🚫 无设备费", key=f"ck_amounts_nosb_{qno}",
                     use_container_width=True):
            am["shebei"] = 0.0
            am["shebei_known"] = True
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
            ctx["qno"] = qno + 1
            st.rerun()


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


def _clear_card_overrides(ctx: dict, card: dict):
    """「本卡用默认值」：清除该卡已收集的覆盖值。"""
    fn = card["fee"]
    kind = card["kind"]
    if kind == "coef":
        ctx.get("coef_overrides", {}).pop(fn, None)
    elif kind == "rate":
        ctx.get("rate_overrides", {}).pop(fn, None)
    elif kind == "service":
        ctx.get("service_selections", {}).pop(fn, None)
    elif kind == "party":
        ctx["jiaoyi_party"] = None
    elif kind == "shencha":
        ctx.get("rate_overrides", {}).pop("施工图审查费", None)
        ctx.get("spec_overrides", {}).pop("施工图审查费", None)
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
    card = cards[pidx]
    label = FEE_CATALOG.get(card["fee"], {}).get("label", card["fee"])
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
    elif kind == "yubei_rate":
        _ask_yubei_rate_card(ctx, card, qno)
    elif kind == "shuibao_comp":
        _ask_shuibao_comp_card(ctx, card, qno)
    elif kind == "spec":
        _ask_spec_card(ctx, card, qno)

    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("✅ 确认并继续", key=f"ck_params_ok_{qno}", use_container_width=True):
            ctx["qno"] = qno + 1
            _params_advance(ctx, pidx, len(cards))
            st.rerun()
    with col2:
        if st.button("⏭ 本卡用默认值", key=f"ck_params_skip_{qno}", use_container_width=True):
            _clear_card_overrides(ctx, card)
            ctx["qno"] = qno + 1
            _params_advance(ctx, pidx, len(cards))
            st.rerun()
    with col3:
        if st.button("⚡ 剩余全部用默认值", key=f"ck_params_alldef_{qno}",
                     use_container_width=True):
            ctx["params_done"] = True
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
    if chosen["key"] != cur:
        ctx.setdefault("rate_overrides", {})["施工图审查费"] = chosen["key"]
        st.rerun()
    if chosen.get("billing") == "area":
        area = st.number_input("建筑面积（m²）", min_value=0.0, value=0.0, step=1000.0,
                               key=f"ck_shencha_area_{qno}")
        if area > 0:
            ctx.setdefault("spec_overrides", {}).setdefault(
                "施工图审查费", {})["area_m2"] = float(area)


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
        default = ["编制报告书"]
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
    ctx.setdefault("service_selections", {})[fee_name] = list(picked)
    if not picked:
        st.warning("未选择任何服务类型，该费种将计为 0 元。")


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
            ev = st.number_input("开采量（m³）", min_value=0.0,
                                 value=float(sbp.get("extract_vol", 0.0) or 0.0),
                                 step=1000.0, key=f"ck_sb_ev_{qno}")
            sbp["extract_vol"] = float(ev)
        with c6:
            mv = st.number_input("材料量（m³）", min_value=0.0,
                                 value=float(sbp.get("material_vol", 0.0) or 0.0),
                                 step=1000.0, key=f"ck_sb_mv_{qno}")
            sbp["material_vol"] = float(mv)
        with c7:
            wv = st.number_input("弃方量（m³）", min_value=0.0,
                                 value=float(sbp.get("waste_vol", 0.0) or 0.0),
                                 step=1000.0, key=f"ck_sb_wv_{qno}")
            sbp["waste_vol"] = float(wv)
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
            v = st.number_input(f"{ps['label']}（{ps.get('unit', '')}）",
                                min_value=0.0, value=float(default_v or 0.0),
                                step=step, key=f"ck_spec_{fn}_{ps['key']}_{qno}")
            params[ps["key"]] = float(v)
    ctx.setdefault("spec_params", {})[fn] = params


def _ask_discount(ctx: dict):
    qno = ctx.get("qno", 1)
    discs = ctx.get("discounts") or {}
    if discs:
        st.markdown("**已设折扣**：")
        for fn, d in discs.items():
            st.markdown(f"- {FEE_CATALOG.get(fn, {}).get('label', fn)}：打 {d:g} 折")
    if ctx.get("_subflow") == "discount_loop":
        _discount_loop(ctx)
        return
    st.markdown("是否有费用需要**打折**？")
    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("🚫 不需要", key=f"ck_disc_no_{qno}", use_container_width=True):
            ctx["discount_answered"] = True
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
            ctx["qno"] = qno + 1
            st.rerun()


def _discount_loop(ctx: dict):
    qno = ctx.get("qno", 1)
    st.markdown("**💸 指定费用打折**")
    fee_names = [f for f in ctx.get("fees") or []
                 if FEE_CATALOG.get(f, {}).get("tier") is not None and f != "预备费"]
    fee_options = [(f, FEE_CATALOG.get(f, {}).get("label", f)) for f in fee_names]
    if not fee_options:
        st.info("当前费种无适用折扣项。")
        ctx["_subflow"] = None
        st.rerun()
    fn = st.selectbox("费种", fee_options, format_func=lambda x: x[1],
                      key=f"ck_dl_fee_{qno}")
    coef = st.number_input("折扣系数（0.8 = 打八折）", min_value=0.01, max_value=2.0,
                           value=0.8, step=0.05, format="%.2f", key=f"ck_dl_coef_{qno}")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("✅ 应用", key=f"ck_dl_ok_{qno}", use_container_width=True):
            ctx.setdefault("discounts", {})[fn[0]] = float(coef)
            ctx["discount_answered"] = True
            ctx["_subflow"] = None
            ctx["qno"] = qno + 1
            st.rerun()
    with c2:
        if st.button("↩ 返回", key=f"ck_dl_back_{qno}", use_container_width=True):
            ctx["_subflow"] = None
            ctx["qno"] = qno + 1
            st.rerun()


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
    col1, col2 = st.columns(2)
    with col1:
        if st.button("🚀 开始计算", key=f"ck_final_go_{qno}",
                     use_container_width=True, type="primary"):
            _finalize_and_render(ctx)
    with col2:
        if st.button("🗑 结束本次计算", key=f"ck_final_cancel_{qno}",
                     use_container_width=True):
            _cancel_task(ctx)


def _finalize_and_render(ctx: dict):
    """结算并进入 done 阶段（结果持久化在 pending_task["payload"]）。"""
    with st.spinner("正在计算..."):
        result = settle_checklist(ctx)
    if result["kind"] == "cascade":
        payload = result["payload"]
        ctx["payload"] = payload
        ctx["phase"] = "done"
        _record_calculation(ctx.get("query", ""), {
            "fee_type": "__cascade__",
            "结果汇总": {"项目总投资(万元)": payload["preview"]["project_total_with_custom"]},
        })
        # 折扣情景：结算后自动生成对比（缓存防重复计算）
        if ctx.get("discount_scenario"):
            ctx["_discount_scenarios"] = calc_discount_scenarios(ctx)
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
    payload = ctx.get("payload")
    if payload:
        preview = payload["preview"]
        st.markdown("## 二类费计算结果（程序精确计算）")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("第一部分工程费", f"{payload['total_part1']:.2f} 万元")
        c2.metric("二类费合计", f"{preview['fee_total_with_custom']:.2f} 万元")
        c3.metric("预备费", f"{preview['yubei_total']:.2f} 万元")
        c4.metric("项目总投资", f"{preview['project_total_with_custom']:.2f} 万元")

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

        # 折扣情景对比（结算后按需生成，缓存防重复计算）
        if ctx.get("_discount_scenarios"):
            _render_comparison_result(ctx["_discount_scenarios"])

        col1, col2 = st.columns(2)
        with col1:
            try:
                excel_bytes = _build_cascade_excel(payload)
                st.download_button(
                    "📥 导出 Excel 汇总表",
                    data=excel_bytes,
                    file_name="二类费计算汇总.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"ck_excel_{ctx.get('qno', 1)}",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"Excel 生成失败：{e}")
        with col2:
            if st.button("📊 折扣方案对比", key=f"ck_scen_btn_{ctx.get('qno', 1)}",
                         use_container_width=True):
                with st.spinner("正在生成折扣情景..."):
                    ctx["_discount_scenarios"] = calc_discount_scenarios(ctx)
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


def _try_start_fee_task(query: str) -> bool:
    """尝试启动清单式任务流。返回 True 表示已接管（调用方立即 rerun）。

    路由规则：
    - fee_calc_intent False（查规则等）→ 落 LLM，返回 False
    - cascade/iteration/comparison 且金额齐 → 旧直算路径（B3 翻转后进清单机）
    - 其余 → build_checklist_meta → pending_task，返回 True
    """
    if not fee_calc_intent(query):
        return False
    ctx = build_checklist_meta(
        query, region=st.session_state.get("selected_region"),
        force=_TASK_FLOW_CASCADE,
    )
    if ctx is None:
        return False
    if not _TASK_FLOW_START and not _TASK_FLOW_CASCADE:
        return False
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
    st.markdown("- 智能问答（已上线）")
    st.markdown("- 二类费计算（已上线）")
    st.markdown("- 指标对比分析（开发中）")
    st.markdown("- 材料价格趋势（开发中）")

    st.divider()

    st.markdown("### 数据状态")
    try:
        engine = get_engine()
        total = len(engine.chunks)
        cats = len(engine.data)
        kb_count = len(engine.knowledge_chunks)
        st.success(f"已加载 {cats} 个类别，共 {total} 条记录")
        st.caption(f"知识库：{kb_count} 个政策文件片段")
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

# ===== 主界面 =====
st.title("🏗️ 工程造价智能问答")
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
    if pending_task and pending_task.get("phase") == "ask":
        _user_msg_appended = True
        st.session_state.messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)
        res = advance_checklist(
            pending_task, prompt, region=st.session_state.get("selected_region"))
        st.session_state.pending_task = pending_task
        if res.get("recognized"):
            st.rerun()
        if res.get("fee_domain_token"):
            # 有费种/金额关键词但没解析出新信息 → 留在任务中提示
            st.session_state.messages.append({
                "role": "assistant",
                "content": "🤔 没有从这句话中识别出新的信息。"
                           "请回复当前步骤所需的内容，或点击「🗑 结束本次计算」重新提问。",
            })
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
    if _try_start_fee_task(prompt):
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