"""
绿化造价智能助手 - Streamlit 网页界面（Phase 1: 问答功能）
"""
import re
import streamlit as st
from rag_engine import get_engine
from fee_engine import detect_and_calculate, calc_jianli, calc_sheji, calc_huanping

# ===== 页面设置 =====
st.set_page_config(
    page_title="绿化造价智能助手",
    page_icon="🌿",
    layout="wide",
)

# ===== 辅助渲染函数 =====


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
        st.caption(f"依据：{fee_result.get('依据', '')}")

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
            # 区分两种步骤格式：粗略估算类（有"步骤"/"公式"/"结果"）vs 分档累进类（有"区间"/"费率"）
            if steps[0].get("步骤"):
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


def _render_sheji_static(fee_result):
    """渲染工程设计费的静态确认框（绕过 LLM）"""
    params = fee_result.get("参数", {})
    prof = params.get("专业调整系数", "")
    complexity = params.get("复杂程度系数", "")
    additional = params.get("附加调整系数", "")
    jifei = params.get("计费额(万元)", "")
    jijia = params.get("收费基价(万元)", "")
    basic = fee_result.get("基本设计收费(万元)")

    result_lines = [
        f"以上为程序依据《工程勘察设计收费管理规定》（计价格[2002]10号）精确计算结果：\n",
        f"- 计费额 **{jifei}** 万元",
        f"- 收费基价（附表一内插）**{jijia}** 万元",
        f"- 专业调整系数（附表二）**{prof}**",
        f"- 复杂程度系数 **{complexity}**",
        f"- 附加调整系数 **{additional}**",
        f"- 基本设计收费 = {jijia} × {prof} × {complexity} × {additional} = **{basic} 万元**",
    ]

    other_items = fee_result.get("其他设计收费明细") or []
    other_total = 0.0
    result_lines.append("")
    result_lines.append("**其他设计收费（可选，按基本设计收费比例计取）**：")
    other_types = [
        ("总体设计费", 0.05),
        ("主体设计协调费", 0.05),
        ("施工图预算编制费", 0.10),
        ("竣工图编制费", 0.08),
    ]
    other_selected = {od.get("项目", ""): od.get("费用(万元)", 0) for od in other_items}
    for name, rate in other_types:
        if name in other_selected:
            fee_val = other_selected[name]
            other_total += fee_val
            result_lines.append(f"  - {name}（{int(rate*100)}%）**{fee_val} 万元** ✓已计算")
        else:
            fee_val = round(basic * rate, 4)
            result_lines.append(f"  - {name}（{int(rate*100)}%）{fee_val} 万元")
    for od in other_items:
        if od.get("项目", "") not in dict(other_types):
            fee_val = od.get("费用(万元)", 0)
            other_total += fee_val
            result_lines.append(f"  - {od.get('项目', '')} **{fee_val} 万元** ✓已计算")

    if other_items:
        benchmark = fee_result.get("结果(万元)", basic + other_total)
        result_lines.append(
            f"- 工程设计收费基准价 = {basic} + {other_total} = **{benchmark} 万元**"
        )
    else:
        result_lines.append(
            "  > 💡 以上为参考值，需在提问时说明（如\"含施工图预算\"）才会计入总计。"
        )

    st.success("\n".join(result_lines))


def _build_sheji_text(fee_result):
    """构建工程设计费的对话历史文本"""
    params = fee_result.get("参数", {})
    prof = params.get("专业调整系数", "")
    complexity = params.get("复杂程度系数", "")
    additional = params.get("附加调整系数", "")
    jifei = params.get("计费额(万元)", "")
    jijia = params.get("收费基价(万元)", "")
    basic = fee_result.get("基本设计收费(万元)")

    text = (
        f"根据计价格[2002]10号，计费额{jifei}万元，"
        f"收费基价{jijia}万元（附表一内插），"
        f"专业调整系数{prof}，"
        f"复杂程度系数{complexity}，"
        f"附加调整系数{additional}，"
        f"基本设计收费 **{basic} 万元**。"
    )
    other_items = fee_result.get("其他设计收费明细") or []
    if other_items:
        other_desc = "；".join(
            f"{od.get('项目', '')}{od.get('费用(万元)', 0)}万元"
            for od in other_items
        )
        benchmark = fee_result.get("结果(万元)")
        text += f"其他设计收费：{other_desc}。工程设计收费基准价 **{benchmark} 万元**。"
    else:
        other_types = [
            ("总体设计费", 0.05), ("主体设计协调费", 0.05),
            ("施工图预算编制费", 0.10), ("竣工图编制费", 0.08),
        ]
        other_lines = "，".join(
            f"{n}({int(r*100)}%)={round(basic*r,4)}万元"
            for n, r in other_types
        )
        text += (
            f"其他设计收费（参考）：{other_lines}。"
            f"请在提问时说明需要哪些（如\"含施工图预算\"），会计入基准价。"
        )
    return text


# ===== 多费种迭代计算渲染函数 =====

def _render_cascade_result(result):
    """渲染模式1：多费种联算结果。"""
    import pandas as pd

    st.markdown("## 多费种联算结果（程序精确计算）")

    params = result["输入参数"]
    col1, col2, col3 = st.columns(3)
    col1.metric("建安工程费", f"{params['建安工程费(万元)']} 万元")
    col2.metric("设备购置费", f"{params['设备购置费(万元)']} 万元")
    col3.metric("项目类型", params["项目类型"])

    st.markdown(
        "**依赖层级说明**：\n"
        "- **T0**（仅依赖建安+设备）：监理费、设计费、勘察费、劳安评审费、场地准备费、工程保险费\n"
        "- **T1**（依赖 T0 结果）：施工图审查费、交易服务费\n"
        "- **T2**（依赖总投资）：建设管理费、可行性研究费、环境影响咨询费\n"
        "- **预备费**：（第一部分工程费+二类费）× 5%"
    )

    rows = result["费种合计"]
    tier_colors = {0: "#e8f5e9", 1: "#fff3e0", 2: "#e3f2fd", 3: "#fce4ec"}
    tier_labels = {0: "Tier 0 — 第一部分工程费相关", 1: "Tier 1 — 勘察设计费相关", 2: "Tier 2 — 总投资相关", 3: "预备费"}

    for tier in [0, 1, 2, 3]:
        tier_rows = [r for r in rows if r["层级"] == tier]
        if tier_rows:
            st.markdown(f"#### {tier_labels[tier]}")
            df = pd.DataFrame(tier_rows)
            df = df[["费种", "金额(万元)"]]
            st.dataframe(df, use_container_width=True, hide_index=True)

    # 汇总
    summary = result["结果汇总"]
    st.markdown("---")
    # 额外费用（用户指定）
    extra_fees = result.get("额外费用", [])
    if extra_fees:
        st.markdown("#### 额外费用（用户指定）")
        for e in extra_fees:
            st.markdown(f"- **{e['名称']}**：{e['金额(万元)']} 万元")

    st.markdown("### 费用汇总")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("第一部分工程费", f"{summary['第一部分工程费(万元)']:.4f} 万元")
    col2.metric("二类费合计", f"{summary['二类费合计(万元)']:.4f} 万元")
    yubei_val = summary.get("预备费(万元)", 0)
    col3.metric("预备费", f"{yubei_val:.4f} 万元")
    col4.metric("项目总投资", f"{summary['项目总投资(万元)']:.4f} 万元")

    # 层级小计
    extra_caption = ""
    if extra_fees:
        extra_total = summary.get("额外费用小计(万元)", 0)
        extra_caption = f" ｜ 额外费用：{extra_total:.4f} 万元"
    st.caption(
        f"T0 小计：{summary['T0小计(万元)']:.4f} 万元 ｜ "
        f"T1 小计：{summary['T1小计(万元)']:.4f} 万元 ｜ "
        f"T2 小计：{summary['T2小计(万元)']:.4f} 万元"
        f"{extra_caption}"
    )

    skipped = result.get("跳过的费种", {})
    if skipped:
        with st.expander("⚠️ 无法自动计算的费种"):
            for name, reason in skipped.items():
                st.warning(f"**{name}**：{reason}")

    with st.expander("查看各费种详细计算步骤"):
        for fee_name, detail in result["明细"].items():
            if not fee_name.startswith("_"):
                _render_engine_card(detail)

    # 构建响应文本
    lines = []
    for r in rows:
        lines.append(f"- **{r['费种']}**：{r['金额(万元)']:.4f} 万元")
    if extra_fees:
        for e in extra_fees:
            lines.append(f"- **{e['名称']}**（用户指定）：{e['金额(万元)']} 万元")
    yubei_text = ""
    yb_val = summary.get("预备费(万元)", 0)
    if yb_val > 0:
        yubei_text = f"\n**预备费（基本预备费）：{yb_val:.4f} 万元**（(一类费+二类费)×5%）"
    return (
        f"## 多费种联算结果\n\n"
        f"计费基数：建安费 {params['建安工程费(万元)']} 万 + 设备费 {params['设备购置费(万元)']} 万 "
        f"= **{params['第一部分工程费(万元)']} 万元**\n\n"
        f"### 各项费用\n\n" + "\n".join(lines) + "\n\n"
        f"**二类费合计：{summary['二类费合计(万元)']:.4f} 万元**"
        f"{yubei_text}\n\n"
        f"**项目总投资：{summary['项目总投资(万元)']:.4f} 万元**"
    )


def _render_iteration_result(result):
    """渲染模式2：迭代收敛结果。"""
    import pandas as pd

    st.markdown("## 迭代计算（总投资收敛）")

    params = result["输入参数"]
    col1, col2 = st.columns(2)
    col1.metric("建安工程费", f"{params['建安工程费(万元)']} 万元")
    col2.metric("设备购置费", f"{params['设备购置费(万元)']} 万元")

    st.info(
        "**迭代原理**：建设管理费、可行性研究费、环境影响咨询费依赖总投资；"
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
            f"预备费：**{yubei_val:.4f} 万元**（(一类费+二类费)×5%），"
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
        st.markdown(f"- **{fee_key}**：{val:.4f} 万元")
    # 预备费单独显示
    yubei_final_val = final.get("预备费(万元)")
    if yubei_final_val is not None and yubei_final_val > 0:
        st.markdown(f"- **预备费**：{yubei_final_val:.4f} 万元")
    proj_final = final.get("项目总投资(万元)")
    if proj_final is not None:
        st.markdown(f"\n**项目总投资（含预备费）：{proj_final:.2f} 万元**")

    with st.expander("查看每轮迭代详细数据"):
        for h in history:
            st.markdown(f"#### 第 {h['迭代次数']} 轮")
            fees = h["各项费用"]
            for fee_key, val in sorted(fees.items()):
                st.markdown(f"- {fee_key}：{val:.4f} 万元")
            st.caption(f"总投资：{h['总投资(万元)']:.2f} 万元 ｜ 变化：{h['变化(万元)']:.4f} 万元")

    # 响应文本
    yb_val = final.get("预备费(万元)", 0)
    proj_total = final.get("项目总投资(万元)", final["总投资(万元)"])
    yb_text = f"\n预备费：**{yb_val:.4f} 万元**（(一类费+二类费)×5%）" if yb_val > 0 else ""
    return (
        f"## 迭代计算结果\n\n"
        f"经过 **{result['迭代次数']}** 次迭代，静态总投资收敛至 "
        f"**{final['总投资(万元)']:.2f} 万元**，"
        f"二类费合计 **{final['二类费合计(万元)']:.2f} 万元**。"
        f"{yb_text}\n"
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
                st.markdown(f"- **{fee_key}**：{val:.4f} 万元")
            st.metric("二类费合计", f"{s['二类费合计(万元)']:.4f} 万元")
            st.metric("总投资", f"{s['总投资(万元)']:.4f} 万元")

    # 响应文本
    return (
        f"## 多方案比选结果\n\n"
        f"扫描参数：{sweep['参数描述']}，共 {len(sweep['值列表'])} 个方案。"
    )


# ===== 侧边栏 =====
with st.sidebar:
    st.title("🌿 绿化造价智能助手")
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
    examples_fee = [
        "建安工程费131万，设备费160万，桥梁工程，勘察费多少？",
        "工程总概算8000万，建设管理费多少？",
        "中标金额6000万工程招标代理费",
        "交易服务费 中标额2000万",
        "监理费 计费额8000万",
        "总投资1.2亿建设管理费和监理费",
    ]
    for i, ex in enumerate(examples_fee):
        if st.button(ex, use_container_width=True, key=f"fee_{i}"):
            st.session_state.current_query = ex

    st.divider()

    st.markdown("### 多费种联算 / 迭代 / 比选")
    examples_multi = [
        "建安费131万，设备费160万，桥梁工程，帮我算全部费用",
        "建安费8000万，工程总概算迭代计算",
        "建安费5000万，设备费3000万，方案比选",
    ]
    for i, ex in enumerate(examples_multi):
        if st.button(ex, use_container_width=True, key=f"multi_{i}"):
            st.session_state.current_query = ex

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
st.title("🌿 绿化工程造价智能问答")
st.caption("基于园林绿化工程指标数据库，提供专业造价问答服务 | v2026-06-22 设计费三系数输出")

# 全局调试：检查 current_query 状态
if "current_query" in st.session_state:
    st.warning(f"**:bug: current_query 已设置:** `{st.session_state.current_query[:80]}`")

# 初始化引擎
with st.spinner("正在加载数据库和 AI 模型..."):
    engine = get_engine()

# 初始化聊天历史
if "messages" not in st.session_state:
    st.session_state.messages = [
        {
            "role": "assistant",
            "content": (
                "你好！我是绿化造价智能助手。\n\n"
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

# ===== 交互式费率选择（持久化在聊天区外） =====
if "pending_rate_select" in st.session_state:
    ctx = st.session_state.pending_rate_select
    fee_result = ctx["fee_result"]
    fee_name = fee_result.get("费种", "")
    ft = fee_result.get("fee_type", "")
    basis = fee_result.get("依据", "")
    desc = fee_result.get("说明", "")
    detail = fee_result.get("费率明细", [])
    steps = fee_result.get("计算步骤", [])
    params = fee_result.get("参数", {})

    rate_opts = [d['费率'] for d in detail]
    fee_map = {d['费率']: d['费用(万元)'] for d in detail}
    mid_idx = len(detail) // 2

    st.divider()

    # ── 卡片容器 ──
    with st.container(border=True):
        # 标题行
        col_title, col_badge = st.columns([3, 1])
        with col_title:
            st.markdown(f"## 🎯 {fee_name}")
        with col_badge:
            st.info(f"共 {len(detail)} 档费率")

        st.caption(f"📜 **依据**：{basis}")

        # 计算过程折叠
        if steps:
            with st.expander("📐 计算过程", expanded=False):
                for i, s in enumerate(steps, 1):
                    st.markdown(
                        f"**{i}. {s.get('步骤', '')}**  \n"
                        f"> {s.get('公式', '')}  \n"
                        f"> → **{s.get('结果', '')}**"
                    )

        st.markdown("---")
        st.markdown("### 📊 选择适用费率")

        # 费率卡片式选择（radio + 可视化费用卡片）
        selected_rate = st.radio(
            "费率（间隔 0.1%）",
            rate_opts,
            index=mid_idx,
            horizontal=True,
            key=f"persist_rate_{ft}",
            label_visibility="collapsed",
        )

        # 选中的费率高亮卡片
        selected_fee = fee_map[selected_rate]
        st.markdown(
            f"""<div style="
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                border-radius: 12px;
                padding: 20px 28px;
                margin: 12px 0;
                color: white;
            ">
                <div style="font-size: 0.85rem; opacity: 0.85; margin-bottom: 4px;">✅ 当前选择</div>
                <div style="display: flex; align-items: baseline; gap: 16px;">
                    <span style="font-size: 1.1rem;">费率</span>
                    <span style="font-size: 2.0rem; font-weight: 700;">{selected_rate}</span>
                    <span style="font-size: 1.1rem; opacity: 0.7;">→ 费用</span>
                    <span style="font-size: 2.0rem; font-weight: 700;">{selected_fee} 万</span>
                </div>
            </div>""",
            unsafe_allow_html=True,
        )

        st.markdown("---")

        # 操作按钮行
        col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 2])
        with col_btn1:
            if st.button("✅ 确认选择", type="primary", use_container_width=True, key=f"confirm_rate_{ft}"):
                response = (
                    f"## {fee_name}\n\n"
                    f"**依据**：{basis}\n\n"
                    f"**选定费率**：{selected_rate}\n\n"
                    f"**费用**：{selected_fee} 万元\n\n"
                    f"{desc}"
                )
                st.session_state.messages.append({"role": "assistant", "content": response})
                del st.session_state.pending_rate_select
                st.rerun()
        with col_btn2:
            if st.button("🗑 取消", use_container_width=True, key=f"cancel_rate_{ft}"):
                del st.session_state.pending_rate_select
                st.rerun()
        with col_btn3:
            pass  # 占位

        # 底部说明
        with st.expander("ℹ️ 费率说明"):
            st.info(desc)

# ===== 交互式系数选择（持久化，在聊天区外渲染） =====
if "pending_coef_select" in st.session_state:
    ctx = st.session_state.pending_coef_select
    meta = ctx["coef_metadata"]
    fee_result = ctx["fee_result"]
    coefs = meta.get("coefs", [])
    base_params = meta.get("base_params", {})
    calc_func = meta.get("calc_func", "")
    fee_name = fee_result.get("费种", "")
    ft = fee_result.get("fee_type", "")
    basis = fee_result.get("依据", "")

    st.divider()

    with st.container(border=True):
        # 标题行
        col_title, col_badge = st.columns([3, 1])
        with col_title:
            st.markdown(f"## 🎛️ {fee_name} — 系数调整")
        with col_badge:
            st.info(f"{len(coefs)} 个系数")

        st.caption(f"📜 **依据**：{basis}")

        # ── 各系数下拉选择器 ──
        selected_coefs: dict = {}
        for i, coef_def in enumerate(coefs):
            key = coef_def["key"]
            param_name = coef_def["param_name"]
            current_val = float(coef_def["current"])
            current_label = coef_def.get("current_label", str(current_val))
            options: list = coef_def.get("options", [])
            desc = coef_def.get("description", "")

            st.markdown(f"### {key}")
            st.caption(f"{desc}")

            # 构建选项标签（含值）
            option_labels = [f"{label}（{val}）" for label, val in options]
            option_values = [val for _, val in options]

            # 找到当前值对应的索引
            try:
                current_idx = option_values.index(current_val)
            except ValueError:
                current_idx = len(option_values)  # 指向"自定义"

            # 添加"自定义"选项
            option_labels.append("✏️ 自定义…")
            option_values.append(-1.0)  # -1 = 自定义标记

            selected_idx = st.selectbox(
                f"选择{key}",
                range(len(option_labels)),
                index=min(current_idx, len(option_labels) - 1),
                format_func=lambda idx, labels=option_labels: labels[idx],
                key=f"coef_sel_{ft}_{param_name}",
                label_visibility="collapsed",
            )

            chosen_val = option_values[selected_idx]

            if chosen_val == -1.0:
                # 自定义：显示数字输入
                custom_val = st.number_input(
                    f"自定义{key}的值",
                    min_value=0.10,
                    max_value=5.00,
                    value=current_val if current_val > 0.1 else 1.0,
                    step=0.05,
                    format="%.2f",
                    key=f"coef_cust_{ft}_{param_name}",
                )
                selected_coefs[param_name] = float(custom_val)
                selected_coefs[f"{param_name}_label"] = f"自定义（{custom_val:.2f}）"
            else:
                selected_coefs[param_name] = float(chosen_val)
                # 提取纯标签（去掉末尾的系数值括号）
                raw_label = option_labels[selected_idx]
                if "（" in raw_label:
                    raw_label = raw_label.rsplit("（", 1)[0]
                selected_coefs[f"{param_name}_label"] = raw_label

        st.markdown("---")

        # ── 根据所选系数实时重算 ──
        recalc_fee = None
        recalc_desc = ""
        recalc_error = ""
        try:
            if calc_func == "calc_jianli":
                prof = selected_coefs.get("professional_coef", 1.0)
                comp = selected_coefs.get("complexity_coef", 1.0)
                elev = selected_coefs.get("elevation_coef", 1.0)
                jianan = base_params.get("jianan")
                shebei = base_params.get("shebei")
                amount_wan = base_params.get("amount_wan")

                if jianan is not None and shebei is not None:
                    recalc = calc_jianli(
                        jianan=jianan, shebei=shebei,
                        professional_coef=prof, complexity_coef=comp,
                        elevation_coef=elev,
                    )
                elif amount_wan is not None:
                    recalc = calc_jianli(
                        amount_wan=amount_wan,
                        professional_coef=prof, complexity_coef=comp,
                        elevation_coef=elev,
                    )
                else:
                    recalc = None
                    recalc_error = "无法确定计费额"

                recalc_fee = recalc["结果(万元)"] if recalc else None
                recalc_desc = recalc.get("说明", "") if recalc else ""

            elif calc_func == "calc_sheji":
                prof = selected_coefs.get("professional_coef", 1.0)
                comp = selected_coefs.get("complexity_coef", 1.0)
                addi = selected_coefs.get("additional_coef", 1.0)
                amount_wan = base_params.get("amount_wan")

                if amount_wan is not None:
                    addi_list = [addi] if abs(addi - 1.0) > 0.005 else None
                    recalc = calc_sheji(amount_wan, prof, comp, additional_coefs=addi_list)
                else:
                    recalc = None
                    recalc_error = "无法确定计费额"

                recalc_fee = recalc["结果(万元)"] if recalc else None
                recalc_desc = recalc.get("说明", "") if recalc else ""

            elif calc_func == "calc_huanping":
                ind = selected_coefs.get("industry_coef", 1.0)
                sens = selected_coefs.get("sensitivity_coef", 1.0)
                amount_wan = base_params.get("amount_wan", 0)
                svc = base_params.get("service_type", "编制报告书")
                ind_name = selected_coefs.get("industry_coef_label", "市政（默认）")

                recalc = calc_huanping(
                    amount_wan, svc,
                    industry_coef=ind, industry_name=ind_name,
                    sensitivity_coef=sens,
                )
                recalc_fee = recalc.get("结果中值(万元)")
                recalc_desc = recalc.get("说明", "")
            else:
                recalc_error = f"未知计算类型：{calc_func}"
        except Exception as e:
            recalc_error = f"计算出错：{e}"
            import traceback
            recalc_error += f"\n```\n{traceback.format_exc()}\n```"

        # ── 结果卡片 ──
        if recalc_fee is not None:
            st.markdown(
                f"""<div style="
                    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                    border-radius: 12px;
                    padding: 20px 28px;
                    margin: 12px 0;
                    color: white;
                ">
                    <div style="font-size: 0.85rem; opacity: 0.85; margin-bottom: 4px;">✅ 调整后费用</div>
                    <div style="display: flex; align-items: baseline; gap: 16px;">
                        <span style="font-size: 2.0rem; font-weight: 700;">{recalc_fee} 万元</span>
                    </div>
                </div>""",
                unsafe_allow_html=True,
            )
            if recalc_desc:
                with st.expander("📐 查看计算过程", expanded=False):
                    st.markdown(recalc_desc)
        elif recalc_error:
            st.error(recalc_error)

        st.markdown("---")

        # ── 操作按钮 ──
        col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 2])
        with col_btn1:
            if st.button("✅ 确认选择", type="primary", use_container_width=True, key=f"confirm_coef_{ft}"):
                if recalc_fee is None:
                    st.warning("请先选择有效的系数值")
                    st.stop()
                coef_summary = "、".join(
                    f"{cd['key']}={selected_coefs.get(cd['param_name'], cd['current'])}"
                    for cd in coefs
                )
                response = (
                    f"## {fee_name}\n\n"
                    f"**依据**：{basis}\n\n"
                    f"**调整后系数**：{coef_summary}\n\n"
                    f"**费用**：{recalc_fee} 万元\n\n"
                    f"---\n{recalc_desc}"
                )
                st.session_state.messages.append({"role": "assistant", "content": response})
                del st.session_state.pending_coef_select
                st.rerun()
        with col_btn2:
            if st.button("🗑 取消", use_container_width=True, key=f"cancel_coef_{ft}"):
                del st.session_state.pending_coef_select
                st.rerun()

# ===== 输入框 =====
st.divider()
st.markdown("### 输入你的问题")

if "current_query" in st.session_state:
    prompt = st.session_state.current_query
    del st.session_state.current_query
else:
    prompt = st.chat_input("请输入你的造价问题，例如：白皮松高度3.5米多少钱？")

if prompt:
    # 新提问时清除旧的待处理选择
    st.session_state.pop("pending_rate_select", None)
    st.session_state.pop("pending_coef_select", None)
    # 添加用户消息
    st.session_state.messages.append({"role": "user", "content": prompt})

    # 调试：将检测结果存入 session_state，跨 rerun 持久化
    try:
        debug_fee = detect_and_calculate(prompt)
        st.session_state.debug_info = {
            "prompt": prompt[:80],
            "fee_type": debug_fee.get("fee_type") if debug_fee else "None",
            "has_amount": debug_fee.get("has_amount") if debug_fee else "N/A",
            "费种": debug_fee.get("费种", "") if debug_fee else "",
            "结果": str(debug_fee.get("结果(万元)", debug_fee.get("结果(元)", ""))) if debug_fee else "",
        }
    except Exception as e:
        st.session_state.debug_info = {"error": str(e)}
        debug_fee = None

    # 关键：fee_result 必须在后续分支中使用
    fee_result = debug_fee

    with st.chat_message("user"):
        st.markdown(prompt)

    # 生成回答
    with st.chat_message("assistant"):
        # DEBUG（用 st.write 确保可见）
        st.write(f"🔍 DEBUG A: prompt='{prompt[:60]}' | len={len(prompt)}")
        try:
            fee_result = detect_and_calculate(prompt)
            st.write(f"🔍 DEBUG B: fee_result is None = {fee_result is None}, type={type(fee_result).__name__}")
            if fee_result:
                st.write(f"🔍 DEBUG C: fee_type={fee_result.get('fee_type')}, has_amount={fee_result.get('has_amount')}")
        except Exception as e:
            st.write(f"🔍 DEBUG EXCEPTION: {e}")
            import traceback
            st.code(traceback.format_exc())
            fee_result = None

        with st.spinner("正在检索数据并生成回答..."):
            # 二类费规则引擎
            # fee_result = detect_and_calculate(prompt)  # 已在上方调用

            if fee_result and fee_result.get("has_amount"):
                # === 多费种迭代模式路由 ===
                mode = fee_result.get("mode")
                if mode == "cascade":
                    response = _render_cascade_result(fee_result)
                elif mode == "iteration":
                    response = _render_iteration_result(fee_result)
                elif mode == "comparison":
                    response = _render_comparison_result(fee_result)

                if mode is None:
                    # === 引擎精确计算：单费种直接展示，不经过 LLM ===
                    is_sheji = fee_result.get("fee_type") == "工程设计费"
                    is_rate_selectable = fee_result.get("is_rate_selectable", False)
                    is_coef_selectable = fee_result.get("is_coef_selectable", False)

                    if is_rate_selectable:
                        # === 交互式费率选择：存入 session state，在聊天区外渲染 ===
                        st.session_state.pending_rate_select = {
                            "fee_result": fee_result,
                            "query": prompt,
                        }
                        fee_name = fee_result.get("费种", "")
                        n_rates = len(fee_result.get("费率明细", []))
                        response = (
                            f"## {fee_name}\n\n"
                            f"> ℹ️ 该费种支持交互式费率选择\n\n"
                            f"请滚动到页面下方 **🎯 费率选择** 区域，"
                            f"从 {n_rates} 档费率中选择适用费率后点击确认。"
                        )
                    elif is_coef_selectable:
                        # === 交互式系数选择：存入 session state，在聊天区外渲染 ===
                        st.session_state.pending_coef_select = {
                            "coef_metadata": fee_result.get("coef_metadata", {}),
                            "fee_result": fee_result,
                            "query": prompt,
                        }
                        fee_name = fee_result.get("费种", "")
                        n_coefs = len(fee_result.get("coef_metadata", {}).get("coefs", []))
                        response = (
                            f"## {fee_name}\n\n"
                            f"> ℹ️ 该费种支持交互式系数调整\n\n"
                            f"请滚动到页面下方 **🎛️ 系数调整** 区域，"
                            f"调整 {n_coefs} 个系数后点击确认。"
                        )
                    else:
                        st.markdown("### 计算结果（程序精确计算）")
                        _render_engine_card(fee_result)

                    if is_rate_selectable or is_coef_selectable:
                        pass  # 已在上面处理完成
                    elif is_sheji:
                        st.divider()
                        _render_sheji_static(fee_result)
                        response = _build_sheji_text(fee_result)
                    else:
                        # 非设计费：也跳过 LLM，用静态确认
                        st.divider()
                        fee_name = fee_result.get("费种", "")
                        result_val = fee_result.get("结果(万元)") or fee_result.get("结果(元)")
                        unit = "万元" if "结果(万元)" in fee_result else "元"
                        basis = fee_result.get("依据", "")
                        desc = fee_result.get("说明", "")
                        ft = fee_result.get("fee_type", "")

                        # 施工图审查费（津价管[2011]46号 + 建市[2007]86号）
                        is_shencha = ft == "施工图审查费"
                        # 环评费（计价格[2002]125号 — 四项服务类型全部输出）
                        is_huanping = ft == "环境影响咨询费"
                        # 可行性研究费（计价格[1999]1283号 — 内插法详细步骤）
                        is_keyan = ft == "可行性研究费"
                        # 粗略估算类费种（《市政工程设计概算编制办法》）
                        is_rough = ft in (
                            "勘察费", "劳动安全卫生评审费",
                            "场地准备费及临时设施费", "工程保险费",
                        )
                        if is_shencha or is_huanping or is_rough or is_keyan:
                            # 构建完整 markdown 响应（跨 rerun 持久化）
                            mid_val = fee_result.get("结果中值(万元)")
                            mid_text = f"（中值约 **{mid_val} 万元**）" if mid_val else ""

                            # 构建计算过程
                            steps = fee_result.get("计算步骤", [])
                            steps_md = ""
                            if steps:
                                steps_md = "### 计算过程\n\n"
                                for i, s in enumerate(steps, 1):
                                    step_name = s.get("步骤", "")
                                    formula = s.get("公式", "")
                                    result_step = s.get("结果", "")
                                    steps_md += f"**{i}. {step_name}**：{formula} → **{result_step}**\n\n"

                            # 构建费率明细表
                            detail = fee_result.get("费率明细", [])
                            detail_md = ""
                            if detail:
                                detail_md = "### 费率-费用对照表\n\n"
                                detail_md += "| 费率 | 费用（万元） |\n"
                                detail_md += "|------|-------------|\n"
                                for d in detail:
                                    detail_md += f"| {d['费率']} | **{d['费用(万元)']}** |\n"
                                detail_md += "\n"

                            if is_shencha:
                                response = (
                                    f"## {fee_name}\n\n"
                                    f"**依据**：{basis}\n\n"
                                    f"{steps_md}"
                                    f"---\n\n"
                                    f"### 计算结果\n\n"
                                    f"审查费：**{result_val} {unit}**\n\n"
                                    f"{desc}"
                                )
                            elif is_huanping:
                                # 四种服务类型结果表
                                all_svc = fee_result.get("全部服务类型结果", {})
                                svc_table = ""
                                if all_svc:
                                    svc_table = "### 四种服务类型全部结果\n\n"
                                    svc_table += "| 服务类型 | 费用范围（万元） | 中值（万元） |\n"
                                    svc_table += "|----------|:--:|:--:|\n"
                                    for svc_name in ["编制报告书", "编制报告表", "评估报告书", "评估报告表"]:
                                        svc_r = all_svc.get(svc_name, {})
                                        svc_table += (
                                            f"| **{svc_name}** "
                                            f"| {svc_r.get('结果(万元)', '-')} "
                                            f"| {svc_r.get('结果中值(万元)', '-')} |\n"
                                        )
                                    svc_table += "\n"
                                response = (
                                    f"## {fee_name}\n\n"
                                    f"**依据**：{basis}\n\n"
                                    f"{steps_md}"
                                    f"---\n\n"
                                    f"{svc_table}"
                                    f"### 计算结果\n\n"
                                    f"{desc}"
                                )
                            elif is_keyan:
                                # 服务类型结果表（2项或4项）
                                all_svc = fee_result.get("全部服务类型结果", {})
                                svc_table = ""
                                if all_svc:
                                    n = len(all_svc)
                                    if n == 4:
                                        title = "### 四种服务类型全部结果"
                                    elif n == 2:
                                        first_key = list(all_svc.keys())[0]
                                        if "可研" in first_key:
                                            title = "### 可研报告相关服务类型结果"
                                        else:
                                            title = "### 项目建议书相关服务类型结果"
                                    else:
                                        title = "### 服务类型结果"
                                    svc_table = f"{title}\n\n"
                                    svc_table += "| 服务类型 | 基准价（万元） | 最终费用（万元） |\n"
                                    svc_table += "|----------|:--:|:--:|\n"
                                    for svc_name in all_svc:
                                        svc_r = all_svc[svc_name]
                                        base = svc_r.get("基准价(万元)", "-")
                                        fee = svc_r.get("结果(万元)", "-")
                                        svc_table += f"| **{svc_name}** | {base} | {fee} |\n"
                                    svc_table += "\n"
                                n_svc = len(all_svc) if all_svc else 0
                                if n_svc > 0:
                                    response = (
                                        f"## {fee_name}\n\n"
                                        f"**依据**：{basis}\n\n"
                                        f"{svc_table}"
                                        f"{steps_md}"
                                        f"---\n\n"
                                        f"### 计算结果\n\n"
                                        f"最终费用：**{result_val} {unit}**\n\n"
                                        f"{desc}"
                                    )
                                else:
                                    response = (
                                        f"## {fee_name}\n\n"
                                        f"**依据**：{basis}\n\n"
                                        f"{steps_md}"
                                        f"---\n\n"
                                        f"### 计算结果\n\n"
                                        f"最终费用：**{result_val} {unit}**\n\n"
                                        f"{desc}"
                                    )
                            else:
                                # === 粗略估算类费种 — 交互式费率选择 ===
                                detail = fee_result.get("费率明细", [])
                                if detail:
                                    rate_opts = [d['费率'] for d in detail]
                                    fee_map = {d['费率']: d['费用(万元)'] for d in detail}
                                    mid_idx = len(detail) // 2

                                    st.markdown("### 请选择适用费率（间隔 0.1%）")
                                    selected_rate = st.selectbox(
                                        "费率",
                                        rate_opts,
                                        index=mid_idx,
                                        key=f"rate_select_{ft}",
                                        label_visibility="collapsed",
                                    )
                                    selected_fee = fee_map[selected_rate]
                                    st.success(
                                        f"选定费率 **{selected_rate}** → "
                                        f"**{fee_name}**：**{selected_fee} 万元**"
                                    )

                                    with st.expander("查看完整费率对照表"):
                                        st.markdown(detail_md)
                                else:
                                    selected_rate = "中值"
                                    selected_fee = mid_val

                                response = (
                                    f"## {fee_name}\n\n"
                                    f"**依据**：{basis}\n\n"
                                    f"{steps_md}"
                                    f"{detail_md}"
                                    f"---\n\n"
                                    f"### 计算结果\n\n"
                                    f"选定费率：**{selected_rate}**\n\n"
                                    f"费用：**{selected_fee} 万元**\n\n"
                                    f"{desc}"
                                )
                        else:
                            st.success(
                                f"以上为程序依据 **{basis}** 精确计算结果。\n\n"
                                f"计算结果：**{result_val} {unit}**\n\n"
                                f"{desc}"
                            )
                            response = (
                                f"根据{basis}，{fee_name}计算结果为 **{result_val} {unit}**。"
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
                    st.caption(f"依据：{fee_result.get('依据', '')}")
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