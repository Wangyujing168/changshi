"""
绿化造价智能助手 - Streamlit 网页界面（Phase 1: 问答功能）
"""
import re
import streamlit as st
from rag_engine import get_engine
from fee_engine import detect_and_calculate

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
    jifei = params.get("计费额(万元)", "")
    jijia = params.get("收费基价(万元)", "")
    basic = fee_result.get("基本设计收费(万元)")

    text = (
        f"根据计价格[2002]10号，计费额{jifei}万元，"
        f"收费基价{jijia}万元（附表一内插），"
        f"专业调整系数{prof}，"
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
    st.caption("Powered by DeepSeek v4")

# ===== 主界面 =====
st.title("🌿 绿化工程造价智能问答")
st.caption("基于园林绿化工程指标数据库，提供专业造价问答服务")

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

# ===== 输入框 =====
st.divider()
st.markdown("### 输入你的问题")

if "current_query" in st.session_state:
    prompt = st.session_state.current_query
    del st.session_state.current_query
else:
    prompt = st.chat_input("请输入你的造价问题，例如：白皮松高度3.5米多少钱？")

if prompt:
    # 添加用户消息
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    # 生成回答
    with st.chat_message("assistant"):
        with st.spinner("正在检索数据并生成回答..."):
            # 二类费规则引擎
            fee_result = detect_and_calculate(prompt)

            if fee_result and fee_result.get("has_amount"):
                # === 引擎精确计算：全部费种直接展示，不经过 LLM ===
                is_sheji = fee_result.get("fee_type") == "工程设计费"

                st.markdown("### 计算结果（程序精确计算）")
                _render_engine_card(fee_result)

                if is_sheji:
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
                    params = fee_result.get("参数", {})
                    desc = fee_result.get("说明", "")
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