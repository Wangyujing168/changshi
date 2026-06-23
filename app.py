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
                            response = (
                                f"## {fee_name}\n\n"
                                f"**依据**：{basis}\n\n"
                                f"{steps_md}"
                                f"{detail_md}"
                                f"---\n\n"
                                f"### 估算结果\n\n"
                                f"估算范围：**{result_val} {unit}** {mid_text}\n\n"
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