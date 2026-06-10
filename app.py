"""
绿化造价智能助手 - Streamlit 网页界面（Phase 1: 问答功能）
"""
import streamlit as st
from rag_engine import get_engine

# ===== 页面设置 =====
st.set_page_config(
    page_title="绿化造价智能助手",
    page_icon=":",
    layout="wide",
)

# ===== 侧边栏 =====
with st.sidebar:
    st.title("绿化造价智能助手")
    st.divider()

    st.markdown("### 功能导航")
    st.markdown("- 智能问答（已上线）")
    st.markdown("- Excel数据入库（开发中）")
    st.markdown("- 指标对比分析（开发中）")
    st.markdown("- 材料价格趋势（开发中）")

    st.divider()

    st.markdown("### 数据状态")
    try:
        engine = get_engine()
        total = len(engine.chunks)
        cats = len(engine.data)
        st.success(f"已加载 {cats} 个类别，共 {total} 条记录")
    except Exception as e:
        st.error(f"数据加载失败：{e}")

    st.divider()

    st.markdown("### 试试这些问题")
    examples = [
        "白皮松高度3.5米的综合指标是多少？",
        "落叶乔木胸径14cm的有哪些品种？",
        "常绿乔木和落叶乔木的综合指标对比",
        "灌木球类中综合指标最低的是哪个？",
        "银杏的综合指标是多少？",
    ]
    for ex in examples:
        if st.button(ex, use_container_width=True):
            st.session_state.current_query = ex

    st.divider()
    st.caption("Powered by DeepSeek v4")

# ===== 主界面 =====
st.title("绿化工程造价智能问答")
st.caption("基于园林绿化工程指标数据库，提供专业造价问答服务")

# 初始化引擎
@st.cache_resource
def init_engine():
    return get_engine()

with st.spinner("正在加载数据库和 AI 模型..."):
    engine = init_engine()

# 初始化聊天历史
if "messages" not in st.session_state:
    st.session_state.messages = [
        {
            "role": "assistant",
            "content": (
                "你好！我是绿化造价智能助手。\n\n"
                "我可以回答以下问题：\n"
                "- 查询具体苗木品种的综合指标、栽植费用\n"
                "- 对比不同规格、不同品种的造价差异\n"
                "- 按类别统计和分析造价数据\n\n"
                "请在下方输入你的问题，例如：白皮松高度3.5米多少钱一株？"
            ),
        }
    ]

# 显示聊天历史
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# ===== 输入框（始终在底部显示）=====
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
            # 构建对话历史
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

    # 刷新页面让新消息显示在输入框上方
    st.rerun()
