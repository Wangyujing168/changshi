"""
RAG 问答引擎 - 检索 + DeepSeek 生成
"""
import re
import json
from openai import OpenAI
from config import (
    DEEPSEEK_API_KEY,
    DEEPSEEK_BASE_URL,
    DEEPSEEK_MODEL,
    SYSTEM_PROMPT,
)
from data_loader import load_all_data, get_text_chunks


def _expand_spec_range(query: str) -> list[str]:
    """
    将用户输入的规格数值展开为对应的数据区间。
    规则：胸径/地径 X cm → 区间 (X-1).0-(X-1).9
    例如："胸径12cm" → ["11.0-11.9"]
          "地径7cm"  → ["6.0-6.9"]
          "高度3.5m" → 直接匹配原文
    """
    expanded = []

    # 匹配 "胸径12cm"、"地径7" 等模式
    for m in re.finditer(r'(胸径|地径)\s*(\d+\.?\d*)\s*(cm|m)?', query):
        try:
            num = float(m.group(2))
        except ValueError:
            continue

        # 胸径/地径: X → (X-1).0 ~ (X-1).9
        lower = num - 1
        expanded.append(f"{lower:.1f}-{lower + 0.9:.1f}")  # 11.0-11.9
        expanded.append(f"{int(lower)}-{int(lower)}.9")    # 11-11.9
        expanded.append(f"{lower:.1f}")                     # 11.0
        expanded.append(f"{int(lower)}")                    # 11

    return expanded


class CostRAGEngine:
    """绿化造价 RAG 问答引擎"""

    def __init__(self):
        # 初始化 DeepSeek 客户端
        self.client = OpenAI(
            api_key=DEEPSEEK_API_KEY,
            base_url=DEEPSEEK_BASE_URL,
        )

        # 加载数据
        self.data = load_all_data()
        self.chunks = get_text_chunks(self.data)

        # 构建品种名索引，用于快速匹配
        self._name_index = {}
        for chunk in self.chunks:
            name = chunk["name"]
            if name not in self._name_index:
                self._name_index[name] = []
            self._name_index[name].append(chunk)

        pass  # Engine initialized

    def search(self, query: str, top_k: int = 10) -> list[dict]:
        """
        检索相关数据块。
        采用关键词匹配 + 品种名模糊匹配 + 规格区间智能映射。
        """
        scored = []
        expanded_ranges = _expand_spec_range(query)

        for chunk in self.chunks:
            score = 0
            text = chunk["text"]
            name = chunk["name"]
            spec = chunk.get("spec", "")

            # 1. 品种名精确匹配（最高权重）
            if name in query:
                score += 10

            # 2. 品种名分词匹配
            name_parts = [name[i:i+2] for i in range(len(name)-1)]
            for part in name_parts:
                if part in query:
                    score += 2

            # 3. 规格区间智能映射匹配
            # 用户说"胸径12"→映射到"11.0-11.9"
            for expanded in expanded_ranges:
                if expanded in spec:
                    score += 8  # 高权重，仅次于精确品种名匹配
                    break

            # 4. 类别关键词匹配
            cat_keywords = {
                "常绿乔木": ["常绿"],
                "落叶乔木": ["落叶"],
                "小乔木": ["小乔", "小乔木"],
                "灌木": ["灌木"],
                "灌木球类": ["球类", "球"],
                "地被": ["地被", "草坪", "麦冬", "玉簪", "蛇莓"],
                "绿篱": ["绿篱", "绿蓠"],
                "花卉": ["花卉", "花", "草花", "宿根"],
            }
            for cat, kws in cat_keywords.items():
                if any(kw in query for kw in kws) and chunk["category"] == cat:
                    score += 3

            # 5. 规格关键词匹配
            spec_kws = ["高度", "胸径", "地径", "冠幅", "H=", "G", "cm", "m"]
            spec_parts = spec.replace(",", " ").split()
            for sp in spec_parts:
                if sp in query:
                    score += 2

            # 6. 单位匹配
            if ("平方" in query or "m2" in query or "m²" in query) and chunk.get("unit", "").startswith("元/m"):
                score += 2

            # 7. 通用关键词
            keywords = ["综合指标", "栽植费用", "苗木价格", "造价", "单价", "多少钱"]
            for kw in keywords:
                if kw in query:
                    score += 1

            if score > 0:
                scored.append((score, chunk))

        # 按分数排序，取 top_k
        scored.sort(key=lambda x: x[0], reverse=True)
        return [c for _, c in scored[:top_k]]

    def build_context(self, query: str, top_k: int = 10) -> str:
        """构建发给 LLM 的上下文"""
        results = self.search(query, top_k)

        if not results:
            return "未在数据库中检索到相关数据。"

        context_parts = ["以下是从绿化工程造价指标数据库中检索到的相关数据：\n"]

        # 按类别分组展示
        by_category = {}
        for r in results:
            cat = r["category"]
            if cat not in by_category:
                by_category[cat] = []
            by_category[cat].append(r)

        for cat, items in by_category.items():
            context_parts.append(f"\n### {cat}")
            for item in items:
                unit = item.get("unit", "元/株")
                seedling_info = f"（其中苗木价格 {item.get('苗木价格', '')} 元）" if item.get('苗木价格') else ""
                context_parts.append(
                    f"- {item['name']}（{item['spec']}）："
                    f"综合指标 **{item['comprehensive']}{unit}** {seedling_info}"
                )

        return "\n".join(context_parts)

    def ask(self, query: str) -> str:
        """
        执行一次问答：
        1. 检索相关数据
        2. 拼接上下文
        3. 调用 DeepSeek 生成回答
        """
        # 检索
        context = self.build_context(query)

        # 构建消息
        messages = [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": f"""请根据以下数据库检索结果回答用户问题。

## 检索结果
{context}

## 用户问题
{query}

请用专业简洁的语言回答，引用具体数据。如果没有查到相关数据，请如实告知。"""}
        ]

        # 调用 DeepSeek
        try:
            response = self.client.chat.completions.create(
                model=DEEPSEEK_MODEL,
                messages=messages,
                temperature=0.3,  # 低温度，保证专业性
                max_tokens=2000,
            )
            return response.choices[0].message.content
        except Exception as e:
            return f"❌ 调用 DeepSeek API 出错：{str(e)}\n\n请检查 API Key 是否正确配置。"

    def chat(self, query: str, history: list[dict] = None) -> str:
        """
        多轮对话模式（保留对话历史）。
        """
        if history is None:
            history = []

        # 每次对话都检索最新数据
        context = self.build_context(query)

        messages = [{"role": "system", "content": SYSTEM_PROMPT}]

        # 添加历史对话
        for msg in history[-6:]:  # 最近 3 轮对话
            messages.append(msg)

        # 添加当前查询和检索结果
        messages.append({
            "role": "user",
            "content": f"""## 数据库检索结果
{context}

## 用户问题
{query}

请根据检索结果回答。检索结果中没有的信息，结合你的专业知识补充说明。"""
        })

        try:
            response = self.client.chat.completions.create(
                model=DEEPSEEK_MODEL,
                messages=messages,
                temperature=0.3,
                max_tokens=2000,
            )
            return response.choices[0].message.content
        except Exception as e:
            return f"❌ 调用 DeepSeek API 出错：{str(e)}"


# 全局单例
_engine = None


def get_engine() -> CostRAGEngine:
    global _engine
    if _engine is None:
        _engine = CostRAGEngine()
    return _engine
