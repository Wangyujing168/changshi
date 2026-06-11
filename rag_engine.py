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
    WEB_SEARCH_ENABLED,
    SCORE_THRESHOLD as CONFIG_SCORE_THRESHOLD,
)
from data_loader import load_all_data, get_text_chunks


def _expand_spec_range(query: str) -> dict[str, list[str]]:
    """
    将用户输入的规格数值展开为对应的数据区间。
    规则：胸径/地径 X cm → 区间 (X-1).0-(X-1).9
    例如："胸径12cm" → ranges: ["11.0-11.9", "11-11.9"], singles: ["11.0", "11"]

    返回两个列表：
    - ranges:  范围格式，匹配全文（name+spec+text）
    - singles: 单数字格式，仅匹配 name+spec（防止在系数等无关数字上假阳性）
    """
    ranges = []
    singles = []

    # 匹配 "胸径12cm"、"地径7" 等模式（高度不适用此公式，直接原文匹配）
    for m in re.finditer(r'(胸径|地径)\s*(\d+\.?\d*)\s*(cm|m)?', query):
        try:
            num = float(m.group(2))
        except ValueError:
            continue

        # 胸径/地径: X → (X-1).0 ~ (X-1).9
        lower = num - 1
        ranges.append(f"{lower:.1f}-{lower + 0.9:.1f}")   # 11.0-11.9
        ranges.append(f"{int(lower)}-{int(lower)}.9")     # 11-11.9
        singles.append(f"{lower:.1f}")                      # 11.0
        singles.append(f"{int(lower)}")                     # 11

    return {"ranges": ranges, "singles": singles}


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

        # 联网搜索缓存
        self._web_cache: dict[str, str] = {}

        pass  # Engine initialized

    def search(self, query: str, top_k: int = 10) -> list[dict]:
        """
        检索相关数据块。
        采用关键词匹配 + 品种名模糊匹配 + 规格区间智能映射。
        """
        scored = []
        expanded = _expand_spec_range(query)

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
            # ranges:  范围格式匹配全文（spec+name+text，灌木规格可能嵌在 name 里）
            # singles: 单数字只匹配 name+spec（防止在系数 1.0794 等处假阳性）
            search_target = f"{spec} {name} {text}"
            name_spec_target = f"{spec} {name}"
            for r in expanded["ranges"]:
                if r in search_target:
                    score += 8
                    break
            for s in expanded["singles"]:
                if s in name_spec_target:
                    score += 5
                    break

            # 3b. 查询中的数字直接匹配原文（用于高度等不适用X-1公式的规格）
            query_nums = re.findall(r'(\d+\.?\d*)', query)
            for num in query_nums:
                if len(num) >= 2 or '.' in num:
                    if num in name_spec_target:
                        score += 7
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
        results = []
        for s, c in scored[:top_k]:
            c = c.copy()
            c["_score"] = s
            results.append(c)
        return results

    @staticmethod
    def _format_context(results: list[dict]) -> str:
        """将检索结果列表格式化为 LLM 上下文文本"""
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

    def build_context(self, query: str, top_k: int = 10) -> str:
        """构建发给 LLM 的上下文"""
        results = self.search(query, top_k)
        return self._format_context(results)

    def _web_search(self, query: str, max_results: int = 5) -> str:
        """联网搜索（Bing）并格式化为 LLM 上下文"""
        if not WEB_SEARCH_ENABLED:
            return ""

        # 检查缓存
        cache_key = query.strip().lower()
        if cache_key in self._web_cache:
            return self._web_cache[cache_key]

        try:
            import requests
            from bs4 import BeautifulSoup

            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                              "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            }
            resp = requests.get(
                "https://www.bing.com/search",
                params={"q": query, "count": max_results},
                headers=headers,
                timeout=10,
            )
            resp.raise_for_status()

            soup = BeautifulSoup(resp.text, "html.parser")
            results = []

            # Bing 搜索结果通常在 li.b_algo 元素中
            for item in soup.select("li.b_algo")[:max_results]:
                title_el = item.select_one("h2 a")
                if not title_el:
                    continue
                title = title_el.get_text(strip=True)
                href = title_el.get("href", "")
                body_el = item.select_one(".b_caption p, .b_lineclamp2, .b_algoSlug")
                body = body_el.get_text(strip=True) if body_el else ""
                if title:
                    results.append({"title": title, "href": href, "body": body})

            if not results:
                self._web_cache[cache_key] = ""
                return ""

            parts = ["## 网络搜索结果"]
            parts.append("以下是从互联网搜索到的相关信息：\n")
            for i, r in enumerate(results, 1):
                parts.append(f"### {i}. {r['title']}")
                if r["href"]:
                    parts.append(f"来源：{r['href']}")
                parts.append(f"{r['body']}\n")

            result = "\n".join(parts)
            self._web_cache[cache_key] = result
            return result
        except Exception:
            return ""  # 静默失败，回退到数据库模式

    @staticmethod
    def _merge_contexts(db_context: str, web_context: str) -> str:
        """合并数据库和网络上下文字段"""
        parts = []
        if db_context and db_context != "未在数据库中检索到相关数据。":
            parts.append(db_context)
        if web_context:
            parts.append(web_context)
        if not parts:
            return "未在数据库中检索到相关数据。"
        return "\n\n".join(parts)

    SCORE_THRESHOLD = CONFIG_SCORE_THRESHOLD  # 低于此分数触发联网搜索

    def ask(self, query: str) -> str:
        """
        执行一次问答：
        1. 检索本地数据库
        2. 判断是否需要联网搜索兜底
        3. 拼接上下文
        4. 调用 DeepSeek 生成回答
        """
        # 1. 检索本地数据库
        db_results = self.search(query, top_k=10)
        db_has_results = len(db_results) > 0
        max_score = max(r.get("_score", 0) for r in db_results) if db_results else 0
        needs_web = WEB_SEARCH_ENABLED and (not db_has_results or max_score < self.SCORE_THRESHOLD)

        # 2. 构建数据库上下文
        db_context = self._format_context(db_results)

        # 3. 必要时联网搜索
        web_context = ""
        if needs_web:
            web_context = self._web_search(query)

        # 4. 合并上下文
        context = self._merge_contexts(db_context, web_context)

        # 5. 构建消息
        source_desc = "## 检索来源说明"
        if web_context:
            source_desc += "\n- [数据库检索结果] 来自本地绿化工程造价指标数据库\n- [网络搜索结果] 来自互联网搜索，作为补充参考\n- 优先使用数据库数据，数据库无数据时可参考网络信息，并标注来源"

        messages = [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": f"""{source_desc}

## 检索结果
{context}

## 用户问题
{query}

请用专业简洁的语言回答。规则：
1. 优先引用数据库中的具体数据
2. 如果数据库没有相关数据，使用网络搜索结果作为参考
3. 如果两者都没有，如实告知并给出通用建议
4. 使用网络信息时请标注来源"""}
        ]

        # 6. 调用 DeepSeek
        try:
            response = self.client.chat.completions.create(
                model=DEEPSEEK_MODEL,
                messages=messages,
                temperature=0.3,
                max_tokens=2000,
            )
            return response.choices[0].message.content
        except Exception as e:
            return f"❌ 调用 DeepSeek API 出错：{str(e)}\n\n请检查 API Key 是否正确配置。"

    def chat(self, query: str, history: list[dict] = None) -> str:
        """
        多轮对话模式（保留对话历史）。
        本地数据库优先，查不到时自动联网搜索。
        """
        if history is None:
            history = []

        # 1. 检索本地数据库
        db_results = self.search(query, top_k=10)
        db_has_results = len(db_results) > 0
        max_score = max(r.get("_score", 0) for r in db_results) if db_results else 0
        needs_web = WEB_SEARCH_ENABLED and (not db_has_results or max_score < self.SCORE_THRESHOLD)

        # 2. 构建数据库上下文
        db_context = self._format_context(db_results)

        # 3. 必要时联网搜索
        web_context = ""
        if needs_web:
            web_context = self._web_search(query)

        # 4. 合并上下文
        context = self._merge_contexts(db_context, web_context)

        # 5. 构建消息
        messages = [{"role": "system", "content": SYSTEM_PROMPT}]

        # 添加历史对话
        for msg in history[-6:]:  # 最近 3 轮对话
            messages.append(msg)

        # 添加当前查询和检索结果
        source_note = ""
        if web_context:
            source_note = "\n（数据库检索结果来自本地绿化工程造价指标库，网络搜索结果来自互联网搜索作为补充。优先使用数据库数据，无数据时参考网络信息并标注来源。）"

        messages.append({
            "role": "user",
            "content": f"""## 检索结果
{context}
{source_note}

## 用户问题
{query}

请根据检索结果回答。规则：
1. 优先引用数据库中的具体数据
2. 数据库无数据时使用网络搜索结果作为参考
3. 两者都没有时如实告知并给出通用建议
4. 使用网络信息时请标注来源"""
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
