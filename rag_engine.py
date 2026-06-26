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
    KNOWLEDGE_BASE_DIR,
)
from data_loader import load_all_data, get_text_chunks
from fee_engine import detect_and_calculate, format_for_llm
from knowledge_loader import (
    load_knowledge_base,
    search_knowledge_base,
    format_knowledge_context,
)
from config import KNOWLEDGE_BASE_DIR



def _expand_spec_range(query: str) -> dict[str, list[str]]:
    """
    将用户输入的规格数值展开为对应的数据区间。
    规则：胸径/地径 X cm → 区间 (X-1).0-(X-1).9
    例如："胸径12cm" → ranges: ["11.0-11.9", "11-11.9"], singles: ["11.0", "11"]

    高度/冠幅：直接提取数值作为 singles，用于匹配 H=、G 等规格文本。

    返回两个列表：
    - ranges:  范围格式，匹配全文（name+spec+text）
    - singles: 单数字格式，仅匹配 name+spec（防止在系数等无关数字上假阳性）
    """
    ranges = []
    singles = []

    # 匹配 "胸径12cm"、"地径7" 等模式
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

    # 高度/冠幅: 提取数值用于直接匹配规格文本
    for m in re.finditer(r'(?:高度|冠幅|H[=]?|G[=]?)\s*(\d+\.?\d*)\s*(?:m|米)?', query):
        try:
            num = float(m.group(1))
        except ValueError:
            continue
        # 生成多种文本形式用于匹配（如 H0.8-1 中的 "1", "0.8-1" 等）
        if num == int(num):
            singles.append(str(int(num)))
        singles.append(f"{num:.1f}")
        singles.append(f"{num:.2f}")

    return {"ranges": ranges, "singles": singles}


def _extract_query_dimensions(query: str) -> dict[str, list[float]]:
    """
    从用户自然语言查询中提取维度参数。

    返回: {"高度": [1.0], "冠幅": [0.8], "胸径": [12.0], "地径": [7.0]}
    数值单位：高度/冠幅为米，胸径/地径为厘米。
    """
    dims: dict[str, list[float]] = {}

    # 高度: "高度1m", "高度1米", "高度1.2", "H=1m", "H1m", "H=1"
    for m in re.finditer(r'(?:高度|H[=]?)\s*(\d+\.?\d*)\s*(?:m|米)?', query):
        dims.setdefault("高度", []).append(float(m.group(1)))

    # 冠幅: "冠幅0.8m", "冠幅0.8米", "G=0.8m", "G0.8", "冠幅0.8"
    for m in re.finditer(r'(?:冠幅|G[=]?)\s*(\d+\.?\d*)\s*(?:m|米)?', query):
        dims.setdefault("冠幅", []).append(float(m.group(1)))

    # 胸径: "胸径12cm", "胸径12厘米", "胸径12"
    for m in re.finditer(r'胸径\s*(\d+\.?\d*)\s*(?:cm|厘米)?', query):
        dims.setdefault("胸径", []).append(float(m.group(1)))

    # 地径: "地径7cm", "地径7厘米", "地径7"
    for m in re.finditer(r'地径\s*(\d+\.?\d*)\s*(?:cm|厘米)?', query):
        dims.setdefault("地径", []).append(float(m.group(1)))

    return dims


def _check_dimension_match(query_dims: dict[str, list[float]], chunk_dims: dict) -> int:
    """
    检查用户查询中的维度是否命中数据记录的维度区间。

    这是"给定一个维度参数，输出其他所有规格"的关键：
    用户说"高度1m" → 检查 1.0 是否落在某条记录的高度区间内（如 H0.8-1）。
    命中则给高分，使该记录排到检索结果前列，LLM 就能看到该记录的全部规格。

    返回匹配分数（0=未命中）。
    """
    if not query_dims or not chunk_dims:
        return 0

    score = 0
    for dim_type, query_values in query_dims.items():
        if dim_type in chunk_dims:
            for q_val in query_values:
                for lo, hi, unit in chunk_dims[dim_type]:
                    # 精确命中：用户值在记录的区间范围内
                    if lo <= q_val <= hi:
                        score += 12
                        break
                    # 接近命中：在 15% 容差范围内
                    margin = (hi - lo) * 0.15 if hi > lo else max(0.1, lo * 0.1)
                    if lo - margin <= q_val <= hi + margin:
                        score += 5
                        break

    return score


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

        # 加载二类费知识库
        self.knowledge = load_knowledge_base(KNOWLEDGE_BASE_DIR)
        

        # 构建品种名索引，用于快速匹配
        self._name_index = {}
        for chunk in self.chunks:
            name = chunk["name"]
            if name not in self._name_index:
                self._name_index[name] = []
            self._name_index[name].append(chunk)

        # 联网搜索缓存
        self._web_cache: dict[str, str] = {}

        # Load knowledge base
        self.knowledge_chunks = load_knowledge_base(KNOWLEDGE_BASE_DIR)

        pass  # Engine initialized

    def search(self, query: str, top_k: int = 10) -> list[dict]:
        """
        检索相关数据块。
        采用关键词匹配 + 品种名模糊匹配 + 规格区间智能映射 + 维度区间匹配。
        """
        scored = []
        expanded = _expand_spec_range(query)
        query_dims = _extract_query_dimensions(query)  # 从查询中提取维度参数

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
                # 放宽限制：1位数字也允许匹配（如高度1m中的"1"）
                if '.' in num or len(num) >= 1:
                    if num in name_spec_target:
                        score += 7
                        break

            # 3c. 维度区间匹配 — "给定一个参数，输出所有其他规格"的核心
            chunk_dims = chunk.get("dims", {})
            dim_score = _check_dimension_match(query_dims, chunk_dims)
            score += dim_score

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

        # ===== 知识库检索 =====
        kc_scored = []
        for kc in self.knowledge_chunks:
            kc_score = 0
            kc_text = kc.get("text", "")
            kc_title = kc.get("title", "")

            # A. 标题关键词匹配
            title_lower = kc_title.lower()
            query_lower = query.lower()
            # 标题整体出现在查询中
            if any(word in query for word in kc_title.replace("-", " ").replace("—", " ").split() if len(word) >= 2):
                kc_score += 8

            # B. 内容关键词匹配
            # 二类费相关关键词
            fee_keywords = ["勘察费", "工程勘察", "建设管理费", "勘察设计费", "监理费", "招标代理费", "可行性研究费",
                          "二类费", "管理费", "设计费", "监理", "招标", "可研", "代建",
                          "费率", "计费", "收费标准", "如何计算", "怎么算", "计算规则",
                          "勘察", "岩土工程", "水文地质",
                          "财建", "发改价格", "计价格"]
            for kw in fee_keywords:
                if kw in query and kw in kc_text:
                    kc_score += 5
                    break

            # C. 查询词直接匹配内容
            query_words = re.findall(r'[一-鿿]{2,}', query)
            for w in query_words:
                if w in kc_text:
                    kc_score += 2

            if kc_score > 0:
                kc_scored.append((kc_score, kc))

        kc_scored.sort(key=lambda x: x[0], reverse=True)

        # 合并两个来源的结果，知识库结果放入后段
        data_results = []
        for s, c in scored:
            c = c.copy()
            c["_score"] = s
            c["_source"] = "database"
            data_results.append(c)

        # 按分数排序后截取 top db_slots（修复：之前未排序就截取，导致取的是文件顺序前N条）
        data_results.sort(key=lambda x: x["_score"], reverse=True)

        kb_results = []
        for s, kc in kc_scored[:5]:  # 知识库最多 5 条
            kc = kc.copy()
            kc["_score"] = s
            kc["_source"] = "knowledge"
            kb_results.append(kc)

        # 合并：前 80% 给数据库，后 20% 给知识库（保证知识库结果能出现）
        db_slots = max(top_k - min(2, len(kb_results)), int(top_k * 0.8))
        data_results = data_results[:db_slots]
        kb_slots = top_k - len(data_results)
        kb_results = kb_results[:kb_slots]

        combined = data_results + kb_results
        combined.sort(key=lambda x: x["_score"], reverse=True)
        return combined[:top_k]

    @staticmethod
    def _format_context(results: list[dict]) -> str:
        """将检索结果列表格式化为 LLM 上下文文本，区分数据库和知识库"""
        if not results:
            return "未检索到相关数据。"

        db_results = [r for r in results if r.get("_source") != "knowledge"]
        kb_results = [r for r in results if r.get("_source") == "knowledge"]

        parts = []

        # ===== 数据库结果 =====
        if db_results:
            parts.append("## 数据库检索结果")
            parts.append("以下是从绿化工程造价指标数据库中检索到的相关数据：\n")

            by_category = {}
            for r in db_results:
                cat = r.get("category", "其他")
                if cat not in by_category:
                    by_category[cat] = []
                by_category[cat].append(r)

            for cat, items in by_category.items():
                parts.append(f"\n### {cat}")
                for item in items:
                    unit = item.get("unit", "元/株")
                    seedling_info = f"（其中苗木价格 {item.get('苗木价格', '')} 元）" if item.get('苗木价格') else ""
                    spec_str = f"（{item['spec']}）" if item.get('spec') else ""
                    parts.append(
                        f"- {item['name']}{spec_str}："
                        f"综合指标 **{item['comprehensive']}{unit}** {seedling_info}"
                    )
        else:
            parts.append("## 数据库检索结果")
            parts.append("未在数据库中检索到相关数据。")

        # ===== 知识库结果 =====
        if kb_results:
            parts.append("\n## 二类费规则参考")
            parts.append("以下是从工程造价二类费知识库中检索到的相关规则：\n")
            for i, item in enumerate(kb_results, 1):
                title = item.get("title", "未知条目")
                content = item.get("content", item.get("text", ""))
                source = item.get("source", "")
                parts.append(f"### {i}. {title}")
                if source:
                    parts.append(f"（来源：{source}）")
                parts.append(f"{content[:1500]}")  # 限制长度
                parts.append("")

        return "\n".join(parts)

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
        # 0. 二类费规则引擎检测
        fee_result = detect_and_calculate(query)
        fee_context = ""
        fee_instruction = ""
        if fee_result:
            fee_context = format_for_llm(fee_result)
            has_amount = fee_result.get("has_amount")
            is_sheji = fee_result.get("fee_type") == "工程设计费"
            if has_amount:
                if is_sheji:
                    fee_instruction = (
                        "\n\n> # ⚠️ 工程设计费 — 最高优先级指令\n\n"
                        "> 程序已完成精确计算，结果直接展示在界面上。\n"
                        "> **你只能确认上述金额和依据文件，禁止输出任何计算过程。**\n"
                        "> 禁止说\"无法计算\"\"数据不足\"\"需查表\"。\n\n"
                        "> ## 🚫 常见错误（禁止在确认中说这些）\n"
                        "> - **桥梁/地铁/隧道属于「交通运输工程」，不是建筑市政**\n"
                        "> - **园林绿化系数是 1.1，不是 1.2**\n"
                        "> - 确认时只说金额和依据，**禁止重新分类工程类型**"
                    )
                else:
                    fee_instruction = (
                        "\n\n> **⚠️ 程序已完成计算。你只能确认上述金额和依据文件。"
                        "禁止输出任何计算过程、查表步骤、内插法、公式。"
                        "禁止说\"无法计算\"\"数据不足\"\"需查表\"。**"
                    )
            else:
                is_kancha = fee_result.get("fee_type") == "勘察费"
                if is_sheji:
                    fee_instruction = (
                        "\n\n> # ⚠️ 工程设计费 — 最高优先级指令\n\n"
                        "> 附表一和附表二已由程序直接展示给用户，你**不需要也不允许**再输出任何表格。\n"
                        "> **严禁重新输出任何系数表或系数值！**\n"
                        "> 你只能做解释性描述，但决不能列出具体数字。\n\n"
                        "> ## 🚫 常见错误（训练数据经常搞错）\n"
                        "> - **桥梁工程属于「交通运输工程」，不是建筑市政**\n"
                        "> - **园林绿化系数是 1.1（不是 1.2）**\n"
                        "> - 用户问系数值 → 回答「请查阅上方程序生成的附表二」\n"
                        "> - 用户问分类归属 → 只能回答大类名称，**严禁说出系数数字**"
                    )
                elif is_kancha:
                    fee_instruction = (
                        "\n\n> # ⚠️ 工程勘察费 — 最高优先级指令\n\n"
                        "> **用户问的是勘察费！不是监理费！**\n"
                        "> 工程勘察费依据《工程勘察设计收费管理规定》（计价格[2002]10号）的"
                        "**工程勘察收费标准**部分计算，不是发改价格[2007]670号监理费！\n"
                        "> 上述费率表和计算规则来自政策文件引擎，你只能据此回答。\n\n"
                        "> ## 🚫 严禁以下行为\n"
                        "> - **严禁回答任何监理费相关内容**（用户没问监理费，建安费+设备费≠监理费）\n"
                        "> - **严禁提发改价格[2007]670号**（那是监理费文件，与勘察费无关）\n"
                        "> - **严禁提40%规则**（那是监理费的规则，勘察费不适用）\n"
                        "> - 严禁输出监理费数字或计算过程\n\n"
                        "> ## ✅ 正确做法\n"
                        "> - 说明勘察费按实物工作量定额计费（不是按投资额比例）\n"
                        "> - 列出计算所需参数：勘察类型、实物工作量、复杂程度\n"
                        "> - 给出行业经验估算范围供参考\n"
                        "> - 如需精确计算，请用户提供勘察类型和实物工作量"
                    )
                else:
                    fee_instruction = (
                        "\n\n> **[最高优先级] 上述费率表和调整系数来自政策文件引擎，数据权威准确。"
                        "你必须逐字引用检索结果中的数据，不得使用模型训练数据覆盖。"
                        "如需精确计算，请让用户提供具体金额参数。**"
                    )

        # 1. 检索本地数据库
        db_results = self.search(query, top_k=10)
        db_has_results = len(db_results) > 0
        max_score = max(r.get("_score", 0) for r in db_results) if db_results else 0
        needs_web = WEB_SEARCH_ENABLED and (not db_has_results or max_score < self.SCORE_THRESHOLD)

        # 2. 检索知识库
        kb_results = search_knowledge_base(query, self.knowledge_chunks, top_k=5)
        kb_context = format_knowledge_context(kb_results)

        # 3. 构建数据库上下文
        db_context = self._format_context(db_results)

        # 4. 必要时联网搜索
        web_context = ""
        if needs_web:
            web_context = self._web_search(query)

        # 5. 合并上下文
        normal_context = self._merge_contexts(db_context, web_context)

        if has_amount := (fee_result and fee_result.get("has_amount")):
            context = fee_context
        else:
            parts = []
            if fee_context:
                parts.append(fee_context)
            if kb_context and kb_context != "未在知识库中检索到相关政策文件。":
                parts.append(kb_context)
            if normal_context and normal_context != "未在数据库中检索到相关数据。":
                parts.append(normal_context)
            context = "\n\n---\n\n".join(parts) if parts else "未检索到相关数据。"

        # 6. 构建消息
        source_desc = "## 检索来源说明"
        if web_context:
            source_desc += "\n- [数据库检索结果] 来自本地绿化工程造价指标数据库\n- [知识库检索结果] 来自政策文件知识库\n- [网络搜索结果] 来自互联网搜索，作为补充参考"
        else:
            source_desc += "\n- [数据库检索结果] 来自本地绿化工程造价指标数据库\n- [知识库检索结果] 来自政策文件知识库"

        messages = [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": f"""{source_desc}

## 检索结果
{context}
{fee_instruction}

## 用户问题
{query}

请用专业简洁的语言回答。重要规则：
1. 【费率/计算规则/收费标准】必须严格使用上方检索结果中的数据，不要使用你训练数据中的旧版本信息
2. 优先引用数据库中的具体数据
3. 如果数据库没有相关数据，使用知识库搜索结果作为参考
4. 如果都没有，如实告知并给出通用建议
5. 使用网络信息时请标注来源"""}
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

        # 0. 二类费规则引擎检测
        fee_result = detect_and_calculate(query)
        fee_context = ""
        fee_instruction = ""
        if fee_result:
            fee_context = format_for_llm(fee_result)
            has_amount = fee_result.get("has_amount")
            is_sheji = fee_result.get("fee_type") == "工程设计费"
            if has_amount:
                if is_sheji:
                    fee_instruction = (
                        "\n\n> # ⚠️ 工程设计费 — 最高优先级指令\n\n"
                        "> 程序已完成精确计算，结果直接展示在界面上。\n"
                        "> **你只能确认上述金额和依据文件，禁止输出任何计算过程。**\n"
                        "> 禁止说\"无法计算\"\"数据不足\"\"需查表\"。\n\n"
                        "> ## 🚫 常见错误（禁止在确认中说这些）\n"
                        "> - **桥梁/地铁/隧道属于「交通运输工程」，不是建筑市政**\n"
                        "> - **园林绿化系数是 1.1，不是 1.2**\n"
                        "> - 确认时只说金额和依据，**禁止重新分类工程类型**"
                    )
                else:
                    fee_instruction = (
                        "\n\n> **⚠️ 程序已完成计算。你只能确认上述金额和依据文件。"
                        "禁止输出任何计算过程、查表步骤、内插法、公式。"
                        "禁止说\"无法计算\"\"数据不足\"\"需查表\"。**"
                    )
            else:
                is_kancha = fee_result.get("fee_type") == "勘察费"
                if is_sheji:
                    fee_instruction = (
                        "\n\n> # ⚠️ 工程设计费 — 最高优先级指令\n\n"
                        "> 附表一和附表二已由程序直接展示给用户，你**不需要也不允许**再输出任何表格。\n"
                        "> **严禁重新输出任何系数表或系数值！**\n"
                        "> 你只能做解释性描述，但决不能列出具体数字。\n\n"
                        "> ## 🚫 常见错误（训练数据经常搞错）\n"
                        "> - **桥梁工程属于「交通运输工程」，不是建筑市政**\n"
                        "> - **园林绿化系数是 1.1（不是 1.2）**\n"
                        "> - 用户问系数值 → 回答「请查阅上方程序生成的附表二」\n"
                        "> - 用户问分类归属 → 只能回答大类名称，**严禁说出系数数字**"
                    )
                elif is_kancha:
                    fee_instruction = (
                        "\n\n> # ⚠️ 工程勘察费 — 最高优先级指令\n\n"
                        "> **用户问的是勘察费！不是监理费！**\n"
                        "> 工程勘察费依据《工程勘察设计收费管理规定》（计价格[2002]10号）的"
                        "**工程勘察收费标准**部分计算，不是发改价格[2007]670号监理费！\n"
                        "> 上述费率表和计算规则来自政策文件引擎，你只能据此回答。\n\n"
                        "> ## 🚫 严禁以下行为\n"
                        "> - **严禁回答任何监理费相关内容**（用户没问监理费，建安费+设备费≠监理费）\n"
                        "> - **严禁提发改价格[2007]670号**（那是监理费文件，与勘察费无关）\n"
                        "> - **严禁提40%规则**（那是监理费的规则，勘察费不适用）\n"
                        "> - 严禁输出监理费数字或计算过程\n\n"
                        "> ## ✅ 正确做法\n"
                        "> - 说明勘察费按实物工作量定额计费（不是按投资额比例）\n"
                        "> - 列出计算所需参数：勘察类型、实物工作量、复杂程度\n"
                        "> - 给出行业经验估算范围供参考\n"
                        "> - 如需精确计算，请用户提供勘察类型和实物工作量"
                    )
                else:
                    fee_instruction = (
                        "\n\n> **[最高优先级] 上述费率表和调整系数来自政策文件引擎，数据权威准确。"
                        "你必须逐字引用检索结果中的数据，不得使用模型训练数据覆盖。"
                        "如需精确计算，请让用户提供具体金额参数。**"
                    )

        # 1. 检索本地数据库
        db_results = self.search(query, top_k=10)
        db_has_results = len(db_results) > 0
        max_score = max(r.get("_score", 0) for r in db_results) if db_results else 0
        needs_web = WEB_SEARCH_ENABLED and (not db_has_results or max_score < self.SCORE_THRESHOLD)

        # 2. 检索知识库
        kb_results = search_knowledge_base(query, self.knowledge_chunks, top_k=5)
        kb_context = format_knowledge_context(kb_results)

        # 3. 构建数据库上下文
        db_context = self._format_context(db_results)

        # 4. 必要时联网搜索
        web_context = ""
        if needs_web:
            web_context = self._web_search(query)

        # 5. 合并上下文
        normal_context = self._merge_contexts(db_context, web_context)

        if fee_result and fee_result.get("has_amount"):
            context = fee_context
        else:
            parts = []
            if fee_context:
                parts.append(fee_context)
            if kb_context and kb_context != "未在知识库中检索到相关政策文件。":
                parts.append(kb_context)
            if normal_context and normal_context != "未在数据库中检索到相关数据。":
                parts.append(normal_context)
            context = "\n\n---\n\n".join(parts) if parts else "未检索到相关数据。"

        # 6. 构建消息
        messages = [{"role": "system", "content": SYSTEM_PROMPT}]

        # 添加历史对话
        for msg in history[-6:]:  # 最近 3 轮对话
            messages.append(msg)

        # 添加当前查询和检索结果
        source_note = ""
        if web_context:
            source_note = "\n（数据库检索结果来自本地绿化工程造价指标库，知识库检索结果来自政策文件，网络搜索结果来自互联网搜索作为补充。优先使用数据库和知识库数据。）"

        messages.append({
            "role": "user",
            "content": f"""## 检索结果
{context}
{source_note}
{fee_instruction}

## 用户问题
{query}

请根据检索结果回答。重要规则：
1. 【费率/计算规则/收费标准】必须严格使用上方检索结果中的数据，不要使用你训练数据中的旧版本信息
2. 优先引用数据库和知识库中的具体数据
3. 数据库和知识库都无数据时使用网络搜索结果作为参考
4. 两者都没有时如实告知并给出通用建议
5. 使用网络信息时请标注来源"""
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
