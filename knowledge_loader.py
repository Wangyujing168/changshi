"""
知识库加载模块 — 加载 knowledge_base/ 中的 Markdown 政策文件
"""
import re
from pathlib import Path


def load_knowledge_base(kb_dir: Path | str) -> list[dict]:
    """
    加载知识库目录中的所有 .md 文件，按标题分块。

    返回 list[dict]，每个 dict:
        - title: 文件标题（取自 # 标题行）
        - content: 当前分块的文本内容
        - source_file: 来源文件名（不含路径）
        - chunk_index: 在同一文件中的分块序号（从 0 开始）
    """
    kb_dir = Path(kb_dir)
    if not kb_dir.exists():
        return []

    chunks: list[dict] = []

    for md_file in sorted(kb_dir.glob("*.md")):
        try:
            raw = md_file.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            raw = md_file.read_text(encoding="utf-8", errors="ignore")

        # 提取文件标题（第一个 # 标题）
        title_match = re.search(r'^#\s+(.+)$', raw, re.MULTILINE)
        file_title = title_match.group(1).strip() if title_match else md_file.stem

        # 按 ## 或 ### 标题分块
        sections = re.split(r'(?=^#{2,3}\s)', raw, flags=re.MULTILINE)

        chunk_index = 0
        for section in sections:
            # 提取节标题
            sec_title_match = re.search(r'^#{2,3}\s+(.+)$', section, re.MULTILINE)
            sec_title = sec_title_match.group(1).strip() if sec_title_match else file_title

            # 节内再按册/章标题切（如「第二册道路工程」「第三章 管道附属构筑物」），
            # 让大文件（预算基价章节说明）的块标题带上册章信息（2026-10-26 命中率修复）。
            # 只认 第X册/第X章/第X篇 行，且标题只取到首个句读为止，防止把正文整行带进标题
            parts = re.split(r'(?=^第[一二三四五六七八九十百]+[册篇章]\s*\S)',
                             section, flags=re.MULTILINE)
            for part in parts:
                chap_match = re.search(
                    r'^第[一二三四五六七八九十百]+[册篇章]\s*([^\s。；;]{1,40})',
                    part, re.MULTILINE)
                part_title = (f"{sec_title} / {chap_match.group(1).strip()}"
                              if chap_match and chap_match.group(1) != sec_title
                              else sec_title)

                content = part.strip()
                if not content:
                    continue

                # 如果分块太长，进一步按段落切分
                if len(content) > 3000:
                    sub_chunks = _split_long_section(content, max_chars=2000)
                    for sub in sub_chunks:
                        chunks.append({
                            "title": f"{file_title} / {part_title}",
                            "content": sub,
                            "source_file": md_file.name,
                            "chunk_index": chunk_index,
                        })
                        chunk_index += 1
                else:
                    chunks.append({
                        "title": f"{file_title} / {part_title}",
                        "content": content,
                        "source_file": md_file.name,
                        "chunk_index": chunk_index,
                    })
                    chunk_index += 1

    return chunks


def _split_long_section(text: str, max_chars: int = 2000) -> list[str]:
    """将过长的段落按语义边界切分为多个小块。"""
    sub_chunks = []
    paragraphs = text.split("\n\n")
    current = ""

    for para in paragraphs:
        if len(current) + len(para) < max_chars:
            current += "\n\n" + para if current else para
        else:
            if current:
                sub_chunks.append(current)
            current = para

    if current:
        sub_chunks.append(current)

    return sub_chunks if sub_chunks else [text]


# 领域同义词表：查询词 → 正文中可能出现的写法（用户 2026-10-26 反馈命中率低后补充，
# 覆盖桥梁定额匹配规则 + 土方系数等天津定额说明口径）
_SYNONYMS: dict[str, list[str]] = {
    "预算基价": ["预算基价", "基价", "定额", "消耗量"],
    "定额": ["定额", "预算基价", "基价", "消耗量"],
    "土方": ["土方", "土石方"],
    "挖方": ["挖方", "挖土", "开挖"],
    "填方": ["填方", "回填", "填压", "还土"],
    "买土": ["买土", "外购土", "借土", "购土"],
    "余土外运": ["运土", "弃运", "弃土", "运(弃)土", "余土", "外运", "弃方"],
    "外运": ["外运", "运土", "弃运", "弃土"],
    "自然方": ["自然方", "自然土方"],
    "压实方": ["压实方", "压实"],
    "松方": ["松方", "虚方"],
    "虚方": ["虚方", "松方"],
    "系数": ["系数", "换算"],
    "凿桩头": ["凿桩头", "凿除桩头", "凿除桩顶", "凿桩"],
    "护筒": ["护筒", "钢护筒", "埋设护筒"],
    "泥浆": ["泥浆", "触变泥浆"],
    "桩基": ["桩基", "桩基础", "灌注桩", "钻孔桩"],
    "灌注桩": ["灌注桩", "钻孔桩", "桩基", "桩基础"],
    "钢绞线": ["钢绞线", "预应力钢绞线"],
    "锚具": ["锚具", "OVM"],
    "泄水管": ["泄水管", "泄水孔", "排水管"],
    "预埋铁件": ["预埋铁件", "铁件"],
    "钢筋笼": ["钢筋笼"],
    "声测管": ["声测管", "检测管"],
    "桥梁": ["桥梁", "桥涵"],
    "桥涵": ["桥涵", "桥梁"],
    "市政": ["市政"],
    "跨径": ["跨径", "束长", "桥长"],
    "钻孔": ["钻孔", "钻机", "成孔"],
    "混凝土": ["混凝土", "砼"],
    "黏土": ["黏土", "粘土"],
    # 道路/市政常用术语（2026-10-26 命中率修复补充）
    "沥青": ["沥青", "柏油"],
    "水泥稳定碎石": ["水泥稳定碎", "水稳", "稳定碎"],
    "基层": ["基层", "底基层"],
    "路缘石": ["路缘石", "侧石", "缘石", "道牙", "路牙", "侧缘石"],
    "沟槽": ["沟槽", "开槽", "挖槽"],
    "管道": ["管道", "管线"],
    "摊铺": ["摊铺", "铺筑", "摊铺机"],
    "桥面": ["桥面", "桥面铺装", "桥面系"],
    "防水": ["防水", "防水层", "防水混凝土"],
    "模板": ["模板", "支模"],
    "支架": ["支架", "满堂支架", "脚手架"],
    "拆除": ["拆除", "破除"],
    "砌筑": ["砌筑", "安砌"],
    "检查井": ["检查井", "砌井"],
    "人行道": ["人行道", "步道"],
    "垫层": ["垫层"],
    "路面": ["路面", "面层"],
}

# 锚词：任何定额说明块都含有的泛化词。查询里同时存在具体术语时，
# 这些词不再参与打分（否则每块都有 +32 噪声地板，把区分词淹没）
_ANCHOR_KEYS = {"定额", "预算基价", "市政", "混凝土", "砼"}


def search_knowledge_base(query: str, chunks: list[dict], top_k: int = 5) -> list[dict]:
    """
    在知识库中搜索与查询相关的内容。

    评分策略（2026-10-26 第二次重写，解决命中率低问题）：
    1. 领域词反向激活：查询里出现同义词表中任一写法（键或变体，如「开挖」「回填」）
       即激活该术语组，命中标题 +8、命中正文 +4（覆盖 定额↔预算基价、
       凿桩头↔凿除桩头、买土↔外购土、松方↔虚方 等写法差异）；
    2. 查询中存在具体术语时，锚词（定额/市政/混凝土等泛化词）不参与打分，
       消除所有定额说明块共有的噪声地板；
    3. 未被术语覆盖的文本按 2~4 字 n-gram 匹配（标题 +2、正文 +1，封顶），
       让 水泥稳定碎石/路缘石/沟槽 等未收进词表的词也能区分块；
    4. 数字精确匹配（如 1.22/1.32/0.5/30）→ +6；
    5. 文件号（10号/670号…）→ +10；费种关键词 → +8，命中来源文件名再 +10；
    6. 地区倾向：查询指明省份时该省文件 +3；未指明省份且问题涉及定额/桥梁/土方等
       天津口径问题时，天津文件 +2（用户项目均适用天津市规则）。
    """
    q_norm = query.replace("粘", "黏")
    # 1) 领域术语反向激活：先按键词原词命中（字面优先），再按变体命中；
    #    同一查询片段只归属一个键词——防止「桩基」被「灌注桩」词组的变体
    #    抢先占位而丢掉字面命中（长词先匹配，覆盖区间防重复计数）
    terms: list[list[str]] = []
    term_keys: list[str] = []
    covered: list[tuple[int, int]] = []

    def _activate(key: str, v: str) -> None:
        for m in re.finditer(re.escape(v), q_norm):
            if any(a <= m.start() < b or a < m.end() <= b
                   for a, b in covered):
                continue
            terms.append([key] + _SYNONYMS[key])
            term_keys.append(key)
            covered.append((m.start(), m.end()))

    for key in sorted(_SYNONYMS, key=len, reverse=True):
        _activate(key, key)
    for key in sorted(_SYNONYMS, key=len, reverse=True):
        for v in sorted({v for v in _SYNONYMS[key] if v != key},
                        key=len, reverse=True):
            _activate(key, v)
    # 同一术语多次命中只保留一份（否则标题/正文重复计分膨胀）
    seen_keys = set()
    terms = [vs for vs, k in zip(terms, term_keys)
             if not (k in seen_keys or seen_keys.add(k))]
    # 具体术语存在时剔除锚词，避免泛化词噪声地板
    specific = any(k not in _ANCHOR_KEYS for k in term_keys)
    if specific:
        kept = [(k, vs) for vs, k in zip(terms, term_keys) if k not in _ANCHOR_KEYS]
    else:
        kept = list(zip(terms, term_keys))
    # 键词字面出现在查询中比仅靠同义变体命中更有区分度
    term_meta = [(k, vs, k in q_norm) for k, vs in kept]
    # 2) 未被术语覆盖的文本 → 2~4 字 n-gram（CJK 连续串内切）
    bg_parts = []
    pos = 0
    for a, b in sorted(covered):
        bg_parts.append(q_norm[pos:a])
        pos = b
    bg_parts.append(q_norm[pos:])
    bg_text = "".join(bg_parts)
    grams = []
    for run in re.findall(r'[一-鿿]+', bg_text):
        for nlen in (2, 3, 4):
            if len(run) >= nlen:
                grams.extend(run[i:i + nlen] for i in range(len(run) - nlen + 1))
    # 通用词停用：这些词在所有定额说明块里都出现，只加噪声不区分块。
    # 按子串过滤（而非精确匹配），堵住 3/4-gram 如「程量计算」「工程量计」这类
    # 由两个停用词拼接出的假命中词
    _GRAM_STOP = {"工程", "程量", "工程量", "计算", "公式", "怎么", "什么", "如何",
                  "请问", "多少", "依据", "应该", "可以", "规定", "标准", "方法",
                  "费用", "收费", "项目", "定额"}
    grams = [g for g in grams if not any(s in g for s in _GRAM_STOP)]
    # 3) 数字（含小数，如 1.22/1.32/0.5）
    nums = [n for n in re.findall(r'\d+\.\d+|\d+', query) if '.' in n or len(n) >= 2]
    file_nums = re.findall(r'\d+号', query)

    # 地区倾向：查询明确指定省份 → 只抬该省；否则问题涉及天津口径专业词时抬天津
    qprov = next((p for p in ("天津", "河北", "河南", "湖北") if p in q_norm), None)
    domain_kws = ("定额", "桥梁", "桥", "桩", "土方", "市政", "钢绞线", "预算基价",
                  "挖方", "填方", "买土", "凿桩", "护筒", "泥浆", "钢筋",
                  "道路", "基层", "沟槽", "管道", "路面", "沥青", "路缘石")

    fee_keywords = {
        "设计费": ["工程设计", "设计收费", "基本设计", "计价格[2002]10"],
        "勘察费": ["工程勘察", "勘察设计", "计价格[2002]10"],
        "监理费": ["监理", "670号", "发改价格[2007]"],
        "建设管理费": ["建设管理", "建设单位", "财建[2016]504"],
        "招标代理": ["招标代理", "1980号"],
        "交易服务": ["交易服务", "979号"],
        "可研": ["可行性研究", "前期工作", "1283号"],
        "施工图": ["施工图审查", "46号"],
        "水土保持": ["水土保持", "22号"],
        "环境影响": ["环境影响", "环评", "125号"],
    }

    scored = []
    for chunk in chunks:
        score = 0
        title = chunk.get("title", "").replace("粘", "黏")
        content = chunk.get("content", "").replace("粘", "黏")
        combined = title + " " + content
        src = chunk.get("source_file", "")

        # 1) 领域术语：标题 +8（字面命中）/+5（仅变体命中）；正文 +4，
        #    字面命中的键词按正文中出现的不同变体数再加 +2/个（封顶 3 个），
        #    正文出现键词原词本身再 +4（比只含泛化变体的块更有说服力）
        score += sum(8 if lit else 5
                     for k, vs, lit in term_meta if any(v in title for v in vs))
        for k, vs, lit in term_meta:
            chits = {v for v in vs if v in content}
            if chits:
                score += 4 + (2 * min(len(chits), 3) if lit else 0)
                if lit and k in content:
                    score += 4

        # 2) n-gram 兜底（标题 +2 封顶 4 个、正文 +1 封顶 6 个）
        bgt = sum(1 for kw in grams if kw in title)
        bgc = sum(1 for kw in grams if kw in content)
        score += min(bgt, 4) * 2 + min(bgc, 6) * 1

        # 3) 数字精确匹配（系数/规格问题：1.22、1.32、0.5、30…）
        for n in nums:
            if n in combined:
                score += 6
                break

        # 4) 文件号 + 费种关键词（命中来源文件名再加 +10，
        #    防止 716号 等只含「监理」字样的文件压过 670号 主文件）
        for fn in file_nums:
            if fn in combined:
                score += 10
        for fee_name, patterns in fee_keywords.items():
            if any(p in query for p in [fee_name] + patterns):
                for p in patterns:
                    if p in combined:
                        score += 8
                for p in patterns:
                    if p in src:
                        score += 10
                        break

        # 5) 定额说明来源加权 + 地区倾向
        quota_keywords = ["定额", "消耗量", "章节说明", "计算规则", "预算基价",
                          "栽植", "养护", "拆除工程", "消耗标准",
                          "凿桩", "护筒", "泥浆", "钢绞线", "买土", "土方", "桩",
                          "基层", "路面", "路缘石", "沟槽", "管道", "沥青", "道路"]
        if any(k in query for k in quota_keywords) \
                and src[:2] in ("天津", "河北", "河南", "湖北"):
            score += 6
        if qprov:
            if src.startswith(qprov):
                score += 3
        elif any(k in q_norm for k in domain_kws) and src.startswith("天津"):
            score += 2

        # 6) 桥梁定额规则权威源：查询涉及桥梁/桩基/土方等领域时，
        # 《天津_桥梁工程定额套用规则》是用户确认口径的规则汇总，优先置顶
        bridge_kws = ("桥梁", "桩", "护筒", "泥浆", "凿桩", "钢绞线", "锚具",
                      "泄水", "声测", "钢筋笼", "买土", "弃运", "外运", "土方",
                      "挖方", "填方", "模板", "预埋铁件", "钻孔", "灌注桩")
        if src == "天津_桥梁工程定额套用规则.md" and any(k in q_norm for k in bridge_kws):
            score += 15

        # 粘/黏 同义模糊：查询提黏土时，正文含「黏土」的分块直接加权
        if "黏" in q_norm and "黏土" in content:
            score += 8

        if score > 0:
            scored.append((score, chunk))

    scored.sort(key=lambda x: x[0], reverse=True)

    # 去重：每个文件最多 2 个分块，优先高分
    seen_files: dict[str, int] = {}
    deduped = []
    for score, chunk in scored:
        src = chunk.get("source_file", "")
        if seen_files.get(src, 0) < 2:
            deduped.append(chunk)
            seen_files[src] = seen_files.get(src, 0) + 1
        if len(deduped) >= top_k:
            break

    return deduped


def format_knowledge_context(results: list[dict]) -> str:
    """将知识库检索结果格式化为 LLM 可读的上下文。"""
    if not results:
        return "未在知识库中检索到相关政策文件。"

    parts = ["## 知识库检索结果（政策文件／定额说明原文）", ""]

    # 按来源文件分组
    by_file: dict[str, list[dict]] = {}
    for r in results:
        src = r.get("source_file", "未知")
        if src not in by_file:
            by_file[src] = []
        by_file[src].append(r)

    for src, items in by_file.items():
        parts.append(f"### 来源文件：{src}")
        for item in items:
            parts.append(f"**{item['title']}**")
            # 截断过长内容
            content = item["content"]
            if len(content) > 2500:
                content = content[:2500] + "\n\n...(内容过长，已截断。完整内容请查阅原文)"
            parts.append(content)
            parts.append("")

    return "\n".join(parts)