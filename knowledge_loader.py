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

            # 去除纯空行和纯标点行
            content = section.strip()
            if not content:
                continue

            # 如果分块太长，进一步按段落切分
            if len(content) > 3000:
                sub_chunks = _split_long_section(content, max_chars=2000)
                for sub in sub_chunks:
                    chunks.append({
                        "title": f"{file_title} / {sec_title}",
                        "content": sub,
                        "source_file": md_file.name,
                        "chunk_index": chunk_index,
                    })
                    chunk_index += 1
            else:
                chunks.append({
                    "title": f"{file_title} / {sec_title}",
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


def search_knowledge_base(query: str, chunks: list[dict], top_k: int = 5) -> list[dict]:
    """
    在知识库中搜索与查询相关的内容。
    使用简单的关键词匹配评分。
    """
    # 提取查询中的关键词（两字组合 + 文件号）
    keywords = re.findall(r'[一-鿿]{2}', query)  # 二字词

    scored = []
    for chunk in chunks:
        score = 0
        title = chunk.get("title", "")
        content = chunk.get("content", "")
        combined = title + " " + content

        # 标题匹配权重最高
        for kw in keywords:
            if kw in title:
                score += 5
            if kw in content:
                score += 1

        # 文件号匹配（如 "10号"、"670号"、"504号"）
        file_nums = re.findall(r'\d+号', query)
        for fn in file_nums:
            if fn in combined:
                score += 10

        # 费种关键词匹配
        fee_keywords = {
            "设计费": ["工程设计", "设计收费", "基本设计", "计价格[2002]10"],
            "监理费": ["监理", "670号", "发改价格[2007]"],
            "建设管理费": ["建设管理", "建设单位", "财建[2016]504"],
            "招标代理": ["招标代理", "1980号"],
            "交易服务": ["交易服务", "979号"],
            "可研": ["可行性研究", "前期工作", "1283号"],
            "施工图": ["施工图审查", "46号"],
            "水土保持": ["水土保持", "22号"],
            "环境影响": ["环境影响", "环评", "125号"],
        }
        for fee_name, patterns in fee_keywords.items():
            if any(p in query for p in [fee_name] + patterns):
                for p in patterns:
                    if p in combined:
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

    parts = ["## 知识库检索结果（政策文件原文）", ""]

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