"""
知识库加载模块 — 解析 Markdown 格式的二类费规则文件
"""
from pathlib import Path


def load_knowledge_base(kb_dir: Path) -> list[dict]:
    """
    加载知识库目录下所有 .md 文件。
    返回结构：
    [
        {
            "title": "建设管理费",
            "content": "## 计算规则\n...",
            "source": "建设管理费.md"
        },
        ...
    ]
    """
    entries = []

    if not kb_dir.exists():
        return entries

    for filepath in sorted(kb_dir.glob("*.md")):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                content = f.read().strip()
        except Exception:
            continue

        if not content:
            continue

        # 从第一行 # 标题提取 title，或使用文件名
        title = filepath.stem
        lines = content.split("\n")
        if lines and lines[0].startswith("# "):
            title = lines[0][2:].strip()

        entries.append({
            "title": title,
            "content": content,
            "source": filepath.name,
        })

    return entries


def get_knowledge_chunks(entries: list[dict]) -> list[dict]:
    """
    将知识库条目转为可检索的 chunk 列表。
    每个条目可能按二级标题拆分为多个 chunk，提高检索精度。
    """
    chunks = []

    for entry in entries:
        title = entry["title"]
        content = entry["content"]
        source = entry["source"]

        # 按 ## 二级标题拆分
        sections = content.split("\n## ")
        if not sections:
            continue

        # 第一个 section 可能是标题行
        for i, section in enumerate(sections):
            section = section.strip()
            if not section:
                continue

            # 提取 section 标题
            section_lines = section.split("\n")
            if section_lines[0].startswith("# "):
                # 跳过主标题行，取下一行
                section_body = "\n".join(section_lines[1:]).strip()
                section_title = title
            elif i == 0 and not section.startswith("#"):
                section_body = section
                section_title = title
            else:
                # 二级标题 section
                first_line = section_lines[0]
                # 去掉可能的 # 前缀
                section_title_raw = first_line.lstrip("#").strip()
                section_body = "\n".join(section_lines[1:]).strip()
                section_title = f"{title} - {section_title_raw}"

            if section_body:
                # 构建可检索的文本
                search_text = f"【{title}】{section_title_raw if i > 0 else ''} {section_body}"
                chunks.append({
                    "text": search_text,
                    "title": section_title,
                    "content": section_body,
                    "source": source,
                    "full_entry": entry,  # 保留原始条目引用，用于显示完整内容
                    "_type": "knowledge",  # 标记类型，区分于指标数据
                })

    return chunks