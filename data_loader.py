"""
数据加载模块 - 解析绿化工程造价 CSV 数据
"""
import re
import pandas as pd
from pathlib import Path

# 数据目录（先读取你已有的 data 目录）
DEFAULT_DATA_DIR = Path(r"c:\Users\wangy\.claude\projects\data\绿化指标")


def _parse_embedded_dimensions(name: str) -> tuple[str, str, dict]:
    """
    从品种名中提取嵌入式规格信息。

    球类格式: "金叶女贞球G0.8-1，H0.8-1" → 纯名="金叶女贞球", 规格="冠幅0.8-1m, 高度0.8-1m"
    灌木格式: "木槿 地径5-5.9cm" → 纯名="木槿", 规格="地径5-5.9cm"
              "贴梗海棠 G0.6-0.8 H1.0-1.2m" → 纯名="贴梗海棠", 规格="冠幅0.6-0.8m, 高度1.0-1.2m"

    Returns:
        (pure_name, spec_string, dimensions_dict)
        dimensions_dict: {"冠幅": [(lo, hi, unit)], "高度": [...], "地径": [...]}
    """
    name = name.strip()
    dims: dict[str, list[tuple[float, float, str]]] = {}

    # === 组合模式（先匹配复合规格，避免被单一模式误匹配） ===

    # Pattern 1: 品种...Gx-y，Hx-y(m) — 球类常见（中文逗号分隔）
    m = re.match(
        r'^(.+?)\s*G(\d+\.?\d*)\s*[-~—]\s*(\d+\.?\d*)\s*[，,]\s*H(\d+\.?\d*)\s*[-~—]\s*(\d+\.?\d*)\s*(m)?\s*$',
        name
    )
    if m:
        pure = m.group(1).strip()
        g_lo, g_hi = float(m.group(2)), float(m.group(3))
        h_lo, h_hi = float(m.group(4)), float(m.group(5))
        unit = m.group(6) or "m"
        spec = f"冠幅{g_lo}-{g_hi}{unit}, 高度{h_lo}-{h_hi}{unit}"
        dims["冠幅"] = [(g_lo, g_hi, unit)]
        dims["高度"] = [(h_lo, h_hi, unit)]
        return pure, spec, dims

    # Pattern 2: 品种... Gx-y Hx-y(m) — 灌木（空格分隔 G+H）
    m = re.match(
        r'^(.+?)\s+G(\d+\.?\d*)\s*[-~—]\s*(\d+\.?\d*)\s+H(\d+\.?\d*)\s*[-~—]\s*(\d+\.?\d*)\s*(m)?\s*$',
        name
    )
    if m:
        pure = m.group(1).strip()
        g_lo, g_hi = float(m.group(2)), float(m.group(3))
        h_lo, h_hi = float(m.group(4)), float(m.group(5))
        unit = m.group(6) or "m"
        spec = f"冠幅{g_lo}-{g_hi}{unit}, 高度{h_lo}-{h_hi}{unit}"
        dims["冠幅"] = [(g_lo, g_hi, unit)]
        dims["高度"] = [(h_lo, h_hi, unit)]
        return pure, spec, dims

    # === 单一模式 ===

    # Pattern 3: 品种... 地径x-y(cm)
    m = re.match(
        r'^(.+?)\s+地径\s*(\d+\.?\d*)\s*[-~—]\s*(\d+\.?\d*)\s*(cm)?\s*$',
        name
    )
    if m:
        pure = m.group(1).strip()
        d_lo, d_hi = float(m.group(2)), float(m.group(3))
        unit = m.group(4) or "cm"
        spec = f"地径{d_lo}-{d_hi}{unit}"
        dims["地径"] = [(d_lo, d_hi, unit)]
        return pure, spec, dims

    # Pattern 4: 品种...Hx-y(m) — 高度范围
    m = re.match(
        r'^(.+?)\s*H[=]?\s*(\d+\.?\d*)\s*[-~—]\s*(\d+\.?\d*)\s*(m)?\s*$',
        name
    )
    if m:
        pure = m.group(1).strip()
        h_lo, h_hi = float(m.group(2)), float(m.group(3))
        unit = m.group(4) or "m"
        spec = f"高度{h_lo}-{h_hi}{unit}"
        dims["高度"] = [(h_lo, h_hi, unit)]
        return pure, spec, dims

    # Pattern 5: 品种...H=x(m) — 高度单值
    m = re.match(
        r'^(.+?)\s*H[=]?\s*(\d+\.?\d*)\s*(m)?\s*$',
        name
    )
    if m:
        pure = m.group(1).strip()
        h_val = float(m.group(2))
        unit = m.group(3) or "m"
        spec = f"高度{h_val}{unit}"
        dims["高度"] = [(h_val, h_val, unit)]
        return pure, spec, dims

    # Pattern 6: 品种...Gx-ym — 冠幅范围
    m = re.match(
        r'^(.+?)\s*G(\d+\.?\d*)\s*[-~—]\s*(\d+\.?\d*)\s*(m)?\s*$',
        name
    )
    if m:
        pure = m.group(1).strip()
        g_lo, g_hi = float(m.group(2)), float(m.group(3))
        unit = m.group(4) or "m"
        spec = f"冠幅{g_lo}-{g_hi}{unit}"
        dims["冠幅"] = [(g_lo, g_hi, unit)]
        return pure, spec, dims

    # Pattern 7: 品种...G xm — 冠幅单值（G前后可有空格，如"接骨木G 1.2m"、"红瑞木G1.2m"）
    m = re.match(
        r'^(.+?)\s*G\s*(\d+\.?\d*)\s*(m)?\s*$',
        name
    )
    if m:
        pure = m.group(1).strip()
        g_val = float(m.group(2))
        unit = m.group(3) or "m"
        spec = f"冠幅{g_val}{unit}"
        dims["冠幅"] = [(g_val, g_val, unit)]
        return pure, spec, dims

    # Pattern 8: 品种...Gxm — 冠幅单值（无空格，如"红瑞木G1-1.2m"会被#6匹配）
    # 但"丛生紫薇 G0.8-1.0m"这种Gx-y带小数的也会被#6匹配，无需额外处理

    # 无嵌入式规格
    return name, "", {}


def load_all_data(data_dir: Path = DEFAULT_DATA_DIR) -> dict:
    """
    加载所有 CSV 数据，返回结构化的数据字典。

    返回格式:
    {
        "常绿乔木": [
            {"品种": "白皮松", "规格": "H=3.0-3.49m", "栽植费用": 268.43, "苗木价格": 1450, "综合指标": 1834},
            ...
        ],
        "落叶乔木": [...],
        "灌木球类": [...]
    }
    """
    all_data = {}

    # 映射文件名 → (类别名, 单位)
    file_mapping = {
        "常绿乔木指标.csv": ("常绿乔木", "元/株"),
        "落叶乔木指标.csv": ("落叶乔木", "元/株"),
        "小乔木指标.csv":   ("小乔木",   "元/株"),
        "灌木指标.csv":     ("灌木",     "元/株"),
        "球类指标.csv":     ("灌木球类", "元/株"),
        "地被指标.csv":     ("地被",     "元/m²"),
        "绿篱指标.csv":     ("绿篱",     "元/m²"),
        "花卉指标.csv":     ("花卉",     "元/m²"),
    }

    for filename, (category, unit) in file_mapping.items():
        filepath = data_dir / filename
        if not filepath.exists():
            continue

        records = _parse_category_csv(filepath, category, unit)
        if records:
            all_data[category] = records

    return all_data


def _parse_category_csv(filepath: Path, category: str, unit: str = "元/株") -> list[dict]:
    """解析单个类别 CSV 文件"""
    records = []
    current_spec = ""  # 当前规格段，如 "H=3.0-3.49m"

    with open(filepath, "r", encoding="utf-8") as f:
        lines = f.readlines()

    for line in lines:
        line = line.strip()
        if not line:
            continue

        parts = [p.strip() for p in line.split(",")]

        # 跳过标题行和说明行
        if not parts or not parts[0]:
            continue
        if parts[0] in ["序号", "单位：元/株", "单位：元/m2"]:
            continue
        if "综合指标" in line and "栽植费用" not in line and "主材取费系数" not in line:
            continue
        if "说明" in line:
            continue

        # 检测规格段标题，如 "（一）,高度H=3.0-3.49m" 或 "（二）,胸径9.0-9.9cm"
        if parts[0].startswith("（"):
            # 提取规格描述
            for p in parts:
                p_clean = p.strip()
                if any(kw in p_clean for kw in ["高度", "胸径", "地径", "冠幅", "G"]):
                    current_spec = p_clean
                    break
                elif "H=" in p_clean or "φ" in p_clean:
                    current_spec = p_clean
                    break
            continue

        # 解析数据行：序号, 品种, 栽植费用, 苗木价格, 主材取费系数, 综合指标, 备注
        if len(parts) >= 6:
            try:
                name = parts[1]

                # 尝试解析数字
                planting_cost = _safe_float(parts[2])
                seedling_price = _safe_float(parts[3])
                coefficient = _safe_float(parts[4])
                comprehensive = _safe_float(parts[5])
                remark = parts[6] if len(parts) > 6 else ""

                if name and comprehensive is not None:
                    # 解析品种名中嵌入的规格（如 G冠幅，H高度，地径）
                    pure_name, embedded_spec, dims = _parse_embedded_dimensions(name)

                    # 合并节标题规格 + 嵌入式规格
                    if embedded_spec and current_spec:
                        final_spec = f"{current_spec}, {embedded_spec}"
                    elif embedded_spec:
                        final_spec = embedded_spec
                    else:
                        final_spec = current_spec

                    records.append({
                        "类别": category,
                        "品种": pure_name,
                        "品种_原始": name,
                        "规格": final_spec,
                        "维度": dims,
                        "栽植费用": planting_cost,
                        "苗木价格": seedling_price,
                        "主材取费系数": coefficient if coefficient else 1.0794,
                        "综合指标": comprehensive,
                        "单位": unit,
                        "备注": remark
                    })
            except (ValueError, IndexError):
                continue

    return records


def _safe_float(val: str) -> float | None:
    """安全转换为浮点数"""
    try:
        return float(val.strip())
    except (ValueError, AttributeError):
        return None


def get_dataframe(data: dict) -> pd.DataFrame:
    """将字典数据转为 pandas DataFrame"""
    all_records = []
    for category, records in data.items():
        all_records.extend(records)
    return pd.DataFrame(all_records)


def get_text_chunks(data: dict) -> list[dict]:
    """
    将数据转为文本块，用于检索。
    每个块是一条苗木的完整信息。
    """
    chunks = []
    for category, records in data.items():
        for r in records:
            unit = r.get("单位", "元/株")
            text = (
                f"【{category}】品种：{r['品种']}，"
                f"规格：{r['规格']}，"
                f"栽植费用：{r['栽植费用']}元，"
                f"苗木价格：{r['苗木价格']}元，"
                f"主材取费系数：{r['主材取费系数']}，"
                f"综合指标：{r['综合指标']}{unit}"
            )
            chunks.append({
                "text": text,
                "category": category,
                "name": r["品种"],
                "spec": r["规格"],
                "dims": r.get("维度", {}),
                "comprehensive": r["综合指标"],
                "苗木价格": r["苗木价格"],
                "栽植费用": r["栽植费用"],
                "主材取费系数": r["主材取费系数"],
                "unit": unit,
            })
    return chunks


if __name__ == "__main__":
    # 测试
    data = load_all_data()
    for cat, records in data.items():
        print(f"\n{'='*50}")
        print(f"类别: {cat}（共 {len(records)} 条）")
        print(f"{'='*50}")
        for r in records[:3]:
            print(f"  {r['品种']} | {r['规格']} | 综合指标: {r['综合指标']}元/株")
