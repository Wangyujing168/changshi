"""
二类费规则引擎 — 纯 Python 计算，不走 LLM，保证费率 100% 准确。

支持的费种：
- 建设管理费（财建[2016]504号）→ 差额分档累进
- 招标代理服务费（计价格[2002]1980号）→ 差额定率累进
- 工程建设交易服务费（津发改价管[2017]979号）→ 分档定额
- 施工监理服务费（发改价格[2007]670号）→ 基价内插 + 调整系数
"""

import re
import math
from typing import Any

# ============================================================
# 费率常量表（hardcoded，不依赖文件读取）
# ============================================================

# 财建[2016]504 号 建设管理费费率表（差额分档累进）
JIANSHE_GUANLI_RATES: list[tuple[float, float]] = [
    (1000, 2.0),       # ≤1000 万：2.0%
    (5000, 1.5),       # 1001-5000 万：1.5%
    (10000, 1.2),      # 5001-10000 万：1.2%
    (50000, 1.0),      # 10001-50000 万：1.0%
    (100000, 0.8),     # 50001-100000 万：0.8%
    (float("inf"), 0.4),  # >100000 万：0.4%
]

# 计价格[2002]1980 号 招标代理服务费费率表（差额定率累进）
# 格式：(上限, 货物招标%, 服务招标%, 工程招标%)
ZHAOBIAO_DAILI_RATES: list[tuple[float, float, float, float]] = [
    (100, 1.5, 1.5, 1.0),
    (500, 1.1, 0.8, 0.7),
    (1000, 0.8, 0.45, 0.55),
    (5000, 0.5, 0.25, 0.35),
    (10000, 0.25, 0.1, 0.2),
    (100000, 0.05, 0.05, 0.05),
    (float("inf"), 0.01, 0.01, 0.01),
]

# 津发改价管[2017]979 号 工程建设交易服务费（分档定额，单位：元）
JIAOYI_FUWU_RATES: list[tuple[float, float]] = [
    (100, 400),
    (500, 1000),
    (1000, 3000),
    (3000, 7000),
    (5000, 15000),
    (8000, 25000),
    (10000, 35000),
    (float("inf"), 50000),
]

# 发改价格[2007]670 号 施工监理服务收费基价表（单位：万元）
# 格式：(计费额万元, 收费基价万元)
JIANLI_BASE_RATES: list[tuple[float, float]] = [
    (500, 16.5),
    (1000, 30.1),
    (3000, 78.1),
    (5000, 120.8),
    (8000, 181.0),
    (10000, 218.6),
    (20000, 393.4),
    (40000, 708.2),
    (60000, 991.4),
    (80000, 1255.8),
    (100000, 1507.0),
    (200000, 2712.5),
    (400000, 4882.6),
    (600000, 6835.6),
    (800000, 8658.4),
    (1000000, 10390.1),
]
# 超过 1000000 万时，按 1.039% 收费率计算
JIANLI_LARGE_RATE = 1.039  # %

# 发改价格[2007]670号 附表三 施工监理服务收费专业调整系数
# 格式：(关键词正则, 系数)
JIANLI_PROFESSIONAL_COEFS: list[tuple[str, float]] = [
    ("园林|绿化", 0.8),
    ("矿山|采选", 0.9),
    ("核能|核电|水电|水库", 1.2),
    ("水运|地铁|桥梁|隧道|索道", 1.1),
    ("农业|林业|农林", 0.9),
    # 默认 1.0：建筑、人防、市政公用、石油化工、火电、送变电、
    #          铁路、公路、城市道路、轻轨、加工冶炼、船舶水工、水利 等
]


# 计价格[2002]10 号 工程设计收费基价表（附表一，单位：万元）
# 格式：(计费额万元, 收费基价万元)
SHEJI_BASE_RATES: list[tuple[float, float]] = [
    (200, 9.0),
    (500, 20.9),
    (1000, 38.8),
    (3000, 103.8),
    (5000, 163.9),
    (8000, 249.6),
    (10000, 304.8),
    (20000, 566.8),
    (40000, 1054.0),
    (60000, 1515.2),
    (80000, 1960.1),
    (100000, 2393.4),
    (200000, 4450.8),
    (400000, 8276.7),
    (600000, 11897.5),
    (800000, 15391.4),
    (1000000, 18793.8),
    (2000000, 34948.9),
]
SHEJI_LARGE_RATE = 1.6  # 超过 2000000 万时，计费额 × 1.6%

# 计价格[2002]10 号 工程设计收费专业调整系数（附表二，按优先级排序）
SHEJI_PROFESSIONAL_COEFS: list[tuple[str, float]] = [
    # 交通运输工程
    ("水运|地铁|桥梁|隧道", 1.1),
    ("索道", 1.3),
    ("机场空管|助航灯光|轻轨", 1.0),
    ("公路|城市道路", 0.9),
    ("机场场道", 0.8),
    # 建筑市政工程（排在石油化工前面，避免"园林绿化"中的"化"+"工"误匹配"化工"）
    ("人防|园林|绿化|广电", 1.1),
    ("邮政工艺", 0.8),
    ("建筑|市政|电信", 1.0),
    # 矿山采选工程
    ("选煤|煤炭工程", 1.3),
    ("采煤|铀矿", 1.2),
    ("矿山|采选|黑色|黄金|化学矿|非金属", 1.1),
    # 加工冶炼工程
    ("核加工", 1.3),
    ("冶炼|热加工|压力加工", 1.2),
    ("船舶水工", 1.1),
    ("冷加工", 1.0),
    # 石油化工工程（核化工必须排在化工前面）
    ("核化工", 1.6),
    ("石油|化工|石化|化纤|医药", 1.2),
    # 水利电力工程（核能必须在核电前面）
    ("核能工程", 1.6),
    ("核电常规岛|水电|水库|送变电", 1.2),
    ("火电", 1.0),
    ("风力发电|水利工程", 0.8),
    # 农业林业工程
    ("农业", 0.9),
    ("林业", 0.8),
]

# ============================================================
# 交互式系数选择 — 选项表（用于前端下拉菜单）
# ============================================================

# 监理费 专业调整系数选项（发改价格[2007]670号 附表三）
JIANLI_PROFESSIONAL_OPTIONS: list[tuple[str, float]] = [
    ("建筑、市政、公路、城市道路等（默认）", 1.0),
    ("园林绿化", 0.8),
    ("矿山采选", 0.9),
    ("农业、林业", 0.9),
    ("水运、地铁、桥梁、隧道、索道", 1.1),
    ("核能、核电、水电、水库", 1.2),
]

# 监理费 复杂程度系数选项（发改价格[2007]670号 1.0.9条）
JIANLI_COMPLEXITY_OPTIONS: list[tuple[str, float]] = [
    ("II级 / 较复杂（默认）", 1.0),
    ("I级 / 简单", 0.85),
    ("III级 / 复杂", 1.15),
]

# 监理费 高程调整系数选项（发改价格[2007]670号 1.0.9条）
JIANLI_ELEVATION_OPTIONS: list[tuple[str, float]] = [
    ("≤2000m（默认）", 1.0),
    ("2001~3000m", 1.1),
    ("3001~4000m", 1.2),
    (">4000m", 1.3),
]

# 设计费 专业调整系数选项（计价格[2002]10号 附表二）
SHEJI_PROFESSIONAL_OPTIONS: list[tuple[str, float]] = [
    ("建筑、市政、电信（默认）", 1.0),
    ("公路、城市道路", 0.9),
    ("水运、地铁、桥梁、隧道", 1.1),
    ("索道", 1.3),
    ("机场场道", 0.8),
    ("机场空管、助航灯光、轻轨", 1.0),
    ("人防、园林、绿化、广电", 1.1),
    ("邮政工艺", 0.8),
    ("石油、化工、石化、化纤、医药", 1.2),
    ("核化工", 1.6),
    ("核能工程", 1.6),
    ("核电常规岛、水电、水库、送变电", 1.2),
    ("火电", 1.0),
    ("风力发电、水利工程", 0.8),
    ("选煤、煤炭工程", 1.3),
    ("采煤、铀矿", 1.2),
    ("矿山、采选、黑色、黄金、化学矿、非金属", 1.1),
    ("冶炼、热加工、压力加工", 1.2),
    ("船舶水工", 1.1),
    ("核加工", 1.3),
    ("冷加工", 1.0),
    ("农业", 0.9),
    ("林业", 0.8),
]

# 设计费 复杂程度系数选项（计价格[2002]10号 1.0.9.2）
SHEJI_COMPLEXITY_OPTIONS: list[tuple[str, float]] = [
    ("II级 / 较复杂（默认）", 1.0),
    ("I级 / 一般", 0.85),
    ("III级 / 复杂", 1.15),
]

# 环评费 行业调整系数选项（计价格[2002]125号 附件二 表1）
HUANPING_INDUSTRY_OPTIONS: list[tuple[str, float]] = [
    ("市政（默认）", 1.0),
    ("建筑", 0.6),
    ("交通、铁道、民航、管线、建材、烟草、兵器", 1.0),
    ("林业、畜牧、渔业、农业", 1.0),
    ("石化、石油、天然气、水利、水电、旅游", 1.1),
    ("化工、冶金、有色、黄金、煤炭、矿产、纺织、化纤、轻工、医药", 1.2),
    ("邮电、广播电视、航空、机械、船舶、航天、电子、勘探、社会服务、火电", 0.8),
    ("粮食、信息产业、仓储", 0.6),
]

# 环评费 环境敏感程度系数选项（计价格[2002]125号 附件二 表2）
HUANPING_SENSITIVITY_OPTIONS: list[tuple[str, float]] = [
    ("未指定（默认）", 1.0),
    ("一般", 0.8),
    ("敏感", 1.2),
]

# 计价格[1999]1283 号 建设项目前期工作咨询收费标准（可行性研究费）
# 格式：{服务类型: [(投资下限_亿, 投资上限_亿, 费用下限_万, 费用上限_万)]}
# 注：<0.3亿项目由各省自行制定标准，此处按地方补充标准分段
KEYAN_BRACKETS: dict[str, list[tuple[float, float, float, float]]] = {
    "编制项目建议书": [
        (0.0,  0.05, 1.3, 1.3),   # <500万：固定 1.3 万
        (0.05, 0.1,  1.3, 2.5),   # 500~1000万：1.3~2.5 万
        (0.1,  0.3,  2.5, 6),     # 1000~3000万：2.5~6 万
        (0.3,  1.0,  6,   14),
        (1.0,  5.0,  14,  37),
        (5.0,  10.0, 37,  55),
        (10.0, 50.0, 55,  100),
        (50.0, float("inf"), 100, 125),
    ],
    "编制可研报告": [
        (0.0,  0.05, 2.5, 2.5),   # <500万：固定 2.5 万
        (0.05, 0.1,  2.5, 5),     # 500~1000万：2.5~5 万
        (0.1,  0.3,  5,   12),    # 1000~3000万：5~12 万
        (0.3,  1.0,  12,  28),
        (1.0,  5.0,  28,  75),
        (5.0,  10.0, 75,  110),
        (10.0, 50.0, 110, 200),
        (50.0, float("inf"), 200, 250),
    ],
    "评估项目建议书": [
        (0.0,  0.05, 1.0, 1.0),   # <500万：固定 1 万
        (0.05, 0.1,  1.0, 1.7),   # 500~1000万：1~1.7 万
        (0.1,  0.3,  1.7, 4.0),   # 1000~3000万：1.7~4 万
        (0.3,  1.0,  4,   8),
        (1.0,  5.0,  8,   12),
        (5.0,  10.0, 12,  15),
        (10.0, 50.0, 15,  17),
        (50.0, float("inf"), 17, 20),
    ],
    "评估可研报告": [
        (0.0,  0.05, 1.3, 1.3),   # <500万：固定 1.3 万
        (0.05, 0.1,  1.3, 2.5),   # 500~1000万：1.3~2.5 万
        (0.1,  0.3,  2.5, 5.0),   # 1000~3000万：2.5~5.0 万
        (0.3,  1.0,  5,   10),
        (1.0,  5.0,  10,  15),
        (5.0,  10.0, 15,  20),
        (10.0, 50.0, 20,  25),
        (50.0, float("inf"), 25, 35),
    ],
}
# 行业调整系数（以 1.0 为基准）
KEYAN_INDUSTRY_COEFS = {
    # 1. 石化、化工、钢铁 — 1.3
    "石化": 1.3, "化工": 1.3, "钢铁": 1.3,
    # 2. 石油、天然气、水利、水电、交通（水运）、化纤 — 1.2
    "石油": 1.2, "天然气": 1.2, "水利": 1.2, "水电": 1.2,
    "交通（水运）": 1.2, "水运": 1.2, "化纤": 1.2,
    # 3. 有色、黄金、纺织、轻工、邮电、广播电视、医药、煤炭、
    #    火电（含核电）、机械（含船舶、航空、航天、兵器）— 1.0
    "有色": 1.0, "黄金": 1.0, "纺织": 1.0, "轻工": 1.0, "邮电": 1.0,
    "广播电视": 1.0, "医药": 1.0, "煤炭": 1.0,
    "火电（含核电）": 1.0, "火电": 1.0, "核电": 1.0,
    "机械（含船舶、航空、航天、兵器）": 1.0, "机械": 1.0,
    "船舶": 1.0, "航空": 1.0, "航天": 1.0, "兵器": 1.0,
    # 4. 林业、商业、粮食、建筑 — 0.8
    "林业": 0.8, "商业": 0.8, "粮食": 0.8, "建筑": 0.8,
    # 5. 建材、交通（公路）、铁道、市政公用工程 — 0.7
    "建材": 0.7, "交通（公路）": 0.7, "公路": 0.7, "铁道": 0.7,
    "市政公用工程": 0.7, "市政": 0.7,
    # 市政公用工程子项（均属市政行业，系数 0.7）
    "轨道交通": 0.7, "公共交通": 0.7, "环境卫生": 0.7, "风景园林": 0.7,
    "给水工程": 0.7, "排水工程": 0.7, "燃气工程": 0.7, "热力工程": 0.7,
    "桥梁工程": 0.7, "道路工程": 0.7, "隧道工程": 0.7,
    "污水处理": 0.7, "垃圾处理": 0.7, "垃圾焚烧": 0.7, "垃圾填埋": 0.7,
    "供热工程": 0.7, "供热管网": 0.7, "热源厂": 0.7, "气源厂": 0.7,
    "净水厂": 0.7, "处理厂": 0.7, "泵站": 0.7, "BRT": 0.7, "快速公交": 0.7,
    "给水": 0.7, "排水": 0.7, "燃气": 0.7, "热力": 0.7,
    "桥梁": 0.7, "道路": 0.7, "隧道": 0.7, "供热": 0.7, "环卫": 0.7,
}

# 津价管[2011]46 号 施工图审查收费标准
# 住宅类：元/m²
SHIGONG_SHENCHA_ZHUZHAI: dict[str, float] = {"大型": 1.9, "中型": 1.7, "小型": 1.3}
# 公建/工业/市政类：以勘察设计费为基数的费率(%)
SHIGONG_SHENCHA_RATES: dict[str, dict[str, float]] = {
    "公建": {"大型": 3.2, "中型": 2.9, "小型": 2.4},
    "工业": {"大型": 3.2, "中型": 3.0, "小型": 2.8},
    "市政": {"大型": 4.8, "中型": 4.0, "小型": 3.2},
}
# 河北省施工图审查费（冀价行费[2018]57号 / 冀建质[2017]1号）
# 除有特殊规定外，按（勘察费+设计费）× 6.5%
HEBEI_SHENCHA_RATE = 6.5
# 幕墙/深基坑等单项：1.6‰，最低 1000 元

# 建市[2007]86号 工程设计资质标准 — 各行业大中小项目划分标准
# 用于施工图审查费的项目规模自动判定（津价管[2011]46号 第七条）
# 格式：(下限, 上限, 大型阈值, 中型上限) — 小于等于下限→小型，大于等于上限→大型，之间→中型

# 建筑行业（建筑工程）— 一般公共建筑：单体建筑面积(m²)
JIANZHU_GONGGONG_M2: tuple[float, float] = (5000, 20000)   # ≤5000小, 5000~20000中, ≥20000大
# 建筑行业（建筑工程）— 一般公共建筑：建筑高度(m)
JIANZHU_GONGGONG_H: tuple[float, float] = (24, 50)          # ≤24小, 24~50中, ≥50大
# 住宅/宿舍：层数
JIANZHU_ZHUZHAI_CENG: tuple[float, float] = (12, 20)        # ≤12小, 12~20中, >20大
# 住宅小区/工厂生活区：总建筑面积(m²)
JIANZHU_XIAOQU_M2: tuple[float, float] = (0, 300000)        # ≤30万中, >30万大（无小型）

# 市政行业 — 各子项规模划分（数据来源：建市[2007]86号 附件3-17 原文）
# 注意：道路/桥梁不在本表中，由 _detect_project_size_86 特殊处理
SHIZHENG_SCALE: dict[str, dict[str, tuple[float, float]]] = {
    "给水": {
        "净水厂": (5, 10),      # 万m³/日: <5小, 5~10中, ≥10大
        "泵站": (5, 20),        # 万m³/日: <5小, 5~20中, ≥20大
        "管道": (1000, 1600),   # 管径mm: <1000小, 1000~1600中, ≥1600大
    },
    "排水": {
        "处理厂": (4, 8),       # 万m³/日: <4小, 4~8中, ≥8大
        "泵站": (5, 10),        # 万m³/日: <5小, 5~10中, ≥10大
        "管道": (1000, 1500),   # 管径mm: ≤1000小, 1000~1500中, ≥1500大
    },
    "燃气": {
        "输配系统": (10000, 10000),  # 万m³/年: ≥10000大, <10000中
        "气源厂": (30, 30),          # 万m³/日: ≥30大, <30中
    },
    "热力": {
        "供热面积": (150, 500),      # 万m²: <150小, 150~500中, ≥500大
    },
    "隧道": {
        "隧道": (1000, 1000),        # 原文：全部大型，但保留数值用于兜底判定
    },
    "风景园林": {
        "风景园林": (100, 1000),     # 万元: ≤100小, 100~1000中, >1000大
    },
    "环境卫生": {
        "填埋": (200, 500),           # 吨/日: <200小, 200~500中, ≥500大
        "转运站": (150, 400),         # 吨/日: <150小, 150~400中, ≥400大
        "医疗废弃物": (5, 5),         # 吨/日: ≥5大, <5中
    },
}

# 工业行业 — 按投资额划分（简化，实际按细分行业有不同标准）
GONGYE_SCALE: tuple[float, float] = (3000, 10000)  # 万元: ≤3000小, 3000~10000中, ≥10000大

# 保监[2005]22 号 水土保持咨询服务费（单位：土建投资_亿元, 费用_万元）
SHUIBAO_TUDI_TOUZI: list[float] = [0.5, 1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0,
                                   11.0, 12.0, 13.0, 14.0, 15.0, 16.0, 17.0, 18.0, 19.0, 20.0]
SHUIBAO_BIANZHI: list[float] = [30, 52, 72, 82, 95, 104, 116, 119, 132, 156, 171,
                                 185, 200, 220, 230, 245, 259, 270, 290, 320, 350]  # 方案编制费
SHUIBAO_JIANCE: list[float] = [30, 60, 90, 140, 180, 220, 275, 310, 350, 385, 420,
                                460, 490, 525, 560, 600, 640, 680, 710, 735, 760]  # 监测费
SHUIBAO_PINGGU: list[float] = [10, 18, 30, 36, 42, 48, 54, 60, 66, 72, 78,
                                84, 107, 111, 116, 119, 126, 130, 144, 150, 160]  # 验收评估费
SHUIBAO_CONSULT: list[float] = [1.0, 1.5, 2.0, 2.5, 2.9, 3.2, 3.5, 3.8, 4.0, 4.8, 5.2,
                                 5.6, 6.0, 6.5, 7.0, 7.5, 7.8, 8.3, 8.5, 9.0, 9.5]  # 技术咨询费


# ============================================================
# 造价咨询费 — 津价房地[2008]136号
# ============================================================

# 差额定率分档累进，分档（万元）
_COST_CONSULTING_BRACKETS: list[float] = [100, 500, 1000, 5000, 10000, float("inf")]

# 费率表（单位：‰）
_COST_CONSULTING_RATES: dict[str, dict] = {
    # ── 编制类（基数 = 工程费用 = 建安+设备）──
    "编制工程量清单": {
        "rates": [3.4, 3.2, 3.0, 2.4, 2.0, 1.6],
        "base_type": "工程费用",
        "category": "编制",
    },
    "编制标底(含清单)": {
        "rates": [3.6, 3.4, 3.1, 2.6, 2.0, 1.7],
        "base_type": "工程费用",
        "category": "编制",
    },
    "编制施工图预算": {
        "rates": [3.6, 3.4, 3.1, 2.6, 2.0, 1.7],
        "base_type": "工程费用",
        "category": "编制",
    },
    "编制竣工结算": {
        "rates": [3.6, 3.4, 3.1, 2.6, 2.0, 1.7],
        "base_type": "工程费用",
        "category": "编制",
    },
    "施工阶段全过程造价控制": {
        "rates": [10.0, 9.0, 8.0, 7.5, 7.0, 6.0],
        "base_type": "工程费用",
        "category": "编制",
    },
    # ── 审核类 ──
    "审核概算": {
        "rates": [3.0, 2.5, 2.0, 1.5, 1.2, 1.0],
        "base_type": "工程总投资",   # 唯一以总投资为基数的子项
        "category": "审核",
    },
    "审核预算、标底": {
        "rates": [3.5, 3.1, 2.2, 1.9, 1.2, 0.9],
        "base_type": "工程费用",
        "category": "审核",
    },
    "审核竣工结算": {
        "rates": [3.5, 3.1, 2.2, 1.9, 1.2, 0.9],
        "base_type": "工程费用",
        "category": "审核",
    },
    # ── 其他（基数 = 建安工程费用）──
    "编制项目投资估算": {
        "rates": [0.8, 0.7, 0.6, 0.5, 0.3, 0.15],
        "base_type": "建安工程费用",
        "category": "编制",
    },
    "编制设计概算": {
        "rates": [1.7, 1.5, 1.2, 0.85, 0.7, 0.4],
        "base_type": "建安工程费用",
        "category": "编制",
    },
}

# 服务类型分组展示顺序
_COST_CONSULTING_SERVICE_ORDER: list[str] = [
    "编制工程量清单",
    "编制标底(含清单)",
    "编制施工图预算",
    "编制竣工结算",
    "施工阶段全过程造价控制",
    "审核概算",
    "审核预算、标底",
    "审核竣工结算",
    "编制项目投资估算",
    "编制设计概算",
]

# ============================================================
# 造价咨询费 — 冀建市研[2017]2号（河北省）
# ============================================================

# 河北省差额定率分档，分档（万元）
_HEBEI_COST_CONSULTING_BRACKETS: list[float] = [200, 500, 2000, 10000, float("inf")]

# 河北省费率表（单位：‰）
# 备注：工程建设设备费用不计入取费基数（建安费为基数，部分类型用总投资/概算额）
_HEBEI_COST_CONSULTING_RATES: dict[str, dict] = {
    # ── 编制类 ──
    "投资估算": {
        "rates": [0.8, 0.7, 0.6, 0.5, 0.3],
        "base_type": "投资估算造价",
        "base_from": "total_investment",
        "category": "编制",
    },
    "经济评价": {
        "rates": [0.8, 0.7, 0.6, 0.5, 0.3],
        "base_type": "投资估算造价",
        "base_from": "total_investment",
        "category": "编制",
    },
    "概算编制": {
        "rates": [3.0, 2.5, 2.0, 1.8, 1.6],
        "base_type": "设计概算造价",
        "base_from": "total_investment",
        "category": "编制",
    },
    "预算编制": {
        "rates": [4.0, 3.5, 3.0, 2.5, 2.0],
        "base_type": "建安费",
        "base_from": "jianan",
        "category": "编制",
    },
    "工程量清单编制(审核)": {
        "rates": [5.0, 4.0, 3.0, 2.2, 1.8],
        "base_type": "建安费",
        "base_from": "jianan",
        "category": "编制",
    },
    "招标控制价编制(审核)": {
        "rates": [2.0, 1.8, 1.6, 1.4, 1.2],
        "base_type": "建安费",
        "base_from": "jianan",
        "category": "编制",
    },
    "结算编制": {
        "rates": [5.0, 4.5, 4.0, 3.5, 3.0],
        "base_type": "建安费",
        "base_from": "jianan",
        "category": "编制",
    },
    "竣工决算编制": {
        "rates": [2.0, 1.5, 1.2, 1.0, 0.8],
        "base_type": "建设项目总投资",
        "base_from": "total_investment",
        "category": "编制",
    },
    # ── 审核类 ──
    "概算审核": {
        "rates": [2.0, 1.8, 1.5, 1.2, 1.0],
        "base_type": "设计概算造价",
        "base_from": "total_investment",
        "category": "审核",
    },
    "预算审核": {
        "rates": [3.2, 2.8, 2.4, 2.0, 1.6],
        "base_type": "建安费",
        "base_from": "jianan",
        "category": "审核",
    },
    "结算审核": {
        "rates": [4.5, 4.0, 3.5, 3.0, 2.5],
        "base_type": "建安费",
        "base_from": "jianan",
        "category": "审核",
        "note": "计费=(1)基本收费+(2)效益收费。效益收费=(核减+核增)×8%",
    },
    # ── 其他 ──
    "投标报价分析(清标)": {
        "rates": [0.6, 0.5, 0.4, 0.3, 0.2],
        "base_type": "最高投标限价",
        "base_from": "jianan",
        "category": "其他",
    },
    # 施工阶段造价咨询：≤500万不适用（费率=0），实际从500万起算
    "施工阶段造价咨询": {
        "rates": [0.0, 0.0, 8.0, 6.0, 4.0],
        "base_type": "建安费",
        "base_from": "jianan",
        "category": "全过程",
        "note": "计费=(1)基本收费+(2)效益收费。X≤500万不建议采用此模式，费率=0；效益收费=(核减+核增)×5%；驻场人员10000~20000元/月",
    },
    "全过程造价咨询": {
        "rates": [0.0, 0.0, 13.0, 10.0, 8.0],
        "base_type": "建安费",
        "base_from": "jianan",
        "category": "全过程",
        "note": "计费=(1)基本收费+(2)效益收费。X≤500万不建议采用，费率=0；驻场人员10000~20000元/月；效益收费=(核减+核增)×5%",
    },
    "工程造价鉴定": {
        "rates": [12.0, 10.0, 8.0, 6.0, 5.0],
        "base_type": "鉴定标的额",
        "base_from": "total_investment",
        "category": "其他",
    },
}

# 河北省造价咨询服务类型展示顺序
_HEBEI_COST_CONSULTING_SERVICE_ORDER: list[str] = [
    "投资估算",
    "经济评价",
    "概算编制",
    "概算审核",
    "预算编制",
    "预算审核",
    "工程量清单编制(审核)",
    "招标控制价编制(审核)",
    "投标报价分析(清标)",
    "结算编制",
    "结算审核",
    "竣工决算编制",
    "施工阶段造价咨询",
    "全过程造价咨询",
    "工程造价鉴定",
]

# 河北省专业工程调整系数（冀建市研[2017]2号 附件2）
_HEBEI_PROFESSIONAL_COEFFICIENTS: dict[str, float] = {
    "房屋建筑工程": 0.8,
    "水利电力工程": 0.9,
    "交通建设工程": 0.7,
    "公路工程": 0.8,
    "铁路工程": 0.8,
    "市政工程": 0.8,
    "园林绿化工程": 0.7,
    "港口工程": 0.8,
    "矿山工程": 1.1,
    "园林景观工程": 1.1,
    "装饰装修工程": 1.2,
    "古建筑工程": 1.2,
    "安装工程": 1.2,
    "其他工程": 1.0,
}

# 河北省造价咨询费最低收费标准（元）
_HEBEI_COST_CONSULTING_MIN_FEE = 3000.0


def _detect_hebei_cost_consulting_type(query: str) -> str | None:
    """从查询中检测河北省造价咨询的具体服务类型（冀建市研[2017]2号）。"""
    type_patterns: list[tuple[str, str]] = [
        ("施工阶段造价咨询", "施工阶段造价咨询"),
        ("全过程造价咨询", "全过程造价咨询"),
        ("全过程.*控制", "全过程造价咨询"),
        ("工程量清单.*编制.*审核|工程量清单.*审核", "工程量清单编制(审核)"),
        ("工程量清单编制", "工程量清单编制(审核)"),
        ("工程量清单.*编制", "工程量清单编制(审核)"),
        ("招标控制价.*编制.*审核|招标控制价.*审核", "招标控制价编制(审核)"),
        ("招标控制价编制", "招标控制价编制(审核)"),
        ("招标控制价", "招标控制价编制(审核)"),
        ("投标报价分析|清标", "投标报价分析(清标)"),
        ("投资估算", "投资估算"),
        ("经济评价|国民经济评价|财务评价", "经济评价"),
        ("概算编[制审]|编制.*概算", "概算编制"),
        ("概算审核", "概算审核"),
        ("概算", "概算编制"),
        ("预算编[制审]|编制.*预算|施工图预算", "预算编制"),
        ("预算审核", "预算审核"),
        ("结算编[制审]|编制.*结算|竣工结算编", "结算编制"),
        ("结算审核|竣工结算审", "结算审核"),
        ("竣工决算编|编制.*决算", "竣工决算编制"),
        ("竣工决算", "竣工决算编制"),
        ("工程造价鉴定|造价鉴定|工程鉴定", "工程造价鉴定"),
        ("造价咨询", "预算编制"),  # 泛化默认
    ]
    for pattern, svc_type in type_patterns:
        if re.search(pattern, query):
            return svc_type
    return None


def calc_cost_consulting_hebei(
    jianan_wan: float,
    service_type: str,
    total_investment: float | None = None,
    professional_coef: float = 1.0,
    discount_coef: float = 1.0,
) -> dict:
    """
    河北省造价咨询费（冀建市研[2017]2号）。

    与津价房地[2008]136号的核心差异：
      - 计费基数为建安费（不含设备费）
      - 差额定率分档不同：200/500/2000/10000
      - 专业工程调整系数（附件2）
      - 最低收费 3000 元
      - 下浮不超过 20%
    """
    config = _HEBEI_COST_CONSULTING_RATES.get(service_type)
    if config is None:
        raise ValueError(f"未知的河北省造价咨询服务类型：{service_type}")

    rates: list[float] = config["rates"]
    base_type: str = config["base_type"]
    base_from: str = config["base_from"]
    note: str = config.get("note", "")

    # 确定计费基数
    if base_from == "total_investment":
        if total_investment is None:
            amount = jianan_wan  # fallback
        else:
            amount = total_investment
    else:
        amount = jianan_wan  # 默认建安费

    # 使用实际费率数量对应的分档（施工阶段/全过程只有3档）
    num_tiers = len(rates)
    brackets = _HEBEI_COST_CONSULTING_BRACKETS[:num_tiers]
    # 最后一个分档需要设为 inf 以保证剩余金额全部计入
    brackets[-1] = float("inf")

    # 差额分档累进
    total_fee = 0.0
    prev_limit = 0.0
    steps: list[dict] = []

    for i, limit in enumerate(brackets):
        if amount <= prev_limit:
            break
        tier_amount = min(amount, limit) - prev_limit
        if tier_amount <= 0:
            prev_limit = limit
            continue
        rate = rates[i]
        tier_fee = tier_amount * rate / 1000.0  # ‰ → 万元
        total_fee += tier_fee
        step_qujian = f"{prev_limit:.0f}~{limit:.0f}" if limit != float("inf") else f">{prev_limit:.0f}"
        steps.append({
            "区间": step_qujian,
            "金额(万元)": round(tier_amount, 2),
            "费率(‰)": rate,
            "费用(万元)": round(tier_fee, 4),
        })
        prev_limit = limit

    # 应用专业调整系数
    total_fee_before_prof = round(total_fee, 4)
    total_fee = round(total_fee * professional_coef, 4)

    # 最低收费检查（3000元 = 0.3万元）
    # 仅当实际计算费用 > 0 且低于最低标准时才适用（费率=0的档位不触发最低收费）
    min_fee_wan = _HEBEI_COST_CONSULTING_MIN_FEE / 10000.0
    applied_min = total_fee > 0 and total_fee < min_fee_wan
    if applied_min:
        total_fee = round(min_fee_wan, 4)

    # 打折系数
    total_fee = round(total_fee * discount_coef, 4)

    # 构建说明
    lines = [f"计费基数 {amount:.0f} 万元（{base_type}），{service_type}"]
    if abs(professional_coef - 1.0) > 0.001:
        lines.append(f"专业调整系数 {professional_coef}，调整前 {total_fee_before_prof:.2f} 万元")
    if applied_min:
        lines.append(f"计算费用不足最低收费标准 {_HEBEI_COST_CONSULTING_MIN_FEE:.0f} 元，按最低标准收取")
    if abs(discount_coef - 1.0) > 0.001:
        lines.append(f"打折系数 {discount_coef:.2f}")
    if note:
        lines.append(f"注：{note}")

    return {
        "费种": f"造价咨询费（{service_type}）",
        "依据": "《河北省建设工程造价咨询服务收费管理暂行办法》（冀建市研[2017]2号）",
        "计算公式": "差额分档累进，基准价可下浮 ≤20%",
        "参数": {
            "计费基数(万元)": round(amount, 4),
            "基数类型": base_type,
            "服务类型": service_type,
            "专业调整系数": professional_coef,
            "打折系数": discount_coef,
            "最低收费(元)": _HEBEI_COST_CONSULTING_MIN_FEE,
        },
        "计算步骤": steps,
        "结果(万元)": total_fee,
        "说明": (
            f"计费基数 {amount:.0f} 万元（{base_type}），{service_type}，\n"
            f"差额分档累进计算，造价咨询费基准价为 **{total_fee:.2f} 万元**"
            f"（可在下浮 ≤20% 幅度内协商浮动）"
        ),
    }


def calc_cost_consulting_multi_hebei(
    selected_services: list[str],
    jianan_wan: float,
    total_investment: float | None = None,
    professional_coef: float = 1.0,
    discount_coef: float = 1.0,
) -> dict:
    """计算多个河北省造价咨询服务子项的费用（每项独立计算，汇总求和）。"""
    details: list[dict] = []
    total = 0.0
    warnings: list[str] = []
    for svc in selected_services:
        try:
            single = calc_cost_consulting_hebei(
                jianan_wan, svc,
                total_investment=total_investment,
                professional_coef=professional_coef,
                discount_coef=1.0,  # 单个不折，最后统一折
            )
            fee = single["结果(万元)"]
            total += fee
            details.append({
                "服务类型": svc,
                "计费基数(万元)": single["参数"]["计费基数(万元)"],
                "基数类型": single["参数"]["基数类型"],
                "费用(万元)": fee,
                "计算步骤": single["计算步骤"],
            })
        except (ValueError, KeyError) as e:
            warnings.append(f"⚠️ **{svc}**：{e}")
    total_before_discount = round(total, 4)
    total = round(total_before_discount * discount_coef, 4)
    desc = f"共 {len(details)} 项服务，合计 **{total_before_discount} 万元**"
    if abs(discount_coef - 1.0) > 0.001:
        desc += f"，打折后 **{total} 万元**（系数 {discount_coef:.2f}）"
    if abs(professional_coef - 1.0) > 0.001:
        desc += f"\n\n专业调整系数：**{professional_coef}**"
    if warnings:
        desc += "\n\n" + "\n\n".join(warnings)
    return {
        "明细": details,
        "合计(万元)": total,
        "合计(打折前)(万元)": total_before_discount,
        "参数": {
            "建安工程造价(万元)": round(jianan_wan, 4),
            "工程总投资(万元)": round(total_investment, 4) if total_investment else None,
            "选中服务": selected_services,
            "专业调整系数": professional_coef,
            "打折系数": discount_coef,
            "警告": warnings,
        },
        "费种": "造价咨询费",
        "依据": "《河北省建设工程造价咨询服务收费管理暂行办法》（冀建市研[2017]2号）",
        "说明": desc,
    }


def calc_cost_consulting_multi(
    selected_services: list[str],
    base_amount_wan: float,
    jianan_only: float | None = None,
    total_investment: float | None = None,
) -> dict:
    """计算多个造价咨询服务子项的费用（每项独立计算，汇总求和）。

    返回结构：
        {"明细": [...], "合计(万元)": float, "参数": {...}}
    """
    details: list[dict] = []
    total = 0.0
    warnings: list[str] = []
    for svc in selected_services:
        try:
            single = calc_cost_consulting(
                base_amount_wan, svc,
                total_investment=total_investment,
                jianan_only=jianan_only,
            )
            fee = single["结果(万元)"]
            total += fee
            details.append({
                "服务类型": svc,
                "计费基数(万元)": single["参数"]["计费基数(万元)"],
                "基数类型": single["参数"]["基数类型"],
                "费用(万元)": fee,
                "计算步骤": single["计算步骤"],
            })
        except (ValueError, KeyError) as e:
            warnings.append(f"⚠️ **{svc}**：{e}")
    total = round(total, 4)
    desc = f"共 {len(details)} 项服务，合计 **{total} 万元**"
    if warnings:
        desc += "\n\n" + "\n\n".join(warnings)
    return {
        "明细": details,
        "合计(万元)": total,
        "参数": {
            "工程费用(万元)": round(base_amount_wan, 4),
            "建安工程费用(万元)": round(jianan_only, 4) if jianan_only else None,
            "工程总投资(万元)": round(total_investment, 4) if total_investment else None,
            "选中服务": selected_services,
            "警告": warnings,
        },
        "费种": "造价咨询费",
        "依据": "《天津市建设工程造价咨询服务项目和价格标准》（津价房地[2008]136号）",
        "说明": desc,
    }


def _detect_cost_consulting_type(query: str) -> str | None:
    """从查询中检测造价咨询的具体服务类型。"""
    # 按关键词长度从长到短匹配
    type_patterns: list[tuple[str, str]] = [
        ("施工阶段全过程造价控制", "施工阶段全过程造价控制"),
        ("全过程造价控制", "施工阶段全过程造价控制"),
        ("编制工程量清单", "编制工程量清单"),
        ("工程量清单编制", "编制工程量清单"),
        ("编制标底.*清单", "编制标底(含清单)"),
        ("编制标底", "编制标底(含清单)"),
        ("编制施工图预算", "编制施工图预算"),
        ("施工图预算编制", "编制施工图预算"),
        ("编制竣工结算", "编制竣工结算"),
        ("竣工结算编制", "编制竣工结算"),
        ("编制项目投资估算", "编制项目投资估算"),
        ("投资估算编制", "编制项目投资估算"),
        ("编制设计概算", "编制设计概算"),
        ("设计概算编制", "编制设计概算"),
        ("审核概算", "审核概算"),
        ("概算审核", "审核概算"),
        ("审核预算.*标底", "审核预算、标底"),
        ("审核标底", "审核预算、标底"),
        ("审核预算", "审核预算、标底"),
        ("审核竣工结算", "审核竣工结算"),
        ("竣工结算审核", "审核竣工结算"),
        ("审核.*结算", "审核竣工结算"),
        # 泛化匹配
        ("编制.*清单", "编制工程量清单"),
        ("清单.*编制", "编制工程量清单"),
        ("编制.*预算", "编制施工图预算"),
        ("编制.*结算", "编制竣工结算"),
    ]
    for pattern, svc_type in type_patterns:
        if re.search(pattern, query):
            return svc_type
    return None


def calc_cost_consulting(
    base_amount_wan: float,
    service_type: str,
    total_investment: float | None = None,
    jianan_only: float | None = None,
) -> dict:
    """
    造价咨询费（津价房地[2008]136号）。

    参数：
        base_amount_wan: 工程费用（万元）= 建安+设备
        service_type: 具体服务类型
        total_investment: 工程总投资（万元），仅"审核概算"需要
        jianan_only: 建安工程费（不含设备），仅"编制投资估算/设计概算"需要
    """
    config = _COST_CONSULTING_RATES.get(service_type)
    if config is None:
        raise ValueError(f"未知的造价咨询服务类型：{service_type}")

    rates: list[float] = config["rates"]
    base_type: str = config["base_type"]

    # 确定计费基数
    if base_type == "工程总投资":
        if total_investment is None:
            raise ValueError("审核概算需要工程总投资，但总投资未知。请提供总投资金额。")
        amount = total_investment
    elif base_type == "建安工程费用" and jianan_only is not None:
        amount = jianan_only
    elif base_type == "建安工程费用":
        amount = base_amount_wan  # fallback
    else:
        amount = base_amount_wan  # 工程费用

    # 差额分档累进
    total_fee = 0.0
    prev_limit = 0.0
    steps: list[dict] = []

    for i, limit in enumerate(_COST_CONSULTING_BRACKETS):
        if amount <= prev_limit:
            break
        tier_amount = min(amount, limit) - prev_limit
        if tier_amount <= 0:
            prev_limit = limit
            continue
        rate = rates[i]
        tier_fee = tier_amount * rate / 1000.0  # ‰ → 万元
        total_fee += tier_fee
        steps.append({
            "区间": f"{prev_limit:.0f}~{limit:.0f}" if limit != float("inf") else f">{prev_limit:.0f}",
            "金额(万元)": round(tier_amount, 2),
            "费率(‰)": rate,
            "费用(万元)": round(tier_fee, 4),
        })
        prev_limit = limit

    total_fee = round(total_fee, 4)

    return {
        "费种": f"造价咨询费（{service_type}）",
        "依据": "《天津市建设工程造价咨询服务项目和价格标准》（津价房地[2008]136号）",
        "计算公式": "差额分档累进，基准价可上下浮动 ±20%",
        "参数": {
            "计费基数(万元)": round(amount, 4),
            "基数类型": base_type,
            "服务类型": service_type,
            "浮动幅度": "±20%",
        },
        "计算步骤": steps,
        "结果(万元)": total_fee,
        "说明": (
            f"计费基数 {amount:.0f} 万元（{base_type}），{service_type}，\n"
            f"差额分档累进计算，造价咨询费基准价为 **{total_fee:.2f} 万元**"
            f"（可在 ±20% 幅度内协商浮动）"
        ),
    }


# ============================================================
# 工具函数
# ============================================================

def _cumulative_tiered(
    amount: float,
    rates: list[tuple[float, float]],
) -> tuple[float, list[dict]]:
    """
    差额分档累进计算（建设管理费模式）。

    返回 (总费用, [各档明细...])
    """
    total = 0.0
    prev_limit = 0.0
    steps: list[dict] = []

    for limit, rate in rates:
        if amount <= prev_limit:
            break
        tier_amount = min(amount, limit) - prev_limit
        if tier_amount <= 0:
            prev_limit = limit
            continue
        tier_fee = tier_amount * rate / 100.0
        total += tier_fee
        steps.append({
            "区间": f"{prev_limit:.0f}~{limit:.0f}" if limit != float("inf") else f">{prev_limit:.0f}",
            "金额(万元)": round(tier_amount, 2),
            "费率(%)": rate,
            "费用(万元)": round(tier_fee, 4),
        })
        prev_limit = limit

    return round(total, 4), steps


def _bracket_fixed(
    amount: float,
    rates: list[tuple[float, float]],
) -> tuple[float, dict | None]:
    """
    分档定额计算（交易服务费模式）— 金额落在哪个档就收固定费用。

    返回 (费用, 匹配档位信息)
    """
    for limit, fee in rates:
        if amount <= limit:
            return fee, {"中标额≤(万元)": limit, "收费标准(元)": fee}
    return 0, None


def _linear_interpolate(
    amount: float,
    table: list[tuple[float, float]],
) -> float:
    """
    线性内插 — 用于监理费基价表。
    计费额在两个档位之间时，按比例计算收费基价。
    超出最大档位时按 1.039% 收费率计算。
    """
    if amount <= table[0][0]:
        # 低于最低档，按最低档比例折算
        return round(amount * table[0][1] / table[0][0], 4)

    for i in range(len(table) - 1):
        x1, y1 = table[i]
        x2, y2 = table[i + 1]
        if x1 <= amount <= x2:
            # 线性内插: y = y1 + (y2-y1)*(x-x1)/(x2-x1)
            result = y1 + (y2 - y1) * (amount - x1) / (x2 - x1)
            return round(result, 4)

    # 超出最大档位 → 按 1.039% 收费率
    return round(amount * JIANLI_LARGE_RATE / 100.0, 4)


# ============================================================
# 查询解析
# ============================================================

def _extract_discount_coefficient(query: str) -> float:
    """
    从查询中提取打折系数。默认 1.0（不打折）。

    支持表述：
    - "打八折" / "打8折" → 0.8
    - "打六五折" / "65折" → 0.65
    - "折扣0.8" / "打折系数0.85" → 0.8 / 0.85
    - "下浮20%" → 0.8
    - "上浮10%" → 1.1
    """
    # 中文数字 → 阿拉伯数字映射（仅用于折扣表达，不修改原始 query）
    _CN_NUM = {'一': '1', '二': '2', '三': '3', '四': '4', '五': '5',
               '六': '6', '七': '7', '八': '8', '九': '9', '零': '0',
               '两': '2'}
    query_norm = query
    for cn, digit in _CN_NUM.items():
        query_norm = query_norm.replace(cn, digit)

    # "打X折" / "打 X 折"
    m = re.search(r'打\s*(\d+\.?\d*)\s*折', query_norm)
    if m:
        val = float(m.group(1))
        if val >= 10:
            return val / 100.0   # "打65折" → 0.65
        return val / 10.0        # "打8折" → 0.8

    # "X折"（独立出现，非"打折"的一部分，非"折扣"）
    m = re.search(r'(?<!打)(?<!\d)\s*(\d+\.?\d*)\s*折(?!扣)', query_norm)
    if m:
        val = float(m.group(1))
        if val >= 10:
            return val / 100.0   # "65折" → 0.65
        return val / 10.0 if val > 1 else val  # "8折" → 0.8

    # "折扣X" / "打折系数X" / "打折系数: X"
    m = re.search(r'(?:折扣|打折系数)\s*[:：]?\s*(\d+\.?\d*)', query)
    if m:
        val = float(m.group(1))
        return val if val <= 1 else val / 100.0

    # "下浮X%" → 1 - X/100
    m = re.search(r'下浮\s*(\d+\.?\d*)\s*%', query)
    if m:
        return round(1.0 - float(m.group(1)) / 100.0, 4)

    # "上浮X%" → 1 + X/100
    m = re.search(r'上浮\s*(\d+\.?\d*)\s*%', query)
    if m:
        return round(1.0 + float(m.group(1)) / 100.0, 4)

    return 1.0


def _extract_amount(query: str) -> float | None:
    """
    从查询文本提取金额（统一转为"万元"）。
    支持：
    - "8000万" → 8000
    - "1.2亿" → 12000
    - "5000000元" → 500
    - "1000万元" → 1000
    - 纯数字（默认视为万元）
    """
    # 亿
    m = re.search(r'(\d+\.?\d*)\s*亿', query)
    if m:
        return float(m.group(1)) * 10000

    # 万元 — 取所有匹配中最大的（当"建安费131万+设备费160万"时取160而非131）
    m_wan = re.findall(r'(\d+\.?\d*)\s*万', query)
    if m_wan:
        return max(float(x) for x in m_wan)

    # …元以上 / …元
    m = re.search(r'(\d+\.?\d*)\s*元', query)
    if m:
        return float(m.group(1)) / 10000

    # 纯数字（默认万元）— 找最大的数字
    nums = re.findall(r'(\d+\.?\d*)', query)
    if nums:
        # 优先取看起来像工程金额的数字（≥10 的）
        candidates = [float(n) for n in nums if float(n) >= 10]
        if candidates:
            return max(candidates)  # 取最大的，通常是总金额

    return None


# 费种检测模式（提取为模块级，供 _detect_fee_type 和 _detect_all_fee_types 共用）
_FEE_PATTERNS: list[tuple[str, str]] = [
    ("建设管理费", r"建设管理费|建设单位管理费|代建管理费|项目建设管理费|财建.*504"),
    ("招标代理费", r"招标代理|招标.*收费|计价格.*1980"),
    ("交易服务费", r"交易服务费|工程建设交易|津发改.*979"),
    ("监理费", r"监理.*(?:费|收费|服务费)|施工监理.*收费|发改价格.*670"),
    ("工程设计费", r"工程(?:勘察)?设计费?[^用]|基本设计收费|勘察设计收费|设计收费基价|设计费.*(?:计费|计算|收费|多少|怎么|如何|专业|调整|系数|表|区分|分类|有哪些|是什么|怎么区分|复杂|Ⅰ|Ⅱ|Ⅲ|I级|II级|III级|\d)|设计费\s*[？?]|设计费\s*$|计价格.*10号"),
    ("施工图审查费", r"施工图(?:设计文件)?审查|图审[费费]|津价管.*46"),
    ("勘察费", r"勘察费|工程勘察(?!设计)|岩土.*勘察.*费|水文地质.*勘察.*费|勘察.*(?:多少|计算|怎么|如何|收费|取费|标准|定额)"),
    ("可行性研究费", r"可行性研究|可研|项目建议书|前期工作咨询|计价格.*1283"),
    ("水土保持费", r"水土保持.*(?:费|方案编制|监测|验收|咨询)|保监.*22"),
    ("环境影响咨询费", r"环境影响(?:咨询|评价).*[费费]|环评[费费]|计价格.*125"),
    ("劳动安全卫生评审费", r"劳动安全卫生评审|安全卫生评审费|安全评审费|劳安评审"),
    ("场地准备费及临时设施费", r"场地准备费|临时设施费|场地准备及临时设施|场地.*准备.*费"),
    ("工程保险费", r"工程保险[费费]|工程一切险|工程险|工程.*保险"),
    ("预备费", r"预备费|基本预备费|工程预备费|预备.*费率"),
    ("造价咨询费", r"造价咨询|造价.*(?:编制|审核|清单|标底|预算|结算|概算|全过程.*控制)|工程量清单.*[费费]|标底.*编制.*[费费]|津价房地.*136"),
]


def _detect_fee_type(query: str) -> str | None:
    """检测查询涉及哪种二类费（返回第一个匹配的费种）。"""
    for fee_type, pattern in _FEE_PATTERNS:
        if re.search(pattern, query):
            return fee_type
    return None


def _detect_all_fee_types(query: str) -> list[str]:
    """检测查询中涉及的所有二类费（支持"监理费和设计费分别为多少"这类多费种提问）。"""
    return [fee_type for fee_type, pattern in _FEE_PATTERNS if re.search(pattern, query)]


# ============================================================
# 多费种迭代计算模式
# ============================================================

_MODES: list[tuple[str, str]] = [
    ("cascade", r"全部费用|所有费用|各项费用|费用汇总|联算|一并计算|算.*二类费|二类费.*算|计算.*二类费"),
    ("iteration", r"迭代.*(?:计算|总投资|总概算)|反复.*计算|循环.*收敛|总投资.*收敛|工程总概算.*计算"),
    ("comparison", r"方案对比|方案比选|敏感性分析|多方案|比选"),
]


def _detect_multi_fee_mode(query: str) -> str | None:
    """检测查询是否为多费种模式（在单费种检测之前调用）。"""
    for mode, pattern in _MODES:
        if re.search(pattern, query):
            return mode
    return None


def _extract_numeric_value(result: dict) -> float:
    """从 calc_* 结果字典中提取可求和的数值（统一转为万元）。"""
    val = result.get("结果(万元)")
    if val is not None and isinstance(val, (int, float)):
        return float(val)
    mid = result.get("结果中值(万元)")
    if mid is not None:
        return float(mid)
    yuan = result.get("结果(元)")
    if yuan is not None:
        return round(float(yuan) / 10000.0, 4)
    return 0.0


# ============================================================
# 各费种计算函数
# ============================================================

def calc_jianshe_guanli(amount_wan: float) -> dict:
    """建设管理费（财建[2016]504号）"""
    total, steps = _cumulative_tiered(amount_wan, JIANSHE_GUANLI_RATES)
    return {
        "费种": "建设管理费（建设单位管理费）",
        "依据": "《基本建设项目建设成本管理规定》（财建[2016]504号）",
        "计算公式": "差额分档累进",
        "参数": {"工程总概算(万元)": amount_wan},
        "计算步骤": steps,
        "结果(万元)": total,
        "说明": (
            f"工程总概算 {amount_wan:.0f} 万元，"
            f"项目建设管理费总额控制数为 **{total:.2f} 万元**"
        ),
    }


def calc_zhaobiao_daili(amount_wan: float, service_type: str = "工程招标") -> dict:
    """招标代理服务费（计价格[2002]1980号）"""
    type_idx = {"货物招标": 1, "服务招标": 2, "工程招标": 3}
    idx = type_idx.get(service_type, 3)

    # 构建该类型的费率表
    type_rates: list[tuple[float, float]] = [
        (limit, rates[idx - 1]) for limit, *rates in
        [(r[0], r[1], r[2], r[3]) for r in ZHAOBIAO_DAILI_RATES]
    ]

    total, steps = _cumulative_tiered(amount_wan, type_rates)
    return {
        "费种": f"招标代理服务费（{service_type}）",
        "依据": "《招标代理业务收费管理暂行办法》（计价格[2002]1980号）",
        "计算公式": "差额定率累进，上下浮动不超过 20%",
        "参数": {"中标金额(万元)": amount_wan, "招标类型": service_type},
        "计算步骤": steps,
        "结果(万元)": total,
        "说明": (
            f"{service_type} 中标金额 {amount_wan:.0f} 万元，"
            f"招标代理服务费为 **{total:.2f} 万元**（可上下浮动 20%）"
        ),
    }


# 招标代理服务费 — 子类型与计费基数映射
_ZHAOBIAO_SUB_TYPES: list[dict] = [
    {"key": "货物招标", "base_label": "设备费", "base_source": "shebei"},
    {"key": "工程招标", "base_label": "建安费", "base_source": "jianan"},
    {"key": "服务招标（勘察）", "base_label": "勘察费", "base_source": "kancha"},
    {"key": "服务招标（设计）", "base_label": "设计费", "base_source": "sheji"},
    {"key": "服务招标（监理）", "base_label": "监理费", "base_source": "jianli"},
]


def calc_zhaobiao_daili_all(
    jianan: float,
    shebei: float = 0.0,
    project_type: str = "建筑",
    query: str = "",
    dependent_configs: dict | None = None,
) -> dict:
    """招标代理服务费 — 全部 5 类自动计算汇总。

    计费基数规则：
    - 货物招标 → 设备费
    - 工程招标 → 建安费
    - 服务招标（勘察）→ 勘察费
    - 服务招标（设计）→ 设计费
    - 服务招标（监理）→ 监理费

    dependent_configs: 用户配置的依赖费种参数。不为 None 时使用用户参数，
    否则使用默认值（所有系数=1.0，勘察费取区间中值）。
    """
    amount_wan = jianan + shebei

    # 1. 先计算依赖费种（监理费、设计费、勘察费）
    if dependent_configs is not None:
        # ── 使用用户配置的参数 ──
        jl_cfg = dependent_configs.get("监理费", {})
        sj_cfg = dependent_configs.get("工程设计费", {})
        kc_cfg = dependent_configs.get("勘察费", {})

        # 监理费 — 传 jianan+shebei 分开以触发 40% 规则
        jl_prof = jl_cfg.get("professional_coef", 1.0)
        jl_comp = jl_cfg.get("complexity_coef", 1.0)
        jl_elev = jl_cfg.get("elevation_coef", 1.0)
        if jianan > 0 or shebei > 0:
            jianli_result = calc_jianli(
                jianan=jianan, shebei=shebei,
                professional_coef=jl_prof, complexity_coef=jl_comp,
                elevation_coef=jl_elev,
            )
        else:
            jianli_result = calc_jianli(
                amount_wan=amount_wan,
                professional_coef=jl_prof, complexity_coef=jl_comp,
                elevation_coef=jl_elev,
            )
        jianli_fee = jianli_result["结果(万元)"]

        # 设计费
        sj_prof = sj_cfg.get("professional_coef", 1.0)
        sj_comp = sj_cfg.get("complexity_coef", 1.0)
        sj_addi = sj_cfg.get("additional_coef", 1.0)
        sj_addi_list = [sj_addi] if abs(sj_addi - 1.0) > 0.005 else None
        sheji_result = calc_sheji(amount_wan, sj_prof, sj_comp, additional_coefs=sj_addi_list)
        sheji_fee = sheji_result["结果(万元)"]

        # 勘察费
        kc_rate = kc_cfg.get("rate")
        kc_ptype = kc_cfg.get("project_type", project_type)
        if kc_rate is not None:
            kancha_fee = round((jianan + shebei) * kc_rate / 100.0, 4)
            kancha_result = {"结果中值(万元)": kancha_fee, "结果(万元)": kancha_fee}
        else:
            kancha_result = calc_kancha_rough(jianan, shebei, kc_ptype)
            kancha_fee = kancha_result["结果中值(万元)"]
            if kancha_fee is None:
                kancha_fee = kancha_result.get("结果(万元)", 0) or 0
    else:
        # ── 默认参数（原有行为，所有系数=1.0）──
        # 修复：传 jianan+shebei 分开以触发 40% 规则，而非合并 amount_wan
        if jianan > 0 or shebei > 0:
            jianli_result = calc_jianli(jianan=jianan, shebei=shebei)
        else:
            jianli_result = calc_jianli(amount_wan=amount_wan)
        jianli_fee = jianli_result["结果(万元)"]

        sheji_result = calc_sheji(amount_wan)
        sheji_fee = sheji_result["结果(万元)"]

        kancha_result = calc_kancha_rough(jianan, shebei, project_type)
        kancha_fee = kancha_result["结果中值(万元)"]
        if kancha_fee is None:
            kancha_fee = kancha_result.get("结果(万元)", 0) or 0

    # 2. 各子类型基数
    bases = {
        "货物招标": shebei,
        "工程招标": jianan,
        "服务招标（勘察）": kancha_fee,
        "服务招标（设计）": sheji_fee,
        "服务招标（监理）": jianli_fee,
    }

    # 3. 逐项计算
    details: list[dict] = []
    total = 0.0
    for sub in _ZHAOBIAO_SUB_TYPES:
        key = sub["key"]
        base = bases[key]
        base_label = sub["base_label"]
        if base <= 0:
            details.append({
                "类型": key,
                "基数(万元)": 0,
                "基数来源": base_label,
                "费用(万元)": 0,
                "说明": f"{base_label}为 0，无法计算",
            })
            continue

        # 映射到计价格[2002]1980号的类型索引
        if key == "货物招标":
            svc = "货物招标"
        elif key == "工程招标":
            svc = "工程招标"
        else:
            svc = "服务招标"

        single = calc_zhaobiao_daili(base, svc)
        fee = single["结果(万元)"]
        total += fee
        details.append({
            "类型": key,
            "基数(万元)": round(base, 4),
            "基数来源": base_label,
            "费用(万元)": fee,
            "计算步骤": single["计算步骤"],
        })

    total = round(total, 4)

    # 构建说明
    desc_parts = []
    for d in details:
        desc_parts.append(
            f"- **{d['类型']}**：基数 {d['基数来源']} {d['基数(万元)']:.2f} 万 → {d['费用(万元)']:.2f} 万元"
        )
    desc = "### 费用明细\n\n" + "\n".join(desc_parts)
    desc += f"\n\n### 💰 合计：**{total} 万元**"

    return {
        "明细": details,
        "合计(万元)": total,
        "依赖费种": {
            "监理费(万元)": round(jianli_fee, 4),
            "设计费(万元)": round(sheji_fee, 4),
            "勘察费(万元)": round(kancha_fee, 4),
        },
        "费种": "招标代理服务费",
        "依据": "《招标代理业务收费管理暂行办法》（计价格[2002]1980号）",
        "说明": desc,
    }


def calc_jiaoyi_fuwu(
    jianan: float | None = None,
    shebei: float | None = None,
    jianli_fee: float | None = None,
    sheji_fee: float | None = None,
    amount_wan: float | None = None,
) -> dict:
    """
    工程建设交易服务费（津发改价管[2017]979号）

    分 4 类服务分别按中标额分档定额，每类 8 档，总计 = 四类之和：
    - 施工：基数 = 建安工程费
    - 设备：基数 = 设备购置费
    - 监理：基数 = 监理费（程序计算）
    - 设计：基数 = 设计费（程序计算）

    若仅提供单一金额 amount_wan，回退到简单按中标额查档模式。
    招标方承担 60%，中标方承担 40%。
    """
    categories: list[dict] = []
    total_fee = 0.0

    label_amt_pairs: list[tuple[str, float | None]] = [
        ("施工", jianan),
        ("设备", shebei),
        ("监理", jianli_fee),
        ("设计", sheji_fee),
    ]

    for label, amt in label_amt_pairs:
        if amt is not None and amt > 0:
            fee, bracket = _bracket_fixed(amt, JIAOYI_FUWU_RATES)
            categories.append({
                "类别": label,
                "基数(万元)": round(amt, 4),
                "费用(元)": fee,
                "档位": f"中标额≤{bracket['中标额≤(万元)']:.0f}万" if bracket else "",
            })
            total_fee += fee

    if not categories and amount_wan is not None:
        # 回退：单金额模式（用户只给了中标额）
        fee, bracket = _bracket_fixed(amount_wan, JIAOYI_FUWU_RATES)
        categories.append({
            "类别": "交易服务费（按中标额）",
            "基数(万元)": amount_wan,
            "费用(元)": fee,
            "档位": f"中标额≤{bracket['中标额≤(万元)']:.0f}万" if bracket else "",
        })
        total_fee = fee

    total_fee = round(total_fee, 2)
    zhaobiao_share = round(total_fee * 0.6, 2)
    zhongbiao_share = round(total_fee * 0.4, 2)

    if len(categories) > 1:
        desc_parts = [f"{c['类别']} {c['费用(元)']:.0f} 元" for c in categories]
        desc = " + ".join(desc_parts) + f" = 合计 **{total_fee:.0f} 元**"
    elif categories:
        desc = f"中标额 {categories[0]['基数(万元)']:.0f} 万元，交易服务费最高 **{total_fee:.0f} 元**"
    else:
        desc = "未能确定任何基数，无法计算"

    return {
        "费种": "工程建设交易服务费",
        "依据": "《市发展改革委关于规范工程建设交易服务收费标准有关问题的通知》（津发改价管[2017]979号）",
        "计算公式": "分 4 类服务（施工/设备/监理/设计）分别按中标额分档定额，合计后招标方 60%、中标方 40%",
        "分项明细": categories,
        "结果(元)": total_fee,
        "分摊": f"招标方 60%: {zhaobiao_share:.0f} 元，中标方 40%: {zhongbiao_share:.0f} 元",
        "说明": desc,
    }


def _calc_jifei_from_components(jianan: float, shebei: float) -> tuple[float, dict]:
    """
    发改价格[2007]670 号 1.0.8 条：设备+联合试运转费占比 > 40% 时的计费额调整。

    参数：
        jianan: 建筑安装工程费（万元）
        shebei:  设备购置费 + 联合试运转费（万元）

    返回：(调整后计费额, 调整说明)
    """
    total = jianan + shebei
    shebei_ratio = shebei / total if total > 0 else 0

    if shebei_ratio <= 0.4:
        # 不触发调整
        return total, {
            "触发调整": False,
            "说明": (
                f"设备+联合试运转费占比 {shebei_ratio:.1%} ≤ 40%，不触发调整。"
                f"计费额 = 建筑安装工程费 + 设备+联合试运转费 = {total:.0f} 万元"
            ),
        }

    # 触发 40% 调整
    adjusted_before_floor = jianan + shebei * 0.4

    # 保底下限：建安费相同 + 设备占比恰好 40% 的假想项目的计费额
    # B' / (A + B') = 0.4  →  B' = (2/3)A  →  计费额 = A + B' = (5/3)A
    floor = jianan * 5 / 3
    floor_triggered = False

    if adjusted_before_floor < floor:
        final_adjusted = floor
        floor_triggered = True
        floor_note = (
            f"打折后计费额为 **{adjusted_before_floor:.0f} 万元**，"
            f"低于保底值 **{floor:.0f} 万元**"
            f"（等建安费且设备占比恰好 40% 的假想项目计费额），"
            f"取保底值，最终计费额 **{floor:.0f} 万元**"
        )
    else:
        final_adjusted = adjusted_before_floor
        floor_note = ""

    return round(final_adjusted, 4), {
        "触发调整": True,
        "设备占比": f"{shebei_ratio:.1%}",
        "原计费额(万元)": total,
        "打折后计费额(万元)": adjusted_before_floor,
        "最终计费额(万元)": final_adjusted,
        "保底触发": floor_triggered,
        "说明": (
            f"设备+联合试运转费占比 {shebei_ratio:.1%} > 40%，触发 670 号文 1.0.8 条调整：\n"
            f"- 建筑安装工程费 {jianan:.0f} 万元，设备+联合试运转费 {shebei:.0f} 万元\n"
            f"- 原计费额 **{total:.0f} 万元**\n"
            f"- 设备费按 40% 打折后计费额 **{adjusted_before_floor:.0f} 万元**"
            + (f"\n- {floor_note}" if floor_note else "")
        ),
    }


def calc_jianli(
    amount_wan: float | None = None,
    professional_coef: float = 1.0,
    complexity_coef: float = 1.0,
    elevation_coef: float = 1.0,
    jianan: float | None = None,
    shebei: float | None = None,
) -> dict:
    """
    施工监理服务费（发改价格[2007]670号）

    两种调用方式：
    1. 直接给计费额：calc_jianli(amount_wan=8000)
    2. 给分项金额（触发 40% 规则）：calc_jianli(jianan=6000, shebei=10000)

    公式：施工监理服务收费基准价 = 收费基价 × 专业调整系数
          × 复杂程度调整系数 × 高程调整系数
    """
    adjustment_info = None

    # 如果给了分项金额，先算计费额（含 40% 规则）
    if jianan is not None and shebei is not None:
        amount_wan, adjustment_info = _calc_jifei_from_components(jianan, shebei)
    elif amount_wan is None:
        raise ValueError("请提供计费额 amount_wan，或分项金额 jianan + shebei")

    base_price = _linear_interpolate(amount_wan, JIANLI_BASE_RATES)
    benchmark = round(base_price * professional_coef * complexity_coef * elevation_coef, 4)

    has_coef = any(c != 1.0 for c in [professional_coef, complexity_coef, elevation_coef])

    params: dict = {
        "计费额(万元)": amount_wan,
        "收费基价(万元)": base_price,
        "专业调整系数": professional_coef,
        "复杂程度系数": complexity_coef,
        "高程调整系数": elevation_coef,
    }
    if adjustment_info and adjustment_info.get("触发调整"):
        params["计费额调整"] = adjustment_info["说明"]

    desc = f"计费额 {amount_wan:.0f} 万元，收费基价 {base_price:.2f} 万元"
    if has_coef:
        coef_parts = []
        if professional_coef != 1.0:
            coef_parts.append(f"专业调整系数 {professional_coef}（{_describe_jianli_professional_coef(professional_coef)}）")
        if complexity_coef != 1.0:
            coef_parts.append(f"复杂程度系数 {complexity_coef}（{_describe_complexity_coef(complexity_coef)}）")
        if elevation_coef != 1.0:
            coef_parts.append(f"高程调整系数 {elevation_coef}（{_describe_elevation_coef(elevation_coef)}）")
        coef_desc = "，".join(coef_parts)
        desc += f"，{coef_desc}，调整后基准价 **{benchmark:.2f} 万元**"
    else:
        desc += f"，基准价 **{benchmark:.2f} 万元**（专业系数1.0、复杂系数1.0、高程系数1.0，均为默认值）"
    desc += "（可上下浮动 20%）"

    return {
        "费种": "施工监理服务费",
        "依据": "《建设工程监理与相关服务收费管理规定》（发改价格[2007]670号）",
        "计算公式": (
            "施工监理服务收费基准价 = 收费基价 × 专业调整系数 "
            "× 工程复杂程度调整系数 × 高程调整系数\n"
            "实际收费 = 基准价 × (1 ± 浮动幅度)，浮动 ≤ 20%\n"
            "（1.0.8条：设备+联合试运转占比>40%时打折计入，且不低于等建安费40%设备占比假想项目）"
        ),
        "参数": params,
        "计费额调整": adjustment_info,
        "结果(万元)": benchmark,
        "说明": desc,
    }


def calc_sheji(
    amount_wan: float,
    professional_coef: float = 1.0,
    complexity_coef: float = 1.0,
    additional_coefs: list[float] | None = None,
    zongti_sheji: bool = False,
    zhuti_xietiao: bool = False,
    shigongtu_yusuan: bool = False,
    jungongtu: bool = False,
    qita_sheji_fee: float = 0.0,
) -> dict:
    """
    工程设计费（计价格[2002]10 号）。

    公式（1.0.3）：
        工程设计收费 = 基准价 × (1 ± 浮动幅度)
        基准价 = 基本设计收费 + 其他设计收费
        基本设计收费 = 收费基价 × 专业系数 × 复杂系数 × 附加系数

    附加项（计入其他设计收费）：
        - 总体设计费（1.0.13）：基本设计收费的 5%
        - 主体设计协调费（1.0.14）：基本设计收费的 5%
        - 施工图预算编制费（1.0.16）：基本设计收费的 10%
        - 竣工图编制费（1.0.16）：基本设计收费的 8%
        - 其他设计收费（1.0.6）：直接指定金额

    计费额定义（1.0.8）：建筑安装工程费 + 设备与工器具购置费 + 联合试运转费
    """
    base_price = _linear_interpolate(amount_wan, SHEJI_BASE_RATES)
    if amount_wan > SHEJI_BASE_RATES[-1][0]:
        base_price = round(amount_wan * SHEJI_LARGE_RATE / 100.0, 4)

    # 附加调整系数合并（1.0.9.3：多个系数不能连乘）
    if additional_coefs and len(additional_coefs) > 1:
        additional_coef = sum(additional_coefs) - len(additional_coefs) + 1
    elif additional_coefs:
        additional_coef = additional_coefs[0]
    else:
        additional_coef = 1.0

    # 基本设计收费
    basic_design = round(base_price * professional_coef * complexity_coef * additional_coef, 4)

    # 其他设计收费
    other_items: list[tuple[str, float]] = []
    if zongti_sheji:
        fee = round(basic_design * 0.05, 4)
        other_items.append(("总体设计费（5%）", fee))
    if zhuti_xietiao:
        fee = round(basic_design * 0.05, 4)
        other_items.append(("主体设计协调费（5%）", fee))
    if shigongtu_yusuan:
        fee = round(basic_design * 0.10, 4)
        other_items.append(("施工图预算编制费（10%）", fee))
    if jungongtu:
        fee = round(basic_design * 0.08, 4)
        other_items.append(("竣工图编制费（8%）", fee))
    if qita_sheji_fee > 0:
        other_items.append(("其他设计收费", qita_sheji_fee))

    other_total = round(sum(f for _, f in other_items), 4)
    benchmark = round(basic_design + other_total, 4)  # 基准价

    params: dict = {
        "计费额(万元)": amount_wan,
        "收费基价(万元)": base_price,
        "专业调整系数": professional_coef,
        "复杂程度系数": complexity_coef,
        "附加调整系数": round(additional_coef, 2),
    }
    if additional_coefs and len(additional_coefs) > 1:
        params["附加系数明细"] = f"{' + '.join(str(c) for c in additional_coefs)} → 合并后 {additional_coef:.2f}"

    desc = f"计费额 {amount_wan:.0f} 万元，收费基价 {base_price:.2f} 万元"
    coef_parts = [
        f"专业系数 {professional_coef}",
        f"复杂系数 {complexity_coef}",
        f"附加系数 {additional_coef}",
    ]
    desc += f"，{'×'.join(coef_parts)}，基本设计收费 **{basic_design:.2f} 万元**"

    if other_items:
        desc += "\n其他设计收费："
        desc += " + ".join(f"{label} {fee:.2f} 万" for label, fee in other_items)
        desc += f" = **{other_total:.2f} 万元**"
        desc += f"\n工程设计收费基准价 = {basic_design:.2f} + {other_total:.2f} = **{benchmark:.2f} 万元**"

    return {
        "费种": "工程设计费",
        "依据": "《工程勘察设计收费管理规定》（计价格[2002]10号）",
        "计算公式": (
            "工程设计收费基准价 = 基本设计收费 + 其他设计收费\n"
            "基本设计收费 = 收费基价 × 专业系数 × 复杂系数 × 附加系数（多个附加系数合并 = 相加 − 个数 + 1）\n"
            "计费额（1.0.8）= 建筑安装工程费 + 设备与工器具购置费 + 联合试运转费"
        ),
        "参数": params,
        "基本设计收费(万元)": basic_design,
        "其他设计收费明细": [{"项目": label, "费用(万元)": fee} for label, fee in other_items] if other_items else [],
        "结果(万元)": benchmark,
        "说明": desc,
    }


def _keyan_interpolate(amount_yi: float, brackets: list[tuple]) -> float:
    """线性内插计算可行性研究费分档基准价。"""
    for inv_lo, inv_hi, fee_lo, fee_hi in brackets:
        if amount_yi <= inv_hi or inv_hi == float("inf"):
            if inv_hi == float("inf") or amount_yi >= inv_hi:
                return fee_hi
            if amount_yi <= inv_lo:
                return fee_lo
            # 线性内插
            ratio = (amount_yi - inv_lo) / (inv_hi - inv_lo)
            return round(fee_lo + ratio * (fee_hi - fee_lo), 4)
    return 0.0


def _detect_keyan_industry(query: str) -> tuple[str, float]:
    """从查询中检测行业并返回可行性研究费行业调整系数。"""
    # 按关键词长度从长到短匹配，避免"建筑"误匹配"建筑材料"
    sorted_industries = sorted(KEYAN_INDUSTRY_COEFS.keys(), key=len, reverse=True)
    for ind in sorted_industries:
        if re.search(ind, query):
            return ind, KEYAN_INDUSTRY_COEFS[ind]
    return "未指定（默认1.0）", 1.0


def calc_keyan(
    amount_yi: float,
    service_type: str = "编制可研报告",
    industry_coef: float | None = None,
    industry_name: str = "",
    complexity_coef: float = 1.0,
) -> dict:
    """
    可行性研究费（计价格[1999]1283号）。

    amount_yi: 估算投资额（亿元）
    service_type: 编制项目建议书 / 编制可研报告 / 评估项目建议书 / 评估可研报告
    industry_coef: 行业调整系数（0.7~1.3），None 则默认 1.0
    industry_name: 行业名称（用于展示）
    complexity_coef: 工程复杂程度调整系数（0.8~1.2，默认 1.0）
    """
    brackets = KEYAN_BRACKETS.get(service_type)
    if brackets is None:
        brackets = KEYAN_BRACKETS["编制可研报告"]

    # 线性内插计算基准价
    base_fee = _keyan_interpolate(amount_yi, brackets)

    # 行业调整系数
    if industry_coef is None:
        industry_coef = 1.0
        industry_name = industry_name or "未指定（默认1.0）"

    # 总调整系数 = 行业 × 复杂程度
    total_coef = round(industry_coef * complexity_coef, 4)

    # 最终费用（基准 × 总调整系数）
    final_fee_mid = round(base_fee * total_coef, 4)

    steps = [
        {"步骤": "估算投资额", "公式": f"{amount_yi:.4f} 亿元", "结果": f"{amount_yi * 10000:.0f} 万元"},
        {"步骤": "确定服务类型", "公式": "", "结果": service_type},
    ]

    # 显示匹配到的分档区间
    for inv_lo, inv_hi, fee_lo, fee_hi in brackets:
        if amount_yi <= inv_hi or inv_hi == float("inf"):
            if amount_yi <= inv_lo and inv_hi < float("inf"):
                steps.append({"步骤": "所在分档区间", "公式": f"<{inv_hi}亿元", "结果": f"基准价固定 {fee_lo} 万元"})
            elif inv_hi == float("inf") or amount_yi >= inv_hi:
                steps.append({"步骤": "所在分档区间", "公式": f"≥{inv_lo}亿元", "结果": f"基准价 {fee_hi} 万元"})
            else:
                steps.append({"步骤": "所在分档区间", "公式": f"{inv_lo}~{inv_hi}亿元", "结果": f"基准价 {fee_lo}~{fee_hi} 万元"})
            break

    steps.extend([
        {"步骤": "线性内插基准价", "公式": f"插值({amount_yi:.4f})", "结果": f"{base_fee:.2f} 万元"},
        {"步骤": "行业调整系数", "公式": f"{industry_name}", "结果": str(industry_coef)},
        {"步骤": "工程复杂程度系数", "公式": f"复杂程度 {complexity_coef}", "result": str(complexity_coef)},
        {"步骤": "总调整系数", "公式": f"{industry_coef} × {complexity_coef}", "结果": str(total_coef)},
        {"步骤": "最终费用", "公式": f"{base_fee:.2f} × {total_coef}", "结果": f"{final_fee_mid:.2f} 万元"},
    ])

    return {
        "费种": f"建设项目前期工作咨询费（{service_type}）",
        "依据": "《建设项目前期工作咨询收费暂行规定》（计价格[1999]1283号）",
        "计算公式": "基准价 = 按估算投资额分档线性内插；最终费用 = 基准价 × 行业调整系数 × 复杂程度系数",
        "参数": {
            "估算投资额(亿元)": amount_yi,
            "服务类型": service_type,
            "行业": industry_name,
            "行业调整系数": industry_coef,
            "复杂程度系数": complexity_coef,
            "总调整系数": total_coef,
        },
        "结果(万元)": final_fee_mid,
        "基准价(万元)": base_fee,
        "说明": (
            f"估算投资额 {amount_yi:.4f} 亿元（{amount_yi * 10000:.0f} 万元），{service_type}\n"
            f"分档线性内插基准价：**{base_fee:.2f} 万元**\n"
            f"行业「{industry_name}」调整系数：**{industry_coef}**\n"
            f"工程复杂程度系数：**{complexity_coef}**\n"
            f"最终费用 = {base_fee:.2f} × {total_coef} = **{final_fee_mid:.2f} 万元**"
        ),
        "计算步骤": steps,
    }


def calc_keyan_multi(
    amount_yi: float,
    selected_services: list[str],
    industry_coef: float = 1.0,
    industry_name: str = "",
    complexity_coef: float = 1.0,
) -> dict:
    """可行性研究费 — 多服务类型选择计算。

    对用户选择的每项服务分别计算，汇总合计。
    """
    details: list[dict] = []
    total = 0.0

    for svc in selected_services:
        r = calc_keyan(
            amount_yi, svc,
            industry_coef=industry_coef,
            industry_name=industry_name,
            complexity_coef=complexity_coef,
        )
        fee = r["结果(万元)"]
        details.append({
            "服务类型": svc,
            "费用(万元)": fee,
            "基准价(万元)": r.get("基准价(万元)", 0),
            "计算步骤": r.get("计算步骤", []),
        })
        total += fee

    total = round(total, 4)

    desc_parts = []
    for d in details:
        desc_parts.append(f"- **{d['服务类型']}**：{d['费用(万元)']} 万元")

    return {
        "费种": "建设项目前期工作咨询费",
        "依据": "《建设项目前期工作咨询收费暂行规定》（计价格[1999]1283号）",
        "明细": details,
        "合计(万元)": total,
        "参数": {
            "估算投资额(亿元)": amount_yi,
            "行业调整系数": industry_coef,
            "行业名称": industry_name or "默认",
            "复杂程度系数": complexity_coef,
            "选中服务": selected_services,
        },
        "说明": "\n".join(desc_parts),
    }


def _is_hebei_project(query: str) -> bool:
    """检测是否为河北省项目。"""
    return bool(re.search(r"河北", query))


def calc_shigong_shencha(
    amount: float,
    project_type: str = "公建",
    size: str = "中型",
    sheji_fee: float | None = None,
    sheji_fee_only: float | None = None,
    kancha_fee_mid: float | None = None,
    kancha_rate_desc: str = "区间中值",
    query: str = "",
) -> dict:
    """
    施工图审查费。

    默认依据：津价管[2011]46号 + 建市[2007]86号
    河北省项目：冀价行费[2018]57号 / 冀建质[2017]1号
      —（勘察费+设计费）× 6.5%

    住宅类：按建筑面积 × 单价（元/m²）
    公建/工业/市政类：按勘察设计费（设计费+勘察费）× 费率(%)
      其中设计费按计价格[2002]10号计算，勘察费按《市政工程设计概算编制办法》百分比法粗略估算
    项目大中小划分：建市[2007]86号《工程设计资质标准》
    """
    size_desc = {"大型": "大型", "中型": "中型", "小型": "小型"}.get(size, size)

    if project_type == "住宅":
        rate_per_m2 = SHIGONG_SHENCHA_ZHUZHAI.get(size, 1.7)
        fee = round(amount * rate_per_m2, 2)
        desc = f"住宅 {size_desc} 项目，建筑面积 {amount:.0f} m² × {rate_per_m2} 元/m²，审查费 **{fee:.2f} 元**"
        return {
            "费种": f"施工图审查费（住宅{size_desc}）",
            "依据": "《市发展改革委关于施工图审查收费标准的通知》（津价管[2011]46号）\n"
                    "项目规模划分依据：《工程设计资质标准》（建市[2007]86号）",
            "计算公式": f"审查费 = 建筑面积 × {rate_per_m2} 元/m²",
            "参数": {"建筑面积(m²)": amount, "项目规模": size_desc, "单价(元/m²)": rate_per_m2},
            "结果(元)": fee,
            "说明": desc,
            "计算步骤": [
                {"步骤": "判定项目类型", "公式": "查询关键词检测", "结果": "住宅类"},
                {"步骤": "判定项目规模", "公式": "建市[2007]86号：住宅按层数（≤12小/12~20中/>20大）", "结果": f"{size_desc}项目"},
                {"步骤": "查找收费标准", "公式": f"津价管[2011]46号 第一条：住宅{size_desc} {rate_per_m2}元/m²", "结果": f"{rate_per_m2} 元/m²"},
                {"步骤": "计算审查费", "公式": f"{amount:.0f} m² × {rate_per_m2} 元/m²", "结果": f"{fee:.2f} 元"},
            ],
        }

    # 河北省项目：除有特殊规定外，（勘察费+设计费）× 6.5%
    if query and _is_hebei_project(query):
        hebei_rate = HEBEI_SHENCHA_RATE
        if sheji_fee is not None:
            fee = round(sheji_fee * hebei_rate / 100.0, 4)
            desc = f"河北省项目，勘察设计费（设计费+勘察费）{sheji_fee:.2f} 万元 × {hebei_rate}%，审查费 **{fee:.2f} 万元**"
            params = {"勘察设计费(万元)": sheji_fee, "= 设计费+勘察费": "", "费率(%)": hebei_rate, "适用地区": "河北省"}
            if sheji_fee_only is not None and kancha_fee_mid is not None:
                steps = [
                    {"步骤": "判定适用地区", "公式": "查询关键词检测", "结果": "河北省"},
                    {"步骤": "计算设计费", "公式": "计价格[2002]10号：收费基价 × 专业系数 × 复杂系数 × 附加系数", "结果": f"{sheji_fee_only:.2f} 万元"},
                    {"步骤": "计算勘察费", "公式": f"《市政工程设计概算编制办法》百分比法（{kancha_rate_desc}）", "结果": f"{kancha_fee_mid:.2f} 万元"},
                    {"步骤": "计算勘察设计费基数", "公式": f"设计费 + 勘察费 = {sheji_fee_only:.2f} + {kancha_fee_mid:.2f}", "结果": f"{sheji_fee:.2f} 万元"},
                    {"步骤": "应用河北省费率", "公式": f"冀价行费[2018]57号：施工图审查费 = (勘察费+设计费) × {hebei_rate}%", "结果": f"{hebei_rate}%"},
                    {"步骤": "计算审查费", "公式": f"{sheji_fee:.2f} 万元 × {hebei_rate}%", "结果": f"{fee:.2f} 万元"},
                ]
            else:
                steps = [
                    {"步骤": "判定适用地区", "公式": "查询关键词检测", "结果": "河北省"},
                    {"步骤": "计算勘察设计费", "公式": "设计费（计价格[2002]10号）+ 勘察费（概算编制办法百分比法）", "结果": f"{sheji_fee:.2f} 万元"},
                    {"步骤": "应用河北省费率", "公式": f"冀价行费[2018]57号：施工图审查费 = (勘察费+设计费) × {hebei_rate}%", "结果": f"{hebei_rate}%"},
                    {"步骤": "计算审查费", "公式": f"{sheji_fee:.2f} 万元 × {hebei_rate}%", "结果": f"{fee:.2f} 万元"},
                ]
        else:
            fee = round(amount * hebei_rate / 100.0, 4)
            desc = f"河北省项目，计费基数 {amount:.0f} 万元 × {hebei_rate}%，审查费 **{fee:.2f} 万元**"
            params = {"计费基数(万元)": amount, "费率(%)": hebei_rate, "适用地区": "河北省"}
            steps = [
                {"步骤": "判定适用地区", "公式": "查询关键词检测", "结果": "河北省"},
                {"步骤": "查找河北省费率", "公式": f"冀价行费[2018]57号：施工图审查费 = (勘察费+设计费) × {hebei_rate}%", "结果": f"{hebei_rate}%"},
                {"步骤": "计算审查费", "公式": f"{amount:.0f} 万元 × {hebei_rate}%", "结果": f"{fee:.2f} 万元"},
            ]
        return {
            "费种": f"施工图审查费（河北省）",
            "依据": "河北省物价局、河北省住房和城乡建设厅《关于规范施工图审查收费有关问题的通知》"
                    "（冀价行费[2018]57号）\n"
                    "河北省住房和城乡建设厅《关于进一步规范施工图审查工作的通知》"
                    "（冀建质[2017]1号）",
            "计算公式": f"审查费 = (勘察费+设计费) × {hebei_rate}%",
            "参数": params,
            "结果(万元)": fee,
            "说明": desc,
            "计算步骤": steps,
        }

    # 默认：津价管[2011]46 号 公建/工业/市政
    rates = SHIGONG_SHENCHA_RATES.get(project_type, SHIGONG_SHENCHA_RATES["公建"])
    rate_pct = rates.get(size, 3.0)

    if sheji_fee is not None:
        fee = round(sheji_fee * rate_pct / 100.0, 4)
        desc = f"{project_type} {size_desc} 项目，勘察设计费（设计费+勘察费）{sheji_fee:.2f} 万元 × {rate_pct}%，审查费 **{fee:.2f} 万元**"
        params = {"勘察设计费(万元)": sheji_fee, "= 设计费+勘察费": "", "项目类型": project_type, "项目规模": size_desc, "费率(%)": rate_pct}
        if sheji_fee_only is not None and kancha_fee_mid is not None:
            steps = [
                {"步骤": "判定项目类型", "公式": "查询关键词检测", "结果": f"{project_type}类"},
                {"步骤": "判定项目规模", "公式": "建市[2007]86号各行业大中小项目划分标准", "结果": f"{size_desc}项目"},
                {"步骤": "计算设计费", "公式": "计价格[2002]10号：收费基价 × 专业系数 × 复杂系数 × 附加系数", "结果": f"{sheji_fee_only:.2f} 万元"},
                {"步骤": "计算勘察费", "公式": f"《市政工程设计概算编制办法》百分比法（{kancha_rate_desc}）", "结果": f"{kancha_fee_mid:.2f} 万元"},
                {"步骤": "计算勘察设计费基数", "公式": f"设计费 + 勘察费 = {sheji_fee_only:.2f} + {kancha_fee_mid:.2f}", "结果": f"{sheji_fee:.2f} 万元"},
                {"步骤": "查找审查费率", "公式": f"津价管[2011]46号：{project_type}类{size_desc} {rate_pct}%", "结果": f"{rate_pct}%"},
                {"步骤": "计算审查费", "公式": f"{sheji_fee:.2f} 万元 × {rate_pct}%", "结果": f"{fee:.2f} 万元"},
            ]
        else:
            steps = [
                {"步骤": "判定项目类型", "公式": "查询关键词检测", "结果": f"{project_type}类"},
                {"步骤": "判定项目规模", "公式": "建市[2007]86号各行业大中小项目划分标准", "结果": f"{size_desc}项目"},
                {"步骤": "计算勘察设计费", "公式": "设计费（计价格[2002]10号）+ 勘察费（概算编制办法百分比法）", "结果": f"{sheji_fee:.2f} 万元"},
                {"步骤": "查找审查费率", "公式": f"津价管[2011]46号：{project_type}类{size_desc} {rate_pct}%", "结果": f"{rate_pct}%"},
                {"步骤": "计算审查费", "公式": f"{sheji_fee:.2f} 万元 × {rate_pct}%", "结果": f"{fee:.2f} 万元"},
            ]
    else:
        fee = round(amount * rate_pct / 100.0, 4)
        desc = f"{project_type} {size_desc} 项目，计费基数 {amount:.0f} 万元 × {rate_pct}%，审查费 **{fee:.2f} 万元**"
        params = {"计费基数(万元)": amount, "项目类型": project_type, "项目规模": size_desc, "费率(%)": rate_pct}
        steps = [
            {"步骤": "判定项目类型", "公式": "查询关键词检测", "结果": f"{project_type}类"},
            {"步骤": "判定项目规模", "公式": "建市[2007]86号各行业大中小项目划分标准", "结果": f"{size_desc}项目"},
            {"步骤": "查找审查费率", "公式": f"津价管[2011]46号：{project_type}类{size_desc} {rate_pct}%", "结果": f"{rate_pct}%"},
            {"步骤": "计算审查费", "公式": f"{amount:.0f} 万元 × {rate_pct}%", "结果": f"{fee:.2f} 万元"},
        ]

    return {
        "费种": f"施工图审查费（{project_type}{size_desc}）",
        "依据": "《市发展改革委关于施工图审查收费标准的通知》（津价管[2011]46号）\n"
                "项目规模划分依据：《工程设计资质标准》（建市[2007]86号）\n"
                "设计费计算依据：《工程勘察设计收费管理规定》（计价格[2002]10号）\n"
                "勘察费估算依据：《市政工程设计概算编制办法》（中国计划出版社）",
        "计算公式": f"审查费 = 勘察设计费（设计费+勘察费）× {rate_pct}%",
        "参数": params,
        "结果(万元)": fee,
        "说明": desc + "\n幕墙/深基坑单项工程按 1.6‰ 计取（最低 1000 元）",
        "计算步骤": steps,
    }


def _linear_interp_table(x: float, xs: list[float], ys: list[float]) -> float:
    """在离散数据点之间线性内插。低于最小点取最小，超过最大点取最大。"""
    if x <= xs[0]:
        return round(ys[0], 4)
    if x >= xs[-1]:
        return round(ys[-1], 4)
    for i in range(len(xs) - 1):
        if xs[i] <= x <= xs[i + 1]:
            y = ys[i] + (ys[i + 1] - ys[i]) * (x - xs[i]) / (xs[i + 1] - xs[i])
            return round(y, 4)
    return round(ys[-1], 4)


def calc_shuibao(amount_yi: float, service_type: str = "方案编制") -> dict:
    """
    水土保持咨询服务费（保监[2005]22号）。

    amount_yi: 主体工程土建投资（亿元）
    service_type: 方案编制 / 施工期监测 / 验收评估 / 技术咨询
    """
    table_map = {
        "方案编制": ("水土保持方案编制费", SHUIBAO_BIANZHI),
        "施工期监测": ("水土保持施工期监测费", SHUIBAO_JIANCE),
        "验收评估": ("水土保持验收评估报告编制费", SHUIBAO_PINGGU),
        "技术咨询": ("水土保持技术咨询费", SHUIBAO_CONSULT),
    }
    name, values = table_map.get(service_type, table_map["方案编制"])
    fee = _linear_interp_table(amount_yi, SHUIBAO_TUDI_TOUZI, values)

    return {
        "费种": name,
        "依据": "《关于开发建设项目水土保持咨询服务费用计列的指导意见》（保监[2005]22号）",
        "计算公式": "按主体工程土建投资内插查表（地貌调整系数：山区 1.2，丘陵及风沙区 1.0，平原区 0.8）",
        "参数": {
            "土建投资(亿元)": amount_yi,
            "服务类型": service_type,
        },
        "结果(万元)": fee,
        "说明": (
            f"土建投资 {amount_yi:.1f} 亿元，{name} **{fee:.2f} 万元**"
            f"（可根据地貌类型乘以调整系数）"
        ),
    }


# ============================================================
# 粗略估算类费种（《市政工程设计概算编制办法》）
# ============================================================

def _detect_project_type(query: str) -> str:
    """检测项目类型（用于勘察费粗略估算的费率选取：建筑 vs 通用）"""
    if re.search(r"建筑|房建|住宅|公建|办公楼|教学楼|医院|商场|酒店|场馆", query):
        return "建筑"
    return "通用"


def _detect_project_size_86(query: str, project_type: str) -> str:
    """
    根据建市[2007]86号《工程设计资质标准》自动判定项目规模（大/中/小）。

    优先级：
    1. 用户显式声明"大型"/"中型"/"小型" → 直接采用
    2. 查询中包含具体技术参数（建筑面积/层数/道路面积等）→ 查表判定
    3. 默认 → 中型
    """
    # 1. 用户显式声明（最高优先级）
    if re.search(r"大[型]|大型", query):
        return "大型"
    if re.search(r"小[型]|小型", query):
        return "小型"

    # 2. 根据项目类型查表
    if project_type == "住宅":
        # 住宅：按层数判定
        m = re.search(r"(\d+)\s*层", query)
        if m:
            ceng = float(m.group(1))
            lo, hi = JIANZHU_ZHUZHAI_CENG
            if ceng > hi:
                return "大型"
            elif ceng > lo:
                return "中型"
            else:
                return "小型"
        # 住宅小区：按总建筑面积判定
        m = re.search(r"(?:总建筑面积|建筑面积|面积)\s*[:：]?\s*(\d+\.?\d*)\s*万?\s*(?:m2|㎡|平米|平方米)?", query)
        if m:
            val = float(m.group(1))
            unit = "万m²" if re.search(r"万\s*(?:m2|㎡|平米|平方米)?", query) else "m²"
            if unit == "m²":
                val = val / 10000  # 转万m²
            # 按小区标准：>30万m² 大型，≤30万m² 中型
            if val > 30:
                return "大型"
            else:
                return "中型"

    elif project_type == "公建" or project_type == "建筑":
        # 公共建筑：按单体建筑面积 或 建筑高度判定
        m_m2 = re.search(r"(?:单体建筑面积|建筑面积|面积)\s*[:：]?\s*(\d+\.?\d*)\s*万?\s*(?:m2|㎡|平米|平方米)?", query)
        m_h = re.search(r"(?:建筑高度|高度|檐高)\s*[:：]?\s*(\d+\.?\d*)\s*m?", query)
        m_ceng = re.search(r"(\d+)\s*层", query)

        size = "中型"  # 默认
        if m_m2:
            val = float(m_m2.group(1))
            if re.search(r"万\s*(?:m2|㎡|平米|平方米)?", query):
                val = val * 10000  # 转m²
            lo, hi = JIANZHU_GONGGONG_M2
            if val >= hi:
                size = "大型"
            elif val > lo:
                size = "中型"
            else:
                size = "小型"
        if m_h:
            val = float(m_h.group(1))
            lo_h, hi_h = JIANZHU_GONGGONG_H
            h_size = "中型"
            if val > hi_h:
                h_size = "大型"
            elif val > lo_h:
                h_size = "中型"
            else:
                h_size = "小型"
            # 取更高级别（面积和高度以高级别为准）
            if h_size == "大型" or size == "大型":
                size = "大型"
            elif h_size == "小型" and size == "小型":
                size = "小型"
        if m_ceng and not m_m2 and not m_h:
            ceng = float(m_ceng.group(1))
            if ceng > 20:
                size = "大型"
            elif ceng > 12:
                size = "中型"
            else:
                size = "小型"
        return size

    elif project_type == "市政":
        # ============================================================
        # 市政行业 — 依据建市[2007]86号 附件3-17 原文
        # ============================================================

        # --- 1. 固定为大型的子项（原文明确"全部为大型项目"）---
        if re.search(r"轨道交通|城市隧道|BRT|快速公交|公交枢纽"
                     r"|电车系统|公共交通专用道"
                     r"|垃圾焚烧|生活垃圾焚烧|危险废弃物", query):
            return "大型"

        # --- 2. 道路工程 — 按道路等级判定（附件3-17 第5项，非面积！）---
        if re.search(r"道路|快速路|主干道|次干路|支路|苜蓿叶|互通.*立交|立交", query):
            if re.search(r"快速路|主干道|苜蓿叶|枢纽.*立交|互通.*立交|全互通", query):
                return "大型"
            elif re.search(r"次干路|简单.*立交", query):
                return "中型"
            elif re.search(r"支路", query):
                return "小型"
            else:
                return "中型"  # 未指定等级，默认中型

        # --- 3. 桥梁工程 — 组合条件判定（附件3-17 第6项，无小型）---
        if re.search(r"桥梁", query):
            dankua = None   # 单孔/单跨跨径
            zongchang = None  # 总长/全长

            # 提取单跨数值
            m_dk = re.search(r"单[孔跨].*?(\d+\.?\d*)", query)
            if not m_dk:
                m_dk = re.search(r"(\d+\.?\d*)\s*(?:m|米).*单[孔跨]", query)
            if m_dk:
                dankua = float(m_dk.group(1))

            # 提取总长/全长数值
            m_zc = re.search(r"(?:总长|全长|总跨|桥长).*?(\d+\.?\d*)", query)
            if not m_zc:
                m_zc = re.search(r"(\d+\.?\d*)\s*(?:m|米).*(?:总长|全长|总跨|桥长)", query)
            if m_zc:
                zongchang = float(m_zc.group(1))

            # 提取通用"米"参数（兜底：单跨≥40 或 总长≥100 → 大型）
            if dankua is None and zongchang is None:
                m_any = re.search(r"(\d+\.?\d*)\s*(?:m|米)(?!\s*(?:万|层|%|％|元))", query)
                if m_any:
                    val = float(m_any.group(1))
                    if val >= 100:
                        return "大型"   # 大概率是总长
                    elif val >= 40:
                        return "大型"   # 大概率是单跨
                    else:
                        return "中型"

            if dankua is not None and dankua >= 40:
                return "大型"
            if zongchang is not None and zongchang >= 100:
                return "大型"
            if dankua is not None or zongchang is not None:
                return "中型"
            return "中型"  # 无技术参数，默认中型

        # --- 4. 环境卫生 — 部分子项固定大型已在上面处理 ---
        # --- 5. SHIZHENG_SCALE 数值查表（给水/排水/燃气/热力/隧道/风景园林等）---
        for broad_type, sub_dict in SHIZHENG_SCALE.items():
            if re.search(broad_type, query):
                # 提取数值（优先匹配带单位的）
                m = re.search(r"(\d+\.?\d*)\s*万?\s*(?:m3|m³|立方米|吨|万吨|mm|毫米|DN\d*|m|米"
                             r"|万m2|万㎡|万平米|万平方米)", query)
                if not m:
                    m = re.search(r"管径\s*[:：]?\s*(\d+\.?\d*)", query)
                if not m:
                    m = re.search(r"DN\s*(\d+\.?\d*)", query)
                if not m:
                    m = re.search(r"(\d+\.?\d*)\s*万", query)
                if m:
                    val = float(m.group(1))
                    # 遍历子类别，匹配最佳阈值
                    for sub_cat, (lo, hi) in sub_dict.items():
                        if sub_cat == broad_type or re.search(sub_cat, query):
                            if val >= hi:
                                return "大型"
                            elif val > lo:
                                return "中型"
                            else:
                                return "小型"
                    # 未匹配子类别，用第一个子类的阈值
                    first_lo, first_hi = list(sub_dict.values())[0]
                    if val >= first_hi:
                        return "大型"
                    elif val > first_lo:
                        return "中型"
                    else:
                        return "小型"
                break
        return "中型"  # 市政默认中型

    elif project_type == "工业":
        # 工业：按投资额判定
        m = re.search(r"(?:投资额|总投资|工程费|计费额)\s*[:：]?\s*(\d+\.?\d*)\s*万?", query)
        if m:
            val = float(m.group(1))
            lo, hi = GONGYE_SCALE
            if val >= hi:
                return "大型"
            elif val > lo:
                return "中型"
            else:
                return "小型"
        return "中型"

    # 3. 默认中型
    return "中型"


def _build_rate_detail(total: float, lo: float, hi: float) -> list[dict]:
    """
    按 0.1% 间隔生成费率-费用对照表。

    参数：
        total: 第一部分工程费（万元）
        lo: 起始费率（%）
        hi: 结束费率（%）

    返回：[{"费率": "0.5%", "费用(万元)": 1.455}, ...]
    """
    detail = []
    rate = lo
    while rate <= hi + 0.001:  # 浮点容差
        fee = round(total * rate / 100.0, 4)
        detail.append({"费率": f"{rate:.1f}%", "费用(万元)": fee})
        rate = round(rate + 0.1, 1)
    return detail


def calc_kancha_rough(
    jianan: float,
    shebei: float = 0,
    project_type: str = "通用",
) -> dict:
    """
    工程勘察费粗略估算（《市政工程设计概算编制办法》，中国计划出版社）。

    - 通用项目：第一部分工程费 × 0.8%, 0.9%, 1.0%, 1.1%
    - 建筑项目：第一部分工程费 × 0.3%, 0.4%, 0.5%

    精确计算应依据计价格[2002]10号，按实物工作量定额计费。
    """
    total = jianan + shebei

    rates_map = {
        "建筑": (0.3, 0.5),
        "通用": (0.8, 1.1),
    }
    lo, hi = rates_map.get(project_type, (0.8, 1.1))

    detail = _build_rate_detail(total, lo, hi)
    fee_lo, fee_hi = detail[0]["费用(万元)"], detail[-1]["费用(万元)"]
    fee_mid = round((fee_lo + fee_hi) / 2, 4)

    return {
        "费种": "工程勘察费（粗略估算）",
        "依据": (
            "粗略估算依据《市政工程设计概算编制办法》（中国计划出版社）；"
            "精确计算依据《工程勘察设计收费管理规定》（计价格[2002]10号）工程勘察收费标准"
        ),
        "计算公式": f"第一部分工程费 × 费率（{project_type}项目，{lo}%~{hi}%，间隔 0.1%）",
        "参数": {
            "第一部分工程费(万元)": total,
            "= 建安工程费(万元)": jianan,
            "+ 设备购置费(万元)": shebei,
            "项目类型": project_type,
            "费率区间": f"{lo}%~{hi}%",
            "间隔": "0.1%",
        },
        "结果(万元)": f"{fee_lo:.2f} ~ {fee_hi:.2f}",
        "结果范围(万元)": f"{fee_lo:.2f} ~ {fee_hi:.2f}",
        "结果中值(万元)": fee_mid,
        "费率明细": detail,
        "计算步骤": [
            {"步骤": "确定第一部分工程费",
             "公式": f"建安工程费 + 设备购置费 = {jianan} + {shebei}",
             "结果": f"{total} 万元"},
            {"步骤": "判定项目类型",
             "公式": "根据查询关键词自动匹配",
             "结果": f"{project_type}项目"},
            {"步骤": "逐费率计算",
             "公式": f"第一部分工程费 × {lo}%~{hi}%，间隔 0.1%",
             "结果": f"共 {len(detail)} 档，详见费率明细表"},
        ],
        "说明": (
            f"第一部分工程费 {total:.0f} 万元（建安 {jianan:.0f} 万 + 设备 {shebei:.0f} 万），"
            f"{project_type}项目，费率 {lo}%~{hi}%（间隔 0.1%），"
            f"勘察费粗略估算范围为 **{fee_lo:.2f} ~ {fee_hi:.2f} 万元**"
            f"（中值约 **{fee_mid:.2f} 万元**）。\n\n"
            f"⚠️ 此为粗略估算，精确计算需按计价格[2002]10号以实物工作量定额计费。"
            f"如需精确计算，请提供勘察类型（工程测量/岩土勘察/水文地质等）和实物工作量"
            f"（钻探米数、测量面积等）。"
        ),
    }


def calc_laodong_anquan(total_wan: float) -> dict:
    """
    劳动安全卫生评审费（《市政工程设计概算编制办法》，中国计划出版社）。

    第一部分工程费用 × 0.1%, 0.2%, 0.3%, 0.4%, 0.5%
    """
    lo, hi = 0.1, 0.5
    detail = _build_rate_detail(total_wan, lo, hi)
    fee_lo, fee_hi = detail[0]["费用(万元)"], detail[-1]["费用(万元)"]
    fee_mid = round((fee_lo + fee_hi) / 2, 4)

    return {
        "费种": "劳动安全卫生评审费",
        "依据": "《市政工程设计概算编制办法》（中国计划出版社）",
        "计算公式": f"第一部分工程费用 × {lo}%~{hi}%（间隔 0.1%）",
        "参数": {"第一部分工程费用(万元)": total_wan, "费率区间": f"{lo}%~{hi}%", "间隔": "0.1%"},
        "结果(万元)": f"{fee_lo:.2f} ~ {fee_hi:.2f}",
        "结果范围(万元)": f"{fee_lo:.2f} ~ {fee_hi:.2f}",
        "结果中值(万元)": fee_mid,
        "费率明细": detail,
        "计算步骤": [
            {"步骤": "确定第一部分工程费用", "公式": "建安工程费 + 设备购置费", "结果": f"{total_wan} 万元"},
            {"步骤": "逐费率计算", "公式": f"费率区间 {lo}%~{hi}%，间隔 0.1%", "结果": f"共 {len(detail)} 档"},
        ],
        "说明": (
            f"第一部分工程费用 {total_wan:.0f} 万元，费率 {lo}%~{hi}%（间隔 0.1%），"
            f"劳动安全卫生评审费估算范围为 **{fee_lo:.2f} ~ {fee_hi:.2f} 万元**"
            f"（中值约 **{fee_mid:.2f} 万元**）。"
        ),
    }


def calc_changdi_zhunbei(total_wan: float) -> dict:
    """
    场地准备费及临时设施费（《市政工程设计概算编制办法》，中国计划出版社）。

    第一部分工程费用 × 0.5%, 0.6%, ..., 2.0%（间隔 0.1%）
    """
    lo, hi = 0.5, 2.0
    detail = _build_rate_detail(total_wan, lo, hi)
    fee_lo, fee_hi = detail[0]["费用(万元)"], detail[-1]["费用(万元)"]
    fee_mid = round((fee_lo + fee_hi) / 2, 4)

    return {
        "费种": "场地准备费及临时设施费",
        "依据": "《市政工程设计概算编制办法》（中国计划出版社）",
        "计算公式": f"第一部分工程费用 × {lo}%~{hi}%（间隔 0.1%）",
        "参数": {"第一部分工程费用(万元)": total_wan, "费率区间": f"{lo}%~{hi}%", "间隔": "0.1%"},
        "结果(万元)": f"{fee_lo:.2f} ~ {fee_hi:.2f}",
        "结果范围(万元)": f"{fee_lo:.2f} ~ {fee_hi:.2f}",
        "结果中值(万元)": fee_mid,
        "费率明细": detail,
        "计算步骤": [
            {"步骤": "确定第一部分工程费用", "公式": "建安工程费 + 设备购置费", "结果": f"{total_wan} 万元"},
            {"步骤": "逐费率计算", "公式": f"费率区间 {lo}%~{hi}%，间隔 0.1%", "结果": f"共 {len(detail)} 档"},
        ],
        "说明": (
            f"第一部分工程费用 {total_wan:.0f} 万元，费率 {lo}%~{hi}%（间隔 0.1%），"
            f"场地准备费及临时设施费估算范围为 **{fee_lo:.2f} ~ {fee_hi:.2f} 万元**"
            f"（中值约 **{fee_mid:.2f} 万元**）。"
        ),
    }


def calc_gongcheng_baoxian(total_wan: float) -> dict:
    """
    工程保险费（《市政工程设计概算编制办法》，中国计划出版社）。

    第一部分工程费用 × 0.3%, 0.4%, 0.5%, 0.6%
    """
    lo, hi = 0.3, 0.6
    detail = _build_rate_detail(total_wan, lo, hi)
    fee_lo, fee_hi = detail[0]["费用(万元)"], detail[-1]["费用(万元)"]
    fee_mid = round((fee_lo + fee_hi) / 2, 4)

    return {
        "费种": "工程保险费",
        "依据": "《市政工程设计概算编制办法》（中国计划出版社）",
        "计算公式": f"第一部分工程费用 × {lo}%~{hi}%（间隔 0.1%）",
        "参数": {"第一部分工程费用(万元)": total_wan, "费率区间": f"{lo}%~{hi}%", "间隔": "0.1%"},
        "结果(万元)": f"{fee_lo:.2f} ~ {fee_hi:.2f}",
        "结果范围(万元)": f"{fee_lo:.2f} ~ {fee_hi:.2f}",
        "结果中值(万元)": fee_mid,
        "费率明细": detail,
        "计算步骤": [
            {"步骤": "确定第一部分工程费用", "公式": "建安工程费 + 设备购置费", "结果": f"{total_wan} 万元"},
            {"步骤": "逐费率计算", "公式": f"费率区间 {lo}%~{hi}%，间隔 0.1%", "结果": f"共 {len(detail)} 档"},
        ],
        "说明": (
            f"第一部分工程费用 {total_wan:.0f} 万元，费率 {lo}%~{hi}%（间隔 0.1%），"
            f"工程保险费估算范围为 **{fee_lo:.2f} ~ {fee_hi:.2f} 万元**"
            f"（中值约 **{fee_mid:.2f} 万元**）。"
        ),
    }


# ============================================================
# 预备费 — 基本预备费
# ============================================================

def calc_yubei(part1_wan: float, erlei_wan: float, rate: float = 5.0) -> dict:
    """
    基本预备费 =（一类费（工程费用）+ 二类费（工程建设其他费））× 预备费率

    默认费率 5%。
    """
    base = part1_wan + erlei_wan
    fee = round(base * rate / 100.0, 4)

    return {
        "费种": "预备费（基本预备费）",
        "依据": "《市政工程设计概算编制办法》",
        "计算公式": f"（一类费 + 二类费）× {rate}%",
        "参数": {
            "一类费(工程费用)(万元)": part1_wan,
            "二类费(工程建设其他费)(万元)": erlei_wan,
            "预备费率(%)": rate,
            "计算基数(万元)": round(base, 4),
        },
        "结果(万元)": fee,
        "计算步骤": [
            {"步骤": "计算基数", "公式": "一类费（工程费用）+ 二类费（工程建设其他费）",
             "结果": f"{part1_wan} + {erlei_wan} = {round(base, 4)} 万元"},
            {"步骤": "计算预备费", "公式": f"基数 × {rate}%",
             "结果": f"{round(base, 4)} × {rate}% = {fee} 万元"},
        ],
        "说明": (
            f"一类费（工程费用）{part1_wan:.2f} 万 + 二类费（工程建设其他费）{erlei_wan:.2f} 万 "
            f"= {round(base, 4):.2f} 万，预备费率 {rate}%，"
            f"预备费为 **{fee:.2f} 万元**。"
        ),
    }


# ============================================================
# 环境影响咨询费 — 计价格[2002]125号
# ============================================================

# 分档定额基准价表（投资额：亿元，费用：万元）
# 注：0~0.1亿区间为固定最低价，0.1~0.3亿区间线性内插
_HUANPING_BRACKETS: list[tuple[float, float, float, float]] = [
    # (投资下限, 投资上限, 费用下限, 费用上限)
    (0,     0.1,    5,   5),   # ≤0.1亿：固定 5 万
    (0.1,   0.3,    5,   6),   # 0.1~0.3亿：5~6 万
    (0.3,   2,      6,  15),
    (2,    10,     15,  35),
    (10,   50,     35,  75),
    (50,  100,     75, 110),
    (100,  float("inf"), 110, 110),
]

_HUANPING_REPORT_TABLE_BRACKETS: list[tuple[float, float, float, float]] = [
    (0,     0.1,    1,   1),   # ≤0.1亿：固定 1 万
    (0.1,   0.3,    1,   2),   # 0.1~0.3亿：1~2 万
    (0.3,   2,      2,   4),
    (2,    10,      4,   7),
    (10,   float("inf"), 7, 7),
]

_HUANPING_EVAL_REPORT_BRACKETS: list[tuple[float, float, float, float]] = [
    (0,     0.1,  0.8, 0.8),   # ≤0.1亿：固定 0.8 万
    (0.1,   0.3,  0.8, 1.5),   # 0.1~0.3亿：0.8~1.5 万
    (0.3,   2,    1.5, 3),
    (2,    10,    3,   7),
    (10,   50,    7,   9),
    (50,  100,    9,  13),
    (100,  float("inf"), 13, 13),
]

_HUANPING_EVAL_TABLE_BRACKETS: list[tuple[float, float, float, float]] = [
    (0,     0.1,  0.5, 0.5),   # ≤0.1亿：固定 0.5 万
    (0.1,   0.3,  0.5, 0.8),   # 0.1~0.3亿：0.5~0.8 万
    (0.3,   2,    0.8, 1.5),
    (2,    10,    1.5, 2),
    (10,   float("inf"), 2, 2),
]

# 行业调整系数（计价格[2002]125号 附件二 表1）
_HUANPING_INDUSTRY_COEF: dict[str, float] = {
    "化工": 1.2, "冶金": 1.2, "有色": 1.2, "黄金": 1.2, "煤炭": 1.2,
    "矿产": 1.2, "纺织": 1.2, "化纤": 1.2, "轻工": 1.2, "医药": 1.2,
    "区域": 1.2,
    "石化": 1.1, "石油": 1.1, "天然气": 1.1, "水利": 1.1, "水电": 1.1, "旅游": 1.1,
    "林业": 1.0, "畜牧": 1.0, "渔业": 1.0, "农业": 1.0, "交通": 1.0,
    "铁道": 1.0, "民航": 1.0, "管线": 1.0, "建材": 1.0, "市政": 1.0,
    "烟草": 1.0, "兵器": 1.0,
    "邮电": 0.8, "广播电视": 0.8, "航空": 0.8, "机械": 0.8, "船舶": 0.8,
    "航天": 0.8, "电子": 0.8, "勘探": 0.8, "社会服务": 0.8, "火电": 0.8,
    "粮食": 0.6, "建筑": 0.6, "信息产业": 0.6, "仓储": 0.6,
}

# 环境敏感程度调整系数（计价格[2002]125号 附件二 表2）
_HUANPING_SENSITIVITY_COEF: dict[str, float] = {
    "敏感": 1.2, "一般": 0.8,
}


def _huanping_interpolate(invest_yi: float, brackets: list[tuple]) -> float:
    """线性内插计算分档定额基准价。"""
    for inv_lo, inv_hi, fee_lo, fee_hi in brackets:
        if invest_yi <= inv_hi or inv_hi == float("inf"):
            if inv_hi == float("inf") or invest_yi >= inv_hi:
                return fee_hi
            if invest_yi <= inv_lo:
                return fee_lo
            # 线性内插
            ratio = (invest_yi - inv_lo) / (inv_hi - inv_lo)
            return round(fee_lo + ratio * (fee_hi - fee_lo), 4)
    return 0.0


def _detect_huanping_industry(query: str) -> tuple[str, float]:
    """从查询中检测行业并返回对应的调整系数。"""
    # 按关键词长度从长到短匹配，避免"建筑"误匹配"建筑材料"
    sorted_industries = sorted(_HUANPING_INDUSTRY_COEF.keys(), key=len, reverse=True)
    for ind in sorted_industries:
        if re.search(ind, query):
            return ind, _HUANPING_INDUSTRY_COEF[ind]
    return "市政（默认）", 1.0


def calc_huanping(
    amount_wan: float,
    service_type: str = "编制报告书",
    industry_coef: float | None = None,
    industry_name: str = "",
    sensitivity_coef: float = 1.0,
) -> dict:
    """
    环境影响咨询费（计价格[2002]125号）。

    参数：
        amount_wan: 估算投资额（万元）
        service_type: 编制报告书 / 编制报告表 / 评估报告书 / 评估报告表
        industry_coef: 行业调整系数，None 则默认 1.0（市政）
        industry_name: 行业名称（用于展示）
        sensitivity_coef: 环境敏感程度调整系数（敏感=1.2, 一般=0.8, 默认1.0）
    """
    invest_yi = amount_wan / 10000.0  # 万元 → 亿元

    # 选择分档表
    bracket_map = {
        "编制报告书": _HUANPING_BRACKETS,
        "编制报告表": _HUANPING_REPORT_TABLE_BRACKETS,
        "评估报告书": _HUANPING_EVAL_REPORT_BRACKETS,
        "评估报告表": _HUANPING_EVAL_TABLE_BRACKETS,
    }
    brackets = bracket_map.get(service_type, _HUANPING_BRACKETS)

    # 线性内插计算基准价
    base_fee = _huanping_interpolate(invest_yi, brackets)

    # 行业调整系数
    if industry_coef is None:
        industry_coef = 1.0
        industry_name = industry_name or "市政（默认）"

    # 总调整系数 = 行业 × 敏感度
    total_coef = round(industry_coef * sensitivity_coef, 4)

    # 最终费用（基准 × 总调整系数）
    final_fee_mid = round(base_fee * total_coef, 4)
    # 上下 20% 协商浮动
    final_fee_lo = round(base_fee * total_coef * 0.8, 4)
    final_fee_hi = round(base_fee * total_coef * 1.2, 4)

    sensitivity_label = {1.2: "敏感", 0.8: "一般", 1.0: "未指定"}.get(sensitivity_coef, str(sensitivity_coef))

    steps = [
        {"步骤": "估算投资额", "公式": f"{amount_wan} 万元", "结果": f"{invest_yi:.2f} 亿元"},
        {"步骤": "确定服务类型", "公式": "", "结果": service_type},
    ]

    # 显示匹配到的分档区间
    for inv_lo, inv_hi, fee_lo, fee_hi in brackets:
        if invest_yi <= inv_hi or inv_hi == float("inf"):
            if invest_yi <= inv_lo:
                steps.append({"步骤": "所在分档区间", "公式": f"≤{inv_lo}亿元", "结果": f"基准价 {fee_lo} 万元"})
            elif inv_hi == float("inf") or invest_yi >= inv_hi:
                steps.append({"步骤": "所在分档区间", "公式": f"≥{inv_lo}亿元", "结果": f"基准价 {fee_hi} 万元"})
            else:
                steps.append({"步骤": "所在分档区间", "公式": f"{inv_lo}~{inv_hi}亿元", "结果": f"基准价 {fee_lo}~{fee_hi} 万元"})
            break

    steps.extend([
        {"步骤": "线性内插基准价", "公式": f"插值({invest_yi:.2f})", "结果": f"{base_fee:.2f} 万元"},
        {"步骤": "行业调整系数", "公式": f"{industry_name}", "结果": str(industry_coef)},
        {"步骤": "环境敏感程度系数", "公式": sensitivity_label, "结果": str(sensitivity_coef)},
        {"步骤": "总调整系数", "公式": f"{industry_coef} × {sensitivity_coef}", "结果": str(total_coef)},
        {"步骤": "最终费用(协商浮动±20%前)", "公式": f"{base_fee:.2f} × {total_coef}", "结果": f"{final_fee_mid:.2f} 万元"},
    ])

    return {
        "费种": f"环境影响咨询费（{service_type}）",
        "依据": "《关于规范环境影响咨询收费有关问题的通知》（计价格[2002]125号）\n"
                "国家计委、国家环境保护总局，2002年1月31日",
        "计算公式": f"最终收费 = 分档定额基准价 × 行业调整系数({industry_coef}) × 环境敏感程度系数({sensitivity_coef}) × (1 ± 20%)",
        "参数": {
            "估算投资额(万元)": amount_wan,
            "估算投资额(亿元)": invest_yi,
            "服务类型": service_type,
            "行业": f"{industry_name}（系数 {industry_coef}）",
            "环境敏感程度": f"{sensitivity_label}（系数 {sensitivity_coef}）",
            "总调整系数": total_coef,
            "协商浮动": "±20%",
        },
        "结果(万元)": f"{final_fee_lo:.2f} ~ {final_fee_hi:.2f}",
        "结果范围(万元)": f"{final_fee_lo:.2f} ~ {final_fee_hi:.2f}",
        "结果中值(万元)": final_fee_mid,
        "基准价(万元)": base_fee,
        "调整系数明细": {"行业系数": industry_coef, "敏感度系数": sensitivity_coef, "总系数": total_coef},
        "计算步骤": steps,
        "说明": (
            f"估算投资额 {invest_yi:.2f} 亿元（{amount_wan:.0f}万元），服务类型「{service_type}」，"
            f"行业「{industry_name}」系数 {industry_coef}，环境{ sensitivity_label }系数 {sensitivity_coef}。\n"
            f"分档定额基准价 **{base_fee:.2f} 万元** × 总调整系数 **{total_coef}**"
            f" = 基准收费 **{final_fee_mid:.2f} 万元**。\n"
            f"可在上下 20% 幅度内协商确定，即 **{final_fee_lo:.2f} ~ {final_fee_hi:.2f} 万元**。"
        ),
    }


def calc_huanping_multi(
    amount_wan: float,
    selected_services: list[str],
    industry_coef: float = 1.0,
    industry_name: str = "",
    sensitivity_coef: float = 1.0,
) -> dict:
    """环境影响咨询费 — 多服务类型选择计算。

    对用户选择的每项服务分别计算，汇总合计（中值）。
    """
    details: list[dict] = []
    total = 0.0

    for svc in selected_services:
        r = calc_huanping(
            amount_wan, svc,
            industry_coef=industry_coef,
            industry_name=industry_name,
            sensitivity_coef=sensitivity_coef,
        )
        mid = r["结果中值(万元)"]
        details.append({
            "服务类型": svc,
            "结果(万元)": r["结果(万元)"],
            "结果中值(万元)": mid,
            "基准价(万元)": r.get("基准价(万元)", 0),
            "计算步骤": r.get("计算步骤", []),
            "调整系数明细": r.get("调整系数明细", {}),
        })
        total += mid

    total = round(total, 4)
    invest_yi = amount_wan / 10000.0

    desc_parts = []
    for d in details:
        desc_parts.append(
            f"- **{d['服务类型']}**：{d['结果(万元)']} 万元"
            f"（中值 **{d['结果中值(万元)']} 万元**）"
        )

    sensitivity_label = {1.2: "敏感", 0.8: "一般", 1.0: "未指定"}.get(sensitivity_coef, str(sensitivity_coef))

    return {
        "费种": "环境影响咨询费",
        "依据": "《关于规范环境影响咨询收费有关问题的通知》（计价格[2002]125号）",
        "明细": details,
        "合计(万元)": total,
        "参数": {
            "估算投资额(万元)": amount_wan,
            "估算投资额(亿元)": invest_yi,
            "行业调整系数": industry_coef,
            "行业名称": industry_name or "市政（默认）",
            "环境敏感程度系数": sensitivity_coef,
            "环境敏感程度": sensitivity_label,
            "总调整系数": round(industry_coef * sensitivity_coef, 4),
            "协商浮动": "±20%",
        },
        "说明": "### 服务类型明细\n\n" + "\n".join(desc_parts)
                + f"\n\n### 💰 合计（中值）：**{total} 万元**",
    }


# ============================================================
# 费种参考信息（无金额时返回费率表/规则说明）
# ============================================================

def _get_fee_reference(fee_type: str) -> dict:
    """返回各费种的费率表和计费规则参考（用户未提供金额时使用）。"""
    refs: dict[str, dict] = {
        "建设管理费": {
            "费种": "建设管理费（建设单位管理费）",
            "依据": "《基本建设项目建设成本管理规定》（财建[2016]504号）",
            "计费方式": "差额分档累进",
            "参数说明": "工程总概算（万元）",
            "费率表": [
                ("工程总概算", "费率"),
                ("≤1000 万元", "2.0%"),
                ("1001~5000 万元", "1.5%"),
                ("5001~10000 万元", "1.2%"),
                ("10001~50000 万元", "1.0%"),
                ("50001~100000 万元", "0.8%"),
                (">100000 万元", "0.4%"),
            ],
            "计算说明": (
                "各档位分段计算后累加。例如总概算 8000 万：\n"
                "1000×2.0% + 4000×1.5% + 3000×1.2% = 20 + 60 + 36 = 116 万元"
            ),
        },
        "招标代理费": {
            "费种": "招标代理服务费",
            "依据": "《招标代理业务收费管理暂行办法》（计价格[2002]1980号）",
            "计费方式": "差额定率累进（上下浮动不超过 20%）",
            "参数说明": "中标金额（万元）+ 招标类型（货物/服务/工程）",
            "费率表": [
                ("档位", "货物招标", "服务招标", "工程招标"),
                ("≤100 万元", "1.5%", "1.5%", "1.0%"),
                ("100~500 万元", "1.1%", "0.8%", "0.7%"),
                ("500~1000 万元", "0.8%", "0.45%", "0.55%"),
                ("1000~5000 万元", "0.5%", "0.25%", "0.35%"),
                ("5000~10000 万元", "0.25%", "0.1%", "0.2%"),
                ("10000~100000 万元", "0.05%", "0.05%", "0.05%"),
                (">100000 万元", "0.01%", "0.01%", "0.01%"),
            ],
        },
        "交易服务费": {
            "费种": "工程建设交易服务费",
            "依据": "《市发展改革委关于规范工程建设交易服务收费标准有关问题的通知》（津发改价管[2017]979号）",
            "计费方式": "分 4 类（施工/设备/监理/设计）分别按中标额分档定额，合计 = 四类之和。招标方 60%，中标方 40%",
            "参数说明": "建安工程费 + 设备购置费（用于内算监理费和设计费基数）",
            "费率表": [
                ("中标额（万元）", "最高收费标准（元）"),
                ("≤100", "400"),
                ("100~500", "1,000"),
                ("500~1000", "3,000"),
                ("1000~3000", "7,000"),
                ("3000~5000", "15,000"),
                ("5000~8000", "25,000"),
                ("8000~10000", "35,000"),
                (">10000", "50,000"),
            ],
            "计算说明": (
                "施工基数 = 建安工程费，设备基数 = 设备购置费，\n"
                "监理基数 = 监理费（670号文计算），设计基数 = 设计费（10号文计算）。\n"
                "每类各自查上表后相加。"
            ),
        },
        "监理费": {
            "费种": "施工监理服务费",
            "依据": "《建设工程监理与相关服务收费管理规定》（发改价格[2007]670号）",
            "计费方式": "收费基价 = 计费额线性内插查表 × 专业调整系数 × 复杂程度系数 × 高程调整系数。可上下浮动 20%",
            "参数说明": "计费额（万元），或建安费+设备费（触发 1.0.8 条 40% 规则）",
            "费率表": [
                ("计费额（万元）", "收费基价（万元）"),
                ("500", "16.5"),
                ("1000", "30.1"),
                ("3000", "78.1"),
                ("5000", "120.8"),
                ("8000", "181.0"),
                ("10000", "218.6"),
                ("20000", "393.4"),
                ("40000", "708.2"),
                ("60000", "991.4"),
                ("80000", "1,255.8"),
                ("100000", "1,507.0"),
                (">1000000", "计费额×1.039%"),
            ],
            "特殊规则": (
                "1.0.8 条（40% 规则）：设备+联合试运转费占比 > 40% 时，\n"
                "设备费按 40% 计入计费额，且不低于等建安费 40% 设备占比假想项目的计费额。\n"
                "保底下限 = 建安费 × 5/3。"
            ),
            "调整系数": (
                "专业调整系数（附表三）：园林绿化 0.8、矿山采选/农林 0.9、"
                "建筑市政公路等 1.0、桥梁隧道地铁 1.1、核电水电水库 1.2\n"
                "复杂程度系数：I级/简单 0.85、II级/较复杂 1.0、III级/复杂 1.15\n"
                "高程调整系数：≤2000m 1.0、2001~3000m 1.1、"
                "3001~3500m 1.2、3501~4000m 1.3、>4000m 协商"
            ),
        },
        "工程设计费": {
            "费种": "工程设计费",
            "依据": "《工程勘察设计收费管理规定》（计价格[2002]10号）",
            "计费方式": "工程设计收费 = 工程设计收费基准价 × (1 ± 浮动幅度值)\n"
                        "工程设计收费基准价 = 基本设计收费 + 其他设计收费\n"
                        "基本设计收费 = 收费基价 × 专业调整系数 × 复杂程度调整系数 × 附加调整系数",
            "参数说明": "计费额（万元）× 专业类型 × 复杂程度 × 可选附加项（总体设计/施工图预算/竣工图等）",
            "费率表": [
                ("计费额（万元）", "收费基价（万元）"),
                ("200", "9.0"),
                ("500", "20.9"),
                ("1000", "38.8"),
                ("3000", "103.8"),
                ("5000", "163.9"),
                ("8000", "249.6"),
                ("10000", "304.8"),
                ("20000", "566.8"),
                ("40000", "1,054.0"),
                ("60000", "1,515.2"),
                ("80000", "1,960.1"),
                ("100000", "2,393.4"),
                (">2000000", "计费额×1.6%"),
            ],
            "调整系数": (
                "工程设计费（计价格[2002]10号）的三个调整系数：\n"
                "专业调整系数（附表二 完整）：\n"
                "  1.矿山采选：黑色/黄金/化学/非金属 1.1、采煤/有色/铀 1.2、选煤/煤炭 1.3\n"
                "  2.加工冶炼：冷加工 1.0、船舶水工 1.1、冶炼/热加工/压力加工 1.2、核加工 1.3\n"
                "  3.石油化工：石油/化工/石化/化纤/医药 1.2、核化工 1.6\n"
                "  4.水利电力：风力发电/水利 0.8、火电 1.0、核电常规岛/水电/水库/送变电 1.2、核能 1.6\n"
                "  5.交通运输：机场场道 0.8、公路/城市道路 0.9、空管/助航灯光/轻轨 1.0、水运/地铁/桥梁/隧道 1.1、索道 1.3\n"
                "  6.建筑市政：邮政工艺 0.8、建筑/市政/电信 1.0、人防/园林绿化/广电 1.1\n"
                "  7.农业林业：农业 0.9、林业 0.8\n"
                "复杂程度系数（1.0.9.2）：I级/一般 0.85、II级/较复杂 1.0、III级/复杂 1.15\n"
                "附加调整系数（1.0.9.3）：多个附加系数不能连乘！"
                "合并公式 = 各系数相加 − 系数个数 + 1\n"
                "⚠️ 工程设计费**不包含**高程调整系数！（高程系数仅用于监理费 发改价格[2007]670号）"
            ),
            "计算说明": (
                "**计费额定义（1.0.8）**：计费额 = 建筑安装工程费 + 设备与工器具购置费 + 联合试运转费。\n"
                "利用原有设备的，按同类设备当期价格；引进设备的，按离岸价折合人民币。\n\n"
                "**收费基价（1.0.7）**：在附表一中查找，计费额处于两个数值区间的采用直线内插法。\n\n"
                "**其他设计收费（1.0.6）**：包括总体设计费、主体设计协调费、采用标准设计和复用设计费、"
                "非标准设备设计文件编制费、施工图预算编制费、竣工图编制费等。\n\n"
                "**常见附加项**：\n"
                "- 总体设计费（1.0.13）：基本设计收费的 5%\n"
                "- 主体设计协调费（1.0.14）：基本设计收费的 5%\n"
                "- 施工图预算编制费（1.0.16）：基本设计收费的 10%\n"
                "- 竣工图编制费（1.0.16）：基本设计收费的 8%\n"
                "- 改扩建项目（1.0.12）：附加调整系数 1.1~1.4\n"
                "- 标准设计/复用设计（1.0.15）：同类新建项目基本设计收费的 30%；"
                "需重新基础设计的按 40%；局部修改协商确定\n"
                "- 非标准设备设计费（1.0.10）：非标准设备计费额 × 非标准设备设计费率（附表三）"
            ),
        },
        "勘察费": {
            "费种": "工程勘察费",
            "依据": "《工程勘察设计收费管理规定》（计价格[2002]10号）— 工程勘察收费标准",
            "计费方式": (
                "工程勘察收费 = 工程勘察收费基准价 × (1 ± 浮动幅度值，≤20%)\n"
                "工程勘察收费基准价 = 工程勘察实物工作收费 + 工程勘察技术工作收费\n"
                "工程勘察实物工作收费 = 实物工作收费基价 × 实物工作量 × 附加调整系数\n"
                "工程勘察技术工作收费 = 工程勘察实物工作收费 × 技术工作收费比例\n\n"
                "⚠️ 工程勘察费按**实物工作量**定额计费，不是按投资额比例计算，"
                "因此仅凭建安费/设备费无法直接算出勘察费——需要知道具体勘察工作量"
                "（如钻探米数、测量面积等）。"
            ),
            "参数说明": (
                "需明确的参数：\n"
                "1. 勘察类型（工程测量/岩土工程勘察/水文地质勘察/工程物探等 16 大类）\n"
                "2. 实物工作量（钻孔深度、测量面积、取样数量等）\n"
                "3. 复杂程度（简单/中等/复杂）\n"
                "4. 附加调整系数（气温/高程/带状/水域等）\n"
                "5. 技术工作收费比例（各类勘察不同，如工程测量 22%、岩土甲级 120% 等）"
            ),
            "费率表": [
                ("勘察类型", "技术工作费比例", "典型实物工作", "说明"),
                ("工程测量", "22%", "km²/比例尺", "地面/水域/地下管线/洞室"),
                ("岩土工程勘察", "甲120%/乙100%/丙80%", "钻孔深度 m", "含钻探/取土/原位测试"),
                ("岩土工程设计检测", "见具体章节", "组日/台班", "验槽/检测/监测"),
                ("水文地质勘察", "简单15%/中等18%/复杂20%", "供水井深度 m", "含抽水试验等"),
                ("工程水文气象勘察", "22%", "—", "水文气象观测与分析"),
                ("工程物探", "22%", "组日/标准点", "电法/地震/测井等"),
                ("室内试验", "10%", "件/组", "土工/岩石/水质分析"),
            ],
            "计算说明": (
                "**16 大类专业工程勘察**（计价格[2002]10号）：\n"
                "1. 总则  2. 工程测量  3. 岩土工程勘察  4. 岩土工程设计与检测监测\n"
                "5. 水文地质勘察  6. 工程水文气象勘察  7. 工程物探  8. 室内试验\n"
                "9. 煤炭  10. 水利水电  11. 电力  12. 长输管道\n"
                "13. 铁路  14. 公路  15. 通信  16. 海洋工程\n\n"
                "**附加调整系数合并公式**（与设计费相同）：多个系数不能连乘，"
                "合并 = 各系数相加 − 个数 + 1\n\n"
                "**气温附加**：≥35°C 或 ≤-10°C 时，气温附加系数 1.2\n"
                "**高程附加**：2000-3000m 1.1，3001-3500m 1.2，3501-4000m 1.3，>4000m 协商\n\n"
                "**粗略估算**（《市政工程设计概算编制办法》，中国计划出版社）：\n"
                "- 通用项目：第一部分工程费 × 0.8%~1.1%（间隔 0.1%）\n"
                "- 建筑项目：第一部分工程费 × 0.3%~0.5%（间隔 0.1%）\n\n"
                "⚠️ 以上为粗略估算方法，精确计算仍需按计价格[2002]10号以实物工作量定额计费。"
            ),
        },
        "可行性研究费": {
            "费种": "建设项目前期工作咨询费",
            "依据": "《建设项目前期工作咨询收费暂行规定》（计价格[1999]1283号）",
            "计费方式": "按估算投资额分档线性内插基准价，乘以行业调整系数（0.7~1.3）和复杂程度系数（0.8~1.2）",
            "参数说明": "估算投资额（亿元）+ 服务类型（编制可研/编制建议书/评估可研/评估建议书）",
            "费率表": [
                ("估算投资额", "编制建议书", "编制可研", "评估建议书", "评估可研"),
                ("<500 万", "1.3 万", "2.5 万", "1.0 万", "1.3 万"),
                ("500~1000 万", "2.5 万", "5.0 万", "1.7 万", "2.5 万"),
                ("1000~3000 万", "2.5~6 万", "5~12 万", "1.7~4 万", "2.5~5 万"),
                ("0.3~1 亿元", "6~14 万", "12~28 万", "4~8 万", "5~10 万"),
                ("1~5 亿元", "14~37 万", "28~75 万", "8~12 万", "10~15 万"),
                ("5~10 亿元", "37~55 万", "75~110 万", "12~15 万", "15~20 万"),
                ("10~50 亿元", "55~100 万", "110~200 万", "15~17 万", "20~25 万"),
                (">50 亿元", "100~125 万", "200~250 万", "17~20 万", "25~35 万"),
            ],
            "计算说明": (
                "行业调整系数：石化/化工/钢铁 1.3，石油/天然气/水利/水电/水运/化纤 1.2，"
                "有色/黄金/纺织/轻工/邮电/广电/医药/煤炭/火电/机械 1.0，"
                "林业/商业/粮食/建筑 0.8，建材/公路/铁道/市政 0.7；"
                "复杂程度系数 0.8~1.2"
            ),
        },
        "施工图审查费": {
            "费种": "施工图审查费",
            "依据": "《市发展改革委关于施工图审查收费标准的通知》（津价管[2011]46号）",
            "计费方式": "住宅类按建筑面积计（元/m²）；公建/工业/市政类按勘察设计费 × 费率",
            "参数说明": "项目类型（住宅/公建/工业/市政）+ 规模（大/中/小）+ 面积或设计费",
            "费率表": [
                ("类别", "大型", "中型", "小型"),
                ("住宅（元/m²）", "1.9", "1.7", "1.3"),
                ("公建（%勘察设计费）", "3.2%", "2.9%", "2.4%"),
                ("工业（%勘察设计费）", "3.2%", "3.0%", "2.8%"),
                ("市政（%勘察设计费）", "4.8%", "4.0%", "3.2%"),
            ],
            "计算说明": "幕墙/深基坑等单项工程按 1.6‰ 计取，最低 1000 元",
        },
        "水土保持费": {
            "费种": "水土保持咨询服务费",
            "依据": "《关于开发建设项目水土保持咨询服务费用计列的指导意见》（保监[2005]22号）",
            "计费方式": "按主体工程土建投资内插查表（可乘以地貌调整系数：山区 1.2、丘陵及风沙区 1.0、平原区 0.8）",
            "参数说明": "土建投资（亿元）+ 服务类型（方案编制/监测/验收评估/技术咨询）",
            "费率表": [
                ("土建投资（亿元）", "方案编制（万元）", "监测费（万元）", "验收评估（万元）", "技术咨询（万元）"),
                ("0.5", "30", "30", "10", "1.0"),
                ("1.0", "52", "60", "18", "1.5"),
                ("2.0", "72", "90", "30", "2.0"),
                ("3.0", "82", "140", "36", "2.5"),
                ("5.0", "104", "220", "48", "3.2"),
                ("10.0", "171", "420", "78", "5.2"),
                ("15.0", "245", "600", "119", "7.5"),
                ("20.0", "350", "760", "160", "9.5"),
            ],
            "计算说明": "土建投资在两个值之间时取线性内插；超过最大值按最后一个值计取",
        },
        "环境影响咨询费": {
            "费种": "环境影响咨询费",
            "依据": "《关于规范环境影响咨询收费有关问题的通知》（计价格[2002]125号）",
            "计费方式": "按估算投资额分档定额（上下浮动 20%），也可按咨询服务工日计费",
            "参数说明": "估算投资额（万元）+ 服务类型（报告书/报告表/评估报告书/评估报告表）",
            "费率表": [
                ("估算投资额（亿元）", "编制报告书（万元）", "编制报告表（万元）", "评估报告书（万元）", "评估报告表（万元）"),
                ("≤0.3", "5~6", "1~2", "0.8~1.5", "0.5~0.8"),
                ("0.3~2", "6~15", "2~4", "1.5~3", "0.8~1.5"),
                ("2~10", "15~35", "4~7", "3~7", "1.5~2"),
                ("10~50", "35~75", "7 以上", "7~9", "2 以上"),
                ("50~100", "75~110", "—", "9~13", "—"),
                (">100", "110 以上", "—", "13 以上", "—"),
            ],
            "计算说明": "⚠️ 125号文精确费率表在原PDF附件中，上表为区间参考值。具体收费在此范围内由双方协商确定。",
        },
        "劳动安全卫生评审费": {
            "费种": "劳动安全卫生评审费",
            "依据": "《市政工程设计概算编制办法》（中国计划出版社）",
            "计费方式": "第一部分工程费用 × 0.1%~0.5%",
            "参数说明": "第一部分工程费用 = 建安工程费 + 设备购置费（万元）",
            "费率表": [
                ("费率下限", "费率上限"),
                ("0.1%", "0.5%"),
            ],
            "计算说明": (
                "劳动安全卫生评审费 = 第一部分工程费用 × 0.1%~0.5%。\n"
                "例如：第一部分工程费用 1000 万元，评审费 ≈ 1~5 万元。"
            ),
        },
        "场地准备费及临时设施费": {
            "费种": "场地准备费及临时设施费",
            "依据": "《市政工程设计概算编制办法》（中国计划出版社）",
            "计费方式": "第一部分工程费用 × 0.5%~2.0%",
            "参数说明": "第一部分工程费用 = 建安工程费 + 设备购置费（万元）",
            "费率表": [
                ("费率下限", "费率上限"),
                ("0.5%", "2.0%"),
            ],
            "计算说明": (
                "场地准备费及临时设施费 = 第一部分工程费用 × 0.5%~2.0%。\n"
                "例如：第一部分工程费用 1000 万元，场地准备费 ≈ 5~20 万元。"
            ),
        },
        "工程保险费": {
            "费种": "工程保险费",
            "依据": "《市政工程设计概算编制办法》（中国计划出版社）",
            "计费方式": "第一部分工程费用 × 0.3%~0.6%",
            "参数说明": "第一部分工程费用 = 建安工程费 + 设备购置费（万元）",
            "费率表": [
                ("费率下限", "费率上限"),
                ("0.3%", "0.6%"),
            ],
            "计算说明": (
                "工程保险费 = 第一部分工程费用 × 0.3%~0.6%。\n"
                "例如：第一部分工程费用 1000 万元，保险费 ≈ 3~6 万元。"
            ),
        },
        "造价咨询费": {
            "费种": "造价咨询费（工程造价咨询服务费）",
            "依据": "《天津市建设工程造价咨询服务项目和价格标准》（津价房地[2008]136号）",
            "计费方式": "差额定率分档累进，基准价可上下浮动 ±20%。\n"
                        "编制类/审核类（除审核概算外）基数 = 工程费用（建安+设备）；\n"
                        "审核概算基数 = 工程总投资；\n"
                        "编制投资估算/设计概算基数 = 建安工程费用。",
            "参数说明": "工程费用（万元）+ 服务类型（编制工程量清单/编制标底/编制施工图预算/编制竣工结算/"
                        "全过程造价控制/审核概算/审核预算标底/审核竣工结算/编制投资估算/编制设计概算）",
            "费率表": [
                ("服务类型", "≤100万", "≤500万", "≤1000万", "≤5000万", "≤10000万", ">10000万"),
                ("编制工程量清单", "3.4‰", "3.2‰", "3.0‰", "2.4‰", "2.0‰", "1.6‰"),
                ("编制标底(含清单)", "3.6‰", "3.4‰", "3.1‰", "2.6‰", "2.0‰", "1.7‰"),
                ("编制施工图预算", "3.6‰", "3.4‰", "3.1‰", "2.6‰", "2.0‰", "1.7‰"),
                ("编制竣工结算", "3.6‰", "3.4‰", "3.1‰", "2.6‰", "2.0‰", "1.7‰"),
                ("全过程造价控制", "10‰", "9.0‰", "8.0‰", "7.5‰", "7.0‰", "6.0‰"),
                ("审核概算", "3.0‰", "2.5‰", "2.0‰", "1.5‰", "1.2‰", "1.0‰"),
                ("审核预算、标底", "3.5‰", "3.1‰", "2.2‰", "1.9‰", "1.2‰", "0.9‰"),
                ("审核竣工结算", "3.5‰", "3.1‰", "2.2‰", "1.9‰", "1.2‰", "0.9‰"),
                ("编制项目投资估算", "0.8‰", "0.7‰", "0.6‰", "0.5‰", "0.3‰", "0.15‰"),
                ("编制设计概算", "1.7‰", "1.5‰", "1.2‰", "0.85‰", "0.7‰", "0.4‰"),
            ],
            "计算说明": (
                "差额分档累进计费。例：编制施工图预算，工程费用 5000 万：\n"
                "100×3.6‰ + 400×3.4‰ + 500×3.1‰ + 4000×2.6‰ = 0.36 + 1.36 + 1.55 + 10.40 = 13.67 万元\n\n"
                "追加收费：审核中审减(增)额超过 ±5% 时，超过部分按 5% 计收。\n"
                "钢筋及预埋件计算：11.00 元/吨。\n"
                "工程造价争议鉴定：标的 ≤500万 按 4%，>500万 按 2%。"
            ),
        },
    }
    return refs.get(fee_type, {
        "费种": fee_type,
        "依据": "请查阅相关政策文件",
        "计费方式": "请提供具体金额进行计算",
        "参数说明": "",
        "费率表": [],
    })

def _build_coef_metadata(fee_type: str, result: dict, query: str) -> dict:
    """为交互式系数选择构建元数据：各系数选项表 + 当前值 + 重算所需参数。"""
    params = result.get("参数", {})

    def _find_label(value: float, options: list[tuple[str, float]]) -> str:
        """在选项表中找到匹配的标签，找不到则返回自定义标签。"""
        for label, val in options:
            if abs(val - value) < 0.005:
                return label
        return f"自定义 ({value})"

    if fee_type == "监理费":
        prof = params.get("专业调整系数", 1.0)
        comp = params.get("复杂程度系数", 1.0)
        elev = params.get("高程调整系数", 1.0)
        amount_wan = params.get("计费额(万元)")
        jianan, shebei = _extract_jianli_components(query)

        return {
            "fee_label": "施工监理服务费",
            "calc_func": "calc_jianli",
            "coefs": [
                {
                    "key": "专业调整系数",
                    "param_name": "professional_coef",
                    "current": prof,
                    "current_label": _find_label(prof, JIANLI_PROFESSIONAL_OPTIONS),
                    "options": JIANLI_PROFESSIONAL_OPTIONS,
                    "description": "发改价格[2007]670号 附表三",
                },
                {
                    "key": "复杂程度系数",
                    "param_name": "complexity_coef",
                    "current": comp,
                    "current_label": _find_label(comp, JIANLI_COMPLEXITY_OPTIONS),
                    "options": JIANLI_COMPLEXITY_OPTIONS,
                    "description": "发改价格[2007]670号 1.0.9条",
                },
                {
                    "key": "高程调整系数",
                    "param_name": "elevation_coef",
                    "current": elev,
                    "current_label": _find_label(elev, JIANLI_ELEVATION_OPTIONS),
                    "options": JIANLI_ELEVATION_OPTIONS,
                    "description": "发改价格[2007]670号 1.0.9条",
                },
            ],
            "base_params": {
                "amount_wan": amount_wan,
                "jianan": jianan,
                "shebei": shebei,
            },
        }

    elif fee_type == "工程设计费":
        prof = params.get("专业调整系数", 1.0)
        comp = params.get("复杂程度系数", 1.0)
        addi = params.get("附加调整系数", 1.0)
        amount_wan = params.get("计费额(万元)")

        # 附加项的元数据
        basic_design = result.get("基本设计收费(万元)", 0)
        other_items = result.get("其他设计收费明细", [])
        other_labels = [it["项目"] for it in other_items]

        # 检查是否有附加系数明细（多个附加系数）
        addi_detail = params.get("附加系数明细", "")

        return {
            "fee_label": "工程设计费",
            "calc_func": "calc_sheji",
            "coefs": [
                {
                    "key": "专业调整系数",
                    "param_name": "professional_coef",
                    "current": prof,
                    "current_label": _find_label(prof, SHEJI_PROFESSIONAL_OPTIONS),
                    "options": SHEJI_PROFESSIONAL_OPTIONS,
                    "description": "计价格[2002]10号 附表二",
                },
                {
                    "key": "复杂程度系数",
                    "param_name": "complexity_coef",
                    "current": comp,
                    "current_label": _find_label(comp, SHEJI_COMPLEXITY_OPTIONS),
                    "options": SHEJI_COMPLEXITY_OPTIONS,
                    "description": "计价格[2002]10号 1.0.9.2",
                },
                {
                    "key": "附加调整系数",
                    "param_name": "additional_coef",
                    "current": addi,
                    "current_label": f"{addi}" + (f"（{addi_detail}）" if addi_detail else ""),
                    "options": [],  # 无预设选项，全靠自定义
                    "description": "计价格[2002]10号 1.0.9.3（多个系数合并 = 相加 − 个数 + 1）",
                },
            ],
            "base_params": {
                "amount_wan": amount_wan,
                "other_labels": other_labels,
            },
        }

    elif fee_type == "环境影响咨询费":
        ind_coef = params.get("行业", "")
        sens_coef_raw = params.get("环境敏感程度", "")
        # 从参数字符串中提取数值
        ind_coef_val = 1.0
        m = re.search(r"系数\s*([\d.]+)", str(ind_coef))
        if m:
            ind_coef_val = float(m.group(1))
        sens_coef_val = 1.0
        m2 = re.search(r"系数\s*([\d.]+)", str(sens_coef_raw))
        if m2:
            sens_coef_val = float(m2.group(1))

        # 获取 service_type
        svc = "编制报告书"
        for label in ["编制报告书", "编制报告表", "评估报告书", "评估报告表"]:
            if label in result.get("费种", ""):
                svc = label
                break
        # 也检测查询
        if re.search(r"报告表|报告书.*表", query):
            svc = "编制报告表"
        elif re.search(r"评估报告书|评估.*报告书", query):
            svc = "评估报告书"
        elif re.search(r"评估报告表|评估.*报告表", query):
            svc = "评估报告表"

        amount_wan = params.get("估算投资额(万元)", 0)

        return {
            "fee_label": "环境影响咨询费",
            "calc_func": "calc_huanping",
            "coefs": [
                {
                    "key": "行业调整系数",
                    "param_name": "industry_coef",
                    "current": ind_coef_val,
                    "current_label": _find_label(ind_coef_val, HUANPING_INDUSTRY_OPTIONS),
                    "options": HUANPING_INDUSTRY_OPTIONS,
                    "description": "计价格[2002]125号 附件二 表1",
                },
                {
                    "key": "环境敏感程度系数",
                    "param_name": "sensitivity_coef",
                    "current": sens_coef_val,
                    "current_label": _find_label(sens_coef_val, HUANPING_SENSITIVITY_OPTIONS),
                    "options": HUANPING_SENSITIVITY_OPTIONS,
                    "description": "计价格[2002]125号 附件二 表2",
                },
            ],
            "base_params": {
                "amount_wan": amount_wan,
                "service_type": svc,
            },
        }

    else:
        return {"fee_label": "", "calc_func": "", "coefs": [], "base_params": {}}


# ============================================================
# 依赖费种交互式配置 — 招标代理费 & 施工图审查费
# ============================================================

def _get_dependent_fee_list(target_fee: str) -> list[str]:
    """返回目标费种需要的依赖费种列表（按计算顺序）。"""
    if target_fee == "招标代理费":
        return ["监理费", "工程设计费", "勘察费"]
    elif target_fee == "施工图审查费":
        return ["工程设计费", "勘察费"]
    return []


def _build_dependent_config_meta(
    target_fee: str,
    base_params: dict,
    query: str,
) -> list[dict]:
    """为每个依赖费种构建前端交互式配置所需的元数据。

    返回列表，每个元素对应一个依赖费种，包含：
    - fee_type / fee_label: 费种标识和显示名
    - config_type: "coef"（系数下拉）或 "rate"（费率单选）
    - coef_metadata（coef 类型）或 rate_options（rate 类型）
    - base_params: 传给对应 calc_* 函数的参数
    - default_result: 默认参数下的计算结果
    - preview_fee: 预览费用（万元）
    """
    deps = []
    jianan = base_params.get("jianan", 0) or 0
    shebei = base_params.get("shebei", 0) or 0
    amount_wan = base_params.get("amount_wan", jianan + shebei)
    project_type = base_params.get("project_type", "建筑")

    dependent_list = _get_dependent_fee_list(target_fee)

    for dep_type in dependent_list:
        if dep_type == "监理费":
            prof = _extract_jianli_professional_coef(query)
            comp = _extract_jianli_complexity_coef(query)
            elev = _extract_jianli_elevation_coef(query)

            # 用提取的系数（或默认值）计算预览
            if jianan > 0 or shebei > 0:
                preview = calc_jianli(
                    jianan=jianan, shebei=shebei,
                    professional_coef=prof, complexity_coef=comp,
                    elevation_coef=elev,
                )
            else:
                preview = calc_jianli(
                    amount_wan=amount_wan,
                    professional_coef=prof, complexity_coef=comp,
                    elevation_coef=elev,
                )

            coef_meta = _build_coef_metadata("监理费", preview, query)

            deps.append({
                "fee_type": "监理费",
                "fee_label": "施工监理服务费",
                "config_type": "coef",
                "coef_metadata": coef_meta,
                "base_params": {
                    "jianan": jianan,
                    "shebei": shebei,
                    "amount_wan": amount_wan,
                },
                "default_result": preview,
                "preview_fee": preview["结果(万元)"],
            })

        elif dep_type == "工程设计费":
            prof = _extract_sheji_professional_coef(query)
            comp = _extract_sheji_complexity_coef(query)
            addi_matches = re.findall(r"附加.*?系数.*?(\d+\.?\d*)", query)
            addi_list = [float(m) for m in addi_matches] if addi_matches else None

            preview = calc_sheji(
                amount_wan,
                professional_coef=prof,
                complexity_coef=comp,
                additional_coefs=addi_list,
            )

            coef_meta = _build_coef_metadata("工程设计费", preview, query)

            deps.append({
                "fee_type": "工程设计费",
                "fee_label": "工程设计费",
                "config_type": "coef",
                "coef_metadata": coef_meta,
                "base_params": {
                    "amount_wan": amount_wan,
                },
                "default_result": preview,
                "preview_fee": preview["结果(万元)"],
            })

        elif dep_type == "勘察费":
            pt = _detect_project_type(query) if project_type == "建筑" else project_type
            preview = calc_kancha_rough(jianan, shebei, pt)

            rates_map = {"建筑": (0.3, 0.5), "通用": (0.8, 1.1)}
            lo, hi = rates_map.get(pt, (0.8, 1.1))
            rate_options: list[dict] = []
            r = lo
            total_for_rate = jianan + shebei
            while r <= hi + 0.001:
                fee_at_rate = round(total_for_rate * r / 100.0, 4)
                rate_options.append({
                    "rate": round(r, 1),
                    "label": f"{r:.1f}%",
                    "fee": fee_at_rate,
                })
                r = round(r + 0.1, 1)

            # 默认取中值
            mid_idx = len(rate_options) // 2

            deps.append({
                "fee_type": "勘察费",
                "fee_label": "工程勘察费（粗略估算）",
                "config_type": "rate",
                "rate_options": rate_options,
                "project_type": pt,
                "base_params": {
                    "jianan": jianan,
                    "shebei": shebei,
                    "project_type": pt,
                },
                "default_result": preview,
                "preview_fee": preview["结果中值(万元)"],
                "default_rate_index": mid_idx,
            })

    return deps


def resolve_dependent_calc(
    target_fee: str,
    configs: dict,
    base_params: dict,
) -> dict:
    """用用户选择的参数计算依赖费种，再汇总计算目标费种。

    configs 结构：
        {"监理费": {"professional_coef": 1.0, ...},
         "工程设计费": {"professional_coef": 1.0, ...},
         "勘察费": {"rate": 0.8, "project_type": "建筑"}}

    返回：目标费种的完整计算结果 dict，附加 _dependent_details 和 _dependent_configs。
    """
    jianan = base_params.get("jianan", 0) or 0
    shebei = base_params.get("shebei", 0) or 0
    amount_wan = base_params.get("amount_wan", jianan + shebei)
    query = base_params.get("query", "")
    project_type = base_params.get("project_type", "建筑")

    # Step 1: 计算各依赖费种
    dep_results: dict = {}
    dep_fee_values: dict = {}  # fee_type → 万元

    if "监理费" in configs:
        cfg = configs["监理费"]
        prof = cfg.get("professional_coef", 1.0)
        comp = cfg.get("complexity_coef", 1.0)
        elev = cfg.get("elevation_coef", 1.0)
        if jianan > 0 or shebei > 0:
            r = calc_jianli(jianan=jianan, shebei=shebei,
                            professional_coef=prof, complexity_coef=comp,
                            elevation_coef=elev)
        else:
            r = calc_jianli(amount_wan=amount_wan,
                            professional_coef=prof, complexity_coef=comp,
                            elevation_coef=elev)
        dep_results["监理费"] = r
        dep_fee_values["监理费"] = r["结果(万元)"]

    if "工程设计费" in configs:
        cfg = configs["工程设计费"]
        prof = cfg.get("professional_coef", 1.0)
        comp = cfg.get("complexity_coef", 1.0)
        addi = cfg.get("additional_coef", 1.0)
        addi_list = [addi] if abs(addi - 1.0) > 0.005 else None
        r = calc_sheji(amount_wan, prof, comp, additional_coefs=addi_list)
        dep_results["工程设计费"] = r
        dep_fee_values["工程设计费"] = r["结果(万元)"]

    if "勘察费" in configs:
        cfg = configs["勘察费"]
        rate = cfg.get("rate")
        pt = cfg.get("project_type", project_type)
        if rate is not None:
            total = jianan + shebei
            fee = round(total * rate / 100.0, 4)
            rates_map = {"建筑": (0.3, 0.5), "通用": (0.8, 1.1)}
            lo, hi = rates_map.get(pt, (0.8, 1.1))
            r = {
                "费种": "工程勘察费（粗略估算）",
                "依据": (
                    "粗略估算依据《市政工程设计概算编制办法》（中国计划出版社）；"
                    "精确计算依据《工程勘察设计收费管理规定》（计价格[2002]10号）工程勘察收费标准"
                ),
                "计算公式": f"第一部分工程费 × {rate}%（{pt}项目，用户选择）",
                "结果(万元)": fee,
                "结果中值(万元)": fee,
                "说明": f"{pt}项目，费率 {rate}%，费用 {fee:.2f} 万元",
            }
        else:
            r = calc_kancha_rough(jianan, shebei, pt)
            fee = r["结果中值(万元)"]
            if fee is None:
                fee = r.get("结果(万元)", 0) or 0
        dep_results["勘察费"] = r
        dep_fee_values["勘察费"] = r.get("结果(万元)", r.get("结果中值(万元)", 0))

    # Step 2: 计算目标费种
    if target_fee == "招标代理费":
        result = calc_zhaobiao_daili_all(
            jianan=jianan,
            shebei=shebei,
            project_type=project_type,
            query=query,
            dependent_configs=configs,
        )
        result["_dependent_details"] = dep_results
        result["_dependent_configs"] = configs
        result["is_zhaobiao_multi"] = True
        return result

    elif target_fee == "施工图审查费":
        sheji_fee_only = dep_fee_values.get("工程设计费", 0)
        kancha_fee_mid = dep_fee_values.get("勘察费", 0)
        sheji_fee = round(sheji_fee_only + kancha_fee_mid, 4)

        ptype = base_params.get("project_type_shencha", "公建")
        size = base_params.get("size", "中型")
        amount = base_params.get("amount", amount_wan)

        # 勘察费费率描述
        kc_cfg = configs.get("勘察费", {})
        kc_rate = kc_cfg.get("rate")
        if kc_rate is not None:
            kancha_rate_desc = f"{kc_rate}%（用户选择）"
        else:
            kancha_rate_desc = "区间中值"

        result = calc_shigong_shencha(
            amount, ptype, size,
            sheji_fee=sheji_fee,
            sheji_fee_only=sheji_fee_only,
            kancha_fee_mid=kancha_fee_mid,
            kancha_rate_desc=kancha_rate_desc,
            query=base_params.get("query", ""),
        )
        result["_dependent_details"] = dep_results
        result["_dependent_configs"] = configs
        return result

    return {}


def detect_and_calculate(query: str, *, fee_type: str | None = None) -> dict | None:
    """
    检测查询是否涉及二类费，如果是则直接计算。

    参数：
        fee_type: 可选，指定费种（跳过检测步骤）。用于 detect_and_calculate_all 内部调用。

    返回：计算结果 dict（包含 'fee_type' 和计算明细），用于注入 LLM 上下文。
    返回 None 表示不是二类费问题。
    """
    if fee_type is None:
        # 先检查多费种模式（在单费种检测之前）
        multi_mode = _detect_multi_fee_mode(query)
        if multi_mode == "cascade":
            return calc_cascade(query)
        elif multi_mode == "iteration":
            return calc_iteration(query)
        elif multi_mode == "comparison":
            return calc_comparison(query)
        fee_type = _detect_fee_type(query)
    if not fee_type:
        # 未命中明确的费种关键词，但可能隐含在上下文中
        # 注意：不能因为看到建安费+设备费就直接判为监理费——需排除其他费种提示词
        jianan_test, shebei_test = _extract_jianli_components(query)
        has_jianli_hint = bool(re.search(r"监理", query))
        has_kancha_hint = bool(re.search(r"勘察(?!设计)", query))
        has_sheji_hint = bool(re.search(r"设计费", query))
        if jianan_test is not None and shebei_test is not None:
            if has_kancha_hint and not has_jianli_hint:
                fee_type = "勘察费"
            elif has_sheji_hint and not has_jianli_hint:
                fee_type = "工程设计费"
            else:
                # 同时有建安费和设备费 → 大概率是监理费（40% 规则）
                fee_type = "监理费"
        elif has_jianli_hint and _extract_amount(query) is not None:
            # 提到"监理"且带了金额 → 监理费（用户可能没写"费"字）
            fee_type = "监理费"
        elif has_kancha_hint:
            fee_type = "勘察费"
        else:
            return None

    amount = _extract_amount(query)
    if amount is None:
        # 没有金额 — 返回该费种的费率表和计费规则参考
        ref = _get_fee_reference(fee_type)
        ref["fee_type"] = fee_type
        ref["has_amount"] = False
        # 自动检测系数（无金额模式下也匹配，避免 LLM 自行猜测）
        if fee_type == "工程设计费":
            prof = _extract_sheji_professional_coef(query)
            comp = _extract_sheji_complexity_coef(query)
            ref["auto_detected_coefs"] = {
                "专业调整系数": prof,
                "复杂程度系数": comp,
            }
        elif fee_type == "监理费":
            prof = _extract_jianli_professional_coef(query)
            comp = _extract_jianli_complexity_coef(query)
            elev = _extract_jianli_elevation_coef(query)
            ref["auto_detected_coefs"] = {
                "专业调整系数": prof,
                "复杂程度系数": comp,
                "高程调整系数": elev,
            }
        return ref

    result: dict[str, Any]

    if fee_type == "建设管理费":
        result = calc_jianshe_guanli(amount)
    elif fee_type == "招标代理费":
        jianan_zb, shebei_zb = _extract_jianli_components(query)
        if jianan_zb is not None:
            # 有建安费 → 需要依赖费种交互式配置
            project_type = _detect_project_type(query)
            _shebei = shebei_zb or 0
            bp = {
                "jianan": jianan_zb,
                "shebei": _shebei,
                "amount_wan": jianan_zb + _shebei,
                "project_type": project_type,
                "query": query,
            }
            result = {
                "fee_type": "招标代理费",
                "费种": "招标代理服务费",
                "has_amount": True,
                "needs_dependent_config": True,
                "target_fee": "招标代理费",
                "target_fee_name": "招标代理服务费",
                "base_params": bp,
                "dependent_fees": _build_dependent_config_meta("招标代理费", bp, query),
            }
        elif amount is not None:
            # 仅有金额未区分建安/设备 → 按旧逻辑单类计算
            if re.search(r"货物", query):
                svc_type = "货物招标"
            elif re.search(r"服务", query):
                svc_type = "服务招标"
            else:
                svc_type = "工程招标"
            result = calc_zhaobiao_daili(amount, svc_type)
        else:
            return None
        # 标记支持多选面板（仅当不是 needs_dependent_config 时由本分支设置）
        if not result.get("needs_dependent_config"):
            result["is_zhaobiao_multi"] = True
    elif fee_type == "交易服务费":
        jianan, shebei = _extract_jianli_components(query)
        if jianan is not None:
            # 4 类分项计算：需先算出监理费和设计费作为基数
            prof = _extract_jianli_professional_coef(query)
            comp = _extract_jianli_complexity_coef(query)
            elev = _extract_jianli_elevation_coef(query)
            if shebei is not None:
                jianli_r = calc_jianli(jianan=jianan, shebei=shebei,
                                       professional_coef=prof, complexity_coef=comp, elevation_coef=elev)
            else:
                jianli_r = calc_jianli(amount_wan=jianan,
                                       professional_coef=prof, complexity_coef=comp, elevation_coef=elev)
            jianli_fee = jianli_r["结果(万元)"]

            total_for_sheji = (jianan + shebei) if shebei else jianan
            sheji_prof = _extract_sheji_professional_coef(query)
            sheji_r = calc_sheji(total_for_sheji, professional_coef=sheji_prof)
            sheji_fee = sheji_r["结果(万元)"]

            result = calc_jiaoyi_fuwu(
                jianan=jianan, shebei=shebei,
                jianli_fee=jianli_fee, sheji_fee=sheji_fee,
            )
        else:
            # 回退：单一中标额模式
            result = calc_jiaoyi_fuwu(amount_wan=amount)
    elif fee_type == "监理费":
        jianan, shebei = _extract_jianli_components(query)
        prof = _extract_jianli_professional_coef(query)
        comp = _extract_jianli_complexity_coef(query)
        elev = _extract_jianli_elevation_coef(query)
        if jianan is not None and shebei is not None:
            result = calc_jianli(
                jianan=jianan, shebei=shebei,
                professional_coef=prof, complexity_coef=comp, elevation_coef=elev,
            )
        else:
            result = calc_jianli(
                amount_wan=amount,
                professional_coef=prof, complexity_coef=comp, elevation_coef=elev,
            )
    elif fee_type == "工程设计费":
        prof = _extract_sheji_professional_coef(query)
        comp = _extract_sheji_complexity_coef(query)
        # 附加调整系数 — 可能多个（1.0.9.3: 不能连乘，需合并）
        addi_matches = re.findall(r"附加.*?系数.*?(\d+\.?\d*)", query)
        addi_list = [float(m) for m in addi_matches] if addi_matches else None
        # 改扩建系数
        gaikuojian = _extract_coef(query, r"改扩建.*?系数.*?(\d+\.?\d*)", 0)
        if gaikuojian > 0:
            if addi_list is None:
                addi_list = []
            addi_list.append(gaikuojian)
        # 常见附加项
        zongti = bool(re.search(r"总体设计", query))
        zhuti = bool(re.search(r"主体.*?协调|主体设计协调", query))
        yusuan = bool(re.search(r"施工图预算|预算编制", query))
        jgt = bool(re.search(r"竣工图", query))
        # 计费额：支持分项写法（建安+设备）
        jianan_s, shebei_s = _extract_jianli_components(query)
        sheji_jifei = amount
        if jianan_s is not None:
            sheji_jifei = jianan_s + (shebei_s or 0)
        result = calc_sheji(
            sheji_jifei, prof, comp,
            additional_coefs=addi_list,
            zongti_sheji=zongti, zhuti_xietiao=zhuti,
            shigongtu_yusuan=yusuan, jungongtu=jgt,
        )
    elif fee_type == "可行性研究费":
        # 可行性研究费金额单位是亿元
        amount_yi = _extract_amount_yi(query)
        if amount_yi is None:
            amount_yi = amount / 10000.0  # 万→亿
        # 检测用户指定了哪种服务类型
        # 三类输出模式：
        #   single: 明确 编制X 或 评估X → 只出单项
        #   pair:   只说 可研报告/项目建议书 没带编制/评估 → 出编制+评估两项
        #   all:    只说"前期工作"没提具体产出物 → 四项全出
        svc = "编制可研报告"
        mode = "all"      # all | pair | single
        pair_type = None  # "可研" | "建议书"（仅 pair 模式）
        if re.search(r"项目建议书.*评估|评估.*项目建议书", query):
            svc = "评估项目建议书"; mode = "single"
        elif re.search(r"可研.*评估|评估.*可研|可行性.*评估|评估.*可行性", query):
            svc = "评估可研报告"; mode = "single"
        elif re.search(r"评估建议书|评估.*建议书", query):
            svc = "评估项目建议书"; mode = "single"
        elif re.search(r"编制.*建议书|建议书.*编制", query):
            svc = "编制项目建议书"; mode = "single"
        elif re.search(r"编制.*可研|可研.*编制|编制.*可行性|可行性.*编制", query):
            svc = "编制可研报告"; mode = "single"
        elif re.search(r"项目建议书|建议书", query):
            svc = "编制项目建议书"; mode = "pair"; pair_type = "建议书"
        elif re.search(r"可研报告|可行性研究|可研", query):
            svc = "编制可研报告"; mode = "pair"; pair_type = "可研"
        elif re.search(r"前期工作", query):
            svc = "编制可研报告"; mode = "all"
        else:
            svc = "编制可研报告"; mode = "all"

        # 自动检测行业调整系数
        ind_name, ind_coef = _detect_keyan_industry(query)
        # 自动检测工程复杂程度系数
        comp_coef = 1.0
        comp_m = re.search(r"复杂程度.*?系数.*?(\d+\.?\d*)", query)
        if comp_m:
            comp_coef = float(comp_m.group(1))
        else:
            if re.search(r"很复杂|非常复杂|特别复杂", query):
                comp_coef = 1.15
            elif re.search(r"较复杂|复杂", query):
                comp_coef = 1.0

        total_coef = round(ind_coef * comp_coef, 4)

        if mode == "single":
            # 明确指定了某一项 → 只计算这一项
            result = calc_keyan(amount_yi, svc,
                                industry_coef=ind_coef, industry_name=ind_name,
                                complexity_coef=comp_coef)
        elif mode == "pair":
            # 只说了"可研报告"或"项目建议书"→ 出编制+评估两项
            if pair_type == "可研":
                pair_svc = ["编制可研报告", "评估可研报告"]
            else:
                pair_svc = ["编制项目建议书", "评估项目建议书"]
            pair_results = {}
            for svc_name in pair_svc:
                r = calc_keyan(amount_yi, svc_name,
                               industry_coef=ind_coef, industry_name=ind_name,
                               complexity_coef=comp_coef)
                pair_results[svc_name] = r

            result = pair_results[svc]
            # 修正步骤2：反映实际提问范围
            pair_label = "编制/评估可研报告" if pair_type == "可研" else "编制/评估项目建议书"
            result["计算步骤"][1] = {"步骤": "确定服务类型", "公式": "", "结果": pair_label}

            lines = []
            for svc_name in pair_svc:
                r = pair_results[svc_name]
                fee = r["结果(万元)"]
                base = r["基准价(万元)"]
                lines.append(f"- **{svc_name}**：基准价 {base:.2f} 万 × 总系数 {total_coef} = **{fee:.2f} 万元**")

            result["全部服务类型结果"] = {
                svc_name: {
                    "结果(万元)": pair_results[svc_name]["结果(万元)"],
                    "基准价(万元)": pair_results[svc_name]["基准价(万元)"],
                }
                for svc_name in pair_svc
            }
            result["说明"] = (
                f"估算投资额 {amount_yi:.4f} 亿元（{amount_yi * 10000:.0f} 万元），"
                f"行业「{ind_name}」系数 {ind_coef}，"
                f"复杂程度系数 {comp_coef}，"
                f"总调整系数 **{total_coef}**。\n\n"
                f"{pair_type}相关服务类型结果：\n" +
                "\n".join(lines)
            )
        else:
            # "前期工作"未指定 → 计算全部四种
            all_svc = ["编制项目建议书", "编制可研报告", "评估项目建议书", "评估可研报告"]
            all_results = {}
            for svc_name in all_svc:
                r = calc_keyan(amount_yi, svc_name,
                               industry_coef=ind_coef, industry_name=ind_name,
                               complexity_coef=comp_coef)
                all_results[svc_name] = r

            result = all_results[svc]
            # 修正步骤2：反映实际提问范围
            result["计算步骤"][1] = {"步骤": "确定服务类型", "公式": "", "结果": "编制/评估项目建议书、编制/评估可研报告"}

            lines = []
            for svc_name in all_svc:
                r = all_results[svc_name]
                fee = r["结果(万元)"]
                base = r["基准价(万元)"]
                lines.append(f"- **{svc_name}**：基准价 {base:.2f} 万 × 总系数 {total_coef} = **{fee:.2f} 万元**")

            result["全部服务类型结果"] = {
                svc_name: {
                    "结果(万元)": all_results[svc_name]["结果(万元)"],
                    "基准价(万元)": all_results[svc_name]["基准价(万元)"],
                }
                for svc_name in all_svc
            }
            result["说明"] = (
                f"估算投资额 {amount_yi:.4f} 亿元（{amount_yi * 10000:.0f} 万元），"
                f"行业「{ind_name}」系数 {ind_coef}，"
                f"复杂程度系数 {comp_coef}，"
                f"总调整系数 **{total_coef}**。\n\n"
                f"四种服务类型全部结果：\n" +
                "\n".join(lines)
            )
        # 标记需要交互式服务类型选择（前端渲染 pending_keyan 面板）
        result["needs_keyan_select"] = True
        result["amount_yi"] = amount_yi
        result["industry_coef"] = ind_coef
        result["industry_name"] = ind_name
        result["complexity_coef"] = comp_coef
    elif fee_type == "施工图审查费":
        # 津价管[2011]46号 + 建市[2007]86号
        if re.search(r"住宅", query):
            ptype = "住宅"
        elif re.search(r"工业", query):
            ptype = "工业"
        elif re.search(r"市政|道路|桥梁|隧道|给水|排水|燃气|热力|轨道交通|风景园林"
                       r"|环境卫生|污水处理|垃圾处理|供热|环卫|填埋|焚烧"
                       r"|净水厂|处理厂|泵站|管网|BRT|快速公交|公交|公共交通"
                       r"|供热面积|热源厂|热网|气源厂|垃圾发电", query):
            ptype = "市政"
        else:
            ptype = "公建"
        # 建市[2007]86号 自动判定项目规模
        size = _detect_project_size_86(query, ptype)
        size_desc = {"大型": "大型", "中型": "中型", "小型": "小型"}.get(size, size)

        if ptype != "住宅":
            # 非住宅类：需要依赖费种交互式配置（设计费 + 勘察费）
            jianan_ss, shebei_ss = _extract_jianli_components(query)
            ss_jifei = amount
            if jianan_ss is not None:
                ss_jifei = jianan_ss + (shebei_ss or 0)
            kc_ptype = _detect_project_type(query)
            bp = {
                "jianan": jianan_ss or 0,
                "shebei": shebei_ss or 0,
                "amount_wan": ss_jifei,
                "amount": amount,
                "project_type": kc_ptype,
                "project_type_shencha": ptype,
                "size": size,
                "query": query,
            }
            result = {
                "fee_type": "施工图审查费",
                "费种": f"施工图审查费（{ptype}{size_desc}）",
                "has_amount": True,
                "needs_dependent_config": True,
                "target_fee": "施工图审查费",
                "target_fee_name": "施工图审查费",
                "base_params": bp,
                "dependent_fees": _build_dependent_config_meta("施工图审查费", bp, query),
            }
        else:
            # 住宅类：提取建筑面积
            m_m2 = re.search(r"(?:建筑面积|面积)\s*[:：]?\s*(\d+\.?\d*)\s*万?\s*(?:m2|㎡|平米|平方米)?", query)
            if m_m2:
                val = float(m_m2.group(1))
                if re.search(r"万\s*(?:m2|㎡|平米|平方米)?", query):
                    amount = val * 10000
                else:
                    amount = val
            result = calc_shigong_shencha(amount, ptype, size, query=query)
    elif fee_type == "水土保持费":
        amount_yi = _extract_amount_yi(query)
        if amount_yi is None:
            amount_yi = amount / 10000.0
        if re.search(r"监测", query):
            svc = "施工期监测"
        elif re.search(r"验收|评估", query):
            svc = "验收评估"
        elif re.search(r"技术咨询|咨询", query):
            svc = "技术咨询"
        else:
            svc = "方案编制"
        result = calc_shuibao(amount_yi, svc)
    elif fee_type == "勘察费":
        # 工程勘察费 — 精确计算需按计价格[2002]10号实物工作量定额
        # 粗略估算按《市政工程设计概算编制办法》百分比法
        jianan_kc, shebei_kc = _extract_jianli_components(query)
        if jianan_kc is not None:
            project_type = _detect_project_type(query)
            result = calc_kancha_rough(jianan_kc, shebei_kc or 0, project_type)
        elif amount is not None:
            project_type = _detect_project_type(query)
            result = calc_kancha_rough(amount, 0, project_type)
        else:
            # 无金额 → 返回参考信息
            ref = _get_fee_reference("勘察费")
            ref["fee_type"] = fee_type
            ref["has_amount"] = False
            return ref
    elif fee_type == "劳动安全卫生评审费":
        jianan_l, shebei_l = _extract_jianli_components(query)
        total_l = (jianan_l or 0) + (shebei_l or 0)
        if total_l > 0:
            result = calc_laodong_anquan(total_l)
        elif amount is not None:
            result = calc_laodong_anquan(amount)
        else:
            ref = _get_fee_reference("劳动安全卫生评审费")
            ref["fee_type"] = fee_type
            ref["has_amount"] = False
            return ref
    elif fee_type == "场地准备费及临时设施费":
        jianan_c, shebei_c = _extract_jianli_components(query)
        total_c = (jianan_c or 0) + (shebei_c or 0)
        if total_c > 0:
            result = calc_changdi_zhunbei(total_c)
        elif amount is not None:
            result = calc_changdi_zhunbei(amount)
        else:
            ref = _get_fee_reference("场地准备费及临时设施费")
            ref["fee_type"] = fee_type
            ref["has_amount"] = False
            return ref
    elif fee_type == "工程保险费":
        jianan_b, shebei_b = _extract_jianli_components(query)
        total_b = (jianan_b or 0) + (shebei_b or 0)
        if total_b > 0:
            result = calc_gongcheng_baoxian(total_b)
        elif amount is not None:
            result = calc_gongcheng_baoxian(amount)
        else:
            ref = _get_fee_reference("工程保险费")
            ref["fee_type"] = fee_type
            ref["has_amount"] = False
            return ref
    elif fee_type == "预备费":
        # 预备费 = （第一部分工程费 + 工程建设其他费）× 费率
        # 尝试从查询中提取费率，默认 5%
        yb_rate = 5.0
        yb_rate_match = re.search(r"预备费.*?(\d+\.?\d*)\s*%|预备费率\s*(\d+\.?\d*)", query)
        if yb_rate_match:
            yb_rate = float(yb_rate_match.group(1) or yb_rate_match.group(2))
        jianan_b, shebei_b = _extract_jianli_components(query)
        part1 = (jianan_b or 0) + (shebei_b or 0)
        if part1 == 0 and amount is not None:
            part1 = amount
        if part1 > 0:
            # 需要先算二类费才能算预备费 — 走联算流程
            cascade_r = calc_cascade(query)
            if cascade_r:
                erlei_total = cascade_r["结果汇总"]["二类费合计(万元)"]
                # 从 cascade 结果中减去额外费用，得到纯二类费
                extra_sum = sum(e["金额(万元)"] for e in cascade_r.get("额外费用", []))
                erlei_pure = erlei_total - extra_sum
                yb_r = calc_yubei(part1, erlei_pure, yb_rate)
                yb_r["fee_type"] = fee_type
                yb_r["has_amount"] = True
                yb_r["二类费合计(万元)"] = erlei_pure
                return yb_r
            else:
                ref = _get_fee_reference("预备费")
                ref["fee_type"] = fee_type
                ref["has_amount"] = False
                return ref
        else:
            ref = _get_fee_reference("预备费")
            ref["fee_type"] = fee_type
            ref["has_amount"] = False
            return ref
    elif fee_type == "环境影响咨询费":
        # 计价格[2002]125号 — 分档定额线性内插
        # 检测用户是否明确指定了服务类型
        explicit_svc = None
        if re.search(r"报告表|报告书.*表", query):
            explicit_svc = "编制报告表"
        elif re.search(r"评估报告书|评估.*报告书", query):
            explicit_svc = "评估报告书"
        elif re.search(r"评估报告表|评估.*报告表", query):
            explicit_svc = "评估报告表"
        elif re.search(r"编制.*报告书|报告书", query):
            explicit_svc = "编制报告书"
        elif re.search(r"大纲", query):
            explicit_svc = "评估报告书"

        # 行业调整系数 + 环境敏感程度系数
        ind_name, ind_coef = _detect_huanping_industry(query)
        if re.search(r"敏感|一级|二类|重要", query):
            sens_coef = 1.2
        elif re.search(r"一般|不敏感|三类", query):
            sens_coef = 0.8
        else:
            sens_coef = 1.0

        if explicit_svc is None:
            # 用户未指定具体服务类型 → 显示多选面板
            # 环评费基数为项目总投资，优先从查询中提取
            estimated_investment = _extract_total_investment(query)
            has_explicit_investment = estimated_investment is not None
            if not has_explicit_investment:
                # 未提供总投资，用查询中的金额作为临时估算
                estimated_investment = amount
            result = {
                "fee_type": "环境影响咨询费",
                "费种": "环境影响咨询费",
                "has_amount": True,
                "needs_huanping_select": True,
                "amount_wan": amount,  # 查询中提取的原始金额（可能是建安费）
                "estimated_investment": estimated_investment,
                "has_explicit_investment": has_explicit_investment,
                "industry_coef": ind_coef,
                "industry_name": ind_name,
                "sensitivity_coef": sens_coef,
                "依据": "《关于规范环境影响咨询收费有关问题的通知》（计价格[2002]125号）",
            }
        else:
            # 用户指定了具体服务类型 → 现有行为（系数可调）
            svc = explicit_svc

            # 计算全部四种服务类型
            all_svc = ["编制报告书", "编制报告表", "评估报告书", "评估报告表"]
            all_results = {}
            total_coef = round(ind_coef * sens_coef, 4)
            for svc_name in all_svc:
                r = calc_huanping(amount, svc_name, industry_coef=ind_coef,
                                  industry_name=ind_name, sensitivity_coef=sens_coef)
                all_results[svc_name] = r

            # 主结果为检测到的服务类型
            result = all_results[svc]

            # 构建四项结果汇总
            lines = []
            for svc_name in all_svc:
                r = all_results[svc_name]
                mid = r["结果中值(万元)"]
                lo_hi = r["结果(万元)"]
                lines.append(f"- **{svc_name}**：{lo_hi} 万元（中值 **{mid} 万元**）")

            # 追加详细说明
            invest_yi = amount / 10000.0
            result["全部服务类型结果"] = {
                svc_name: {"结果(万元)": all_results[svc_name]["结果(万元)"],
                           "结果中值(万元)": all_results[svc_name]["结果中值(万元)"],
                           "基准价(万元)": all_results[svc_name]["基准价(万元)"]}
                for svc_name in all_svc
            }
            result["说明"] = (
                f"估算投资额 {invest_yi:.2f} 亿元（{amount:.0f}万元），"
                f"行业「{ind_name}」系数 {ind_coef}，"
                f"环境{'敏感' if sens_coef==1.2 else '一般' if sens_coef==0.8 else '未指定'}系数 {sens_coef}，"
                f"总调整系数 **{total_coef}**。\n\n"
                f"四种服务类型全部结果（协商浮动 ±20% 后）：\n" +
                "\n".join(lines)
            )
    elif fee_type == "造价咨询费":
        # ── 河北省：冀建市研[2017]2号 ──
        if query and _is_hebei_project(query):
            svc_type = _detect_hebei_cost_consulting_type(query)
            if svc_type is None:
                svc_type = "预算编制"  # 默认

            jianan_zj, shebei_zj = _extract_jianli_components(query)
            # 河北省基数 = 建安费（不含设备费）
            jianan_only = jianan_zj if jianan_zj is not None else amount
            base_amount = jianan_only  # Hebei uses 建安费 only

            total_invest = _extract_total_investment(query)
            if total_invest is None and svc_type in ("投资估算", "经济评价", "概算编制", "概算审核", "竣工决算编制", "工程造价鉴定"):
                cascade_r = calc_cascade(query)
                if cascade_r:
                    total_invest = cascade_r["结果汇总"]["项目总投资(万元)"]

            try:
                result = calc_cost_consulting_hebei(
                    jianan_only, svc_type,
                    total_investment=total_invest,
                )
            except ValueError as e:
                result = {
                    "费种": f"造价咨询费（{svc_type}）",
                    "依据": "《河北省建设工程造价咨询服务收费管理暂行办法》（冀建市研[2017]2号）",
                    "参数": {
                        "服务类型": svc_type,
                        "建安工程造价(万元)": round(jianan_only, 4),
                    },
                    "结果(万元)": None,
                    "计算步骤": [],
                    "说明": str(e),
                    "_error": str(e),
                }
            # 保存原始输入值，供前端多选面板使用
            result["_jianan"] = jianan_zj
            result["_shebei"] = shebei_zj
            result["_total_invest"] = total_invest
            result["_base_amount"] = base_amount
            result["_is_hebei"] = True  # 标记河北项目，前端据此渲染
        else:
            # ── 津价房地[2008]136号（天津市）──
            svc_type = _detect_cost_consulting_type(query)
            if svc_type is None:
                svc_type = "编制施工图预算"  # 默认

            jianan_zj, shebei_zj = _extract_jianli_components(query)
            if jianan_zj is not None:
                base_amount = jianan_zj + (shebei_zj or 0)
            else:
                base_amount = amount

            # 提取总投资（优先从查询文本自动捕捉，其次通过级联计算）
            total_invest = _extract_total_investment(query)
            if total_invest is None and svc_type == "审核概算":
                cascade_r = calc_cascade(query)
                if cascade_r:
                    total_invest = cascade_r["结果汇总"]["项目总投资(万元)"]

            try:
                result = calc_cost_consulting(
                    base_amount, svc_type,
                    total_investment=total_invest,
                    jianan_only=jianan_zj,
                )
            except ValueError as e:
                # 审核概算总投资未知时，返回提示
                result = {
                    "费种": f"造价咨询费（{svc_type}）",
                    "依据": "《天津市建设工程造价咨询服务项目和价格标准》（津价房地[2008]136号）",
                    "参数": {
                        "服务类型": svc_type,
                        "工程费用(万元)": round(base_amount, 4),
                    },
                    "结果(万元)": None,
                    "计算步骤": [],
                    "说明": str(e),
                    "_error": str(e),
                }
            # 保存原始输入值，供前端多选面板使用
            result["_jianan"] = jianan_zj
            result["_shebei"] = shebei_zj
            result["_total_invest"] = total_invest
            result["_base_amount"] = base_amount
    else:
        return None

    result["fee_type"] = fee_type
    result["has_amount"] = True
    # 打折系数（从查询中自动提取，默认 1.0 不打折；前端可覆盖）
    result["_discount_coef"] = _extract_discount_coefficient(query)
    # 标记支持交互式费率选择的费种（前端会渲染费率下拉菜单）
    if fee_type in ("勘察费", "劳动安全卫生评审费", "场地准备费及临时设施费", "工程保险费"):
        result["is_rate_selectable"] = True
    # 标记支持交互式系数选择的费种（前端会渲染系数下拉菜单）
    if fee_type in ("监理费", "工程设计费", "环境影响咨询费"):
        if not result.get("needs_huanping_select"):
            result["is_coef_selectable"] = True
            result["coef_metadata"] = _build_coef_metadata(fee_type, result, query)
    return result


def _extract_coef(query: str, pattern: str, default: float) -> float:
    """从查询中提取调整系数"""
    m = re.search(pattern, query)
    return float(m.group(1)) if m else default


def _describe_jianli_professional_coef(coef: float) -> str:
    """将监理费专业调整系数映射回工程类型描述"""
    mapping = {0.8: "园林绿化工程", 0.9: "矿山采选/农林工程", 1.0: "建筑/市政/公路等一般工程",
               1.1: "水运/地铁/桥梁/隧道/索道工程", 1.2: "核能/水电/水库工程"}
    return mapping.get(coef, f"系数{coef}")


def _describe_complexity_coef(coef: float) -> str:
    """复杂程度系数 → 描述"""
    mapping = {0.85: "I级/一般", 1.0: "II级/较复杂", 1.15: "III级/复杂"}
    return mapping.get(coef, f"系数{coef}")


def _describe_elevation_coef(coef: float) -> str:
    """高程调整系数 → 描述"""
    mapping = {1.0: "海拔≤2000m", 1.1: "海拔2001-3000m", 1.2: "海拔3001-3500m", 1.3: "海拔3501-4000m"}
    return mapping.get(coef, f"系数{coef}")


def _extract_jianli_professional_coef(query: str) -> float:
    """自动匹配监理费专业调整系数（发改价格[2007]670号 附表三）。
    优先显式数字（如"专业系数0.8"），其次关键词匹配，默认 1.0。"""
    # 1. 显式数字
    m = re.search(r'专业.*?系数.*?(\d+\.?\d*)', query)
    if m:
        return float(m.group(1))
    # 2. 关键词匹配
    for pattern, coef in JIANLI_PROFESSIONAL_COEFS:
        if re.search(pattern, query):
            return coef
    # 3. 默认
    return 1.0


def _extract_jianli_complexity_coef(query: str) -> float:
    """自动匹配监理费工程复杂程度调整系数（发改价格[2007]670号 1.0.9条）。
    优先显式数字，其次关键词匹配，默认 1.0（II级/较复杂）。
    匹配顺序 III → II → I，避免罗马数字子串误匹配（如 I级 匹配 III级）。"""
    # 1. 显式数字
    m = re.search(r'复杂.*?系数.*?(\d+\.?\d*)', query)
    if m:
        return float(m.group(1))
    # 2. 关键词匹配（III → II → I 顺序，不可颠倒）
    if re.search(r'III级|Ⅲ级', query):
        return 1.15
    if re.search(r'II级|Ⅱ级|较复杂', query):
        return 1.0
    if re.search(r'(?<!较)复杂', query):
        return 1.15
    if re.search(r'I级|Ⅰ级|简单', query):
        return 0.85
    return 1.0


def _extract_sheji_complexity_coef(query: str) -> float:
    """自动匹配工程设计费复杂程度调整系数（计价格[2002]10号 1.0.9.2）。
    优先显式数字，其次关键词匹配，默认 1.0（II级/较复杂）。
    匹配顺序 III → II → I，避免罗马数字子串误匹配（如 I级 匹配 III级）。"""
    # 1. 显式数字：复杂程度系数 1.15 / 工程复杂程度系数 0.85 等
    m = re.search(r'复杂.*?(?:系数|程度).*?(\d+\.?\d*)', query)
    if m:
        return float(m.group(1))
    # 2. 关键词匹配（III → II → I 顺序，不可颠倒；II级必须在裸"复杂"之前）
    if re.search(r'III级|Ⅲ级', query):
        return 1.15
    if re.search(r'II级|Ⅱ级|较复杂', query):
        return 1.0
    if re.search(r'(?<!较)复杂', query):
        return 1.15
    if re.search(r'I级|Ⅰ级|简单|一般', query):
        return 0.85
    return 1.0


def _extract_jianli_elevation_coef(query: str) -> float:
    """自动匹配监理费高程调整系数（发改价格[2007]670号 1.0.9条）。
    优先显式数字，其次海拔高度推断，默认 1.0。"""
    # 1. 显式数字
    m = re.search(r'高程.*?系数.*?(\d+\.?\d*)', query)
    if m:
        return float(m.group(1))
    # 2. 海拔高度推断
    m = re.search(r'海拔\s*(\d+)\s*m?', query)
    if m:
        alt = float(m.group(1))
        if alt <= 2000:
            return 1.0
        elif alt <= 3000:
            return 1.1
        elif alt <= 3500:
            return 1.2
        elif alt <= 4000:
            return 1.3
        else:
            return 1.3  # 超过 4000m 双方协商，暂取 1.3
    return 1.0


def _extract_sheji_professional_coef(query: str) -> float:
    """自动匹配工程设计费专业调整系数（计价格[2002]10号 附表二）。
    优先显式数字，其次关键词匹配（长/优先），默认 1.0。"""
    m = re.search(r'专业.*?系数.*?(\d+\.?\d*)', query)
    if m:
        return float(m.group(1))
    # 按优先级逐条匹配（长词优先，具体子类优先于大类）
    for pattern, coef in SHEJI_PROFESSIONAL_COEFS:
        if re.search(pattern, query):
            return coef
    return 1.0


def _extract_jianli_components(query: str) -> tuple[float | None, float | None]:
    """
    从查询中提取监理费的建安费和设备费（用于 40% 规则和交易服务费）。

    建安费支持的表述：
    - "建安费6000万"、"建筑安装工程费 8000万"、"建安 6000"
    - "建筑工程费117万，安装工程费14万"（分开写，自动求和）
    - "建筑工程费500万"（单独出现也识别）

    设备费支持的表述：
    - "设备购置费3000万"、"设备费 1000万"、"设备+联合试运转 10000万"
    - "设备购置费和联合试运转费 10000 万"
    """
    jianan = None
    shebei = None

    # 建筑安装工程费 / 建安工程费 / 建安费 / 建安（合并表述，长优先）
    m = re.search(r'(?:建筑安装工程费|建安工程费|建安费|建安)\s*[:：]?\s*(\d+\.?\d*)\s*万?', query)
    if m:
        jianan = float(m.group(1))
    else:
        # 分开表述：建筑工程费 + 安装工程费（自动求和）
        jianzhu = None
        anzhuang = None
        m1 = re.search(r'建筑工程费\s*[:：]?\s*(\d+\.?\d*)\s*万?', query)
        if m1:
            jianzhu = float(m1.group(1))
        m2 = re.search(r'安装工程费\s*[:：]?\s*(\d+\.?\d*)\s*万?', query)
        if m2:
            anzhuang = float(m2.group(1))
        if jianzhu is not None or anzhuang is not None:
            jianan = (jianzhu or 0) + (anzhuang or 0)

    # 设备购置费 + 联合试运转费 / 设备费
    # 匹配顺序：长词组优先，避免"设备"过早截断"设备购置费"
    m = re.search(
        r'(?:设备购置费[和与及]联合试运转费'
        r'|设备购置费和联合试运转费'
        r'|设备\+联合试运转'
        r'|设备购置费'
        r'|联合试运转费'
        r'|设备费'
        r'|设备)'
        r'\s*[:：]?\s*(\d+\.?\d*)\s*万?',
        query,
    )
    if m:
        shebei = float(m.group(1))

    return jianan, shebei


def _extract_amount_yi(query: str) -> float | None:
    """从查询中提取以'亿元'为单位的金额。"""
    m = re.search(r'(\d+\.?\d*)\s*亿', query)
    return float(m.group(1)) if m else None


def _extract_total_investment(query: str) -> float | None:
    """从查询中提取工程总投资（万元）。

    匹配："总投资为1429.87万元"、"总投资1429.87万"、"工程总投资约1500万元" 等。
    """
    # 带明确标签的投资
    m = re.search(r'(?:工程)?总投资\D*?(\d+\.?\d*)\s*万', query)
    if m:
        return float(m.group(1))

    # "总概算" / "总造价" 等
    m = re.search(r'(?:总概算|总造价|概算总投资)\D*?(\d+\.?\d*)\s*万', query)
    if m:
        return float(m.group(1))

    # 亿元单位
    m = re.search(r'(?:工程)?总投资\D*?(\d+\.?\d*)\s*亿', query)
    if m:
        return float(m.group(1)) * 10000

    return None


def _format_rate_table(rows: list[tuple]) -> str:
    """将费率表转为 Markdown 表格字符串"""
    if not rows:
        return ""
    lines = []
    # 表头
    header = "| " + " | ".join(str(c) for c in rows[0]) + " |"
    lines.append(header)
    # 分隔线
    lines.append("|" + "|".join("---" for _ in rows[0]) + "|")
    # 数据行
    for row in rows[1:]:
        lines.append("| " + " | ".join(str(c) for c in row) + " |")
    return "\n".join(lines)


def format_for_llm(result: dict) -> str:
    """将计算结果格式化为 LLM 可直接使用的上下文文本"""
    if not result.get("has_amount"):
        # 无金额 → 展示费率表和计费规则
        lines = [
            "## 二类费规则引擎 — 计费依据参考",
            "",
            f"**费种**：{result.get('费种', '')}",
            f"**依据文件**：{result.get('依据', '')}",
            f"**计费方式**：{result.get('计费方式', '')}",
            "",
        ]
        if result.get("参数说明"):
            lines.append(f"**所需参数**：{result['参数说明']}")
            lines.append("")
        rate_table = result.get("费率表", [])
        if rate_table:
            lines.append("**费率表**：")
            lines.append("")
            lines.append(_format_rate_table(rate_table))
            lines.append("")
        if result.get("计算说明"):
            lines.append(f"**计算示例**：{result['计算说明']}")
            lines.append("")
        if result.get("特殊规则"):
            lines.append(f"**特殊规则**：{result['特殊规则']}")
            lines.append("")
        # 工程设计费：调整系数表已由程序直接展示，LLM 不需要也不应该输出任何系数表
        is_sheji = result.get("费种") == "工程设计费" or result.get("fee_type") == "工程设计费"
        if result.get("调整系数") and not is_sheji:
            lines.append(f"**调整系数说明**：{result['调整系数']}")
            lines.append("")
        elif result.get("调整系数") and is_sheji:
            lines.append(
                "**调整系数说明**：工程设计费的专业调整系数表（附表二）已由程序直接展示给用户，"
                "你**不需要**也不应该重新输出任何系数表。只解释如何使用这三个系数即可。"
            )
            lines.append("")
        # 自动检测的系数（引擎根据关键词匹配的结果，最高优先级）
        auto_coefs = result.get("auto_detected_coefs", {})
        if auto_coefs:
            lines.append("**[引擎自动检测] 当前查询匹配到的调整系数**：")
            lines.append("")
            lines.append("| 系数类型 | 系数值 |")
            lines.append("|---|---|")
            for k, v in auto_coefs.items():
                lines.append(f"| {k} | **{v}** |")
            lines.append("")
            lines.append(
                "> **[重要] 上述自动检测的系数为最终答案，"
                "必须逐字引用，不得用训练数据或其他值覆盖！**"
            )
            lines.append("")
        lines.append(
            "> 请提供具体金额参数，程序将根据上述政策文件精确计算。"
        )
        return "\n".join(lines)

    lines = [
        "## 二类费规则引擎计算结果（已由程序精确计算，请直接引用，不要重新计算）",
        "",
        f"**费种**：{result.get('费种', '')}",
        f"**依据文件**：{result.get('依据', '')}",
        f"**计算公式**：{result.get('计算公式', '')}",
        "",
    ]

    # 计费额调整说明（监理费 40% 规则等）
    adjustment = result.get("计费额调整")
    if adjustment and adjustment.get("触发调整"):
        lines.append("### 计费额调整（670号文 1.0.8 条）")
        lines.append(adjustment["说明"])
        lines.append("")

    params = result.get("参数", {})
    if params:
        lines.append("**输入参数**：")
        for k, v in params.items():
            lines.append(f"- {k}：{v}")
        lines.append("")

    # 交易服务费分项明细
    items = result.get("分项明细")
    if items:
        lines.append("**分项明细**：")
        lines.append("")
        lines.append("| 类别 | 基数（万元） | 费用（元） | 档位 |")
        lines.append("|---|---|---|---|")
        for item in items:
            lines.append(
                f"| {item.get('类别', '')} "
                f"| {item.get('基数(万元)', '')} "
                f"| {item.get('费用(元)', '')} "
                f"| {item.get('档位', '')} |"
            )
        lines.append("")

    steps = result.get("计算步骤")
    if steps:
        lines.append("**分档计算明细**：")
        lines.append("")
        lines.append("| 区间（万元） | 金额（万元） | 费率 | 费用（万元） |")
        lines.append("|---|---|---|---|")
        for s in steps:
            lines.append(
                f"| {s.get('区间', '')} "
                f"| {s.get('金额(万元)', '')} "
                f"| {s.get('费率(%)', '')}% "
                f"| {s.get('费用(万元)', '')} |"
            )
        lines.append("")

    # 工程设计费分项
    basic_design = result.get("基本设计收费(万元)")
    if basic_design is not None:
        lines.append(f"**基本设计收费**：{basic_design} 万元")
    other_details = result.get("其他设计收费明细")
    if other_details:
        lines.append("**其他设计收费**：")
        for od in other_details:
            lines.append(f"- {od['项目']}：{od['费用(万元)']} 万元")

    result_val = result.get("结果(万元)") or result.get("结果(元)")
    unit = "元" if "结果(元)" in result else "万元"
    lines.append(f"**计算结果**：**{result_val} {unit}**")

    if "分摊" in result:
        lines.append(f"**费用分摊**：{result['分摊']}")

    lines.append("")
    lines.append(f"**结论**：{result.get('说明', '')}")
    lines.append("")
    # 工程设计费：明确禁止提及高程调整系数
    if result.get("费种") == "工程设计费":
        lines.append(
            "> **[重要] 工程设计费的三个调整系数为: 专业调整系数 + 复杂程度调整系数 + 附加调整系数。"
            "不存在高程调整系数(高程系数仅用于监理费 发改价格[2007]670号)。"
            "请只引用上述三个系数, 不要编造高程系数。**"
        )
    lines.append(
        "> **⚠️ 以上所有数字均由程序依据政策文件精确计算。"
        "请逐字引用上述结论中的数字，不要自行验算或重新计算。"
        "不要输出任何算式（如 ×、÷、+ 号），只陈述结果。**"
    )

    return "\n".join(lines)


# ============================================================
# 迭代计算核心引擎 + 三种模式
# ============================================================

# 不可自动计算的费种（需额外参数）
_SKIP_FEES: dict[str, str] = {
    "招标代理费": "需要中标金额，无法根据建安+设备自动计算",
    "水土保持费": "需要土建投资额，无法根据建安+设备自动计算",
    "造价咨询费": "需要选择具体服务子项（预算编制/结算审核等）",
}

# 费种依赖层级
_TIER_MAP: dict[str, int] = {
    "监理费": 0, "工程设计费": 0, "勘察费": 0,
    "劳动安全卫生评审费": 0, "场地准备费及临时设施费": 0, "工程保险费": 0,
    "交易服务费": 1, "施工图审查费": 1,
    "建设管理费": 2, "可行性研究费": 2, "环境影响咨询费": 2,
    "预备费": 3,
}

# 费种依赖关系（用于 UI 提示，非计算引擎逻辑）
_TIER_DEPS: dict[str, list[str]] = {
    "交易服务费": ["监理费", "工程设计费"],
    "施工图审查费": ["工程设计费", "勘察费"],
}

# 费种显示标签
_FEE_LABELS: dict[str, str] = {
    "监理费": "施工监理服务费",
    "工程设计费": "工程设计费",
    "勘察费": "工程勘察费",
    "劳动安全卫生评审费": "劳动安全卫生评审费",
    "场地准备费及临时设施费": "场地准备费及临时设施费",
    "工程保险费": "工程保险费",
    "交易服务费": "交易服务费（招标代理相关）",
    "施工图审查费": "施工图审查费",
    "建设管理费": "建设管理费（建设单位管理费）",
    "可行性研究费": "建设项目前期工作咨询费",
    "环境影响咨询费": "环境影响咨询费",
    "预备费": "预备费（基本预备费）",
    "招标代理费": "招标代理服务费",
    "造价咨询费": "工程造价咨询服务费",
}


def _calc_all_fees(
    jianan: float,
    shebei: float,
    project_type: str = "通用",
    query: str = "",
    total_investment_override: float | None = None,
    param_overrides: dict | None = None,
    skip_fees: set | None = None,
    coef_overrides: dict | None = None,
) -> dict:
    """
    核心级联引擎：计算所有可自动计算的二类费，按依赖层级 T0→T1→T2。

    skip_fees: 用户取消选中的费种集合。这些费种仍正常计算（以维持依赖链），
               但在最终结果中会被移除，汇总值也会相应重新计算。
    coef_overrides: 用户调整的系数覆盖值，格式：
                   {"监理费": {"professional_coef": 0.8, ...}, ...}
    """
    if param_overrides is None:
        param_overrides = {}

    total_part1 = jianan + shebei
    numerical: dict[str, float] = {}
    raw_results: dict[str, dict] = {}

    # ============ Tier 0：仅需 建安+设备 ============

    # 监理费
    jl_overrides = (coef_overrides or {}).get("监理费", {})
    prof = jl_overrides.get("professional_coef")
    if prof is None:
        prof = _extract_jianli_professional_coef(query)
    comp = jl_overrides.get("complexity_coef")
    if comp is None:
        comp = _extract_jianli_complexity_coef(query)
    elev = jl_overrides.get("elevation_coef")
    if elev is None:
        elev = _extract_jianli_elevation_coef(query)
    jianli_r = calc_jianli(jianan=jianan, shebei=shebei,
                           professional_coef=prof, complexity_coef=comp,
                           elevation_coef=elev)
    raw_results["监理费"] = jianli_r
    numerical["监理费(万元)"] = _extract_numeric_value(jianli_r)

    # 设计费
    sj_overrides = (coef_overrides or {}).get("工程设计费", {})
    sheji_prof = sj_overrides.get("professional_coef")
    if sheji_prof is None:
        sheji_prof = _extract_sheji_professional_coef(query)
    sheji_comp = sj_overrides.get("complexity_coef")
    if sheji_comp is None:
        sheji_comp = _extract_sheji_complexity_coef(query)
    ss_addi = re.findall(r"附加.*?系数.*?(\d+\.?\d*)", query)
    ss_addi_list = [float(m) for m in ss_addi] if ss_addi else None
    # 附加调整系数覆盖
    sj_addi_override = sj_overrides.get("additional_coef")
    if sj_addi_override is not None and sj_addi_override != 1.0:
        ss_addi_list = [sj_addi_override]
    sheji_r = calc_sheji(total_part1, professional_coef=sheji_prof,
                         complexity_coef=sheji_comp, additional_coefs=ss_addi_list)
    raw_results["工程设计费"] = sheji_r
    numerical["工程设计费(万元)"] = _extract_numeric_value(sheji_r)

    # 勘察费
    kancha_rate = param_overrides.get("勘察费费率")
    if kancha_rate is None:
        # 从查询中检测用户指定的勘察费费率
        kc_rate_match = re.search(r"勘察费.*?(\d+\.?\d*)\s*%", query)
        if kc_rate_match:
            kancha_rate = float(kc_rate_match.group(1))
    if kancha_rate is not None:
        fee = round(total_part1 * kancha_rate / 100.0, 4)
        kancha_r = {
            "费种": "工程勘察费（用户指定费率）",
            "依据": "用户指定勘察费费率",
            "计算公式": f"第一部分工程费 × {kancha_rate}%",
            "结果中值(万元)": fee,
        }
        raw_results["勘察费"] = kancha_r
        numerical["勘察费(万元)"] = fee
    else:
        kancha_r = calc_kancha_rough(jianan, shebei, project_type)
        raw_results["勘察费"] = kancha_r
        numerical["勘察费(万元)"] = _extract_numeric_value(kancha_r)

    # 劳动安全卫生评审费
    laoan_rate = param_overrides.get("劳动安全卫生评审费费率")
    if laoan_rate is not None:
        fee = round(total_part1 * laoan_rate / 100.0, 4)
        laoan_r = {
            "费种": "劳动安全卫生评审费（用户指定费率）",
            "依据": "用户指定费率",
            "计算公式": f"第一部分工程费 × {laoan_rate}%",
            "结果中值(万元)": fee,
        }
        raw_results["劳动安全卫生评审费"] = laoan_r
        numerical["劳动安全卫生评审费(万元)"] = fee
    else:
        laoan_r = calc_laodong_anquan(total_part1)
        raw_results["劳动安全卫生评审费"] = laoan_r
        numerical["劳动安全卫生评审费(万元)"] = _extract_numeric_value(laoan_r)

    # 场地准备费
    changdi_rate = param_overrides.get("场地准备费费率")
    if changdi_rate is not None:
        fee = round(total_part1 * changdi_rate / 100.0, 4)
        changdi_r = {
            "费种": "场地准备费及临时设施费（用户指定费率）",
            "依据": "用户指定费率",
            "计算公式": f"第一部分工程费 × {changdi_rate}%",
            "结果中值(万元)": fee,
        }
        raw_results["场地准备费及临时设施费"] = changdi_r
        numerical["场地准备费及临时设施费(万元)"] = fee
    else:
        changdi_r = calc_changdi_zhunbei(total_part1)
        raw_results["场地准备费及临时设施费"] = changdi_r
        numerical["场地准备费及临时设施费(万元)"] = _extract_numeric_value(changdi_r)

    # 工程保险费
    baoxian_rate = param_overrides.get("工程保险费费率")
    if baoxian_rate is not None:
        fee = round(total_part1 * baoxian_rate / 100.0, 4)
        baoxian_r = {
            "费种": "工程保险费（用户指定费率）",
            "依据": "用户指定费率",
            "计算公式": f"第一部分工程费 × {baoxian_rate}%",
            "结果中值(万元)": fee,
        }
        raw_results["工程保险费"] = baoxian_r
        numerical["工程保险费(万元)"] = fee
    else:
        baoxian_r = calc_gongcheng_baoxian(total_part1)
        raw_results["工程保险费"] = baoxian_r
        numerical["工程保险费(万元)"] = _extract_numeric_value(baoxian_r)

    t0_total = sum(numerical.values())

    # ── 招标代理费（依赖 T0：监理费 + 设计费 + 勘察费）──
    # 该费种在 _SKIP_FEES 中，但当选中的依赖费种都有时，可以在联算中自动计算
    if skip_fees is None or "招标代理费" not in skip_fees:
        jl_fee_zb = numerical.get("监理费(万元)", 0)
        sj_fee_zb = numerical.get("工程设计费(万元)", 0)
        kc_fee_zb = numerical.get("勘察费(万元)", 0)
        if jl_fee_zb > 0 and sj_fee_zb > 0 and kc_fee_zb > 0:
            # 使用已经计算的依赖费种值（含用户系数调整），构建 dependent_configs
            _zb_dep_configs = {
                "监理费": {
                    "professional_coef": (coef_overrides or {}).get("监理费", {}).get("professional_coef", 1.0),
                    "complexity_coef": (coef_overrides or {}).get("监理费", {}).get("complexity_coef", 1.0),
                    "elevation_coef": (coef_overrides or {}).get("监理费", {}).get("elevation_coef", 1.0),
                },
                "工程设计费": {
                    "professional_coef": (coef_overrides or {}).get("工程设计费", {}).get("professional_coef", 1.0),
                    "complexity_coef": (coef_overrides or {}).get("工程设计费", {}).get("complexity_coef", 1.0),
                    "additional_coef": (coef_overrides or {}).get("工程设计费", {}).get("additional_coef", 1.0),
                },
                "勘察费": {
                    "rate": param_overrides.get("勘察费费率"),
                    "project_type": project_type,
                },
            }
            try:
                zhaobiao_r = calc_zhaobiao_daili_all(
                    jianan=jianan, shebei=shebei,
                    project_type=project_type, query=query,
                    dependent_configs=_zb_dep_configs,
                )
                raw_results["招标代理费"] = zhaobiao_r
                numerical["招标代理费(万元)"] = zhaobiao_r.get("合计(万元)", 0)
            except Exception:
                pass  # 计算失败则跳过

    # ============ Tier 1：需要 T0 结果 ============

    # 交易服务费（需要 监理费 + 设计费）
    jiaoyi_r = calc_jiaoyi_fuwu(
        jianan=jianan, shebei=shebei,
        jianli_fee=numerical["监理费(万元)"],
        sheji_fee=numerical["工程设计费(万元)"],
    )
    raw_results["交易服务费"] = jiaoyi_r
    numerical["交易服务费(万元)"] = _extract_numeric_value(jiaoyi_r)

    # 施工图审查费（需要 设计费 + 勘察费）
    _shencha_query = query
    if re.search(r"住宅", _shencha_query):
        shencha_ptype = "住宅"
    elif re.search(r"工业", _shencha_query):
        shencha_ptype = "工业"
    elif re.search(r"市政|道路|桥梁|隧道|给水|排水|燃气|热力|轨道交通|风景园林"
                   r"|环境卫生|污水处理|垃圾|供热|环卫|填埋|焚烧"
                   r"|净水厂|处理厂|泵站|管网|BRT|快速公交|公交|公共交通", _shencha_query):
        shencha_ptype = "市政"
    else:
        shencha_ptype = "公建"

    if shencha_ptype != "住宅":
        sheji_kancha_sum = numerical["工程设计费(万元)"] + numerical["勘察费(万元)"]
        size = _detect_project_size_86(_shencha_query, shencha_ptype)
        shencha_r = calc_shigong_shencha(
            amount=total_part1,
            project_type=shencha_ptype,
            size=size,
            sheji_fee=round(sheji_kancha_sum, 4),
            sheji_fee_only=numerical["工程设计费(万元)"],
            kancha_fee_mid=numerical["勘察费(万元)"],
            query=_shencha_query,
        )
        raw_results["施工图审查费"] = shencha_r
        numerical["施工图审查费(万元)"] = _extract_numeric_value(shencha_r)

    t1_total = sum(
        v for k, v in numerical.items()
        if k.replace("(万元)", "") in ["交易服务费", "施工图审查费", "招标代理费"]
    )

    # ============ 总投资 ============
    initial_total = total_part1 + t0_total + t1_total
    total_investment = (
        total_investment_override
        if total_investment_override is not None
        else initial_total
    )

    # ============ Tier 2：需要总投资 ============

    # 建设管理费 — 基数 = 项目总投资 − 建管费自身（财建[2016]504号）
    # 迭代求解：建管费 = f(总投 − 建管费)
    _gl_guess = 0.0
    for __ in range(10):
        _gl_base = total_investment - _gl_guess
        _gl_r = calc_jianshe_guanli(_gl_base)
        _gl_new = _extract_numeric_value(_gl_r)
        if abs(_gl_new - _gl_guess) < 0.001:
            break
        _gl_guess = _gl_new
    gl_r = _gl_r
    raw_results["建设管理费"] = gl_r
    numerical["建设管理费(万元)"] = _gl_new

    # 可行性研究费
    keyan_ind, keyan_coef = _detect_keyan_industry(query)
    amount_yi = total_investment / 10000.0
    # 取可研报告中值
    keyan_r_all = calc_keyan(amount_yi, service_type="编制可研报告",
                             industry_coef=keyan_coef, industry_name=keyan_ind)
    raw_results["可行性研究费"] = keyan_r_all
    numerical["可行性研究费(万元)"] = _extract_numeric_value(keyan_r_all)

    # 环境影响咨询费
    hp_overrides = (coef_overrides or {}).get("环境影响咨询费", {})
    huanping_ind, huanping_coef = _detect_huanping_industry(query)
    hp_ind_coef = hp_overrides.get("industry_coef")
    if hp_ind_coef is None:
        hp_ind_coef = huanping_coef
    hp_sens_coef = hp_overrides.get("sensitivity_coef")
    if hp_sens_coef is None:
        hp_sens_coef = 1.0
    huanping_r_all = calc_huanping(total_investment, service_type="编制报告书",
                                   industry_coef=hp_ind_coef, industry_name=huanping_ind,
                                   sensitivity_coef=hp_sens_coef)
    raw_results["环境影响咨询费"] = huanping_r_all
    numerical["环境影响咨询费(万元)"] = _extract_numeric_value(huanping_r_all)

    t2_total = (
        numerical["建设管理费(万元)"]
        + numerical["可行性研究费(万元)"]
        + numerical["环境影响咨询费(万元)"]
    )

    fee_total = t0_total + t1_total + t2_total

    # ============ 预备费 (Tier 3)：基于（一类费 + 二类费） ============
    # 预备费率：用户指定 > param_overrides > 默认 5%
    yubei_rate = param_overrides.get("预备费率")
    if yubei_rate is None:
        # 从查询中检测用户指定的预备费率
        yb_rate_match = re.search(r"预备费.*?(\d+\.?\d*)\s*%|预备费率.*?(\d+\.?\d*)", query)
        if yb_rate_match:
            yubei_rate = float(yb_rate_match.group(1) or yb_rate_match.group(2))
    if yubei_rate is None:
        yubei_rate = 5.0  # 默认 5%
    yubei_fee = round((total_part1 + fee_total) * yubei_rate / 100.0, 4)
    yubei_rate_source = "用户指定" if (param_overrides.get("预备费率") or re.search(r"预备费.*?(\d+\.?\d*)\s*%|预备费率.*?(\d+\.?\d*)", query)) else "默认"
    yubei_r = {
        "费种": "预备费（基本预备费）",
        "依据": "一类费（工程费用）+ 二类费（工程建设其他费）",
        "计算公式": f"（{total_part1} + {round(fee_total, 4)}）× {yubei_rate}%（{yubei_rate_source}）",
        "结果(万元)": yubei_fee,
        "预备费率(%)": yubei_rate,
        "预备费率来源": yubei_rate_source,
    }
    raw_results["预备费"] = yubei_r
    numerical["预备费(万元)"] = yubei_fee

    t3_total = yubei_fee

    # 项目总投资 = 一类费 + 二类费 + 预备费
    project_total = total_part1 + fee_total + yubei_fee
    # 静态总投资（不含预备费）
    static_investment = total_part1 + fee_total

    # ── skip_fees 过滤：移除用户取消选中的费种 ──
    skipped = dict(_SKIP_FEES)
    if skip_fees:
        for fn in skip_fees:
            key = f"{fn}(万元)"
            if key in numerical:
                del numerical[key]
            raw_results.pop(fn, None)
            if fn in _TIER_MAP:
                skipped[fn] = "用户未选择"
        # 重新计算汇总值
        t0_keys = [k for k, v in _TIER_MAP.items() if v == 0]
        t1_keys = [k for k, v in _TIER_MAP.items() if v == 1] + ["招标代理费"]
        t2_keys = [k for k, v in _TIER_MAP.items() if v == 2]
        t0_total = sum(numerical.get(f"{k}(万元)", 0) for k in t0_keys)
        t1_total = sum(numerical.get(f"{k}(万元)", 0) for k in t1_keys)
        t2_total = sum(numerical.get(f"{k}(万元)", 0) for k in t2_keys)
        t3_total = numerical.get("预备费(万元)", 0)
        fee_total = t0_total + t1_total + t2_total
        project_total = total_part1 + fee_total + t3_total
        static_investment = total_part1 + fee_total

    return {
        "建安工程费(万元)": jianan,
        "设备购置费(万元)": shebei,
        "第一部分工程费(万元)": total_part1,
        "项目类型": project_type,
        "_数值": numerical,
        "原始结果": raw_results,
        "T0小计(万元)": round(t0_total, 4),
        "T1小计(万元)": round(t1_total, 4),
        "T2小计(万元)": round(t2_total, 4),
        "预备费小计(万元)": round(t3_total, 4),
        "总投资(万元)": round(static_investment, 4),
        "项目总投资(万元)": round(project_total, 4),
        "二类费合计(万元)": round(fee_total, 4),
        "_层级": dict(_TIER_MAP),
        "_跳过的费种": skipped,
    }


def _build_fee_summary(result: dict) -> list[dict]:
    """把 _calc_all_fees 结果转为摘要行列表。"""
    numerical = result["_数值"]
    tiers = result["_层级"]
    rows = []
    for fee_name, tier in sorted(tiers.items(), key=lambda x: (x[1], x[0])):
        val = numerical.get(f"{fee_name}(万元)")
        if val is not None:
            rows.append({"费种": fee_name, "金额(万元)": round(val, 4), "层级": tier})
    return rows


def _build_fee_selection_meta(
    engine_result: dict,
    query: str,
) -> list[dict]:
    """构建全费用选择面板的费种元数据列表。

    从 _calc_all_fees 的完整结果中提取各费种的默认值、
    系数配置、依赖关系等，供 UI 渲染勾选框和系数控件。
    """
    numerical = engine_result["_数值"]
    definitions: list[dict] = []

    # 先添加 TIER_MAP 中的费种
    for fee_name, tier in sorted(_TIER_MAP.items(), key=lambda x: (x[1], x[0])):
        default_val = numerical.get(f"{fee_name}(万元)", 0)
        label = _FEE_LABELS.get(fee_name, fee_name)
        deps = _TIER_DEPS.get(fee_name, [])
        has_coefs = fee_name in ("监理费", "工程设计费", "环境影响咨询费",
                                 "可行性研究费")
        has_rates = fee_name in ("勘察费", "劳动安全卫生评审费",
                                 "场地准备费及临时设施费", "工程保险费")
        has_services = fee_name in ("环境影响咨询费", "可行性研究费")

        entry: dict = {
            "name": fee_name,
            "label": label,
            "tier": tier,
            "has_coefs": has_coefs,
            "has_rates": has_rates,
            "has_services": has_services,
            "coef_config": None,
            "rate_config": None,
            "service_config": None,
            "depends_on": deps,
            "default_value_wan": round(default_val, 4),
        }

        # 为有系数的费种构建简化系数配置
        if has_coefs:
            entry["coef_config"] = _get_coef_config_simple(fee_name, query)

        # 为有费率选择的费种构建费率选项
        if has_rates:
            entry["rate_config"] = _get_rate_config_simple(
                fee_name, engine_result, query)

        # 为环评费 / 可行性研究费构建服务类型选项
        if has_services:
            if fee_name == "环境影响咨询费":
                entry["service_config"] = {
                    "services": [
                        {"name": "编制报告书", "label": "编制环境影响报告书（含大纲）"},
                        {"name": "编制报告表", "label": "编制环境影响报告表"},
                        {"name": "评估报告书", "label": "评估环境影响报告书（含大纲）"},
                        {"name": "评估报告表", "label": "评估环境影响报告表"},
                    ],
                    "default_selected": ["编制报告书"],
                }
            elif fee_name == "可行性研究费":
                entry["service_config"] = {
                    "services": [
                        {"name": "编制项目建议书", "label": "编制项目建议书"},
                        {"name": "编制可研报告", "label": "编制可行性研究报告"},
                        {"name": "评估项目建议书", "label": "评估项目建议书"},
                        {"name": "评估可研报告", "label": "评估可行性研究报告"},
                    ],
                    "default_selected": ["编制可研报告"],
                }

        definitions.append(entry)

    # 追加 _SKIP_FEES 中可通过依赖费种计算的费种
    # 招标代理费：依赖监理费+设计费+勘察费（Tier 0 已计算），可自动联算
    zhaobiao_val = numerical.get("招标代理费(万元)", 0)
    definitions.append({
        "name": "招标代理费",
        "label": _FEE_LABELS.get("招标代理费", "招标代理服务费"),
        "tier": 1,  # 放在 Tier 1（与交易服务费同级，依赖 Tier 0）
        "has_coefs": False,
        "has_rates": False,
        "has_services": False,
        "coef_config": None,
        "rate_config": None,
        "service_config": None,
        "depends_on": ["监理费", "工程设计费", "勘察费"],
        "default_value_wan": round(zhaobiao_val, 4) if zhaobiao_val else 0,
        "is_from_skip": True,  # 标记为来自 _SKIP_FEES
    })

    # 造价咨询费：需要选择具体服务子项（预算编制/结算审核等）
    # 支持天津（津价房地[2008]136号）和河北（冀建市研[2017]2号）两套规则
    _cc_tj_services = [
        {"name": "编制施工图预算", "label": "编制施工图预算（基数=工程费用）"},
        {"name": "编制工程量清单", "label": "编制工程量清单（基数=工程费用）"},
        {"name": "编制标底(含清单)", "label": "编制标底，含清单（基数=工程费用）"},
        {"name": "编制竣工结算", "label": "编制竣工结算（基数=工程费用）"},
        {"name": "施工阶段全过程造价控制", "label": "施工阶段全过程造价控制（基数=工程费用）"},
        {"name": "审核概算", "label": "审核概算（基数=总投资）"},
        {"name": "审核预算、标底", "label": "审核预算、标底（基数=工程费用）"},
        {"name": "审核竣工结算", "label": "审核竣工结算（基数=工程费用）"},
        {"name": "编制项目投资估算", "label": "编制项目投资估算（基数=建安费）"},
        {"name": "编制设计概算", "label": "编制设计概算（基数=建安费）"},
    ]
    _cc_hb_services = [
        {"name": "预算编制", "label": "预算编制（基数=建安费）"},
        {"name": "结算编制", "label": "结算编制（基数=建安费）"},
        {"name": "结算审核", "label": "结算审核（基数=建安费）"},
        {"name": "概算编制", "label": "概算编制（基数=设计概算造价）"},
        {"name": "概算审核", "label": "概算审核（基数=设计概算造价）"},
        {"name": "投资估算", "label": "投资估算（基数=投资估算造价）"},
        {"name": "经济评价", "label": "经济评价（基数=投资估算造价）"},
        {"name": "工程量清单编制(审核)", "label": "工程量清单编制/审核（基数=建安费）"},
        {"name": "招标控制价编制(审核)", "label": "招标控制价编制/审核（基数=建安费）"},
        {"name": "竣工决算编制", "label": "竣工决算编制（基数=总投资）"},
        {"name": "预算审核", "label": "预算审核（基数=建安费）"},
        {"name": "投标报价分析(清标)", "label": "投标报价分析/清标（基数=最高投标限价）"},
        {"name": "施工阶段造价咨询", "label": "施工阶段造价咨询（基数=建安费）"},
        {"name": "全过程造价咨询", "label": "全过程造价咨询（基数=建安费）"},
        {"name": "工程造价鉴定", "label": "工程造价鉴定（基数=鉴定标的额）"},
    ]
    _cc_hebei = _is_hebei_project(query)
    definitions.append({
        "name": "造价咨询费",
        "label": _FEE_LABELS.get("造价咨询费", "工程造价咨询服务费"),
        "tier": 0,  # 仅需建安+设备费
        "has_coefs": _cc_hebei,  # 河北项目有专业调整系数（附件2）
        "has_rates": False,
        "has_services": True,
        "coef_config": _get_coef_config_simple("造价咨询费", query) if _cc_hebei else None,
        "rate_config": None,
        "service_config": {
            "services_tianjin": _cc_tj_services,
            "services_hebei": _cc_hb_services,
            "default_selected_tianjin": ["编制施工图预算"],
            "default_selected_hebei": ["预算编制"],
        },
        "depends_on": [],
        "default_value_wan": 0,
        "is_from_skip": True,  # 标记为来自 _SKIP_FEES
    })

    return definitions


def _get_coef_config_simple(fee_name: str, query: str) -> dict | None:
    """为费种选择面板构建简化的系数配置元数据。

    与 _build_coef_metadata 不同，此函数不需要完整的 result 字典，
    仅从 query 中提取默认系数值，用于在折叠面板中渲染系数选择器。
    """
    if fee_name == "监理费":
        prof = _extract_jianli_professional_coef(query)
        comp = _extract_jianli_complexity_coef(query)
        elev = _extract_jianli_elevation_coef(query)
        return {
            "calc_func": "calc_jianli",
            "coefs": [
                {
                    "key": "专业调整系数",
                    "param_name": "professional_coef",
                    "current": prof,
                    "options": list(JIANLI_PROFESSIONAL_OPTIONS),
                    "description": "发改价格[2007]670号 附表三",
                },
                {
                    "key": "复杂程度系数",
                    "param_name": "complexity_coef",
                    "current": comp,
                    "options": list(JIANLI_COMPLEXITY_OPTIONS),
                    "description": "发改价格[2007]670号 1.0.9条",
                },
                {
                    "key": "高程调整系数",
                    "param_name": "elevation_coef",
                    "current": elev,
                    "options": list(JIANLI_ELEVATION_OPTIONS),
                    "description": "发改价格[2007]670号 1.0.9条",
                },
            ],
        }
    elif fee_name == "工程设计费":
        prof = _extract_sheji_professional_coef(query)
        comp = _extract_sheji_complexity_coef(query)
        # 附加调整系数：从 query 中尝试提取
        addi_matches = re.findall(r"附加.*?系数.*?(\d+\.?\d*)", query)
        addi = float(addi_matches[0]) if addi_matches else 1.0
        return {
            "calc_func": "calc_sheji",
            "coefs": [
                {
                    "key": "专业调整系数",
                    "param_name": "professional_coef",
                    "current": prof,
                    "options": list(SHEJI_PROFESSIONAL_OPTIONS),
                    "description": "计价格[2002]10号 附表二",
                },
                {
                    "key": "复杂程度系数",
                    "param_name": "complexity_coef",
                    "current": comp,
                    "options": list(SHEJI_COMPLEXITY_OPTIONS),
                    "description": "计价格[2002]10号 1.0.9.2",
                },
                {
                    "key": "附加调整系数",
                    "param_name": "additional_coef",
                    "current": addi,
                    "options": [],  # 无预设选项，自定义输入
                    "description": "计价格[2002]10号 1.0.9.3（多个系数合并 = 相加 − 个数 + 1）",
                },
            ],
        }
    elif fee_name == "环境影响咨询费":
        ind_name, ind_coef = _detect_huanping_industry(query)
        return {
            "calc_func": "calc_huanping",
            "coefs": [
                {
                    "key": "行业调整系数",
                    "param_name": "industry_coef",
                    "current": ind_coef,
                    "options": list(HUANPING_INDUSTRY_OPTIONS),
                    "description": "计价格[2002]125号 附件二 表1",
                },
                {
                    "key": "环境敏感程度系数",
                    "param_name": "sensitivity_coef",
                    "current": 1.0,
                    "options": list(HUANPING_SENSITIVITY_OPTIONS),
                    "description": "计价格[2002]125号 附件二 表2",
                },
            ],
        }
    elif fee_name == "造价咨询费":
        # 河北省专业工程调整系数（冀建市研[2017]2号 附件2）
        return {
            "calc_func": "calc_cost_consulting_multi_hebei",
            "coefs": [
                {
                    "key": "专业调整系数",
                    "param_name": "professional_coef",
                    "current": 1.0,
                    "options": list(_HEBEI_PROFESSIONAL_COEFFICIENTS.items()),
                    "description": "冀建市研[2017]2号 附件2：不同专业工程适用不同调整系数",
                },
            ],
        }
    elif fee_name == "可行性研究费":
        ind_name, ind_coef = _detect_keyan_industry(query)
        # 行业系数按标准5档分组
        _ky_options = [
            ("石化、化工、钢铁", 1.3),
            ("石油、天然气、水利、水电、交通（水运）、化纤", 1.2),
            ("有色、黄金、纺织、轻工、邮电、广播电视、医药、煤炭、"
             "火电（含核电）、机械（含船舶、航空、航天、兵器）", 1.0),
            ("林业、商业、粮食、建筑", 0.8),
            ("建材、交通（公路）、铁道、市政公用工程", 0.7),
        ]
        return {
            "calc_func": "calc_keyan",
            "coefs": [
                {
                    "key": "行业调整系数",
                    "param_name": "industry_coef",
                    "current": ind_coef,
                    "options": _ky_options,
                    "description": "计价格[1999]1283号：不同行业适用不同调整系数",
                },
                {
                    "key": "复杂程度系数",
                    "param_name": "complexity_coef",
                    "current": 1.0,
                    "options": [("简单", 0.8), ("一般", 1.0), ("复杂", 1.2)],
                    "description": "计价格[1999]1283号：工程复杂程度调整系数",
                },
            ],
        }
    return None


def _get_rate_config_simple(
    fee_name: str,
    engine_result: dict,
    query: str,
) -> dict | None:
    """为费种选择面板构建费率选择元数据。

    从 engine_result 的原始结果中提取费率明细，供 UI 渲染费率选择器。
    """
    raw_results = engine_result.get("原始结果", {})
    fee_result = raw_results.get(fee_name)
    if not fee_result:
        return None

    rate_detail = fee_result.get("费率明细", [])
    if not rate_detail:
        return None

    # 费率选项：{"费率": "0.3%", "费用(万元)": 15.0}
    rate_options = [{"rate": d["费率"], "fee_wan": d["费用(万元)"]} for d in rate_detail]

    # 默认选中值（取中值对应的费率）
    mid_idx = len(rate_detail) // 2
    default_rate = rate_detail[mid_idx]["费率"]

    # param_overrides 键名
    param_key_map = {
        "勘察费": "勘察费费率",
        "劳动安全卫生评审费": "劳动安全卫生评审费费率",
        "场地准备费及临时设施费": "场地准备费费率",
        "工程保险费": "工程保险费费率",
    }

    return {
        "param_key": param_key_map.get(fee_name, ""),
        "rate_options": rate_options,
        "default_rate": default_rate,
        "basis": fee_result.get("依据", ""),
        "desc": fee_result.get("说明", ""),
    }


def _extract_extra_fees(query: str, known_fees: set | None = None) -> list[dict]:
    """从查询中提取用户额外指定的费用（如"旧桥检测费15万"、"增加XX费YY万"）。"""
    known = known_fees or set()
    extra = []
    # 匹配格式："增加/额外/另加/外加/另计 XX费 YY万"
    pattern = r"(?:增加|额外|另加|外加|另计|新增|加)\s*(\S{2,8}?(?:费|检测|试验|评估|监测|加固|拆除|迁改|修复))\s*(\d+\.?\d*)\s*万"
    for m in re.finditer(pattern, query):
        name = m.group(1).strip()
        amount = float(m.group(2))
        # 去重：同一名称只取第一次出现的
        if name not in known and not any(e["名称"] == name for e in extra):
            extra.append({"名称": name, "金额(万元)": amount})
            known.add(name)
    return extra


# ==================== 模式1：多费种联算 ====================

def calc_cascade(query: str) -> dict | None:
    """模式1：给定建安+设备费，一次性计算所有可自动计算的二类费。"""
    jianan, shebei = _extract_jianli_components(query)
    if jianan is None:
        amount = _extract_amount(query)
        if amount is not None:
            jianan = amount
            shebei = 0.0
        else:
            return None
    shebei = shebei or 0.0

    project_type = _detect_project_type(query)
    engine_result = _calc_all_fees(jianan, shebei, project_type, query)

    # 提取用户额外指定的费用
    known_names = set(engine_result["_层级"].keys())
    known_names.update(engine_result["_跳过的费种"].keys())
    extra_fees = _extract_extra_fees(query, known_names)
    extra_total = sum(e["金额(万元)"] for e in extra_fees)

    fee_summary = _build_fee_summary(engine_result)

    base_total = engine_result["二类费合计(万元)"]
    base_invest = engine_result["总投资(万元)"]
    project_total = engine_result["项目总投资(万元)"]
    yubei_fee = engine_result.get("预备费小计(万元)", 0)

    return {
        "mode": "cascade",
        "fee_type": "__cascade__",
        "has_amount": True,
        "输入参数": {
            "建安工程费(万元)": jianan,
            "设备购置费(万元)": shebei,
            "第一部分工程费(万元)": jianan + shebei,
            "项目类型": project_type,
        },
        "费种合计": fee_summary,
        "结果汇总": {
            "第一部分工程费(万元)": engine_result["第一部分工程费(万元)"],
            "T0小计(万元)": engine_result["T0小计(万元)"],
            "T1小计(万元)": engine_result["T1小计(万元)"],
            "T2小计(万元)": engine_result["T2小计(万元)"],
            "预备费(万元)": yubei_fee,
            "额外费用小计(万元)": round(extra_total, 4),
            "二类费合计(万元)": round(base_total + extra_total, 4),
            "项目总投资(万元)": round(project_total + extra_total, 4),
        },
        "明细": engine_result["原始结果"],
        "额外费用": extra_fees,
        "跳过的费种": engine_result["_跳过的费种"],
        "_engine_raw": engine_result,       # 供 UI 初始化费种选择面板
        "_数值": engine_result["_数值"],     # 供 UI 获取各费种默认值
    }


# ==================== 模式2：迭代收敛 ====================

def _snapshot_fees(numerical: dict[str, float]) -> dict[str, float]:
    """快照当前费用数值。"""
    return dict(numerical)


def calc_iteration(query: str) -> dict | None:
    """模式2：迭代收敛计算建设管理费（处理循环依赖）。"""
    jianan, shebei = _extract_jianli_components(query)
    if jianan is None:
        amount = _extract_amount(query)
        if amount is not None:
            jianan = amount
            shebei = 0.0
        else:
            return None
    shebei = shebei or 0.0

    project_type = _detect_project_type(query)
    threshold = 0.01  # 万元
    max_iter = 20

    # 第一轮：初始总投资
    result = _calc_all_fees(jianan, shebei, project_type, query)
    iteration_history = [{
        "迭代次数": 0,
        "总投资(万元)": result["总投资(万元)"],
        "二类费合计(万元)": result["二类费合计(万元)"],
        "变化(万元)": 0.0,
        "各项费用": _snapshot_fees(result["_数值"]),
    }]

    for i in range(1, max_iter + 1):
        prev_total = result["总投资(万元)"]
        result = _calc_all_fees(
            jianan, shebei, project_type, query,
            total_investment_override=prev_total,
        )
        current_total = result["总投资(万元)"]
        diff = round(current_total - prev_total, 6)

        iteration_history.append({
            "迭代次数": i,
            "总投资(万元)": current_total,
            "二类费合计(万元)": result["二类费合计(万元)"],
            "变化(万元)": diff,
            "各项费用": _snapshot_fees(result["_数值"]),
        })

        if abs(diff) < threshold:
            break

    # 提取用户额外指定的费用（不参与迭代，直接加到最终结果）
    known_names = set(result["_层级"].keys())
    known_names.update(result["_跳过的费种"].keys())
    extra_fees = _extract_extra_fees(query, known_names)
    extra_total = sum(e["金额(万元)"] for e in extra_fees)

    final_history = iteration_history[-1]
    # 预备费（基于收敛后的结果）
    final_numerical = result["_数值"]
    yubei_fee = final_numerical.get("预备费(万元)", 0)
    static_total = final_history["总投资(万元)"]  # 不含预备费
    final_fee_total = final_history["二类费合计(万元)"]
    return {
        "mode": "iteration",
        "fee_type": "__iteration__",
        "has_amount": True,
        "输入参数": {
            "建安工程费(万元)": jianan,
            "设备购置费(万元)": shebei,
            "项目类型": project_type,
            "收敛阈值(万元)": threshold,
        },
        "迭代过程": iteration_history,
        "迭代次数": len(iteration_history) - 1,
        "已收敛": abs(final_history["变化(万元)"]) < threshold,
        "收敛阈值(万元)": threshold,
        "收敛结果": {
            **final_history,
            "预备费(万元)": yubei_fee,
            "总投资(万元)": round(static_total + extra_total, 4),
            "项目总投资(万元)": round(static_total + yubei_fee + extra_total, 4),
            "二类费合计(万元)": round(final_fee_total + extra_total, 4),
        },
        "额外费用": extra_fees,
        "明细": result["原始结果"],
        "跳过的费种": result["_跳过的费种"],
    }


# ==================== 模式3：多方案比选 ====================

def _detect_sweep_parameter(query: str) -> dict:
    """从查询中检测扫描参数，默认扫描勘察费费率 0.8%~1.1%。"""
    return {
        "参数名称": "勘察费费率",
        "值列表": [0.8, 0.9, 1.0, 1.1],
        "参数描述": "勘察费费率",
        "单位": "%",
    }


def calc_comparison(query: str) -> dict | None:
    """模式3：多方案比选/敏感性分析。"""
    jianan, shebei = _extract_jianli_components(query)
    if jianan is None:
        amount = _extract_amount(query)
        if amount is not None:
            jianan = amount
            shebei = 0.0
        else:
            return None
    shebei = shebei or 0.0

    project_type = _detect_project_type(query)
    sweep = _detect_sweep_parameter(query)

    scenarios = []
    all_fee_keys = set()
    for pv in sweep["值列表"]:
        result = _calc_all_fees(
            jianan, shebei, project_type, query,
            param_overrides={sweep["参数名称"]: pv},
        )
        fees = result["_数值"]
        scenario = {
            "方案名称": f"{sweep['参数描述']} {pv}{sweep['单位']}",
            "参数值": pv,
            "各项费用": fees,
            "二类费合计(万元)": result["二类费合计(万元)"],
            "总投资(万元)": result["总投资(万元)"],
            "项目总投资(万元)": result["项目总投资(万元)"],
        }
        scenarios.append(scenario)
        all_fee_keys.update(fees.keys())
    all_fee_keys.add("二类费合计(万元)")
    all_fee_keys.add("总投资(万元)")
    all_fee_keys.add("项目总投资(万元)")

    # 提取用户额外指定的费用（对所有方案一致）
    engine_result0 = _calc_all_fees(jianan, shebei, project_type, query)
    known_names = set(engine_result0["_层级"].keys())
    known_names.update(engine_result0["_跳过的费种"].keys())
    extra_fees = _extract_extra_fees(query, known_names)
    extra_total = sum(e["金额(万元)"] for e in extra_fees)

    # 将额外费用加到每个方案
    for s in scenarios:
        s["额外费用小计(万元)"] = round(extra_total, 4)
        s["各项费用"]["额外费用小计(万元)"] = round(extra_total, 4)
        s["二类费合计(万元)"] = round(s["二类费合计(万元)"] + extra_total, 4)
        s["总投资(万元)"] = round(s["总投资(万元)"] + extra_total, 4)
        s["项目总投资(万元)"] = round(s["项目总投资(万元)"] + extra_total, 4)

    if extra_fees:
        all_fee_keys.add("额外费用小计(万元)")

    # 构建对比表行
    comparison_rows = []
    for fee_key in sorted(all_fee_keys):
        row: dict = {"费用名称": fee_key}
        for i, s in enumerate(scenarios):
            if fee_key in s["各项费用"]:
                row[f"方案{i+1}"] = round(s["各项费用"][fee_key], 4)
            elif fee_key == "二类费合计(万元)":
                row[f"方案{i+1}"] = s["二类费合计(万元)"]
            elif fee_key == "总投资(万元)":
                row[f"方案{i+1}"] = s["总投资(万元)"]
            elif fee_key == "项目总投资(万元)":
                row[f"方案{i+1}"] = s["项目总投资(万元)"]
            else:
                row[f"方案{i+1}"] = ""
        comparison_rows.append(row)

    return {
        "mode": "comparison",
        "fee_type": "__comparison__",
        "has_amount": True,
        "输入参数": {
            "建安工程费(万元)": jianan,
            "设备购置费(万元)": shebei,
            "项目类型": project_type,
        },
        "扫描参数": sweep,
        "方案列表": scenarios,
        "对比表": comparison_rows,
        "额外费用": extra_fees,
    }


def detect_and_calculate_all(query: str) -> list[dict]:
    """
    检测并计算查询中涉及的所有二类费（支持"监理费和设计费分别为多少"等多费种提问）。

    返回：计算结果列表，可能为空。
    如果只检测到一种费种，返回单元素列表（与原 detect_and_calculate 行为兼容）。
    """
    # 1. 检测所有涉及的费种
    fee_types = _detect_all_fee_types(query)
    # 2. 回退：无明确费种关键词时，尝试隐含检测
    if not fee_types:
        jianan_test, shebei_test = _extract_jianli_components(query)
        if jianan_test is not None and shebei_test is not None:
            fee_types = ["监理费"]
        elif re.search(r"监理", query) and _extract_amount(query) is not None:
            fee_types = ["监理费"]
    # 3. 逐个计算
    results = []
    for ft in fee_types:
        r = detect_and_calculate(query, fee_type=ft)
        if r:
            results.append(r)
    return results