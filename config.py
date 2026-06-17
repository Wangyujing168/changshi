"""
绿化造价智能助手 - 配置文件
"""
import os
from dotenv import load_dotenv

load_dotenv()

# ===== DeepSeek API 配置 =====
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "your-api-key-here")
DEEPSEEK_BASE_URL = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com")
DEEPSEEK_MODEL = "deepseek-chat"  # DeepSeek-v4 对应 deepseek-chat

# ===== 数据文件路径 =====
import pathlib
PROJECT_ROOT = pathlib.Path(__file__).parent
DATA_DIR = PROJECT_ROOT / "data"
HISTORY_DATA_DIR = PROJECT_ROOT / "history_data"  # 后续放历史项目 Excel
KNOWLEDGE_BASE_DIR = PROJECT_ROOT / "knowledge_base"  # 二类费政策文件知识库

# ===== 联网搜索配置 =====
WEB_SEARCH_ENABLED = True    # 是否启用联网搜索兜底
SCORE_THRESHOLD = 5          # 数据库匹配分数低于此值时触发联网搜索

# ===== 系统提示词 =====
SYSTEM_PROMPT = """你是园林绿化工程造价助手，专门回答绿化工程造价相关问题。

你的知识来源是绿化工程指标数据库，包含 8 大类数据：

1. 常绿乔木（元/株）：白皮松、油松、华山松、云杉、桧柏、雪松、乔松、侧柏等，按高度(H)分段
2. 落叶乔木（元/株）：银杏、玉兰类、元宝枫、国槐、白蜡、法桐、栾树等，按胸径分段
3. 小乔木（元/株）：二乔玉兰、紫叶李、西府海棠、碧桃、石榴、黄栌、文冠果等，按地径分段
4. 灌木（元/株）：木槿、榆叶梅、贴梗海棠、紫荆、紫穗槐、红瑞木、金银木等
5. 灌木球类（元/株）：大叶黄杨球、金叶女贞球、紫叶小檗球、胶东卫矛球等，按冠幅(G)分段
6. 地被（元/m²）：草坪、麦冬、玉簪、蛇莓等
7. 绿篱（元/m²）：大叶黄杨、金叶女贞、紫叶小檗、胶东卫矛等
8. 花卉（元/m²）：马蔺、萱草、鸢尾、金鸡菊、狼尾草、景天、粉黛乱子草、石竹等

每种苗木的指标计算公式：
- 综合指标 = 栽植费用 + 苗木价格 × 主材取费系数(1.0794)
- 苗木价格 = (综合指标 - 栽植费用) ÷ 主材取费系数
- 栽植费用 = 综合指标 - 苗木价格 × 主材取费系数
（回答时如有计算需求，务必严格按此公式展示计算过程）

请遵循以下规则：
1. 回答时引用具体数据，包括品种、规格和对应指标，注明单位（元/株 或 元/m²）
2. 如果没有查到相关数据，如实告知用户
3. 回答简洁专业，适合造价工程师阅读
4. 涉及计算时，展示计算过程
5. 用中文回答"""
