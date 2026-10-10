# -*- coding: utf-8 -*-
"""概算助手 AI 规划器：动态决定二类费计算前需要向用户确认的信息。

纯函数模块（不依赖 Streamlit），供 app.py 的 AI 对话式提问流程使用：
- run_planner：调用 DeepSeek，输出 {think, questions:[...]} 计划
- available_keys / question_done / build_ctx_summary：问题库可用性与上下文摘要
"""
import json
import re

from openai import OpenAI

from config import DEEPSEEK_API_KEY, DEEPSEEK_BASE_URL, DEEPSEEK_MODEL
from fee_engine import FEE_CATALOG, param_cards_for


class PlannerError(Exception):
    """规划器调用失败（API 异常 / 输出格式非法），上层据此降级传统清单模式。"""


_client = None


def _get_client():
    global _client
    if _client is None:
        _client = OpenAI(api_key=DEEPSEEK_API_KEY, base_url=DEEPSEEK_BASE_URL)
    return _client


_ZD_RE = re.compile(r"征地|拆迁|建设用地|土地征用|用地费")


def _zd_amount(custom_fees) -> float:
    return round(sum(
        float(e.get("amount_wan") or e.get("金额(万元)") or 0)
        for e in (custom_fees or [])
        if _ZD_RE.search(str(e.get("名称") or e.get("name") or ""))
    ), 4)


def _ai_needs_jianan(ctx: dict) -> bool:
    """镜像 fee_engine._needs_jianan（有派生费种或造价咨询费才收集建安费）。"""
    for f in ctx.get("fees") or []:
        cat = FEE_CATALOG.get(f)
        if cat is not None and cat["tier"] is not None:
            return True
    return False


# ── 问题库定义：key → (describe 供AI, done, avail)；渲染/写入在 app.py ──
_BASE_BANK: dict[str, tuple[str, object, object]] = {
    "fees_preset_confirm": (
        "全量预设费种确认（用户可排除某些费种）",
        lambda c: bool(c.get("fees_confirmed")),
        lambda c: bool(c.get("preset_all")) and not c.get("fees_confirmed")),
    "fees_select": (
        "选择要计算的费种（多选）",
        lambda c: bool(c.get("fees_confirmed")),
        lambda c: not c.get("fees_confirmed") and not c.get("preset_all")),
    "amounts_jianan": (
        "建安工程费金额（万元）",
        lambda c: (c.get("amounts") or {}).get("jianan") is not None,
        lambda c: _ai_needs_jianan(c) and (c.get("amounts") or {}).get("jianan") is None),
    "amounts_shebei": (
        "设备购置费金额（万元，可为 0）",
        lambda c: bool((c.get("amounts") or {}).get("shebei_known")),
        lambda c: (c.get("amounts") or {}).get("jianan") is not None
                  and not (c.get("amounts") or {}).get("shebei_known")),
    "zd_exists": (
        "是否有征地拆迁费",
        lambda c: bool(c.get("zd_asked")) or _zd_amount(c.get("custom_fees")) > 0,
        lambda c: not c.get("zd_asked") and _zd_amount(c.get("custom_fees")) == 0),
    "zd_placement": (
        "征地拆迁费如何列示（单独列项 / 计入工程建设其他费）",
        lambda c: bool(c.get("zd_asked"))
                  and (not c.get("zd_exists") or c.get("zd_placement") is not None),
        lambda c: bool(c.get("zd_asked")) and bool(c.get("zd_exists"))
                  and c.get("zd_placement") is None),
    "zd_refine": (
        "征地拆迁费是否逐项细化填写",
        lambda c: bool(c.get("zd_asked"))
                  and (not c.get("zd_exists") or c.get("zd_refine_asked")),
        lambda c: bool(c.get("zd_asked")) and bool(c.get("zd_exists"))
                  and not c.get("zd_refine_asked")),
    "custom_need": (
        "是否需要添加自定义费用或合同价/合同费率",
        lambda c: bool(c.get("custom_answered")),
        lambda c: not c.get("custom_answered")),
    "discount_need": (
        "是否有费用需要打折",
        lambda c: bool(c.get("discount_answered")),
        lambda c: not c.get("discount_answered")),
}


_CARD_LABEL = {
    "coef": "系数设置", "rate": "费率档位", "shencha": "施工图审查档位",
    "service": "服务类型", "party": "交易服务费计费方", "yubei_rate": "预备费率",
    "shuibao_comp": "水土保持补偿费参数", "spec": "参数设置",
}


def _card_filled(ctx: dict, card: dict) -> bool:
    """参数卡是否已填（镜像 fee_engine.advance_checklist 的 filled 判定；
    party/yubei_rate 增加 AI 模式私有完成标记，显式选默认值也算完成）。
    _ai_card_done：用户已明确关闭该卡（✅已设置/⏭用默认值）→ 后续轮次不再问。"""
    if _card_key(card) in (ctx.get("_ai_card_done") or []):
        return True
    kind = card.get("kind")
    fee = card.get("fee")
    if kind == "rate":
        return fee in ctx.get("rate_overrides", {})
    if kind == "shencha":
        return "施工图审查费" in ctx.get("rate_overrides", {})
    if kind == "coef":
        return bool(ctx.get("coef_overrides", {}).get(fee))
    if kind == "service":
        return fee in ctx.get("service_selections", {})
    if kind == "party":
        return bool(ctx.get("_ai_party_done")) or ctx.get("jiaoyi_party") is not None
    if kind == "yubei_rate":
        return bool(ctx.get("_ai_yubei_done")) or ctx.get("yubei_rate", 5.0) != 5.0
    if kind == "shuibao_comp":
        return "水土保持补偿费" in ctx.get("spec_overrides", {})
    if kind == "spec":
        return fee in ctx.get("spec_params", {})
    return False


def _card_key(card: dict) -> str:
    kind = card["kind"]
    if kind in ("coef", "rate", "service", "spec"):
        return f"card:{kind}:{card['fee']}"
    return f"card:{kind}"


def available_keys(ctx: dict) -> list[tuple[str, str]]:
    """当前可用的问题键（按计算依赖排序）→ (key, describe)。"""
    out: list[tuple[str, str]] = []
    for key, (desc, _done, avail) in _BASE_BANK.items():
        if avail(ctx) and not _done(ctx):
            out.append((key, desc))
    for card in param_cards_for(ctx):
        if _card_filled(ctx, card):
            continue
        fee_label = FEE_CATALOG.get(card["fee"], {}).get("label", card["fee"])
        out.append((_card_key(card), f"{fee_label}—{_CARD_LABEL.get(card['kind'], card['kind'])}"))
    out.append(("free", "开放问题（自由描述）"))
    return out


def question_done(ctx: dict, key: str) -> bool:
    """问题是否已解决（已答或已完成），app.py 过滤用。"""
    if key in _BASE_BANK:
        return _BASE_BANK[key][1](ctx)
    if key.startswith("card:"):
        parts = key.split(":", 2)
        kind = parts[1]
        fee_map = {"party": "交易服务费", "yubei_rate": "预备费",
                   "shuibao_comp": "水土保持补偿费", "shencha": "施工图审查费"}
        if kind in ("coef", "rate", "service", "spec"):
            card = {"kind": kind, "fee": parts[2]}
        else:
            card = {"kind": kind, "fee": fee_map.get(kind, "")}
        return _card_filled(ctx, card)
    return False


# ── 上下文摘要 ──

def _card_remaining(ctx: dict) -> str:
    cards_txt = []
    for c in param_cards_for(ctx):
        if _card_filled(ctx, c):
            continue
        label = FEE_CATALOG.get(c["fee"], {}).get("label", c["fee"])
        cards_txt.append(f"{label}({c['kind']})")
    return "、".join(cards_txt) or "无"


def build_ctx_summary(ctx: dict, available: list[tuple[str, str]]) -> str:
    """把 ctx 压成紧凑中文摘要，供规划器判断还缺什么。"""
    fees = ctx.get("fees") or []
    fees_txt = "、".join(FEE_CATALOG.get(f, {}).get("label", f) for f in fees) or "未选"
    confirmed = "已确认" if ctx.get("fees_confirmed") else "未确认"
    preset = "是" if ctx.get("preset_all") else "否"

    am = ctx.get("amounts") or {}
    bits = []
    if am.get("jianan") is not None:
        bits.append(f"建安费 {am['jianan']} 万元")
    if am.get("shebei_known"):
        bits.append(f"设备费 {am.get('shebei', 0)} 万元")
    if am.get("total_investment") is not None:
        bits.append(f"项目总投资 {am['total_investment']} 万元")
    amount_txt = "；".join(bits) or "未提供"

    cf = ctx.get("custom_fees") or []
    custom_txt = "；".join(
        f"{e.get('名称')} {e.get('金额(万元)')}万元" for e in cf) or "无"

    co = ctx.get("contract_overrides") or {}
    if co:
        ct = []
        for f, o in co.items():
            label = FEE_CATALOG.get(f, {}).get("label", f)
            if o.get("type") == "price":
                ct.append(f"{label} 一口价 {o.get('amount_wan')}万")
            else:
                ct.append(f"{label} 合同费率 {o.get('rate')}")
        contract_txt = "；".join(ct)
    else:
        contract_txt = "无"

    if ctx.get("zd_asked"):
        if ctx.get("zd_exists"):
            pl = {"separate": "单独列项", "inside": "计入工程建设其他费"}.get(
                ctx.get("zd_placement"), "未定")
            rf = "已细化" if ctx.get("zd_refine_asked") else "未细化"
            zd_txt = f"有（{pl}；{rf}）"
        else:
            zd_txt = "无"
    elif _zd_amount(cf) > 0:
        zd_txt = "有（已随自定义费用提供）"
    else:
        zd_txt = "未询问"

    party = ctx.get("jiaoyi_party") or "未设置"
    yb = ctx.get("yubei_rate", 5.0)
    discs = ctx.get("discounts") or {}
    if not ctx.get("discount_answered"):
        disc_txt = "未确认"
    else:
        disc_txt = "；".join(
            f"{FEE_CATALOG.get(f, {}).get('label', f)}×{v}" for f, v in discs.items()) or "不打折"

    notes = (ctx.get("ai_notes") or [])[-5:]
    notes_txt = "；".join(str(n)[:100] for n in notes) or "无"

    hist = []
    for rd in (ctx.get("_ai_rounds") or [])[-3:]:
        qa = "；".join(
            f"问「{qa.get('text')}」→ 答「{qa.get('answer')}」" for qa in rd.get("qa") or [])
        hist.append(f"（think：{rd.get('think') or ''}）{qa}")
    hist_txt = "\n".join(hist) or "无"

    return (f"【地区】{ctx.get('region') or '未选'}\n"
            f"【费种】{fees_txt}（{confirmed}；全量预设：{preset}）\n"
            f"【金额】{amount_txt}\n"
            f"【自定义费用】{custom_txt}\n"
            f"【合同价/合同费率】{contract_txt}\n"
            f"【征地拆迁】{zd_txt}\n"
            f"【参数】剩余卡：{_card_remaining(ctx)}\n"
            f"【交易服务费计费方】{party}\n"
            f"【预备费率】{yb}%\n"
            f"【折扣】{disc_txt}\n"
            f"【用户补充说明】{notes_txt}\n"
            f"【对话历史】\n{hist_txt}")


# ── 规划器调用与解析 ──

_AI_SYSTEM_PROMPT = (
    "你是「概算助手」的信息收集规划器。用户正在填报工程建设二类费计算所需的信息，"
    "你需要判断哪些信息已齐、还缺什么，然后用自然语言向用户确认。\n\n"
    "只输出一个 JSON 对象，不要输出任何其他文字或代码围栏。格式：\n"
    '{"think":"≤50字，一句话说明你的判断","questions":['
    '{"type":"buttons|number|text","text":"自然、简短的问题",'
    '"options":["选项A","选项B"],"unit":"万元","key":"可用问题键"}]}\n\n'
    "规则：\n"
    "1. 每轮只问 1-3 个最关键的问题；已提供/已确认的信息绝不重复问。\n"
    "2. 有明确有限选项的用 buttons（options 为字符串数组）；需要填数值的用 number"
    "（unit 填单位）；需要用户自由描述的用 text（key 固定为 \"free\"）。\n"
    "3. key 必须从「可用问题键」中选取，禁止编造；可用键为空时 questions 返回 []。\n"
    "4. think 用一句话（≤50字）说明你打算确认什么、为什么。\n"
    "5. 提问顺序按计算依赖：费种与金额 → 征地拆迁/自定义费用/合同价 → 参数（费率/系数/服务等）→ 折扣。\n"
    "6. 金额单位统一为万元，费率单位为 %。"
)

_REPAIR_HINT = "（你上次的输出格式不正确。请只输出一个 JSON 对象，不要包含任何其他文字或代码围栏。）"


def _build_user_prompt(summary: str, available: list[tuple[str, str]],
                       repair_hint: str | None = None) -> str:
    keys_txt = "\n".join(f"- {k}：{d}" for k, d in available)
    parts = [f"【当前信息】\n{summary}", f"【可用问题键】\n{keys_txt}", "请输出 JSON："]
    if repair_hint:
        parts.insert(0, repair_hint)
    return "\n\n".join(parts)


def _parse_plan(raw: str, available: list[tuple[str, str]]) -> dict:
    if not isinstance(raw, str) or not raw.strip():
        raise PlannerError("输出为空")
    s = raw.strip()
    m = re.search(r"```(?:json)?\s*(.*?)```", s, re.S)
    if m:
        s = m.group(1).strip()
    lo, hi = s.find("{"), s.rfind("}")
    if lo < 0 or hi < lo:
        raise PlannerError("输出中未找到 JSON")
    try:
        data = json.loads(s[lo:hi + 1])
    except Exception as e:
        raise PlannerError(f"JSON 解析失败：{e}") from e
    if not isinstance(data, dict):
        raise PlannerError("顶层不是对象")
    think = str(data.get("think") or "")[:50]
    qs = data.get("questions")
    if not isinstance(qs, list):
        raise PlannerError("questions 不是数组")
    valid_keys = {k for k, _ in available}
    out, seen = [], set()
    for q in qs[:3]:
        if not isinstance(q, dict):
            raise PlannerError("问题项不是对象")
        qtype = q.get("type")
        text = str(q.get("text") or "").strip()
        key = str(q.get("key") or "")
        if qtype not in ("buttons", "number", "text"):
            raise PlannerError(f"非法 type：{qtype}")
        if not text:
            raise PlannerError("问题 text 为空")
        if qtype == "text":
            key = "free"
        if key not in valid_keys:
            raise PlannerError(f"未知问题键：{key}")
        if key in seen:
            continue  # 同一轮重复问题去重（含 free）
        item = {"type": qtype, "text": text, "key": key}
        if qtype == "buttons":
            opts = q.get("options")
            if not isinstance(opts, list) or not (2 <= len(opts) <= 6) \
                    or not all(isinstance(o, str) and o.strip() for o in opts):
                raise PlannerError("buttons 需要 2-6 个非空字符串 options")
            item["options"] = [o.strip() for o in opts]
        if qtype == "number":
            item["unit"] = str(q.get("unit") or "万元")[:10]
        out.append(item)
        seen.add(key)
    return {"think": think, "questions": out}


def run_planner(summary: str, available: list[tuple[str, str]],
                repair_hint: str | None = None) -> dict:
    """调用 DeepSeek 生成下一轮提问计划；失败抛 PlannerError。"""
    msgs = [
        {"role": "system", "content": _AI_SYSTEM_PROMPT},
        {"role": "user", "content": _build_user_prompt(summary, available, repair_hint)},
    ]
    for attempt in range(2):
        raw = ""
        try:
            resp = _get_client().chat.completions.create(
                model=DEEPSEEK_MODEL, messages=msgs,
                temperature=0.2, max_tokens=800, timeout=20)
            raw = resp.choices[0].message.content or ""
            return _parse_plan(raw, available)
        except PlannerError:
            # 输出格式非法：带修复提示重试一次
            msgs.append({"role": "assistant", "content": raw[:2000]})
            msgs.append({"role": "user", "content": _REPAIR_HINT})
        except Exception as e:
            if attempt == 1:
                raise PlannerError(f"DeepSeek API 调用失败：{e}") from e
    raise PlannerError("规划器连续两次输出格式非法")
