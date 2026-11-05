# -*- coding: utf-8 -*-
"""把统计结果组装成钉钉互动卡片参数与 markdown 文本。"""
import datetime

from excel_stats import Stats, THRESHOLD


def _role_str(p) -> str:
    return f"（{p.role}）" if p.role else ""


def build_card_param_map(stats: Stats, date: datetime.date,
                         show_detail_names: bool = True) -> dict:
    """生成卡片模板变量。卡片模板在钉钉卡片平台创建，变量名与此处对应。

    show_detail_names=False 时主消息只报人数，不列任何人名（人名只在 @提醒里出现）。
    """
    date_str = f"{date.month}月{date.day}日"
    title = f"📊【{date_str} 昨日助手消息已读情况提醒】"

    overall = f"{stats.overall_rate * 100:.2f}%"
    summary = f"📈 昨日使用率（≥{THRESHOLD:.0%}）汇总：整体使用率 {overall}"

    unused = stats.unused
    # 主消息只报人数，不逐人列出——未使用人员的逐人 @提醒在第二条消息里，避免重复
    if unused:
        unused_section = f"⛔ 未使用（<{THRESHOLD:.0%}）{len(unused)}人（详见下一条 @提醒）"
    else:
        unused_section = f"⛔ 未使用（<{THRESHOLD:.0%}）0人"

    finished = stats.finished
    # 已全部读完只展示前 5 人，其余省略（名单太长会刷屏）
    MAX_FINISHED_SHOW = 5
    if finished and show_detail_names:
        shown = "、".join(p.name for p in finished[:MAX_FINISHED_SHOW])
        if len(finished) > MAX_FINISHED_SHOW:
            shown += f" 等{len(finished)}人"
        finished_section = f"🎉 已全部读完（{len(finished)}人）：\n  {shown}"
    elif finished:
        finished_section = f"🎉 已全部读完（{len(finished)}人）"
    else:
        finished_section = "🎉 已全部读完（0人）"

    partial = stats.partial_sorted
    if partial and show_detail_names:
        lines = [f"• {p.name}{_role_str(p)} 已读 {p.read}/{p.total}，完成率 {p.rate:.1%}"
                 for p in partial]
        partial_section = f"📋 部分已读（{len(partial)}人）：\n\n" + "\n\n".join(lines)
    elif partial:
        partial_section = f"📋 部分已读（{len(partial)}人）"
    else:
        partial_section = f"📋 部分已读（0人）"

    footer = "🚀 请大家及时查看管理驾驶舱消息，不要一键已读"

    return {
        "title": title,
        "summary": summary,
        "unused_section": unused_section,
        "finished_section": finished_section,
        "partial_section": partial_section,
        "footer": footer,
        "date": date_str,
    }


def build_markdown(stats: Stats, date: datetime.date,
                   show_detail_names: bool = True) -> str:
    """markdown 兜底文本（不配置卡片模板时使用）。"""
    pm = build_card_param_map(stats, date, show_detail_names=show_detail_names)
    return "\n\n".join([
        pm["title"],
        pm["summary"],
        pm["unused_section"],
        pm["finished_section"],
        pm["partial_section"],
        pm["footer"],
    ])


def build_at_message(stats: Stats, date: datetime.date) -> tuple:
    """未使用人员单独提醒消息：markdown 文本 + 需要 @ 的手机号列表。

    手机号缺失时退化为直接点名（@不出可点击样式，但内容完整）。
    """
    unused = stats.unused
    if not unused:
        return None, []
    date_str = f"{date.month}月{date.day}日"
    lines = [f"⛔ 未使用人员提醒（{len(unused)}人）："]
    at_mobiles = []
    for p in unused:
        info = f"已读 {p.read}/{p.total}，使用率 {p.rate_str}"
        who = f"{p.name}·{p.role}" if p.role else p.name
        if p.mobile:
            lines.append(f"• @{p.mobile}（{who}）{info}")
            at_mobiles.append(p.mobile)
        else:
            lines.append(f"• {who} {info}")
    # 钉钉 markdown 单个 \n 不换行，行间必须空一行
    text = f"【{date_str} 昨日助手消息已读情况提醒】\n\n" + "\n\n".join(lines)
    return text, at_mobiles
