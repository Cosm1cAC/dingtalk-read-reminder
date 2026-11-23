# -*- coding: utf-8 -*-
"""主入口。

子命令:
  make-template  生成 Excel 数据模板 (data/已读统计.xlsx)
  list-groups    查询机器人所在的群（拿 openConversationId）
  push           立即推送一次统计提醒（卡片 + 未使用人员 @提醒）
  daemon         常驻模式：每天 push_time 时间自动推送（不想用 Windows 计划任务时用）
  daily-preview  解析每日报表目录，终端预览三指标四块通报并生成预览文件（不发钉钉）
  daily-push     推送三指标四块每日通报到钉钉群（含 @责任人）
"""
import argparse
import datetime
import os
import sys
import time

import yaml

from card_builder import build_card_param_map, build_markdown, build_at_message
from daily_report import build_daily_messages, load_daily_report, render_at_names
from dingtalk_api import (DingTalkClient, MSG_KEY_CARD, MSG_KEY_MARKDOWN,
                          send_actioncard_via_webhook, send_via_webhook)
from excel_stats import load_people, make_template

EXCEL_PATH = "data/当前统计.xlsx"  # 可被 config.yaml 的 data.data_file 覆盖


def load_config(path="config.yaml") -> dict:
    if not os.path.exists(path):
        raise SystemExit(
            "缺少 config.yaml。请先复制配置样例：\n"
            "  macOS/Linux:  cp config.example.yaml config.yaml\n"
            "  Windows CMD:  copy config.example.yaml config.yaml\n"
            "然后填入 webhook 地址与加签密钥（见 README「快速开始」第 2 步）"
        )
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def push(cfg: dict):
    excel_path = (cfg.get("data") or {}).get("data_file") or EXCEL_PATH
    stats_list = load_people(excel_path)
    from excel_stats import Stats
    stats = Stats(all_people=stats_list)
    # 标题/正文日期跟随数据本身（meta sheet），没有 meta 才退回"今天-1"
    from excel_stats import load_data_date
    today = load_data_date(excel_path) or (datetime.date.today() - datetime.timedelta(days=1))
    show_names = (cfg.get("card") or {}).get("show_detail_names", True)

    webhook_cfg = cfg.get("webhook") or {}
    webhook_url = webhook_cfg.get("webhook_url") or ""

    # ===== webhook 模式（零权限，群自定义机器人）=====
    if webhook_url and "你的" not in webhook_url:
        client = None
        card_map = build_card_param_map(stats, today, show_detail_names=show_names)
        # 主统计消息用 actionCard（卡片样式渲染），@提醒必须用 markdown（actionCard 不支持 @）
        send_actioncard_via_webhook(webhook_url, webhook_cfg.get("secret", ""),
                                    card_map["title"], build_markdown(stats, today, show_detail_names=show_names))
        print("[OK] 主统计消息已按 actionCard 卡片样式发送")
        text, at_mobiles = build_at_message(stats, today)
        if text:
            send_via_webhook(webhook_url, webhook_cfg.get("secret", ""),
                             "未使用人员提醒", text, at_mobiles=at_mobiles)
            print("[OK] 未使用人员 @提醒已发送（@%s）" % ", ".join(at_mobiles))
        else:
            print("[OK] 无未使用人员，跳过 @提醒")
        return

    # ===== app 模式（企业内部应用，支持互动卡片按钮）=====
    app = cfg["dingtalk"]
    conversation_id = app.get("open_conversation_id")
    if not app.get("app_key") or "你的" in str(app.get("app_key")):
        raise SystemExit("config.yaml 里 webhook 和应用凭证都没填。没有开发权限就填 webhook（群自定义机器人），"
                         "有开发权限就填 app_key/app_secret，见 README")
    if not conversation_id or "你的" in conversation_id:
        raise SystemExit("请先执行 main.py list-groups 查群列表，把 openConversationId 填进 config.yaml")

    client = DingTalkClient(app["app_key"], app["app_secret"], app["robot_code"])

    # 1. 主统计消息
    template_id = app.get("card_template_id")
    card_map = build_card_param_map(stats, today, show_detail_names=show_names)
    if template_id and "你的" not in template_id:
        client.send_group_message(conversation_id, MSG_KEY_CARD, {
            "cardTemplateId": template_id,
            "cardData": {"cardParamMap": card_map},
        })
        print("[OK] 互动卡片已发送（模板: %s）" % template_id)
    else:
        # 未配置卡片模板时用 markdown 兜底（无按钮）
        client.send_markdown(conversation_id, card_map["title"], build_markdown(stats, today, show_detail_names=show_names))
        print("[OK] markdown 消息已发送（未配置卡片模板，无按钮；见 README 配置互动卡片）")

    # 2. 未使用人员 @提醒
    text, at_mobiles = build_at_message(stats, today)
    if text:
        client.send_markdown(conversation_id, "未使用人员提醒", text, at_mobiles=at_mobiles)
        print("[OK] 未使用人员 @提醒已发送（@%s）" % ", ".join(at_mobiles))
    else:
        print("[OK] 无未使用人员，跳过 @提醒")


def cmd_list_groups(cfg: dict):
    app = cfg["dingtalk"]
    client = DingTalkClient(app["app_key"], app["app_secret"], app["robot_code"])
    groups = client.list_chat_groups()
    if not groups:
        print("机器人不在任何群里。先把机器人拉进目标群，再重试。")
        return
    print("机器人所在群列表（把目标群的 openConversationId 填进 config.yaml）:\n")
    for g in groups:
        print(f"  群名称: {g.get('title') or g.get('openGroupName')}")
        print(f"  openConversationId: {g.get('openConversationId')}")
        print("  ---")


def cmd_daemon(cfg: dict):
    push_time = cfg.get("schedule", {}).get("push_time", "17:00")
    print(f"常驻模式已启动，每天 {push_time} 自动推送（Ctrl+C 退出）")
    pushed_date = None
    while True:
        now = datetime.datetime.now()
        hh, mm = map(int, push_time.split(":"))
        target = now.replace(hour=hh, minute=mm, second=0, microsecond=0)
        if now >= target and pushed_date != now.date():
            try:
                push(cfg)
                pushed_date = now.date()
            except Exception as e:
                print(f"[ERROR] 推送失败: {e}")
        time.sleep(30)


def _resolve_app_client(cfg: dict) -> tuple:
    app = cfg.get("dingtalk") or {}
    conversation_id = app.get("open_conversation_id")
    if not app.get("app_key") or "你的" in str(app.get("app_key")):
        raise SystemExit("config.yaml 里 webhook 和应用凭证都没填。没有开发权限就填 webhook（群自定义机器人），"
                         "有开发权限就填 app_key/app_secret，见 README")
    if not conversation_id or "你的" in conversation_id:
        raise SystemExit("请先执行 main.py list-groups 查群列表，把 openConversationId 填进 config.yaml")
    client = DingTalkClient(app["app_key"], app["app_secret"], app["robot_code"])
    return app, conversation_id, client


def cmd_daily_preview(cfg: dict):
    data = load_daily_report(cfg)
    msg = build_daily_messages(data, cfg)

    os.makedirs("data", exist_ok=True)
    preview_path = os.path.join("data", f"每日通报预览_{data.daily_date.isoformat()}.md")
    preview_md = (
        f"## 消息 1：主通报卡片（actionCard）\n\n{msg['card_text']}\n\n"
        f"---\n\n"
        f"## 消息 2：整改 @提醒（markdown）\n\n{msg['at_text']}\n"
    )
    with open(preview_path, "w", encoding="utf-8") as f:
        f.write(preview_md)

    print("匹配到的源文件:")
    for k, fn in data.source_files.items():
        print(f"  - {k}: {fn}")
    print("\n" + "=" * 60)
    print(f"[消息 1 · actionCard] 标题: {msg['title']}")
    print("=" * 60)
    print(msg["card_text"])
    print("\n" + "=" * 60)
    print(f"[消息 2 · markdown @提醒] 标题: {msg['at_title']}")
    print("=" * 60)
    print(render_at_names(msg["at_text"], msg["mobile_to_name"]))
    print("=" * 60)
    print(f"[预览已写入] {preview_path}")
    print(f"[@汇总] 已匹配手机号 {len(msg['at_mobiles'])} 人；缺失手机号 {len(msg['missing_phone_names'])} 人")
    if msg["missing_phone_names"]:
        print("请用以下命令补齐缺失人员手机号（补一次永久复用）:")
        for name in msg["missing_phone_names"]:
            print(f"  python contacts.py {name} <手机号>")
    if msg.get("unmatched_teams"):
        print("以下末位班组暂未自动匹配到组长，如需 @ 具体负责人请在 config.yaml 的 report.team_owners 中补充:")
        for team in msg["unmatched_teams"]:
            print(f'  "{team}": ["组长姓名"]')


def cmd_daily_push(cfg: dict):
    data = load_daily_report(cfg)
    msg = build_daily_messages(data, cfg)
    at_mobiles = msg["at_mobiles"] or None
    missing = msg["missing_phone_names"]

    webhook_cfg = cfg.get("webhook") or {}
    webhook_url = webhook_cfg.get("webhook_url") or ""
    if webhook_url and "你的" not in webhook_url:
        secret = webhook_cfg.get("secret", "")
        send_actioncard_via_webhook(webhook_url, secret, msg["title"], msg["card_text"])
        send_via_webhook(webhook_url, secret, msg["at_title"], msg["at_text"], at_mobiles=at_mobiles)
        print(f"[OK] 每日通报（actionCard 主卡片 + markdown @提醒）已通过 webhook 发送（@人数: {len(msg['at_mobiles'])}）")
    else:
        _, conversation_id, client = _resolve_app_client(cfg)
        client.send_markdown(conversation_id, msg["title"], msg["card_text"])
        client.send_markdown(conversation_id, msg["at_title"], msg["at_text"], at_mobiles=at_mobiles)
        print(f"[OK] 每日通报双消息已通过企业应用发送（@人数: {len(msg['at_mobiles'])}）")

    if missing:
        print(f"[WARN] 以下 {len(missing)} 人暂缺手机号，已退化为文本点名: {', '.join(missing)}")


def main():
    parser = argparse.ArgumentParser(description="钉钉中心管理助手-已读情况提醒机器人")
    parser.add_argument(
        "command",
        choices=["make-template", "list-groups", "push", "daemon", "daily-preview", "daily-push"],
    )
    args = parser.parse_args()

    if args.command == "make-template":
        # 优先按配置里的数据路径生成，避免和 push 读的路径不一致
        target = EXCEL_PATH
        if os.path.exists("config.yaml"):
            target = (load_config().get("data") or {}).get("data_file") or target
        make_template(target)
        print(f"[OK] 模板已生成: {target}（填好姓名/手机号/职位/应读/已读后再执行 push）")
        return

    cfg = load_config()
    if args.command == "list-groups":
        cmd_list_groups(cfg)
    elif args.command == "push":
        push(cfg)
    elif args.command == "daemon":
        cmd_daemon(cfg)
    elif args.command == "daily-preview":
        cmd_daily_preview(cfg)
    elif args.command == "daily-push":
        cmd_daily_push(cfg)


if __name__ == "__main__":
    sys.exit(main())
