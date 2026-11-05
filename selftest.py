# -*- coding: utf-8 -*-
"""本地自测：webhook 加签、actionCard 消息体、完整 push 流程、数据日期、回调模块导入。

不发真实请求：全部 mock 掉 dingtalk_api.SESSION.post。
用法: python selftest.py
"""
from unittest import mock

import dingtalk_api


class FakeResp:
    ok = True
    status_code = 200

    @staticmethod
    def json():
        return {"errcode": 0}


captured = []


def fake_post(url, json=None, timeout=None, headers=None, params=None):
    captured.append((url, json))
    return FakeResp()


def test_webhook_markdown():
    captured.clear()
    with mock.patch.object(dingtalk_api.SESSION, "post", fake_post):
        dingtalk_api.send_via_webhook(
            "https://oapi.dingtalk.com/robot/send?access_token=x",
            "SECtest", "标题", "正文", at_mobiles=["13800000001"])

    url, body = captured[0]
    assert "timestamp=" in url and "sign=" in url, url
    assert body["msgtype"] == "markdown"
    assert body["at"]["atMobiles"] == ["13800000001"]
    assert "@13800000001" in body["markdown"]["text"]
    print("1. webhook markdown 加签+消息体 OK")


def test_webhook_actioncard():
    captured.clear()
    with mock.patch.object(dingtalk_api.SESSION, "post", fake_post):
        dingtalk_api.send_actioncard_via_webhook(
            "https://oapi.dingtalk.com/robot/send?access_token=x",
            "SECtest", "标题", "正文")

    url, body = captured[0]
    assert "sign=" in url
    assert body["msgtype"] == "actionCard"
    assert body["actionCard"]["title"] == "标题"
    assert "singleTitle" not in body["actionCard"]  # 无按钮
    print("2. webhook actionCard 消息体 OK（无按钮）")


def test_no_duplicate_at():
    """正文已含 @手机号时，不应在末尾再追加一排重复 @。"""
    captured.clear()
    text = "提醒：\n\n• @13800000001（某人）已读 0/3"
    with mock.patch.object(dingtalk_api.SESSION, "post", fake_post):
        dingtalk_api.send_via_webhook(
            "https://oapi.dingtalk.com/robot/send?access_token=x",
            "", "标题", text, at_mobiles=["13800000001"])
    assert captured[0][1]["markdown"]["text"] == text
    print("3. 不重复追加 @ OK")


def test_push_flow():
    import main
    import yaml

    cfg = yaml.safe_load(open("config.yaml", encoding="utf-8"))
    cfg.setdefault("webhook", {})
    cfg["webhook"]["webhook_url"] = "https://oapi.dingtalk.com/robot/send?access_token=FAKE"

    captured.clear()
    with mock.patch.object(dingtalk_api.SESSION, "post", fake_post):
        main.push(cfg)

    types = [b.get("msgtype") for _, b in captured]
    assert types[0] == "actionCard", types          # 主统计卡片
    assert set(types[1:]) <= {"markdown"}, types    # @提醒（无未使用人员时可能只有 1 条）
    print(f"4. 完整 push 流程 OK，共发送 {len(captured)} 条: {types}")


def test_data_date():
    from excel_stats import load_data_date
    import os
    import yaml
    import datetime

    cfg = yaml.safe_load(open("config.yaml", encoding="utf-8"))
    path = (cfg.get("data") or {}).get("data_file")
    if not path or not os.path.exists(path):
        print("5. 跳过数据日期检查（数据文件不存在）")
        return
    d = load_data_date(path)
    assert isinstance(d, datetime.date), f"数据文件缺少 meta 日期: {path}"
    print(f"5. 数据日期 meta OK: {d}")


def test_callback_import():
    import callback_server  # noqa: F401
    print("6. callback_server 模块导入 OK (dingtalk-stream SDK 正常)")


if __name__ == "__main__":
    test_webhook_markdown()
    test_webhook_actioncard()
    test_no_duplicate_at()
    test_push_flow()
    test_data_date()
    test_callback_import()
    print("\n全部自测通过")
