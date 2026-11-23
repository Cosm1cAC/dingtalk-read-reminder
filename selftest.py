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


def test_daily_briefing():
    import datetime
    import os
    import yaml
    import contacts
    import daily_report
    import main

    # Part A: 使用脱敏示例数据验证四块双消息拼装与去重规则（无本地报表也能跑）
    sample_data = daily_report.DailyReportData(
        province="示例省区",
        daily_date=datetime.date(2026, 11, 20),
        improve_date=datetime.date(2026, 11, 15),
        improve_period="2026-11-09 00:00:00~2026-11-15",
        mgr_total=10,
        mgr_used=8,
        mgr_rate=80.0,
        mgr_centers=[
            daily_report.DeptStat(center="中心A", name="中心A", total=5, achieved=4),
            daily_report.DeptStat(center="中心B", name="中心B", total=5, achieved=4),
        ],
        mgr_bottom_people=[
            daily_report.ManagerPersonStat(
                name="张三", center="中心A", role="安全管理员",
                read_cnt=0, msg_cnt=10, rate=0.0, emp_id="0001", dept_name="安全监察",
            ),
        ],
        mgr_unused_count=2,
        imp_nat_total=100,
        imp_nat_ok=68,
        imp_nat_rate=68.0,
        imp_prov_total=20,
        imp_prov_ok=11,
        imp_prov_rate=55.0,
        imp_prov_rank=28,
        imp_total_provinces=32,
        imp_bottom_depts=[
            daily_report.DeptStat(
                center="中心A", name="中心A-一班_分拣1组",
                total=3, achieved=1, unimproved_names=["李四", "王五"],
            ),
        ],
        imp_bottom_people=[
            daily_report.OperatorAbnormalStat(
                name="李四", center="中心A", dept_name="一班_分拣1组",
                metric_name="破损率", value=120.0, base=70.0,
            ),
        ],
        imp_has_perf_detail=True,
        team_leaders_exact={("中心A", "一班_分拣1组"): [("组长甲", "操作组长")]},
    )
    sample_cfg = yaml.safe_load(open("config.example.yaml", encoding="utf-8"))
    with mock.patch.object(contacts, "all_contacts", return_value={"张三": "13800000001", "组长甲": "13800000002"}):
        msg = daily_report.build_daily_messages(sample_data, sample_cfg)

    for marker in ["口径说明", "① 整体结论", "② 末位部门", "③ 末位人员",
                   "中心管理助手 · 使用率", "中心管理助手 · 问题解决率", "操作工助手 · 改善率"]:
        assert marker in msg["card_text"], f"主卡片缺少模块/指标: {marker}"
    assert "④ 整改要求" in msg["at_text"] and "④ 整改要求" not in msg["card_text"]
    assert "今日动作" in msg["at_text"] and "今日动作" not in msg["card_text"]
    assert "未改善名单：李四、王五" in msg["card_text"] and "未改善名单：李四、王五" not in msg["at_text"]
    assert "中心A-安全监察" in msg["card_text"]
    for forbidden in ["安全管理员", "操作组长", "复查时间", "（张三）", "（组长甲）"]:
        assert forbidden not in msg["card_text"] and forbidden not in msg["at_text"], f"不应含: {forbidden}"
    assert "@13800000001" in msg["at_text"] and "@13800000002" in msg["at_text"]
    assert msg["at_mobiles"] == ["13800000001", "13800000002"]
    rendered = daily_report.render_at_names(msg["at_text"], msg["mobile_to_name"])
    assert "@张三" in rendered and "@组长甲" in rendered

    # Part B: 若本地存在 config.yaml 与每日报表目录，则校验真实数据完整推送流程（不泄露真实姓名）
    if not os.path.exists("config.yaml"):
        print("7. 每日四块双消息通报 OK（示例数据验证通过，跳过本地报表）")
        return
    cfg = yaml.safe_load(open("config.yaml", encoding="utf-8"))
    report_dir = (cfg.get("report") or {}).get("report_dir") or "每日报表"
    if not os.path.isdir(report_dir):
        print(f"7. 每日四块双消息通报 OK（示例数据验证通过，跳过不存在目录: {report_dir}）")
        return

    real_data = daily_report.load_daily_report(cfg)
    real_msg = daily_report.build_daily_messages(real_data, cfg)
    assert "④ 整改要求" in real_msg["at_text"] and "④ 整改要求" not in real_msg["card_text"]
    assert "今日动作" in real_msg["at_text"] and "今日动作" not in real_msg["card_text"]
    assert "复查时间" not in real_msg["card_text"] and "复查时间" not in real_msg["at_text"]
    for title_word in ["安全管理员", "操作组长", "中心安全监察主管", "中心工艺工程主管"]:
        assert title_word not in real_msg["card_text"] and title_word not in real_msg["at_text"]
    for name in real_msg["mobile_to_name"].values():
        assert f"（{name}）" not in real_msg["at_text"], f"不应重复姓名括号: {name}"

    cfg.setdefault("webhook", {})
    cfg["webhook"]["webhook_url"] = "https://oapi.dingtalk.com/robot/send?access_token=FAKE"
    captured.clear()
    with mock.patch.object(dingtalk_api.SESSION, "post", fake_post):
        main.cmd_daily_push(cfg)

    types = [b.get("msgtype") for _, b in captured]
    assert types == ["actionCard", "markdown"], f"应发送 actionCard + markdown 两条，实际: {types}"
    body = captured[1][1]
    sent_at = body["at"]["atMobiles"]
    assert sent_at == real_msg["at_mobiles"] and len(sent_at) > 0, sent_at
    assert body["markdown"]["text"] == real_msg["at_text"]
    print(
        f"7. 每日四块双消息通报 OK（示例+本地日期 {real_data.daily_date}, 类型: {types}, "
        f"@人数: {len(real_msg['at_mobiles'])}, 缺号: {len(real_msg['missing_phone_names'])}）"
    )


if __name__ == "__main__":
    test_webhook_markdown()
    test_webhook_actioncard()
    test_no_duplicate_at()
    test_push_flow()
    test_data_date()
    test_callback_import()
    test_daily_briefing()
    print("\n全部自测通过")
