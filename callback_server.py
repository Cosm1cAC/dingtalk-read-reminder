# -*- coding: utf-8 -*-
"""卡片按钮回调监听（Stream 模式，无需公网 IP / 域名）。

常驻运行后，卡片上"收到/回复"按钮的点击事件会实时推到这里：
  - 收到：往群里发一条 "👍 xxx 已收到"
  - 回复：往群里发一条回复引导

运行:  python callback_server.py
"""
import json
import logging
from datetime import datetime

import dingtalk_stream

import yaml
from dingtalk_api import DingTalkClient

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s")
LOG = logging.getLogger("card-callback")

# 不同版本 SDK 的卡片回调 topic 略有差异，两个都注册，收到啥处理啥
CARD_CALLBACK_TOPICS = [
    "/v1.0/card/instance/callback",
    "/v1.0/card/instances/callback",
]


def _dig(d, *keys, default=None):
    """在多层嵌套 dict 里按顺序找 key。"""
    for k in keys:
        if isinstance(d, dict):
            d = d.get(k)
        else:
            return default
    return d if d is not None else default


class CardCallbackHandler(dingtalk_stream.CallbackHandler):
    def __init__(self, logger=None, config=None, rest_client=None):
        super().__init__(logger)
        self.config = config or {}
        self.rest_client = rest_client

    async def process(self, callback: dingtalk_stream.CallbackMessage):
        data = callback.data or {}
        LOG.info("收到卡片回调事件: %s", json.dumps(data, ensure_ascii=False)[:2000])

        content = _dig(data, "content", default={}) or {}
        # 按钮自定义值：卡片模板里按钮回调值配 {"action": "received"} / {"action": "reply"}
        value = _dig(content, "value", default={}) or _dig(content, "inputValue", default={}) or {}
        if isinstance(value, str):
            try:
                value = json.loads(value)
            except (ValueError, TypeError):
                value = {"action": value}
        action = str(_dig(value, "action", default="") or "")
        sender_id = (_dig(data, "senderStaffId", default="")
                     or _dig(data, "dtUserId", default="")
                     or _dig(data, "userId", default="") or "未知用户")
        conversation_id = (_dig(data, "openConversationId", default="")
                           or _dig(self.config, "open_conversation_id"))

        if not conversation_id or not self.rest_client:
            LOG.warning("缺少 openConversationId 或钉钉客户端配置，忽略该事件")
            return dingtalk_stream.AckMessage(
                dingtalk_stream.AckMessage.STATUS_OK, "OK")

        name = self._lookup_name(sender_id)
        try:
            if action == "received":
                self.rest_client.send_markdown(
                    conversation_id, "收到确认",
                    f"👍 **{name}** 已收到 {datetime.now().strftime('%H:%M:%S')}")
            elif action == "reply":
                self.rest_client.send_markdown(
                    conversation_id, "回复提醒",
                    f"💬 {name} 想说点什么？请直接在群里回复或 @管理助手 提建议～")
            else:
                LOG.info("未识别的按钮动作: %s", action)
        except Exception as e:  # 回调处理不能抛异常，否则钉钉会反复重推
            LOG.exception("处理按钮回调失败: %s", e)
        return dingtalk_stream.AckMessage(dingtalk_stream.AckMessage.STATUS_OK, "OK")

    def _lookup_name(self, user_id: str) -> str:
        """优先用通讯录接口查姓名，失败则回退显示 userid。"""
        try:
            headers = {"x-acs-dingtalk-access-token": self.rest_client.get_token()}
            from dingtalk_api import SESSION
            resp = SESSION.get(
                f"https://api.dingtalk.com/v1.0/contact/users/{user_id}",
                headers=headers, timeout=5)
            if resp.ok:
                return resp.json().get("nick") or user_id
        except Exception:
            pass
        return user_id


def load_config(path="config.yaml") -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def main():
    cfg = load_config()
    app = cfg["dingtalk"]
    if not app.get("app_key") or "你的" in str(app.get("app_key")):
        raise SystemExit("请先在 config.yaml 里填写钉钉应用凭证（app_key/app_secret/robot_code）")

    rest_client = DingTalkClient(app["app_key"], app["app_secret"], app["robot_code"])
    credential = dingtalk_stream.Credential(app["app_key"], app["app_secret"])
    client = dingtalk_stream.DingTalkStreamClient(credential)

    handler = CardCallbackHandler(logger=LOG, config=cfg, rest_client=rest_client)
    for topic in CARD_CALLBACK_TOPICS:
        client.register_callback_handler(topic, handler)

    LOG.info("按钮回调监听已启动（常驻运行，按钮才有响应）...")
    client.start_forever()


if __name__ == "__main__":
    main()
