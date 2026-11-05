# -*- coding: utf-8 -*-
"""钉钉 API：accessToken、群消息发送、机器人所在群查询、群自定义机器人 webhook。"""
import base64
import hashlib
import hmac
import json
import time
import urllib.parse

import requests

# 钉钉接口国内直连即可。trust_env=False 让请求无视系统代理/环境变量代理，
# 避免机器上残留的失效代理（如 192.168.137.104:7897）导致请求超时。
SESSION = requests.Session()
SESSION.trust_env = False

TOKEN_URL = "https://api.dingtalk.com/v1.0/oauth2/accessToken"
SEND_URL = "https://api.dingtalk.com/v1.0/robot/groupMessages/send"
LIST_GROUPS_URL = "https://api.dingtalk.com/v1.0/robot/chatGroups"

MSG_KEY_MARKDOWN = "sampleMarkdown"
MSG_KEY_CARD = "templateInteractiveCard"


class DingTalkClient:
    def __init__(self, app_key: str, app_secret: str, robot_code: str):
        self.app_key = app_key
        self.app_secret = app_secret
        self.robot_code = robot_code
        self._token = None
        self._token_expire = 0

    # ---------- token ----------
    def get_token(self) -> str:
        if self._token and time.time() < self._token_expire - 60:
            return self._token
        resp = SESSION.post(TOKEN_URL, json={
            "appKey": self.app_key,
            "appSecret": self.app_secret,
        }, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        if "accessToken" not in data:
            raise RuntimeError(f"获取 accessToken 失败: {data}")
        self._token = data["accessToken"]
        self._token_expire = time.time() + int(data.get("expireIn", 7200))
        return self._token

    # ---------- 发消息 ----------
    def send_group_message(self, open_conversation_id: str, msg_key: str,
                           msg_param: dict, at_mobiles: list = None) -> dict:
        """企业机器人发群消息。

        文档: https://open.dingtalk.com/document/orgapp/the-robot-sends-a-group-message
        msg_key:
          - sampleMarkdown: markdown 文本
          - templateInteractiveCard: 卡片平台模板卡片
        """
        body = {
            "robotCode": self.robot_code,
            "openConversationId": open_conversation_id,
            "msgKey": msg_key,
            "msgParam": json.dumps(msg_param, ensure_ascii=False),
        }
        if at_mobiles:
            body["at"] = {"atMobiles": at_mobiles, "isAtAll": False}
        headers = {
            "x-acs-dingtalk-access-token": self.get_token(),
            "Content-Type": "application/json",
        }
        resp = SESSION.post(SEND_URL, json=body, headers=headers, timeout=10)
        data = resp.json()
        if not resp.ok or (isinstance(data, dict) and data.get("code") not in (None, "ok", "0")):
            raise RuntimeError(f"发送消息失败: HTTP {resp.status_code} {data}")
        return data

    def send_markdown(self, open_conversation_id: str, title: str, text: str,
                      at_mobiles: list = None) -> dict:
        return self.send_group_message(
            open_conversation_id, MSG_KEY_MARKDOWN,
            {"title": title, "text": text}, at_mobiles=at_mobiles)

    def send_card(self, open_conversation_id: str, card_template_id: str,
                  card_param_map: dict, out_track_id: str = None) -> dict:
        """发送卡片平台模板互动卡片。

        card_template_id: 卡片平台模板编号（templateId，形如 xxx.yyy）
        收到/回复按钮的回调事件由 callback_server.py（Stream 模式）处理。
        """
        msg_param = {
            "cardTemplateId": card_template_id,
            "cardData": {"cardParamMap": card_param_map},
        }
        if out_track_id:
            msg_param["outTrackId"] = out_track_id
        return self.send_group_message(open_conversation_id, MSG_KEY_CARD, msg_param)

    # ---------- 查群 ----------
    def list_chat_groups(self, max_results: int = 50) -> list:
        """查询机器人所在的群（用于拿 openConversationId）。"""
        headers = {"x-acs-dingtalk-access-token": self.get_token()}
        resp = SESSION.get(LIST_GROUPS_URL, headers=headers, timeout=10,
                            params={"maxResults": max_results})
        data = resp.json()
        if not resp.ok:
            raise RuntimeError(f"查询群列表失败: HTTP {resp.status_code} {data}")
        return data.get("result", data.get("records", []))


WEBHOOK_SEND_URL = "https://oapi.dingtalk.com/robot/send"


def send_via_webhook(webhook_url: str, secret: str, title: str, text: str,
                     at_mobiles: list = None) -> dict:
    """群自定义机器人（零权限模式）发送 markdown 消息。

    webhook_url: 添加自定义机器人时给的 https://oapi.dingtalk.com/robot/send?access_token=xxx
    secret: 加签密钥（SEC 开头）；未开启加签则传空串
    at_mobiles: 要 @ 的手机号，同时需在 markdown 文本中包含 "@手机号" 才会显示 @ 样式
    注意：自定义机器人不支持互动卡片/按钮，也没有回调能力。
    """
    if secret:
        timestamp = str(round(time.time() * 1000))
        sign_str = f"{timestamp}\n{secret}"
        sign = base64.b64encode(
            hmac.new(secret.encode(), sign_str.encode(), hashlib.sha256).digest())
        webhook_url += f"&timestamp={timestamp}&sign={urllib.parse.quote_plus(sign)}"

    if at_mobiles:
        # 仅当正文中还没有对应 "@手机号" 时才在文末补 @文本，避免出现两排重复 @
        missing = [m for m in at_mobiles if f"@{m}" not in text]
        if missing:
            text += "\n\n" + " ".join(f"@{m}" for m in missing)

    body = {
        "msgtype": "markdown",
        "markdown": {"title": title, "text": text},
        "at": {"atMobiles": at_mobiles or [], "isAtAll": False},
    }
    resp = SESSION.post(webhook_url, json=body, timeout=10)
    data = resp.json()
    if not resp.ok or data.get("errcode") != 0:
        raise RuntimeError(f"webhook 发送失败: HTTP {resp.status_code} {data}")
    return data


def send_actioncard_via_webhook(webhook_url: str, secret: str, title: str, text: str,
                                single_title: str = None, single_url: str = None) -> dict:
    """webhook 发送 actionCard：整条消息渲染成卡片样式（标题+markdown 正文）。

    不带按钮时只渲染标题和正文；要加按钮需给 single_title/single_url（只能跳转网页，
    没有点击回调能力）。注意：actionCard 不支持 @人，@提醒消息请继续用 send_via_webhook。
    """
    if secret:
        timestamp = str(round(time.time() * 1000))
        sign_str = f"{timestamp}\n{secret}"
        sign = base64.b64encode(
            hmac.new(secret.encode(), sign_str.encode(), hashlib.sha256).digest())
        webhook_url += f"&timestamp={timestamp}&sign={urllib.parse.quote_plus(sign)}"

    card = {"title": title, "text": text, "btnOrientation": "0"}
    if single_title and single_url:
        card["singleTitle"] = single_title
        card["singleURL"] = single_url

    body = {"msgtype": "actionCard", "actionCard": card}
    resp = SESSION.post(webhook_url, json=body, timeout=10)
    data = resp.json()
    if not resp.ok or data.get("errcode") != 0:
        raise RuntimeError(f"webhook actionCard 发送失败: HTTP {resp.status_code} {data}")
    return data
