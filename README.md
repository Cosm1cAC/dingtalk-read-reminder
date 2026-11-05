# 管理驾驶舱「已读情况提醒」钉钉机器人

每天定时向钉钉群推送「昨日员工消息已读情况」统计，并单独 @ 未使用人员。
支持两种运行模式：**零权限的群自定义机器人（webhook）** 和 **企业内部应用（互动卡片）**。

纯 Python 脚本，无服务端依赖，数据源是本地 Excel。

---

## 功能

- **定时推送**：每天固定时间发两条消息
  1. **主统计卡片**（actionCard 卡片样式）：整体使用率 + 未使用/已全部读完/部分已读 的人数
  2. **未使用人员 @提醒**：逐人列出，可点击 @，被 @ 的人收到钉钉通知
- **统计口径**
  - 使用率 = 已读条数 / 应读条数
  - 未使用：使用率 < 50%（进入 @ 提醒名单）
  - 已全部读完：已读 == 应读
  - 部分已读：≥ 50% 但未读完（按完成率降序）
  - 整体使用率 = Σ已读 / Σ应读
- **姓名 → 手机号永久缓存**：SQLite 存储姓名-手机号映射，Excel 里不写手机号也能自动补全，不必每次重填
- **数据日期跟随数据本身**：Excel 内嵌 meta 日期，凌晨补发也不会把标题日期算错

## 两种模式对比

| | webhook 模式（零权限） | app 模式（完整版） |
|---|---|---|
| 前置条件 | 群内添加「自定义」机器人即可，**无需开发权限** | 需企业管理员分配开发者权限/凭证 |
| 消息样式 | actionCard 卡片样式（无红绿配色、无按钮） | 卡片平台互动卡片（可还原配色 + 「收到/回复」按钮） |
| @人 | ✅（按手机号） | ✅（可按手机号或 userid） |
| 定时推送 | ✅ | ✅ |
| 按钮交互 | ❌ | ✅（Stream 模式回调，无需公网 IP） |

> 无开发权限时**不要尝试「快捷应用」**——它只是把一个现成网页挂到工作台入口，不下发 AppKey/AppSecret，与机器人能力无关。

## 项目结构

| 文件 | 作用 |
|---|---|
| `main.py` | 主入口：`make-template` / `list-groups` / `push` / `daemon` |
| `update_data.py` | 每日数据更新：从原始导出表筛指定日期 → 写统计表 → 交叉验证 |
| `excel_stats.py` | 读取 Excel、计算使用率与分类、数据日期 meta 读写 |
| `card_builder.py` | 组装卡片变量、markdown 文本、@ 提醒文本 |
| `dingtalk_api.py` | 钉钉 API 客户端：token、发消息、actionCard、webhook 加签、查群 |
| `contacts.py` | SQLite 姓名-手机号缓存（`data/contacts.db`） |
| `callback_server.py` | Stream 模式监听卡片按钮（仅 app 模式需要，常驻） |
| `config.yaml` | 凭证与配置（**不入库**，见 `config.example.yaml`） |
| `selftest.py` | 本地自测，不连钉钉 |

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

国内网络/代理异常时：

```bash
python -m pip install -i https://mirrors.aliyun.com/pypi/simple/ -r requirements.txt
```

> 代码内所有请求走 `requests.Session(trust_env=False)`，无视系统代理，避免残留的失效代理导致超时。

### 2. 配置

```bash
cp config.example.yaml config.yaml
```

填 `config.yaml` 的两个部分（webhook 有值时优先走 webhook 模式）：

**webhook 模式**：群设置 → 机器人 → 添加机器人 → 自定义（Webhook 接入）→ 安全设置选「加签」

```yaml
webhook:
  webhook_url: "https://oapi.dingtalk.com/robot/send?access_token=xxxx"
  secret: "SECxxxx"
```

**app 模式**（可选，需要开发者权限）：open-dev.dingtalk.com 创建企业内部应用，机器人消息接收模式选 **Stream**，填 `app_key` / `app_secret` / `robot_code`，`python main.py list-groups` 查 `openConversationId` 填入；如需「收到/回复」按钮，再在卡片平台建模板填 `card_template_id`。

### 3. 准备数据

```bash
python main.py make-template     # 生成 data/当前统计.xlsx 模板
```

五列固定表头：`姓名 | 手机号 | 职位 | 应读条数 | 已读条数`。
手机号可留空——`contacts.py` 会自动按姓名从 `data/contacts.db` 补齐。

### 4. 推送验证

```bash
python main.py push
```

### 5. 定时

Windows 计划任务（推荐）：

```bat
schtasks /Create /TN "钉钉已读提醒" /TR "cmd /c cd /d D:\AI\钉钉机器人 && python main.py push" /SC DAILY /ST 17:00
```

或常驻进程自己管时间：`python main.py daemon`（读 `config.yaml` 的 `push_time`）。

app 模式的按钮回调需常驻：`python callback_server.py`。

---

## 每日使用流程

原始数据来自「部长助手」员工明细导出（含每人两行：日期类型「日」= 当天、「月」= 当月累计）。

```bash
# 一条命令：筛指定日期（只取「日」类型）→ 写 data/当前统计.xlsx（含 meta 日期）→ 逐人交叉验证
python update_data.py <员工明细.xlsx> 2026-11-04

python main.py push        # 推送两条消息到群
```

`update_data.py` 不传日期时，自动取表内最新的「日」类型日期。

### 接入其他省区 / 其他系统导出的表

`update_data.py` 顶部集中定义了源表列号，格式不同只改这一处：

```python
COL_DATE, COL_TYPE = 0, 1                  # 日期、类型（日/月）
COL_NAME, COL_ROLE, COL_MSG, COL_READ = 8, 9, 10, 11   # 姓名、职位、消息量、已读量
TYPE_DAY = "日"                            # 只取当天，排除「月」累计行
```

不想换数据源也行：用 `python main.py make-template` 生成模板，手工填五列（姓名/手机号/职位/应读/已读），再 `push`。两条路互不影响。

## 多团队共用

每个团队各自部署一份即可，互不干扰，也没有共享服务端：

1. 各群自己在「群设置 → 机器人」里建**自定义机器人**，各自的 webhook + 加签密钥
2. 各自一份 `config.yaml`、各自的 `data/`（Excel + `contacts.db`）
3. 定时任务各配各的（Windows 计划任务名不同即可）
4. 代码完全不需要改，只要 `data.data_file` 指向自己的表

## 关于 @ 不到人

webhook 模式**只能按手机号 @**，且该手机号必须是对方**钉钉账号实际绑定的号码**。
人事系统里登记的「电话」字段常常不是绑定号（典型表现：消息里名字是黑色、不是蓝色可点击）。
遇到这种情况只有两条路：

1. 让对方在「钉钉 → 我 → 头像 → 个人信息」里查看绑定手机号（最快）
2. 管理员在 oa.dingtalk.com 通讯录中查看/导出真实绑定号

即使没有手机号，消息仍会点名显示，不会报错——只是对方收不到通知。

## 常见问题

- **发消息失败 / 无权限**：webhook 地址或加签密钥错；app 模式下则为应用未发布、机器人未拉进群、群与应用不同组织。
- **消息挤成一坨**：钉钉 markdown 单个 `\n` 不换行，必须用空行（`\n\n`）分隔，代码已按此处理。
- **@ 显示成黑色不可点**：手机号不是对方绑定号，见上一节。
- **Excel 写入报 Permission denied**：文件正被 Excel/WPS 打开；或换 `config.yaml` 的 `data.data_file` 指向新文件。
- **标题日期不对**：确保用 `update_data.py` 生成数据（会写入 meta 日期）；手写 Excel 无 meta 时退回「今天-1」。

## 安全说明

以下文件含敏感信息，已在 `.gitignore` 中排除，**不要提交**：

- `config.yaml`（webhook 地址、加签密钥、AppSecret）
- `data/`（真实员工姓名、手机号、数据库、原始导出表）
- `20??-??-??.md`（工作日志）

## License

MIT，见 [LICENSE](LICENSE)。可自由使用、修改、二次分发，保留版权声明即可。
