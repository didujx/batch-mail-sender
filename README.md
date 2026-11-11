# batch-mail-sender

[![](https://img.shields.io/badge/powered_by-dsh-4D6BFE?style=flat-square&logo=deepseek&logoColor=white)](https://github.com/deepseek-ai/deepseek-harness)

按名单**逐人人格化**批量发送邮件：A 列邮箱 + B 列内容 → 每人一封、正文各不相同的邮件，一次跑完。

不是群发。每封信的 `To:` 只有当前这一个收件人，彼此不可见；正文取该行自己的数据。

## 典型场景

| 场景 | A 列 | B 列 | 正文示例 |
|---|---|---|---|
| 门票码/邀请码 | 邮箱 | `Hw47@073` | `baoxi你好，现在发送门票码给你Hw47@073，请查收。` |
| 报价单 | 邮箱 | 公司名 | `{数据}的报价单已更新，请查收。` |
| 成绩/工号通知 | 邮箱 | 工号 | `{姓名}你好，你的工号是{数据}。` |

## 快速开始

```bash
pip install openpyxl
copy config.example.ini config.ini     # 然后填写服务器、账号、模板
python send_emails.py --dry-run        # 预演：只打印，不发信
python send_emails.py --no-dry-run --limit 3   # 先发 3 封验收
python send_emails.py --no-dry-run     # 全量
```

名单表格（`收件人名单.xlsx`，第 1 行表头）：

| 邮箱 | 门票码 |
|---|---|
| baoxi@example.com | Hw47@073 |
| shimeiqi@example.com | Hw61@817 |

## 占位符

| 占位符 | 含义 |
|---|---|
| `{姓名}` / `{name}` | 姓名列；未配置姓名列时取邮箱 @ 前缀 |
| `{数据}` / `{value}` | B 列的值 |
| `{邮箱}` / `{email}` | 收件人邮箱 |
| `{序号}` | 表格行号 |

## 安全与可靠性

- **默认只预演**，不加 `--no-dry-run` 绝不出网。
- **行级校验**：邮箱格式错 / B 列为空 / 名单内重复 → 跳过并说明原因。
- **断点续发**：已发记录存 `send_state.json`，中断后重跑只补剩下的。
- **模板变更保护**：主题或正文模板一改，旧的已发记录自动作废，避免漏发。
- **逐行日志**：`send_log_时间戳.csv` 记录每封的成功/失败。
- **失败重试**：`--retry-failed` 只补发上次失败的收件人。
- **节奏控制**：`interval` 秒间隔，降低被风控/进垃圾箱的概率。

## 兼容性

- 发信：任意标准 SMTP（QQ、163、腾讯/阿里企业邮、Outlook 365、Gmail、内部 Exchange）。
- 加密：`starttls`(587) / `none`(25，内网) / `ssl`(465)；STARTTLS 失败自动退回明文。
- 认证：`auto` 依次尝试 `域\\账号` 与 `账号@域`，Exchange 场景实测可用。
- 名单：`.xlsx` / `.xlsm` / `.csv`（UTF-8 或 GBK）。

## 作为 DSH skill 使用

把整个目录放到任一处即被自动识别：

- 项目级：`<项目>/.dsh/skills/batch-mail-sender/SKILL.md`
- 用户级：`<DSH Home>/skills/batch-mail-sender/SKILL.md`

之后对它说“用 batch-mail-sender 帮我把这张表发出去”即可。
skill 名必须保持 kebab-case，与目录名一致。

## 常见报错

| 报错 | 原因 | 处理 |
|---|---|---|
| `501 5.1.7 Invalid address` | MAIL FROM 只有账号名 | 已内置 `resolve_sender()` 补全为 `账号@域` |
| `535 5.7.3 Authentication failed` | 认证名/密码不对 | 认 Exchange 用 `域名\账号`；普通邮箱用授权码 |
| `SSL: WRONG_VERSION_NUMBER` | 465 端口不支持隐式 SSL | 改 587+starttls 或 25+none |
| `550 5.7.1 Client does not have permissions` | 发件地址非本人邮箱 | 改 `from_address` |

## 安全声明

`config.ini` 含明文密码，已被 `.gitignore` 排除，请勿提交。发完建议清空密码并删除
`send_state.json` / `send_log_*.csv` 中的收件人信息。
