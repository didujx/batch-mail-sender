---
name: batch-mail-sender
description: 按名单逐人人格化批量发邮件（一人一封，正文/称呼取自表格对应行）。当用户有一列邮箱+一列内容（门票码、邀请码、报价、通知正文等）想一次性发出、且不想用群发（抄送/密送所有人可见或内容相同）时使用；也用于内部 Exchange/SMTP 发信调试（501 Invalid address、535 认证失败、STARTTLS 等）。Use when a user wants to send one individualized email per row of an Excel/CSV list, or debug SMTP/Exchange delivery failures.
---

# 按行个性化批量发邮件（一人一封）

把“A 列邮箱 + B 列内容”的表格，变成每人一封、正文各不相同的邮件，一次跑完。

## 必须先确认的四件事（缺一个就别发）

1. **名单文件**：.xlsx 或 .csv；A 列邮箱、B 列内容，第 1 行是表头。
2. **SMTP 服务器 / 端口 / 加密方式**：邮箱商（QQ/163/企业邮）用 465+SSL，
   内部 Exchange 常见 587+STARTTLS、25+明文。
3. **登录账号**：普通邮箱用完整地址 + **授权码**（不是网页登录密码）；
   Exchange 用 域账号，认证名往往是 \`域名\\账号\`，不是裸账号名。
4. **主题与正文模板**：正文里用占位符引用表格数据，不要手写每人一句。

## 执行流程（严格按顺序）

### 步骤 1：准备脚本与配置

把本 skill 目录下的 \`send_emails.py\` 复制到工作目录，复制 \`config.example.ini\` 为 \`config.ini\` 并填写。
若 Config 里带 \`<skill_resources>\` 目录，直接用其中的文件；否则按 [参考：核心实现](#参考核心实现) 自建。

\`\`\`bash
pip install openpyxl        # 只用 .csv 时可跳过
\`\`\`

### 步骤 2：模板占位符

| 占位符 | 取值 |
|---|---|
| \`{姓名}\` / \`{name}\` | \`name_column\` 指定的姓名列；未配置该列时**自动取邮箱 @ 前缀** |
| \`{数据}\` / \`{value}\` | B 列的值（门票码、邀请码等） |
| \`{邮箱}\` / \`{email}\` | 收件人地址 |
| \`{序号}\` | 表格行号 |
| `{列表}` | 可选：多个码的编号清单；不写则自动附在正文末尾 |

例：\`{姓名}你好，现在发送门票码给你{数据}，请查收。\` + \`baoxi@corp.com\` + \`Hw47@073\`
→ \`baoxi你好，现在发送门票码给你Hw47@073，请查收。\`

注意：数据里若含 \`@\`（如 \`Hw47@073\`），Python 侧原样输出；个别严格策略的
Exchange 可能把它当邮箱地址重写或拒投。**先 \`--limit 1\` 发一封自查**。

### 步骤 2.5：一人多码怎么发（默认合并）

同一个邮箱在名单里出现多次 = 一个人有多个码。**默认合并成一封**，不要发多封：

```
baoxi@corp.com  Hw47@073
baoxi@corp.com  Hw61@817
```

→ 一封邮件：

```
baoxi你好，现在发送门票码给你Hw47@073，请查收。

清单：
1. Hw47@073
2. Hw61@817

共 2 个，请核对。
```

规则：
- 单码用户正文与旧版逐字一致，不会出现“清单”字样
- 同人同码的重复行自动去重
- 只有对方系统强制“一码一邮件”时才用 `--split`（或 `multi_code = split`），
  此时标题自动加 `(1/2)` `(2/2)`，让收件人知道不是重发
- **不要**对同一个人连发两封同标题邮件 —— 会被当成重复/垃圾而丢掉第二个码

### 步骤 3：预演（必须做）

\`\`\`bash
python send_emails.py --dry-run
\`\`\`

核对三件事：主题是否正确、正文里占位符是否都替换干净、被跳过的行是否合理
（邮箱格式错 / B 列为空），并确认“[已发]”列出的人是断点记录，不是漏发。

### 步骤 4：小批量验收，然后全量

\`\`\`bash
python send_emails.py --no-dry-run --limit 3     # 先发 3 封给自己/真实收件人
python send_emails.py --no-dry-run               # 全量，已发过的人不会重发
python send_emails.py --no-dry-run --retry-failed # 只补发失败的人
\`\`\`

## 排错手册

| 现象 | 原因 | 处理 |
|---|---|---|
| \`501 5.1.7 Invalid address ... 'jiangxin'\` | MAIL FROM 只有账号名，没有域名 | 发件地址补成 \`账号@域\`（脚本 \`resolve_sender()\` 已处理） |
| \`535 5.7.3 Authentication failed\` | 认证名写法不对 | Exchange 试 \`域名\\账号\`；普通邮箱用授权码而非登录密码 |
| \`SSL: WRONG_VERSION_NUMBER\`（连 465） | 服务器不支持隐式 SSL | 改 587 + STARTTLS 或 25 + 明文 |
| STARTTLS 报证书错误 | 内网自签证书 | 关闭主机名校验（内网限定），脚本已内置 |
| \`550 5.7.1 Client does not have permissions\` | 发件地址不是本账号的邮箱 | 修 \`from_address\` |
| 大量进垃圾箱 | 频率太高/内容雷同 | 加大 \`interval\`；正文务必逐人不同 |
| 改了模板后担心漏发 | 旧的“已发”记录会挡住 | 脚本按主题+模板指纹自动作废旧记录 |

## 硬性规则

- **绝不跳过预演**：先 \`--dry-run\`，再 \`--limit 3\`，最后全量。
- **同人不多发**：同一个邮箱的多个码合并进一封，绝不连发多封同标题邮件。
- **一人一封**：\`sendmail\` 的收件人列表只放当前一个人，禁止把多个地址塞进一个列表冒充“批量”。
- **凭据不入库**：密码只写在本地 config.ini，发送完提醒用户清空，并确保 .gitignore 覆盖。
- **不猜服务商参数**：端口/加密/认证写法拿不准时，先探测再发——连接 587/25/465 看 banner，
  用一次登录确认认证名，不要直接对真实收件人试错。

## 参考：核心实现

发送循环（其余校验、日志、断点续发见 \`send_emails.py\`）：

\`\`\`python
sender = resolve_sender(cfg)                     # 必须是完整邮箱，否则 Exchange 501
server = make_sender(cfg)                        # STARTTLS 失败自动退回明文
for idx, email, name, value in todo:
    subject_i = render(subject, email, idx, name, value)
    body_i = render(template, email, idx, name, value)
    msg = build_message(cfg, email, subject_i, body_i, sender)
    server.sendmail(sender, [email], msg.as_string())   # 只投一个人
    save_state(email)                                   # 断点续发
    time.sleep(interval)                                # 控制节奏
\`\`\`

[![](https://img.shields.io/badge/powered_by-dsh-4D6BFE?style=flat-square&logo=deepseek&logoColor=white)](https://github.com/deepseek-ai/deepseek-harness)
