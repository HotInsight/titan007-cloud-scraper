# Titan007 单场数据云端抓取器

每天泰国时间 11:00 在 GitHub Actions 中运行，抓取“即时比分 / Crow* / 亚+欧 / 单场”，生成十列 Excel，并通过 Gmail 发送附件。

## 必需的 Repository Secrets

在仓库 `Settings → Secrets and variables → Actions` 中新增：

- `GMAIL_USERNAME`：作为发件人的完整 Gmail 地址
- `GMAIL_APP_PASSWORD`：开启 Google 两步验证后生成的 16 位应用专用密码
- `EMAIL_TO`：接收 Excel 的邮箱地址（可以与发件地址相同）

不要把邮箱密码直接写入任何代码文件。

## 手动测试

进入仓库 `Actions → Titan007 Daily Scraper → Run workflow`，保留“发送结果邮件”为启用状态，然后运行。

成功时：收到 Excel 邮件，并可在该次运行页面底部的 Artifacts 下载备份。

失败时：会尝试发送失败通知；运行页面底部会出现 `titan007-failure-*`，内含日志，并可能包含失败页面的 PNG 和 HTML。

## 定时规则

工作流使用 `timezone: Asia/Bangkok`，每天泰国时间 11:00 触发。GitHub 高负载时可能有几分钟延迟。
