# 邮件 & 系统漏单审核机器人

本机运行的邮件解析、人工复核和工单比对工具。当前发布基于 2026.09.30-r11 / UI 03.4，提供网页版运行控制台，并保留桌面入口。

## 功能

- 按日期读取邮件、过滤内部邮件，解析正文和 ZIP/RAR 中的表格附件。
- 结构化规则优先提取公司、国家和项目，结合可配置的 LLM 进行识别与复检。
- 三栏人工复核工作台，保留左右拖动、原始证据预览、文本复制及逐条确认。
- 德国 WEEE 与德国电池法独立分块；品牌/品类草稿、确认状态分别保存。
- 同一天重新提取时替换该日提取及复核状态；日历显示全部处理完成的日期。
- 导出确认结果，供第二阶段进行工单查询比对。

## Windows 源码启动

建议使用 Python 3.11 或 3.12，并安装 Microsoft Edge。

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
Copy-Item config.example.yaml config.yaml
```

在本机 `config.yaml` 填写 IMAP、工单地址和账号。LLM 密钥通过 `DEEPSEEK_API_KEY` 环境变量提供，不写入代码或提交到仓库。配置中的域名均为占位符，不能直接用于连接业务系统。

内部邮箱规则通过环境变量设置（英文逗号分隔），必须改成部署方自己的域名：

```powershell
$env:MAIL_INTERNAL_DOMAINS = 'internal.example.invalid'
$env:MAIL_AUDIT_ALIASES = 'audit@internal.example.invalid'
$env:MAIL_INTERNAL_DOMAIN_MARKERS = ''
.\.venv\Scripts\python.exe web_app.py
```

也可双击 `启动网页版.bat`。默认控制台为 http://127.0.0.1:8765/run，人工复核页为 http://127.0.0.1:8765/。桌面界面使用 `启动.bat`。

在控制台导入自己的内部邮箱表、代理邮箱表和项目表。RAR 解析需要系统安装可用的 7-Zip；OCR 依赖需单独安装或配置。仓库不包含 OCR 二进制、模型和浏览器运行时。

## 发布边界

- 只提供脱敏源码、测试与配置模板；不包含真实账号、密码、API 密钥、客户邮件、附件、审核结果及历史缓存。
- 回归样例中的公司、邮箱和代理已替换；业务域名规则可通过上述环境变量恢复到部署方配置。
- 此版本电池分类仍为原有手填方式；随后提出的五种固定下拉选项尚未包含，避免把未完成的修改标成已交付。
- 只监听本机，未实现面向公网的用户认证，不应直接暴露到局域网或公网。
- 启用 LLM 后邮件/附件证据可能发送到所配置的模型服务；上线前确认权限和数据处理要求。
- 人工确认不等于已向工单平台提交业务；工单平台适配与登录有效性需部署方验证。

## 离线测试

```powershell
.\.venv\Scripts\python.exe -m pip install pytest
.\.venv\Scripts\python.exe -B -m pytest tests -q -p no:cacheprovider
```

测试使用临时匿名样例，不需要真实邮箱或工单账号。不要以测试通过替代实际系统联调验收。
