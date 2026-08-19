# MDAC Auto V2

这是一个面向马来西亚数字入境卡（MDAC）处理的 Windows 桌面控制台。当前版本保留两条核心流程：使用 Excel 驱动 MDAC 网页自动化，以及从 Gmail 自动提取移民局发送的 PIN 并回填 Excel。

## 当前功能

| 功能区域 | 作用 |
|---|---|
| MDAC 控制 | 读取 Excel 中状态为 `PENDING` 的客户，自动填写 MDAC 表单、处理滑块并提交；支持测试模式、暂停和停止 |
| Gmail PIN 获取 | 定时读取来自 `mdac@imi.gov.my` 的邮件，提取姓名、护照号和 PIN，匹配 Excel 后写入 H 列，并将状态改为 `COMPLETED` |

项目已移除 Telegram 图片识别、PDF MRZ 批处理、PaddleOCR 和 MRZ 解析包。客户资料应通过当前采用的其他业务流程准备并写入 Excel，之后由 MDAC 控制流程继续处理。

## Excel 数据约定

MDAC 控制流程从活动工作表第 2 行开始读取资料。主要字段如下：

| 列 | 字段 | 用途 |
|---|---|---|
| A | Name | MDAC 姓名 |
| B | Passport | 护照号码 |
| C | Date of Birth | 出生日期 |
| D | Sex | `男` 或 `女` |
| E | Date of Expiry | 护照有效期 |
| F | Date of Arrival | 抵达日期 |
| G | Date of Departure | 离境日期 |
| H | PIN | Gmail PIN 回填位置 |
| I | Status | 控制处理状态 |
| J | Last Check | 状态更新时间 |
| K | Remark | 备注、错误或人工检查说明 |
| L | Nationality | 国籍 |

Tab 1 只会自动处理 I 列为 `PENDING` 的记录。常见状态包括 `PROCESSING`、`TESTED`、`REGISTERED`、`MANUAL_CHECK`、`ERROR` 和 `COMPLETED`。

## MDAC 控制设置

Tab 1 的固定资料设置包括 Excel 路径、Email、手机号、区域代码、航班或船名、住宿地址和邮编。`region` 已改为面板可设置变量，默认值为 `60`，并会保存到 `mdac_settings.json`。面板中的区域代码应填写 MDAC 网页下拉选项对应的 value。

网页日期字段要求使用 `DD/MM/YYYY` 格式，例如：

```text
18/10/1981
23/04/2036
```

Excel 日期会先转换为 Python 日期对象，再在提交网页前转换为 MDAC 要求的 `DD/MM/YYYY` 字符串。测试模式下程序不会点击 Submit，便于先在浏览器中检查资料。

## Gmail PIN 获取

Gmail 模块通过 IMAP 定时扫描来自 `mdac@imi.gov.my` 的邮件，并使用 Message-ID 文件避免重复处理。需要在界面配置 Gmail 地址、Google 应用专用密码、检查间隔，以及用于异常报警的 Telegram Bot Token 和 Chat ID。

当邮件中的护照号能在 Excel B 列找到时，PIN 会写入 H 列，I 列状态更新为 `COMPLETED`，J 列记录更新时间。如果找不到对应护照号，程序会通过配置的 Telegram 报警 Bot 提醒人工核查。

## 安装

需要 Python 3.10 或更高版本、Windows Chrome 浏览器以及 Playwright 运行环境：

```bash
pip install -r requirements.txt
playwright install chromium
```

程序使用可见浏览器运行 MDAC 自动化。请确保 Chrome 浏览器可正常启动，并避免在自动化写入期间手动打开同一个 Excel 文件，以减少文件锁冲突。

## 运行

```bash
python main_console.py
```

首次运行后，在 MDAC 控制面板填写固定资料并选择 Excel 文件；在 Gmail PIN 面板填写 Gmail 和报警配置。修改配置后可点击保存，启动 MDAC 自动化时程序也会保存当前面板设置。

## 测试

运行兼容性测试：

```bash
python -m unittest discover -s tests -v
```

## 当前范围

本项目当前不包含护照图片识别、iPhone Live Text 文本解析、Telegram 图片接收、PDF 护照批处理或 MRZ OCR。后续如采用新的客户资料输入方案，应确保最终结果写入 Excel，并按照 MDAC 控制流程要求准备完整字段及 `PENDING` 状态。

---

*Developed for 粉肠哥*
