# MDAC All-in-One Console (MDAC 全自动中央控制台)

这是一个为马来西亚入境卡 (MDAC) 自动化处理而设计的集成控制台。它结合了 Playwright 网页自动化、Telegram 机器人数据采集、OCR 护照识别、**Gmail PIN 码自动获取**以及 Excel 并发管理，旨在大幅提升 MDAC 的处理效率。

## 🚀 核心功能

*   **四标签页 UI 控制台**:
    *   **MDAC 控制页**: 一键启动自动化填表，实时监控 Excel 状态。
    *   **Telegram 监听页**: 开启机器人监听，自动从聊天中提取护照信息。
    *   **Gmail PIN 获取页**: 自动从 Gmail 收集移民局发送的 PIN 码并匹配回填。
    *   **护照 PDF MRZ 页**: 使用独立 Telegram Bot 批量处理 PDF 每页护照，并回传进度、失败页和重复跳过页。
*   **智能 Telegram 机器人**:
    *   自动接收用户发送的护照照片。
    *   **内存 OCR**: 使用 PaddleOCR 在内存中直接处理图片，不留本地缓存，保护隐私。
    *   **MRZ 自动解析**: 自动提取护照号、姓名、国籍、出生日期及有效期。
*   **Gmail 自动化 (NEW)**:
    *   **IMAP 实时监控**: 定时扫描来自 `mdac@imi.gov.my` 的邮件。
    *   **Message-ID 去重**: 记录已处理邮件，防止重复抓取，支持人工提前查阅。
    *   **智能匹配**: 基于护照号精准匹配 Excel 记录，自动回填 PIN 码并更新状态为 `COMPLETED`。
    *   **异常报警**: 若邮件中的护照号在 Excel 中缺失，自动通过 Telegram 推送详细报警。
*   **高效自动化执行**:
    *   **Playwright 驱动**: 模拟真实浏览器行为，稳定填表。
    *   **ddddocr 验证码识别**: 自动处理滑块验证码。
*   **并发安全管理**:
    *   基于 `filelock` 的 Excel 管理器，确保 UI 刷新、机器人写入、邮件回填与自动化执行之间的数据一致性。
    *   **自动日期格式化**: 严格执行 `d-Mon-yy` (如 `1-May-26`) 的日期格式要求。

## 🛠️ 安装要求

在运行之前，请确保你的环境中已安装以下组件：

1.  **Python 3.10+**
2.  **Chrome 浏览器** (建议使用系统原生安装的 Chrome)
3.  **依赖库安装**:
    ```bash
    pip install playwright ddddocr pyTelegramBotAPI openpyxl filelock pillow pymupdf paddleocr paddlepaddle
    playwright install chromium
    ```

## ⚙️ 配置说明

首次运行程序会生成 `mdac_settings.json` 配置文件，你需要在 UI 界面或文件中配置以下内容：

*   **Excel 路径**: 存放 MDAC 数据的 `.xlsx` 文件路径。
*   **Telegram Token**: 你的 Telegram Bot API Token。
*   **Gmail 配置**: 需开启 Google 账号的“两步验证”并生成 **16 位应用专用密码**。
*   **固定信息**: 如住宿地址、联系方式等默认填表信息。

## 📖 使用流程

1.  **准备数据**: 在 Excel 中维护客户基本信息。
2.  **开启监听**: 在控制台切换到 "Telegram 监听" 标签页，点击“开启监听”。
3.  **采集信息**: 客户向机器人发送护照照片，机器人会自动解析并写入 Excel，状态标记为 `缺少日期`。
4.  **手动确认**: 管理员在 Excel 中补全日期并将状态改为 `PENDING`。
5.  **开始填表**: 在控制台点击“开始执行 MDAC”，程序将自动完成注册。
6.  **PDF 批处理**: 在 Tab 4 选择同格式 Excel、填写独立 Bot Token，启动后发送 PDF；程序逐页识别并回传进度，完成后返回失败页和重复跳过页。
7.  **自动收码**: 开启 "Gmail PIN 获取" 监听，脚本会自动抓取移民局发来的 PIN 码并填入 Excel H 列，状态自动更新为 `COMPLETED`。

## 🔍 MRZ OCR 模块

MRZ 识别已独立拆分到 `mrz/` 包，Telegram 图片和 PDF 页面共用同一套 TD3 Passport 流程：预处理、多个 ROI/增强版本 OCR、两行重建、ICAO 7-3-1 校验位、字段感知 OCR 纠错和 confidence 评分。只有两行各 44 字符、五类 checksum 全部通过且 confidence 达到 90 分时，结果才会进入原有 Excel 写入流程；未通过时保持原有识别失败路径，不会自动写入 MDAC 数据。

主要模块包括 `preprocess.py`、`ocr.py`、`parser.py`、`checksum.py`、`corrector.py`、`engine.py` 和 `legacy.py`。`legacy.py` 保留现有 `MRZParser.parse_image()` 返回协议，因此 Playwright、Excel、Telegram、PDF、Gmail 和 MDAC 提交流程无需改写。`personal_number` 会保留在结构化识别结果和诊断信息中，但当前不新增 Excel 列。

运行 MRZ 回归测试：

```bash
python3 -m unittest discover -s tests -v
```

## ⚠️ 注意事项

*   **护照识别**: 请确保拍摄的护照照片清晰，尤其是底部的 MRZ 区域。
*   **Gmail 密码**: 必须使用 Google **应用专用密码**，普通登录密码无效。
*   **文件锁定**: 运行期间请勿手动打开 Excel 文件，以免造成写入冲突。

---
*Developed by AI网站生成师 - 粉肠哥*
