# MDAC All-in-One Console (MDAC 全自动中央控制台)

这是一个为马来西亚入境卡 (MDAC) 自动化处理而设计的集成控制台。它结合了 Playwright 网页自动化、Telegram 机器人数据采集、OCR 护照识别以及 Excel 并发管理，旨在大幅提升 MDAC 的处理效率。

## 🚀 核心功能

*   **双标签页 UI 控制台**:
    *   **MDAC 控制页**: 一键启动自动化填表，实时监控 Excel 状态。
    *   **Telegram 监听页**: 开启机器人监听，自动从聊天中提取护照信息。
*   **智能 Telegram 机器人**:
    *   自动接收用户发送的护照照片。
    *   **内存 OCR**: 使用 RapidOCR 在内存中直接处理图片，不留本地缓存，保护隐私。
    *   **MRZ 自动解析**: 自动提取护照号、姓名、国籍、出生日期及有效期。
*   **高效自动化执行**:
    *   **Playwright 驱动**: 模拟真实浏览器行为，稳定填表。
    *   **dddocr 验证码识别**: 自动处理滑块验证码。
*   **并发安全管理**:
    *   基于 `filelock` 的 Excel 管理器，确保 UI 刷新、机器人写入与自动化执行之间的数据一致性。
    *   **自动日期格式化**: 严格执行 `d-Mon-yy` (如 `1-May-26`) 的日期格式要求。

## 🛠️ 安装要求

在运行之前，请确保你的环境中已安装以下组件：

1.  **Python 3.10+**
2.  **Chrome 浏览器** (建议使用系统原生安装的 Chrome)
3.  **依赖库安装**:
    ```bash
    pip install playwright tkinterdnd2 ddddocr rapidocr_onnxruntime telebot openpyxl filelock
    playwright install chromium
    ```

## ⚙️ 配置说明

首次运行程序会生成 `mdac_settings.json` 配置文件，你需要在 UI 界面或文件中配置以下内容：

*   **Excel 路径**: 存放 MDAC 数据的 `.xlsx` 文件路径。
*   **Telegram Token**: 你的 Telegram Bot API Token。
*   **固定信息**: 如住宿地址、联系方式等默认填表信息。

## 📖 使用流程

1.  **准备数据**: 在 Excel 中维护客户基本信息。
2.  **开启监听**: 在控制台切换到 "Telegram 监听" 标签页，点击“开启监听”。
3.  **采集信息**: 客户向机器人发送护照照片，机器人会自动解析并写入 Excel，状态标记为 `缺少日期`。
4.  **手动确认**: 管理员在 Excel 中补全日期并将状态改为 `PENDING`。
5.  **开始填表**: 在控制台点击“开始执行 MDAC”，程序将自动处理所有 `PENDING` 状态的任务。

## ⚠️ 注意事项

*   **护照识别**: 请确保拍摄的护照照片清晰，尤其是底部的 MRZ 区域。
*   **日期处理**: 程序会自动处理跨世纪的出生年份（如 80 对应 1980，20 对应 2020）。
*   **文件锁定**: 运行期间请勿手动打开 Excel 文件，以免造成写入冲突。

---
*Developed by AI网站生成师 - 粉肠哥*
