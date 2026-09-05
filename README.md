# MDAC Slider Script

本仓库当前仅保留 MDAC 网页滑块验证处理脚本，不再包含原有的表单自动填写、Excel 状态管理、Gmail PIN 获取、Telegram 通知、桌面界面或旧测试代码。

## 文件

- `slider_script.py`：基于 Playwright 页面对象的 MDAC canvas 滑块识别与拖动逻辑。
- `requirements.txt`：滑块识别所需依赖。

## 调用方式

`solve_mdac_slider(page, log_func=print, max_retries=3)` 接收一个已经打开 MDAC 页面并定位到滑块的 Playwright `Page` 对象，返回布尔值表示验证是否成功。

示例：

```python
from slider_script import solve_mdac_slider

success = solve_mdac_slider(page)
```

调用方负责创建 Playwright 页面、打开 MDAC 页面并填写其他表单字段。本脚本只处理 canvas 图片识别、距离计算、模拟拖动和验证结果检查。

## 安装

```bash
pip install -r requirements.txt
```
