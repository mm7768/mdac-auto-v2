"""MDAC canvas slider solver.

The solver expects an existing Playwright page already opened on the MDAC form.
"""

import base64
import io
import random
import time

import ddddocr
from PIL import Image


def solve_mdac_slider(page, log_func=print, max_retries=3):
    """Solve the MDAC canvas slider on an existing Playwright page.

    Args:
        page: A Playwright Page containing the MDAC slider.
        log_func: Callable receiving diagnostic messages.
        max_retries: Maximum number of slider attempts.

    Returns:
        True when the slider succeeds; otherwise False.
    """
    for attempt in range(max_retries):
        log_func(f"正在尝试第 {attempt + 1} 次滑块验证...")
        try:
            page.wait_for_selector("canvas", timeout=10000)
            page.wait_for_timeout(1500)

            bg_base64 = page.evaluate(
                "document.querySelectorAll('canvas')[0].toDataURL('image/png')"
            )
            block_base64 = page.evaluate(
                "document.querySelectorAll('canvas')[1].toDataURL('image/png')"
            )

            bg_bytes = base64.b64decode(bg_base64.split(",")[1])
            block_bytes = base64.b64decode(block_base64.split(",")[1])

            block_img = Image.open(io.BytesIO(block_bytes))
            bbox = block_img.getbbox()
            img_start_x = bbox[0] if bbox else 0

            detector = ddddocr.DdddOcr(det=False, ocr=False, show_ad=False)
            result = detector.slide_match(
                block_bytes, bg_bytes, simple_target=True
            )
            distance = result["target"][0] - img_start_x

            scale_info = page.evaluate(
                """
                () => {
                    const canvas = document.querySelectorAll('canvas')[0];
                    return {
                        internal: canvas.width,
                        display: canvas.getBoundingClientRect().width
                    };
                }
                """
            )
            scale = (
                scale_info["display"] / scale_info["internal"]
                if scale_info["internal"]
                else 1
            )
            final_distance = distance * scale

            def generate_track(total_distance):
                track = []
                current = 0
                steps = random.randint(30, 40)
                for index in range(1, steps + 1):
                    progress = index / steps
                    ease_progress = (
                        1 if progress == 1 else 1 - (2 ** (-10 * progress))
                    )
                    move = total_distance * ease_progress
                    step_move = move - current
                    current = move
                    track.append(step_move)
                return track

            slider_handle = page.locator(".slider").first
            box = slider_handle.bounding_box()
            if not box:
                raise RuntimeError("未找到滑块手柄位置")

            handle_start_x = box["x"] + box["width"] / 2
            handle_start_y = box["y"] + box["height"] / 2

            slider_handle.hover()
            page.mouse.down()
            page.wait_for_timeout(random.randint(100, 200))

            actual_move = final_distance - 14
            current_x = handle_start_x
            for step_x in generate_track(actual_move):
                current_x += step_x
                page.mouse.move(
                    current_x,
                    handle_start_y + random.uniform(-1.5, 1.5),
                )
                time.sleep(random.uniform(0.01, 0.02))

            page.wait_for_timeout(random.randint(300, 500))
            page.mouse.up()
            page.wait_for_timeout(2000)

            success = page.evaluate(
                """
                () => document.querySelector('.sliderContainer') !== null
                    && document.querySelector('.sliderContainer')
                        .classList.contains('sliderContainer_success')
                """
            )
            if success:
                log_func("滑块验证成功！")
                return True

            log_func(f"第 {attempt + 1} 次验证失败，等待滑块重置...")
            page.wait_for_timeout(2500)

        except Exception as error:
            log_func(f"自动滑块处理发生错误: {error}")
            try:
                page.mouse.up()
            except Exception:
                pass
            page.wait_for_timeout(2000)

    log_func(f"连续 {max_retries} 次滑块验证失败！")
    return False


__all__ = ["solve_mdac_slider"]
