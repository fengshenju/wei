#!/usr/bin/env python3
"""
DrissionPage 爬虫测试项目主入口
"""

import random
from DrissionPage import Chromium, ChromiumOptions

from config.settings import settings
from utils.util_time import get_random_wait
from app_config import CONFIG
import time


def generate_smooth_trajectory(start_x, start_y, end_x, end_y, steps=30):
    """生成平滑的贝塞尔曲线轨迹"""
    # 添加控制点形成弧形
    ctrl_x = (start_x + end_x) / 2 + random.randint(-30, 30)
    ctrl_y = (start_y + end_y) / 2 + random.randint(-20, 20)

    points = []
    for i in range(steps):
        t = i / (steps - 1)
        # 二次贝塞尔曲线
        x = (1 - t) ** 2 * start_x + 2 * (1 - t) * t * ctrl_x + t ** 2 * end_x
        y = (1 - t) ** 2 * start_y + 2 * (1 - t) * t * ctrl_y + t ** 2 * end_y

        # 添加随机噪声
        x += random.randint(-2, 2)
        y += random.randint(-2, 2)
        points.append((x, y))

    return points


def move_with_trajectorybak(tab, points):
    """按轨迹点分段变速移动"""
    total = len(points)
    for i, (x, y) in enumerate(points):
        # 分段速度：开始快，中间慢，结束快
        if i < total * 0.3:
            duration = 0.05  # 快
        elif i < total * 0.7:
            duration = 0.15  # 慢（故意很慢观察效果）
        else:
            duration = 0.08  # 中等

        tab.actions.move_to((x, y), duration=duration)


def main():
    co1 = ChromiumOptions().set_address(CONFIG['chrome_address'])
    tab1 = Chromium(addr_or_opts=co1).latest_tab

    # --- 修改开始 ---
    print("正在唤醒 Chrome 窗口...")
    try:
        # 1. 先恢复“普通”大小 (这一步是解开最小化/全屏死锁的关键)
        tab1.set.window.normal()

        # 2. 稍微等待一下操作系统的动画反应 (0.2~0.5秒即可)
        time.sleep(0.5)

        # 3. 再执行最大化 (此时窗口已经是激活状态，最大化不会报错)
        # tab1.set.window.max()
        tab1.set.window.full()

    except Exception as e:
        # 如果上面流程由于某种极端情况失败，打印日志并尝试强行最大化
        print(f"窗口状态调整遇到小问题 (可忽略): {e}")
        tab1.set.window.max()
    # --- 修改结束 ---
    # 确保Chrome窗口激活和最大化
    # tab1.set.window.full()

    tab1.wait(get_random_wait(3, jitter=0.2))
    tab1.get(CONFIG['base_url'])
    input("按回车键退出...")


if __name__ == "__main__":
    main()