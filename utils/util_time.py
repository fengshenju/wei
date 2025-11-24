# 文件名: utils.py
import random


def get_random_wait(base_time, jitter=0.3):
    """
    生成带有随机抖动的等待时间
    :param base_time: 基础等待时间（秒）
    :param jitter: 抖动幅度（百分比），默认 0.3 表示上下浮动 30%
    :return: 随机时间
    """
    if base_time <= 0:
        return 0

    # 计算最小和最大时间
    min_time = base_time * (1 - jitter)
    max_time = base_time * (1 + jitter)

    # 生成随机浮点数
    wait_time = random.uniform(min_time, max_time)
    # 【修复核心】强制兜底，绝不返回负数
    if wait_time < 0:
        wait_time = 0
    # 保留两位小数，显得更自然
    return round(wait_time, 2)