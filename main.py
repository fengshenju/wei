#!/usr/bin/env python3
"""
DrissionPage 爬虫项目主入口 - 单据自动化处理版本
"""

import os
import time
import glob
from DrissionPage import Chromium, ChromiumOptions

# 导入配置和工具
from app_config import CONFIG
from utils.util_time import get_random_wait
from utils.llm_utils import extract_data_from_image
from utils.data_manager import DataManager

# 定义提取数据的 Prompt (根据您的单据内容调整)
PROMPT_INSTRUCTION = """
这是一张采购单据。请提取以下信息并以JSON返回：
1. 交付日期 (delivery_date)
2. 采购商名称 (buyer_name)
3. 供应商名称 (supplier_name)
4. 商品明细列表 (items)，包含：款号(style_no), 数量(qty), 单价(price), 单位(unit)
"""


def process_single_bill_rpa(tab, data_json, file_name):
    """
    【核心业务逻辑】
    针对单个单据的 RPA 操作。
    这里拿到的是一整张单据的数据（含表头和明细列表）。
    """
    print(f"\n--- 开始执行 RPA 任务: {file_name} ---")

    # 1. 示例：打印表头信息
    delivery_date = data_json.get('delivery_date', '未知日期')
    buyer = data_json.get('buyer_name', '未知买家')
    print(f"正在处理单据头: 日期[{delivery_date}] - 客户[{buyer}]")

    # TODO: 在这里编写实际的网页操作代码
    # 比如:
    # tab.ele('#date_field').input(delivery_date)
    # tab.ele('#buyer_field').input(buyer)

    # 2. 示例：循环处理明细
    items = data_json.get('items', [])
    print(f"该单据包含 {len(items)} 条明细，开始录入...")

    for index, item in enumerate(items, 1):
        style_no = item.get('style_no')
        qty = item.get('qty')
        print(f"  [{index}/{len(items)}] 录入款号: {style_no}, 数量: {qty}")

        # TODO: 在这里编写明细行的录入代码
        # tab.ele('#add_line_btn').click()
        # tab.ele('#style_input').input(style_no)

        # 模拟一点随机延迟，像真人一样
        time.sleep(get_random_wait(0.5, jitter=0.2))

    print(f"--- RPA 任务完成: {file_name} ---\n")


def main():
    # 1. 初始化数据管理器
    storage_path = CONFIG.get('data_storage_path')

    # 简单的防御性编程：如果没配路径，报错提示
    if not storage_path:
        print("错误: 请在 app_config.py 中配置 'data_storage_path'")
        return
    db_manager = DataManager(storage_path_str=storage_path)

    # 2. 扫描图片文件夹
    source_dir = CONFIG.get('image_source_dir')
    if not source_dir or not os.path.exists(source_dir):
        print(f"错误: 图片源目录不存在，请检查 app_config.py 配置: {source_dir}")
        return

    # 支持扫描 jpg, jpeg, png
    image_files = []
    for ext in ['*.jpg', '*.jpeg', '*.png', '*.JPG', '*.PNG']:
        image_files.extend(glob.glob(os.path.join(source_dir, ext)))

    if not image_files:
        print(f"目录 {source_dir} 下没有找到图片文件。")
        return

    print(f"扫描到 {len(image_files)} 个图片文件，准备开始处理...")

    # 3. 连接浏览器 (只连一次)
    co = ChromiumOptions().set_address(CONFIG['chrome_address'])
    tab = Chromium(addr_or_opts=co).latest_tab

    # 唤醒窗口 (使用之前测试过的稳定组合拳)
    print(">>> 正在唤醒 Chrome...")
    try:
        tab.set.window.normal()
        time.sleep(0.5)
        tab.set.window.full()
    except Exception as e:
        print(f"窗口调整轻微异常(可忽略): {e}")
        tab.set.window.max()

    # 4. 核心循环：遍历文件
    for img_path in image_files:
        file_name = os.path.basename(img_path)

        # --- A. 状态检查 (去重) ---
        if db_manager.is_processed(file_name):
            print(f"[跳过] 文件已处理过: {file_name}")
            continue

            # --- B. LLM 解析 ---
        print(f"[处理中] 正在解析: {file_name} ...")
        # 调用您之前提供的 llm_utils
        parsed_data = extract_data_from_image(img_path, PROMPT_INSTRUCTION)

        # 简单错误检查
        if not parsed_data or 'error' in parsed_data:
            print(f"!!! 解析失败，跳过: {file_name}, 原因: {parsed_data.get('error', '未知')}")
            continue

        # --- C. 数据持久化 ---
        # 只有保存成功了，才代表这个单据算“处理了一半”，可以进 RPA
        # 如果您希望“RPA执行完才算处理完”，可以将这步移到 D 步骤之后
        # 但建议放在这里，防止RPA报错导致数据又得重新花钱解析一遍
        saved = db_manager.save_data(file_name, parsed_data)
        if not saved:
            print("!!! 数据保存失败，中断处理本条")
            continue

        # --- D. 执行 RPA ---
        try:
            process_single_bill_rpa(tab, parsed_data, file_name)
        except Exception as e:
            print(f"!!! RPA 执行出错: {file_name}, 错误: {e}")
            # 注意：如果 RPA 失败了，您可能需要手动删除对应的 json 文件，或者在这里加逻辑删除 json
            # 这样下次才会重试。目前逻辑是只要解析成功就视为 Completed。

    print(">>> 所有任务处理完毕。")
    input("按回车键退出...")


if __name__ == "__main__":
    main()