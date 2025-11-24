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
from utils.util_llm import extract_data_from_image
from utils.data_manager import DataManager,load_style_db_from_excel

# 从配置文件获取提示词
PROMPT_INSTRUCTION = CONFIG.get('prompt_instruction', '')

# 假设这是你的本地款号库（从数据库或文件加载）
LOCAL_STYLE_DB = {"T8821", "H2005", "X3002", "D5001"}

# 根据规则确定最终款号
def determine_final_style(json_data, style_db):
    """
    根据业务规则从 candidates 中选出最终款号
    :param json_data: LLM 返回的 JSON
    :param style_db: 本地款号库 (Set集合)
    """
    candidates = json_data.get("style_candidates", [])
    if not candidates:
        return None

    # --- 规则 1: 识别明细中存在的款号 (位置在表格内且在库中) ---
    for item in candidates:
        clean_text = item['text'].strip()
        # 这里直接用 style_db 进行 O(1) 复杂度的极速查找
        if "表格" in item.get('position', '') and clean_text in style_db:
            print(f"✅ 命中规则1 (表格内且在库): {clean_text}")
            return clean_text

    # --- 规则 2: 识别红色字体 (若存在且在库中) ---
    red_candidates = [c for c in candidates if c.get('is_red') == True]
    for item in red_candidates:
        clean_text = item['text'].strip()
        if clean_text in style_db:
            print(f"✅ 命中规则2 (红色字体且在库): {clean_text}")
            return clean_text

    # --- 规则 3: 兜底搜索 (符合 T/H/X/D 开头规律) ---
    for item in candidates:
        text = item['text'].strip().upper()
        # 这里逻辑可根据需求调整：如果不在库里，但长得像，要不要？
        # 下面代码假设：只要长得像就提取，作为最后的保底
        if text.startswith(('T', 'H', 'X', 'D', 'L', 'S', 'F')):  # 我看你csv里还有L/S/F开头
            print(f"✅ 命中规则3 (符合命名规律): {text}")
            return text

    return None

def process_single_bill_rpa(tab,data_json, file_name):
    """
    针对单个单据的完整 RPA 流程
    包括：连接 -> 唤醒窗口 -> 执行业务
    """
    print(f"\n--- [RPA阶段] 开始处理: {file_name} ---")

    # 1. 【新增】在任务内部初始化连接
    # 每次调用这个函数，都会重新获取一次最新的浏览器对象的引用
    # co = ChromiumOptions().set_address(CONFIG['chrome_address'])
    # tab = Chromium(addr_or_opts=co).latest_tab

    # 2. 【新增】把“唤醒/最大化”逻辑移到这里
    # 这意味着：每处理一张单据，都会强制检查一遍窗口状态，确保它在最前面
    print(f"[{file_name}] 正在激活浏览器窗口...")
    try:
        tab.set.window.normal()
        time.sleep(0.5)
        # 强制全屏/最大化
        tab.set.window.full()
    except Exception as e:
        # 兜底方案：如果上面报错，尝试直接最大化
        print(f"窗口调整警告: {e}")
        tab.set.window.max()

    # 3. 开始具体的业务操作 (和之前一样)
    try:
        # 示例：打开目标网页（如果每个单据都要从头开始填，这里可能需要 get 一下）
        # tab.get(CONFIG['base_url'])

        # 填表头...
        delivery_date = data_json.get('delivery_date', '')
        print(f"[{file_name}] 正在录入日期: {delivery_date}")
        # tab.ele('#date').input(delivery_date)

        # 填明细...
        items = data_json.get('items', [])
        for item in items:
            print(f"[{file_name}] 录入明细: {item.get('style_no')}")
            # ...

    except Exception as e:
        print(f"!!! [{file_name}] RPA执行出错: {e}")
        raise e  # 抛出异常，让主循环知道这就出错了

def main():
    # 1. 初始化阶段：加载款号库 (Excel)
    excel_path = CONFIG.get('style_db_path')
    col_name = CONFIG.get('style_db_column', '款式编号')

    # >>> 关键调用 <<<
    LOCAL_STYLE_DB = load_style_db_from_excel(excel_path, col_name)

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
        final_style = determine_final_style(parsed_data, LOCAL_STYLE_DB)
        parsed_data['final_selected_style'] = final_style
        print(f"最终判定款号: {final_style}")

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