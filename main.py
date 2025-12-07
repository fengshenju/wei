#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DrissionPage 爬虫项目主入口 - 重构版本
"""
import os
import glob
import asyncio
import datetime
from concurrent.futures import ThreadPoolExecutor
from DrissionPage import Chromium, ChromiumOptions

# ---------------------------------------------------------
# 导入配置和工具
# ---------------------------------------------------------
try:
    from app_config import CONFIG
    from utils.data_manager import DataManager, load_style_db_with_cache, load_supplier_db_with_cache
    from utils.report_generator import collect_result_data, update_html_report
    from services.image_service import ImageService
    from services.rpa_service import RpaService
except ImportError as e:
    print(f"!!! 导入模块失败: {e}")
    print("请检查项目结构是否包含 app_config.py, services/ 和 utils/ 目录")
    exit(1)


def process_complete_rpa(browser, result, rpa_service):
    """处理完整的RPA+LLM匹配流程并收集报告数据"""
    try:
        if result['success']:
            file_name = result.get('file_name', '')
            parsed_data = result.get('parsed_data', {})
            img_path = result.get('img_path', '')

            # 执行完整RPA+LLM流程
            # 注意: process_single_bill_rpa 现在是 RpaService 的实例方法
            match_prompt, match_result, original_records, retry_count = rpa_service.process_single_bill_rpa(
                browser, parsed_data, file_name, img_path
            )

            # 收集报告数据
            return collect_result_data(
                image_name=file_name,
                parsed_data=parsed_data,
                final_style=result.get('final_style', ''),
                match_prompt=match_prompt,
                match_result=match_result,
                original_records=original_records,
                image_path=result.get('img_path', ''),
                retry_count=retry_count,
                failure_reason=result.get('failure_reason', '')
            )
        else:
            # 解析失败，生成失败报告
            return collect_result_data(
                image_name=result.get('file_name', ''),
                parsed_data=result.get('parsed_data', {}),
                final_style=result.get('final_style', ''),
                match_prompt="",
                match_result=None,
                original_records=[],
                image_path=result.get('img_path', ''),
                retry_count=1,
                failure_reason=result.get('failure_reason', '')
            )
    except Exception as e:
        print(f"!!! 完整流程异常: {result.get('file_name', '')}, 错误: {e}")
        return collect_result_data(
            image_name=result.get('file_name', ''),
            parsed_data=result.get('parsed_data', {}),
            final_style=result.get('final_style', ''),
            match_prompt="",
            match_result=None,
            original_records=[],
            image_path=result.get('img_path', ''),
            retry_count=1,
            failure_reason=str(e)
        )


async def async_main():
    # 1. 初始化阶段：加载款号库 (Excel + 缓存加速)
    excel_path = CONFIG.get('style_db_path')
    col_name = CONFIG.get('style_db_column', '款式编号')
    LOCAL_STYLE_DB = load_style_db_with_cache(excel_path, col_name)
    
    # 1.1 加载月结供应商目录 (Excel + 缓存加速)
    supplier_excel_path = CONFIG.get('supplier_db_path')
    LOCAL_SUPPLIER_DB = load_supplier_db_with_cache(supplier_excel_path) if supplier_excel_path else set()
    print(f">>> 月结供应商目录加载完成，共 {len(LOCAL_SUPPLIER_DB)} 个供应商")

    # 2. 初始化数据管理器
    storage_path = CONFIG.get('data_storage_path')
    if not storage_path:
        print("错误: 请在 app_config.py 中配置 'data_storage_path'")
        return
    db_manager = DataManager(storage_path_str=storage_path)

    # 3. 扫描图片文件夹
    source_dir = CONFIG.get('image_source_dir')
    if not source_dir or not os.path.exists(source_dir):
        print(f"错误: 图片源目录不存在，请检查 app_config.py 配置: {source_dir}")
        return

    image_files = []
    for ext in ['*.jpg', '*.jpeg', '*.png', '*.JPG', '*.PNG']:
        image_files.extend(glob.glob(os.path.join(source_dir, ext)))

    if not image_files:
        print(f"目录 {source_dir} 下没有找到图片文件。")
        return

    print(f"扫描到 {len(image_files)} 个图片文件，准备开始处理...")

    # 4. 连接浏览器 (只连一次)
    co = ChromiumOptions().set_address(CONFIG['chrome_address'])
    browser = Chromium(addr_or_opts=co)

    # 5. 初始化 RPA 服务
    rpa_service = RpaService()

    # 6. 阶段1：并发图片解析
    print("\n=== 阶段1：并发图片解析 ===")
    parsing_concurrency = CONFIG.get('image_parsing_concurrency', 3)
    semaphore = asyncio.Semaphore(parsing_concurrency)

    async def parse_with_semaphore(img_path):
        async with semaphore:
            loop = asyncio.get_event_loop()
            with ThreadPoolExecutor() as executor:
                # 调用 ImageService 的静态方法
                return await loop.run_in_executor(
                    executor, 
                    ImageService.parse_single_image, 
                    img_path, 
                    db_manager, 
                    LOCAL_STYLE_DB, 
                    LOCAL_SUPPLIER_DB
                )

    parse_tasks = [parse_with_semaphore(img_path) for img_path in image_files]
    parse_results = await asyncio.gather(*parse_tasks, return_exceptions=True)

    # 过滤有效结果
    valid_results = [r for r in parse_results if r is not None and not isinstance(r, Exception)]
    success_results = [r for r in valid_results if r.get('success', False)]

    print(f"解析完成: 总计 {len(image_files)} 张，有效 {len(valid_results)} 张，成功 {len(success_results)} 张")

    # 分离成功和失败的结果
    failed_results = [r for r in valid_results if not r.get('success', False)]
    print(f"款号解析失败: {len(failed_results)} 张")

    # 7. 阶段2：并发RPA+LLM匹配 (只处理成功解析款号的结果)
    print("\n=== 阶段2：并发RPA+LLM匹配 ===")
    rpa_concurrency = CONFIG.get('rpa_concurrency', 3)
    rpa_semaphore = asyncio.Semaphore(rpa_concurrency)

    async def process_with_semaphore(result):
        async with rpa_semaphore:
            loop = asyncio.get_event_loop()
            with ThreadPoolExecutor() as executor:
                return await loop.run_in_executor(
                    executor, 
                    process_complete_rpa, 
                    browser, 
                    result, 
                    rpa_service
                )

    # 只对成功解析款号的结果进行RPA处理
    process_tasks = [process_with_semaphore(result) for result in success_results]
    rpa_results = await asyncio.gather(*process_tasks, return_exceptions=True)

    # 为失败的结果生成报告数据
    failed_report_results = []
    for failed_result in failed_results:
        report_data = collect_result_data(
            image_name=failed_result.get('file_name', ''),
            parsed_data=failed_result.get('parsed_data', {}),
            final_style=failed_result.get('final_style', ''),
            match_prompt="",
            match_result=None,
            original_records=[],
            image_path=failed_result.get('img_path', ''),
            retry_count=failed_result.get('parsed_data', {}).get('retry_count', 1),
            failure_reason=failed_result.get('failure_reason') or failed_result.get('error', '款号没有解析到')
        )
        failed_report_results.append(report_data)

    # 合并所有结果
    final_results = rpa_results + failed_report_results

    print(f"RPA+LLM匹配完成: {len(final_results)} 张")

    # 8. 统一生成报告
    print("\n=== 生成最终报告 ===")
    report_path = CONFIG.get('report_output_path')
    if report_path:
        try:
            # 生成带时间戳的报告文件名
            timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            report_filename = f'report_{timestamp}.html'

            # 确保目录存在
            if not os.path.exists(report_path):
                 # 如果路径是文件而不是目录，取其父目录，或者根据配置创建目录
                 # 这里假设 report_output_path 是一个目录路径
                 os.makedirs(report_path, exist_ok=True)
            
            report_file = os.path.join(report_path, report_filename)

            # 逐个更新报告
            for result_data in final_results:
                if not isinstance(result_data, Exception) and result_data:
                    update_html_report(os.path.dirname(report_file), result_data['image_name'], result_data)

            # 重命名最终生成的report.html为带时间戳的文件名
            # update_html_report 默认会生成在目录下叫 report.html
            default_report = os.path.join(os.path.dirname(report_file), 'report.html')
            if os.path.exists(default_report):
                if os.path.exists(report_file):
                    os.remove(report_file)
                os.rename(default_report, report_file)

            print(f"✅ 报告已生成: {report_file}")
        except Exception as e:
            print(f"⚠️ 报告生成失败: {e}")

    print(">>> 所有任务处理完毕。")


def main():
    """主入口函数"""
    asyncio.run(async_main())


if __name__ == "__main__":
    main()
