#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DrissionPage 爬虫项目主入口 - 单据自动化处理版本
"""
import datetime
import os
import glob
import asyncio
from concurrent.futures import ThreadPoolExecutor
from DrissionPage import Chromium, ChromiumOptions
import shutil

try:
    from app_config import CONFIG
    from utils.data_manager import DataManager, load_style_db_with_cache, load_supplier_db_with_cache, load_material_deduction_db_with_cache, apply_material_deduction
    from utils.report_generator import collect_result_data, update_html_report
    from utils.util_llm import extract_data_from_image, extract_data_from_image_dmx
    from rpa_executor import RPAExecutor  # 导入拆分后的执行器
except ImportError as e:
    print(f"!!! 导入模块失败: {e}")
    print("请检查项目结构是否包含 app_config.py, utils/ 和 rpa_executor.py")
    exit(1)

# 从配置文件获取提示词
PROMPT_INSTRUCTION = CONFIG.get('prompt_instruction', '')


def archive_success_image(img_path, parsed_data):
    """
    [新增] 成功图片归档处理
    格式：采购助理_品牌_供应商_合计金额_序号.png
    """
    try:
        archive_dir = CONFIG.get('success_archive_path')
        if not archive_dir or not os.path.exists(img_path):
            return

        if not os.path.exists(archive_dir):
            os.makedirs(archive_dir)

        # 1. 获取字段
        purchaser = parsed_data.get('purchaser', '未知助理').strip()
        supplier = parsed_data.get('supplier_name', '未知供应商').strip()
        amount = str(parsed_data.get('total_amount', '0')).replace(',', '').strip()

        # 2. 识别品牌 (复用业务规则)
        style = parsed_data.get('final_selected_style', '').upper()
        brand = "其他品牌"
        if style.startswith('T'):
            brand = "CHENXIHE"
        elif style.startswith('X'):
            brand = "CHENXIHE抖音"
        elif style.startswith('H'):
            brand = "SUNONEONE"
        elif style.startswith('D'):
            brand = "SUNONEONE抖音"

        # 3. 生成序号 (基于当日归档目录下的文件数量 + 1)
        # 获取当天日期字符串，确保跨天不重复(虽然文件名没要求带日期，但为了计算当日递增，建议简单计算目录下总数)
        # 这里为了简单且保证文件名唯一性，直接计算该目录下现有文件数 + 1
        current_count = len(glob.glob(os.path.join(archive_dir, "*.png"))) + 1

        # 4. 组装文件名
        new_name = f"{purchaser}_{brand}_{supplier}_{amount}_{current_count}.png"
        # 处理文件名中可能存在的非法字符
        new_name = new_name.replace('/', '_').replace('\\', '_')

        dest_path = os.path.join(archive_dir, new_name)

        # 5. 复制文件
        shutil.copy2(img_path, dest_path)
        print(f"✅ [归档成功] 图片已归档至: {new_name}")

    except Exception as e:
        print(f"⚠️ [归档失败] {e}")

def should_use_dmx_for_date_check(delivery_date_str):
    """
    检查交付日期是否异常，如果与当前日期相差超过阈值天数则返回True
    """
    try:
        if not delivery_date_str: return False
        if len(delivery_date_str) == 10 and delivery_date_str.count('-') == 2:
            delivery_date = datetime.datetime.strptime(delivery_date_str, '%Y-%m-%d').date()
        else:
            for fmt in ['%Y/%m/%d', '%m/%d/%Y', '%d/%m/%Y', '%Y.%m.%d']:
                try:
                    delivery_date = datetime.datetime.strptime(delivery_date_str, fmt).date()
                    break
                except ValueError:
                    continue
            else:
                return False

        current_date = datetime.datetime.now().date()
        date_diff = abs((delivery_date - current_date).days)
        threshold_days = CONFIG.get('delivery_date_threshold_days', 7)

        if date_diff > threshold_days:
            print(
                f">>> 交付日期异常检测: 当前日期 {current_date}, 交付日期 {delivery_date}, 差值 {date_diff} 天 > 阈值 {threshold_days} 天")
            return True
        else:
            return False
    except Exception as e:
        print(f">>> 交付日期解析失败: {delivery_date_str}, 错误: {e}")
        return False


def smart_clean_with_db(text, style_db):
    """基于款号库的智能款号清理"""
    if not text: return text
    if text in style_db: return text

    for db_style in style_db:
        if db_style in text and len(db_style) > 3:
            count = text.count(db_style)
            if count > 1:
                return db_style
            elif count == 1:
                return db_style

    ocr_char_map = {'J': '1', 'O': '0', 'I': '1', 'S': '5', 'Z': '2'}
    for wrong_char, correct_char in ocr_char_map.items():
        if wrong_char in text:
            corrected_text = text.replace(wrong_char, correct_char)
            if corrected_text in style_db:
                return corrected_text

    text_clean_hash = text.rstrip('#') if text.endswith('#') else text
    text_clean_kuan = text.rstrip('款') if text.endswith('款') else text

    cleaning_strategies = [
        text_clean_hash,
        text_clean_kuan,
        text.rstrip('款型式号'),
        text.split('款')[0].strip(),
        text.split('型')[0].strip(),
        text.split('式')[0].strip(),
        text.replace(' ', ''),
        text.upper(),
        text.upper().rstrip('款型式号'),
        text.upper().rstrip('#'),
        text.upper().rstrip('款'),
    ]

    for candidate in cleaning_strategies:
        if candidate and candidate in style_db:
            return candidate

    return text.rstrip('款型式号').strip()


def normalize_supplier_name(extracted_name, supplier_db):
    """供应商名称标准化匹配"""
    if not extracted_name or not supplier_db: return None
    extracted_name = str(extracted_name).strip()
    clean_extracted = extracted_name.replace("商行", "").replace("有限公司", "").replace("布行", "").strip()
    if not extracted_name: return None

    # 1. 精确匹配
    if extracted_name in supplier_db:
        return extracted_name

    # 2. 清理后缀匹配
    if clean_extracted in supplier_db:
        return clean_extracted

    # 3. 包含匹配
    for db_name in supplier_db:
        if extracted_name in db_name: return db_name
        if len(db_name) >= 2 and db_name in extracted_name: return db_name
        if len(db_name) >= 2 and db_name in clean_extracted: return db_name

    # 4. OCR字符修正（风/凤互换）
    corrected_name = clean_extracted
    if "风" in clean_extracted:
        corrected_name = clean_extracted.replace("风", "凤")
    elif "凤" in clean_extracted:
        corrected_name = clean_extracted.replace("凤", "风")
    
    if corrected_name != clean_extracted:
        if corrected_name in supplier_db:
            print(f"   [OCR修正] {extracted_name} -> {corrected_name}")
            return corrected_name
        # 修正后包含匹配
        for db_name in supplier_db:
            if len(db_name) >= 2 and db_name in corrected_name: return db_name
            if len(corrected_name) >= 2 and corrected_name in db_name: return db_name

    def levenshtein_distance(s1, s2):
        if len(s1) < len(s2): return levenshtein_distance(s2, s1)
        if len(s2) == 0: return len(s1)
        previous_row = list(range(len(s2) + 1))
        for i, c1 in enumerate(s1):
            current_row = [i + 1]
            for j, c2 in enumerate(s2):
                insertions = previous_row[j + 1] + 1
                deletions = current_row[j] + 1
                substitutions = previous_row[j] + (c1 != c2)
                current_row.append(min(insertions, deletions, substitutions))
            previous_row = current_row
        return previous_row[-1]

    best_match = None
    min_distance = float('inf')
    for standard_name in supplier_db:
        dist = levenshtein_distance(extracted_name, standard_name)
        length = max(len(extracted_name), len(standard_name))
        if length <= 3:
            threshold = 1.0
        else:
            threshold = length * 0.4

        if dist <= threshold and dist < min_distance:
            min_distance = dist
            best_match = standard_name

    if best_match:
        print(f"   [匹配] 模糊命中(距离{min_distance}): {extracted_name} -> {best_match}")
    return best_match


def determine_final_style(json_data, style_db):
    """根据业务规则从 candidates 中选出最终款号"""
    candidates = json_data.get("style_candidates", [])
    if not candidates: return None

    for item in candidates:
        clean_text = item['text'].strip()
        if "表格" in item.get('position', '') and clean_text in style_db:
            cleaned = smart_clean_with_db(clean_text, style_db)
            print(f"✅ 命中规则1 (表格内且在库): {cleaned}")
            return cleaned

    red_candidates = [c for c in candidates if c.get('is_red') == True]
    for item in red_candidates:
        clean_text = item['text'].strip()
        if clean_text in style_db:
            cleaned = smart_clean_with_db(clean_text, style_db)
            print(f"✅ 命中规则2 (红色字体且在库): {cleaned}")
            return cleaned

    for item in candidates:
        text = item['text'].strip().upper()
        if text.startswith(('T', 'H', 'X', 'D', 'L', 'S', 'F')):
            cleaned = smart_clean_with_db(text, style_db)
            print(f"✅ 命中规则3 (符合命名规律): {cleaned}")
            return cleaned

    return None


def parse_single_image(img_path, db_manager, LOCAL_STYLE_DB, LOCAL_SUPPLIER_DB, LOCAL_MATERIAL_DEDUCTION_DB=None, deduction_suppliers=None, LOCAL_PURCHASER_MAPPING=None):
    """解析单张图片"""
    file_name = os.path.basename(img_path)
    try:
        if db_manager.is_processed(file_name):
            if CONFIG.get('test_use_existing_json', False):
                print(f"[测试模式] 复用已有数据: {file_name}")
                parsed_data = db_manager.load_data(file_name)
                if parsed_data:
                    final_style = parsed_data.get('final_selected_style')
                    return {
                        'file_name': file_name,
                        'img_path': img_path,
                        'success': True,
                        'parsed_data': parsed_data,
                        'final_style': final_style,
                        'failure_reason': parsed_data.get('failure_reason', '')
                    }
            print(f"[跳过] 文件已处理过: {file_name}")
            return None

        print(f"[处理中] 正在解析: {file_name} ...")
        supplier_list_str = "、".join(LOCAL_SUPPLIER_DB) if LOCAL_SUPPLIER_DB else "无已知供应商，请自行识别"
        try:
            current_prompt = PROMPT_INSTRUCTION.format(known_suppliers=supplier_list_str)
        except KeyError as e:
            current_prompt = PROMPT_INSTRUCTION.replace("{known_suppliers}", supplier_list_str)

        if CONFIG.get('use_llm_image_parsing', True):
            parsed_data = extract_data_from_image_dmx(img_path, current_prompt)
        else:
            # 测试桩数据
            parsed_data = {
                'buyer_name': '素本服饰',
                'delivery_date': '2025-11-22',
                'final_selected_style': 'H1635A-B',
                'items': [{'price': 18.0, 'qty': 11.0, 'raw_style_text': '00101020105', 'unit': '并'}],
                'style_candidates': [{'is_red': False, 'position': '表格备注栏', 'text': 'H1635A-B'}],
                'supplier_name': '杭州楼国忠辅料(辅料城仓库店)'
            }

        final_style = determine_final_style(parsed_data, LOCAL_STYLE_DB)
        parsed_data['final_selected_style'] = final_style

        supplier_name = normalize_supplier_name(parsed_data.get('supplier_name'), LOCAL_SUPPLIER_DB)

        if supplier_name:
            parsed_data['supplier_name'] = supplier_name
            # 添加采购助理信息
            if LOCAL_PURCHASER_MAPPING:
                purchaser = LOCAL_PURCHASER_MAPPING.get(supplier_name, "不存在")
                parsed_data['purchaser'] = purchaser
                print(f"匹配供应商: {supplier_name}, 采购助理: {purchaser}")
            else:
                print(f"匹配供应商: {supplier_name}")
        else:
            print(f"!!! 首次供应商匹配失败: [{parsed_data.get('supplier_name')}]，正在使用 DMX 进行重试识别...")
            dmx_parsed_data = extract_data_from_image_dmx(img_path, current_prompt, 0)
            dmx_success = False
            if dmx_parsed_data and 'error' not in dmx_parsed_data:
                retry_supplier_name = normalize_supplier_name(dmx_parsed_data.get('supplier_name'), LOCAL_SUPPLIER_DB)
                if retry_supplier_name:
                    print(f"✅ DMX 重试挽回成功! 匹配到供应商: {retry_supplier_name}")
                    parsed_data = dmx_parsed_data
                    parsed_data['supplier_name'] = retry_supplier_name
                    # 添加采购助理信息
                    if LOCAL_PURCHASER_MAPPING:
                        purchaser = LOCAL_PURCHASER_MAPPING.get(retry_supplier_name, "不存在")
                        parsed_data['purchaser'] = purchaser
                    final_style = determine_final_style(parsed_data, LOCAL_STYLE_DB)
                    parsed_data['final_selected_style'] = final_style
                    dmx_success = True
                else:
                    print(f"!!! DMX 重试后供应商依然无法匹配: [{dmx_parsed_data.get('supplier_name')}]")

            if not dmx_success:
                return {
                    'error': '供应商匹配失败',
                    'file_name': file_name,
                    'img_path': img_path,
                    'original_supplier': parsed_data.get('supplier_name'),
                    'failure_reason': '供应商匹配失败(含DMX重试)',
                    'parsed_data': parsed_data,
                    'final_style': parsed_data.get('final_selected_style', '')
                }

        delivery_date = parsed_data.get('delivery_date', '')
        if delivery_date and should_use_dmx_for_date_check(delivery_date):
            print(f">>> 交付日期异常: {delivery_date}，使用DMX重新校验...")
            dmx_date_data = extract_data_from_image_dmx(img_path, current_prompt, 0)
            if dmx_date_data and 'error' not in dmx_date_data:
                print(">>> DMX日期校验返回成功，更新数据")
                temp_supplier = normalize_supplier_name(dmx_date_data.get('supplier_name'), LOCAL_SUPPLIER_DB)
                if temp_supplier:
                    parsed_data = dmx_date_data
                    parsed_data['supplier_name'] = temp_supplier
                    # 添加采购助理信息
                    if LOCAL_PURCHASER_MAPPING:
                        purchaser = LOCAL_PURCHASER_MAPPING.get(temp_supplier, "不存在")
                        parsed_data['purchaser'] = purchaser
                    final_style = determine_final_style(parsed_data, LOCAL_STYLE_DB)
                    parsed_data['final_selected_style'] = final_style
                    parsed_data['used_dmx_for_date_check'] = True
                else:
                    print("   ⚠️ DMX日期校验数据的供应商无法匹配，仅更新日期字段")
                    parsed_data['delivery_date'] = dmx_date_data.get('delivery_date')
            else:
                print(">>> DMX日期校验失败，保留原数据")
                parsed_data['dmx_recheck_failed'] = True

        valid_prefixes = tuple(CONFIG.get('valid_style_prefixes', ['T', 'H', 'X', 'D']))
        max_retries = CONFIG.get('image_recognition_max_retries', 3)
        retry_count = 1
        failure_reason = None

        if not final_style or (final_style and not final_style.upper().startswith(valid_prefixes)):
            if not final_style:
                print(">>> 未识别到款号，开始重试识别...")
            else:
                print(f">>> 识别的款号{final_style}不符合规律，开始重试识别...")

            for retry_attempt in range(1, max_retries + 1):
                print(f">>> 第{retry_attempt}次重试(款号)...")
                retry_parsed_data = extract_data_from_image_dmx(img_path, current_prompt)
                retry_count = retry_attempt + 1
                if retry_parsed_data and 'error' not in retry_parsed_data:
                    retry_final_style = determine_final_style(retry_parsed_data, LOCAL_STYLE_DB)
                    if retry_final_style and retry_final_style.upper().startswith(valid_prefixes):
                        s_name = normalize_supplier_name(retry_parsed_data.get('supplier_name'), LOCAL_SUPPLIER_DB)
                        if s_name:
                            print(f">>> 款号重试成功: {retry_final_style}, 供应商: {s_name}")
                            parsed_data = retry_parsed_data
                            final_style = retry_final_style
                            parsed_data['final_selected_style'] = final_style
                            parsed_data['supplier_name'] = s_name
                            # 添加采购助理信息
                            if LOCAL_PURCHASER_MAPPING:
                                purchaser = LOCAL_PURCHASER_MAPPING.get(s_name, "不存在")
                                parsed_data['purchaser'] = purchaser
                            failure_reason = None
                            break
                        else:
                            print(f">>> 重试款号成功但供应商匹配失败，继续重试...")
                            continue
                else:
                    print(f">>> 第{retry_attempt}次重试识别失败")
            else:
                failure_reason = "款号没有解析到"
                print(f">>> 所有重试均失败，{failure_reason}")

        if not final_style or not final_style.upper().startswith(valid_prefixes):
            print(f"!!! 款号解析最终失败: {file_name}, 款号: {final_style or 'None'}")
            return {
                'file_name': file_name,
                'img_path': img_path,
                'success': False,
                'error': '款号没有解析到',
                'parsed_data': parsed_data,
                'final_style': final_style or '',
                'failure_reason': '款号没有解析到'
            }

        parsed_data['retry_count'] = retry_count
        if failure_reason: parsed_data['failure_reason'] = failure_reason

        # 应用物料扣减价格调整
        if LOCAL_MATERIAL_DEDUCTION_DB and deduction_suppliers:
            try:
                adjusted_data, has_adjustment = apply_material_deduction(
                    parsed_data, LOCAL_MATERIAL_DEDUCTION_DB, deduction_suppliers
                )
                if has_adjustment:
                    parsed_data = adjusted_data
                    parsed_data['price_adjusted'] = True
            except Exception as price_adjust_error:
                print(f"!!! 价格调整异常: {price_adjust_error}")
                parsed_data['price_adjustment_error'] = str(price_adjust_error)

        saved = db_manager.save_data(file_name, parsed_data)
        if not saved:
            return {
                'file_name': file_name,
                'img_path': img_path,
                'success': False,
                'error': '数据保存失败',
                'parsed_data': parsed_data,
                'final_style': final_style,
                'failure_reason': parsed_data.get('failure_reason', '数据保存失败')
            }

        return {
            'file_name': file_name,
            'img_path': img_path,
            'success': True,
            'parsed_data': parsed_data,
            'final_style': final_style,
            'failure_reason': parsed_data.get('failure_reason', '')
        }

    except Exception as e:
        print(f"!!! 解析图片异常: {file_name}, 错误: {e}")
        return {
            'file_name': file_name,
            'img_path': img_path,
            'success': False,
            'error': str(e),
            'parsed_data': None,
            'final_style': None,
            'failure_reason': str(e)
        }


def process_complete_rpa(rpa_executor, result):
    """处理完整的RPA+LLM匹配流程并收集报告数据"""
    try:
        if result['success']:
            file_name = result.get('file_name', '')
            parsed_data = result.get('parsed_data', {})
            img_path = result.get('img_path', '')

            # 使用 executor 实例方法
            match_prompt, match_result, original_records, retry_count = rpa_executor.run_process(parsed_data, file_name,
                                                                                                 img_path)

            # -------------------------------------------------------------------------
            # [核心修复]：检查 RPA 过程中的失败标记
            # -------------------------------------------------------------------------
            if parsed_data.get('processing_failed', False):
                failure_msg = parsed_data.get('failure_reason', 'RPA处理失败')
                print(f"!!! RPA处理失败: {file_name}, 原因: {failure_msg}")

                # ⚠️ 关键点：如果 RPA 失败，强制将 match_result 的状态改为 fail
                # 这样报告生成器才会将其渲染为失败/红色，而不是因为 LLM 匹配成功就显示成功
                if match_result:
                    match_result['status'] = 'fail'
                    # 将 RPA 的失败原因追加到 reason 中
                    old_reason = match_result.get('reason', '')
                    match_result['reason'] = f"{old_reason} | RPA执行失败: {failure_msg}"
                else:
                    # 如果本来就没有 match_result，手动造一个失败的结果
                    match_result = {
                        'status': 'fail',
                        'reason': f"RPA执行失败: {failure_msg}"
                    }

                return collect_result_data(
                    image_name=file_name,
                    parsed_data=parsed_data,
                    final_style=result.get('final_style', ''),
                    match_prompt=match_prompt or "",
                    match_result=match_result,  # 传入修改后的失败结果
                    original_records=original_records or [],
                    image_path=img_path,
                    retry_count=retry_count,
                    failure_reason=failure_msg
                )
            # -------------------------------------------------------------------------

            if parsed_data.get('total_amount'):
                archive_success_image(img_path, parsed_data)

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
    excel_path = CONFIG.get('style_db_path')
    col_name = CONFIG.get('style_db_column', '款式编号')
    LOCAL_STYLE_DB = load_style_db_with_cache(excel_path, col_name)

    supplier_excel_path = CONFIG.get('supplier_db_path')
    supplier_column = CONFIG.get('supplier_name_column')
    purchaser_column = CONFIG.get('supplier_purchaser_column')
    
    if supplier_excel_path:
        supplier_data = load_supplier_db_with_cache(
            supplier_excel_path, 
            supplier_column=supplier_column,
            purchaser_column=purchaser_column
        )
        
        # 兼容处理：如果返回的是字典，提取供应商集合；如果是集合，直接使用
        if isinstance(supplier_data, dict):
            LOCAL_SUPPLIER_DB = supplier_data.get('suppliers', set())
            LOCAL_PURCHASER_MAPPING = supplier_data.get('purchaser_mapping', {})
            print(f">>> 月结供应商目录加载完成，共 {len(LOCAL_SUPPLIER_DB)} 个供应商，{len(LOCAL_PURCHASER_MAPPING)} 个采购助理映射")
        else:
            # 向后兼容：旧格式返回set
            LOCAL_SUPPLIER_DB = supplier_data
            LOCAL_PURCHASER_MAPPING = {}
            print(f">>> 月结供应商目录加载完成，共 {len(LOCAL_SUPPLIER_DB)} 个供应商")
    else:
        LOCAL_SUPPLIER_DB = set()
        LOCAL_PURCHASER_MAPPING = {}

    material_deduction_excel_path = CONFIG.get('material_deduction_db_path')
    material_deduction_name_col = CONFIG.get('material_deduction_name_column', '物料名称')
    material_deduction_amount_col = CONFIG.get('material_deduction_amount_column', '扣减金额')
    LOCAL_MATERIAL_DEDUCTION_DB = load_material_deduction_db_with_cache(material_deduction_excel_path, material_deduction_name_col, material_deduction_amount_col) if material_deduction_excel_path else {}
    print(f">>> 物料扣减列表加载完成，共 {len(LOCAL_MATERIAL_DEDUCTION_DB)} 个物料")

    deduction_suppliers = CONFIG.get('material_deduction_suppliers', [])
    print(f">>> 扣减供应商列表: {deduction_suppliers}")

    storage_path = CONFIG.get('data_storage_path')
    if not storage_path:
        print("错误: 请在 app_config.py 中配置 'data_storage_path'")
        return
    db_manager = DataManager(storage_path_str=storage_path)

    source_dir = CONFIG.get('image_source_dir')
    if not source_dir or not os.path.exists(source_dir):
        print(f"错误: 图片源目录不存在，请检查配置: {source_dir}")
        return

    image_files = []
    for ext in ['*.jpg', '*.jpeg', '*.png', '*.JPG', '*.PNG']:
        image_files.extend(glob.glob(os.path.join(source_dir, ext)))

    # 去重处理，解决Windows下大小写扩展名重复匹配问题
    image_files = list(set(image_files))

    if not image_files:
        print(f"目录 {source_dir} 下没有找到图片文件。")
        return
    print(f"扫描到 {len(image_files)} 个图片文件，准备开始处理...")

    # 连接浏览器
    co = ChromiumOptions().set_address(CONFIG['chrome_address'])
    browser = Chromium(addr_or_opts=co)

    # 初始化 RPA 执行器
    rpa_executor = RPAExecutor(browser)

    # 阶段1：并发图片解析
    print("\n=== 阶段1：并发图片解析 ===")
    parsing_concurrency = CONFIG.get('image_parsing_concurrency', 3)
    semaphore = asyncio.Semaphore(parsing_concurrency)

    async def parse_with_semaphore(img_path):
        async with semaphore:
            loop = asyncio.get_event_loop()
            with ThreadPoolExecutor() as executor:
                return await loop.run_in_executor(executor, parse_single_image, img_path, db_manager, LOCAL_STYLE_DB,
                                                  LOCAL_SUPPLIER_DB, LOCAL_MATERIAL_DEDUCTION_DB, deduction_suppliers, LOCAL_PURCHASER_MAPPING)

    parse_tasks = [parse_with_semaphore(img_path) for img_path in image_files]
    parse_results = await asyncio.gather(*parse_tasks, return_exceptions=True)

    valid_results = [r for r in parse_results if r is not None and not isinstance(r, Exception)]
    success_results = [r for r in valid_results if r.get('success', False)]

    print(f"解析完成: 总计 {len(image_files)} 张，有效 {len(valid_results)} 张，成功 {len(success_results)} 张")
    failed_results = [r for r in valid_results if not r.get('success', False)]
    print(f"款号解析失败: {len(failed_results)} 张")

    # 阶段2：并发RPA+LLM匹配
    print("\n=== 阶段2：并发RPA+LLM匹配 ===")
    rpa_concurrency = CONFIG.get('rpa_concurrency', 3)
    rpa_semaphore = asyncio.Semaphore(rpa_concurrency)

    async def process_with_semaphore(result):
        async with rpa_semaphore:
            loop = asyncio.get_event_loop()
            with ThreadPoolExecutor() as pool:
                # 传递 rpa_executor 实例
                return await loop.run_in_executor(pool, process_complete_rpa, rpa_executor, result)

    process_tasks = [process_with_semaphore(result) for result in success_results]
    rpa_results = await asyncio.gather(*process_tasks, return_exceptions=True)

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

    final_results = rpa_results + failed_report_results
    print(f"RPA+LLM匹配完成: {len(final_results)} 张")

    print("\n=== 生成最终报告 ===")
    report_path = CONFIG.get('report_output_path')
    if report_path:
        try:
            timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            report_filename = f'report_{timestamp}.html'
            if os.path.isdir(report_path):
                report_file = os.path.join(report_path, report_filename)
            else:
                report_file = os.path.join(report_path, report_filename)

            for result_data in final_results:
                if not isinstance(result_data, Exception) and result_data:
                    update_html_report(os.path.dirname(report_file), result_data['image_name'], result_data)

            default_report = os.path.join(os.path.dirname(report_file), 'report.html')
            if os.path.exists(default_report):
                os.rename(default_report, report_file)

            print(f"✅ 报告已生成: {report_file}")
        except Exception as e:
            print(f"⚠️ 报告生成失败: {e}")

    print(">>> 所有任务处理完毕。")


if __name__ == "__main__":
    asyncio.run(async_main())