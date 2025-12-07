import os
from app_config import CONFIG
from utils.util_llm import extract_data_from_image, extract_data_from_image_dmx
from services.data_processor import DataProcessor

class ImageService:
    @staticmethod
    def parse_single_image(img_path, db_manager, style_db, supplier_db):
        """解析单张图片"""
        file_name = os.path.basename(img_path)
        
        # 从配置获取提示词
        PROMPT_INSTRUCTION = CONFIG.get('prompt_instruction', '')

        try:
            # 状态检查 (去重)
            if db_manager.is_processed(file_name):
                print(f"[跳过] 文件已处理过: {file_name}")
                return None

            print(f"[处理中] 正在解析: {file_name} ...")

            supplier_list_str = "、".join(supplier_db) if supplier_db else "无已知供应商，请自行识别"

            try:
                current_prompt = PROMPT_INSTRUCTION.format(known_suppliers=supplier_list_str)
            except KeyError as e:
                current_prompt = PROMPT_INSTRUCTION.replace("{known_suppliers}", supplier_list_str)

            # 1. 首次解析
            if CONFIG.get('use_llm_image_parsing', True):
                parsed_data = extract_data_from_image(img_path, current_prompt)
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

            final_style = DataProcessor.determine_final_style(parsed_data, style_db)
            parsed_data['final_selected_style'] = final_style

            # 2. 供应商匹配与重试逻辑 (核心修改区域)
            # -------------------------------------------------------------------------
            supplier_name = DataProcessor.normalize_supplier_name(parsed_data.get('supplier_name'), supplier_db)

            if supplier_name:
                # 首次匹配成功
                parsed_data['supplier_name'] = supplier_name
                print(f"匹配供应商: {supplier_name}")
            else:
                # 首次匹配失败 -> 尝试 DMX 重试
                print(f"!!! 首次供应商匹配失败: [{parsed_data.get('supplier_name')}]，正在使用 DMX 进行重试识别...")

                # 调用 DMX 接口 (retry_count=0)
                dmx_parsed_data = extract_data_from_image_dmx(img_path, current_prompt, 0)

                dmx_success = False
                if dmx_parsed_data and 'error' not in dmx_parsed_data:
                    # DMX 识别成功，再次尝试匹配供应商
                    retry_supplier_name = DataProcessor.normalize_supplier_name(dmx_parsed_data.get('supplier_name'), supplier_db)

                    if retry_supplier_name:
                        print(f"✅ DMX 重试挽回成功! 匹配到供应商: {retry_supplier_name}")

                        # 用 DMX 的高质量数据替换原始数据
                        parsed_data = dmx_parsed_data
                        parsed_data['supplier_name'] = retry_supplier_name

                        # ⚠️ 数据变了，重新判定一下款号，防止漏掉
                        final_style = DataProcessor.determine_final_style(parsed_data, style_db)
                        parsed_data['final_selected_style'] = final_style
                        print(f"   -> DMX重试后的款号判定: {final_style}")

                        dmx_success = True
                    else:
                        print(f"!!! DMX 重试后供应商依然无法匹配: [{dmx_parsed_data.get('supplier_name')}]")
                else:
                    print("!!! DMX 接口调用失败或未返回有效数据")

                # 如果 DMX 也没救回来，才返回失败
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
            # -------------------------------------------------------------------------

            # 3. 检查交付日期异常 (原有逻辑保持不变)
            # 注意：如果上面 DMX 已经替换了 parsed_data，这里检查的就是新数据的日期，逻辑是通顺的
            delivery_date = parsed_data.get('delivery_date', '')
            if delivery_date and DataProcessor.should_use_dmx_for_date_check(delivery_date):
                # 只有当还没有使用过 DMX 时才再次调用 (避免重复计费/耗时)
                # 我们可以通过检查 parsed_data 是否有特定标记来判断，或者直接根据日期逻辑走
                # 这里简单起见，如果日期异常，即使刚才为了供应商调用过，这里可能还是要处理(虽然概率低)
                # 优化：通常 DMX 的日期也是准的，所以如果上面已经替换过 parsed_data，这里的日期大概率是准的。

                print(f">>> 交付日期异常: {delivery_date}，使用DMX重新校验...")
                # 再次调用 DMX (如果是为了日期)
                dmx_date_data = extract_data_from_image_dmx(img_path, current_prompt, 0)

                if dmx_date_data and 'error' not in dmx_date_data:
                    print(">>> DMX日期校验返回成功，更新数据")
                    # 这里比较微妙，如果为了日期更新了数据，供应商名字可能会变回去(变成未标准化的)
                    # 所以要重新走一遍供应商匹配，或者只更新日期字段。
                    # 稳妥起见，替换数据后重新匹配供应商

                    temp_supplier = DataProcessor.normalize_supplier_name(dmx_date_data.get('supplier_name'), supplier_db)
                    if temp_supplier:
                        parsed_data = dmx_date_data
                        parsed_data['supplier_name'] = temp_supplier
                        final_style = DataProcessor.determine_final_style(parsed_data, style_db)
                        parsed_data['final_selected_style'] = final_style
                        parsed_data['used_dmx_for_date_check'] = True
                        print(f"   -> DMX日期修正后供应商: {temp_supplier}")
                    else:
                        print("   ⚠️ DMX日期校验数据的供应商无法匹配，仅更新日期字段")
                        parsed_data['delivery_date'] = dmx_date_data.get('delivery_date')
                else:
                    print(">>> DMX日期校验失败，保留原数据")
                    parsed_data['dmx_recheck_failed'] = True

            # 4. 款号重试识别逻辑 (原有逻辑保持不变)
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
                    # 使用 DMX 重试
                    retry_parsed_data = extract_data_from_image_dmx(img_path, current_prompt)
                    retry_count = retry_attempt + 1

                    if retry_parsed_data and 'error' not in retry_parsed_data:
                        retry_final_style = DataProcessor.determine_final_style(retry_parsed_data, style_db)
                        if retry_final_style and retry_final_style.upper().startswith(valid_prefixes):
                            # 重试成功，还需要检查供应商匹配
                            s_name = DataProcessor.normalize_supplier_name(retry_parsed_data.get('supplier_name'), supplier_db)
                            if s_name:
                                print(f">>> 款号重试成功: {retry_final_style}, 供应商: {s_name}")
                                parsed_data = retry_parsed_data
                                final_style = retry_final_style
                                parsed_data['final_selected_style'] = final_style
                                parsed_data['supplier_name'] = s_name
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

            # 5. 最终检查
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
            if failure_reason:
                parsed_data['failure_reason'] = failure_reason

            # 6. 数据持久化
            saved = db_manager.save_data(file_name, parsed_data)
            if not saved:
                print("!!! 数据保存失败，中断处理本条")
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
