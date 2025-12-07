import datetime
from app_config import CONFIG

class DataProcessor:
    @staticmethod
    def should_use_dmx_for_date_check(delivery_date_str):
        """
        检查交付日期是否异常，如果与当前日期相差超过阈值天数则返回True
        :param delivery_date_str: 交付日期字符串，格式如 "2025-12-05"
        :return: bool, True表示需要用DMX重新识别
        """
        try:
            if not delivery_date_str:
                return False

            # 解析交付日期
            if len(delivery_date_str) == 10 and delivery_date_str.count('-') == 2:
                # YYYY-MM-DD 格式
                delivery_date = datetime.datetime.strptime(delivery_date_str, '%Y-%m-%d').date()
            else:
                # 尝试其他常见格式
                for fmt in ['%Y/%m/%d', '%m/%d/%Y', '%d/%m/%Y', '%Y.%m.%d']:
                    try:
                        delivery_date = datetime.datetime.strptime(delivery_date_str, fmt).date()
                        break
                    except ValueError:
                        continue
                else:
                    # 所有格式都失败
                    print(f">>> 无法解析交付日期格式: {delivery_date_str}")
                    return False

            # 获取当前日期
            current_date = datetime.datetime.now().date()

            # 计算日期差值（绝对值）
            date_diff = abs((delivery_date - current_date).days)

            # 获取阈值
            threshold_days = CONFIG.get('delivery_date_threshold_days', 7)

            if date_diff > threshold_days:
                print(
                    f">>> 交付日期异常检测: 当前日期 {current_date}, 交付日期 {delivery_date}, 差值 {date_diff} 天 > 阈值 {threshold_days} 天")
                return True
            else:
                print(f">>> 交付日期正常: 当前日期 {current_date}, 交付日期 {delivery_date}, 差值 {date_diff} 天")
                return False

        except Exception as e:
            print(f">>> 交付日期解析失败: {delivery_date_str}, 错误: {e}")
            return False

    @staticmethod
    def smart_clean_with_db(text, style_db):
        """基于款号库的智能款号清理"""
        if not text:
            return text

        # 1. 直接匹配（原文就在库中）
        if text in style_db:
            return text

        # 2. 检测重复款号（如H1591A-AH1591A-B）
        # 先检查是否包含库中的款号
        for db_style in style_db:
            if db_style in text and len(db_style) > 3:  # 避免匹配过短的款号
                # 计算该款号在文本中出现的次数
                count = text.count(db_style)
                if count > 1:
                    # 重复出现，返回单个款号
                    print(f">>> 检测到重复款号，提取: {db_style}")
                    return db_style
                elif count == 1:
                    # 只出现一次，可能是正确的
                    return db_style

        # 3. OCR字符纠错：处理常见的OCR识别错误
        ocr_char_map = {
            'J': '1',  # HJ643C → H1643C
            'O': '0',  # HO123A → H0123A
            'I': '1',  # HI456B → H1456B
            'S': '5',  # H5789C → H5789C (S→5)
            'Z': '2'  # HZ321A → H2321A
        }

        for wrong_char, correct_char in ocr_char_map.items():
            if wrong_char in text:
                corrected_text = text.replace(wrong_char, correct_char)
                if corrected_text in style_db:
                    print(f">>> OCR纠错成功: {text} → {corrected_text}")
                    return corrected_text

        # 4. 尝试多种清理策略
        # 先处理常见的末尾字符
        text_clean_hash = text.rstrip('#') if text.endswith('#') else text
        text_clean_kuan = text.rstrip('款') if text.endswith('款') else text

        cleaning_strategies = [
            text_clean_hash,  # 去除#结尾
            text_clean_kuan,  # 去除款字结尾
            text.rstrip('款型式号'),  # 移除常见后缀
            text.split('款')[0].strip(),  # 取"款"字前的部分
            text.split('型')[0].strip(),  # 取"型"字前的部分
            text.split('式')[0].strip(),  # 取"式"字前的部分
            text.replace(' ', ''),  # 去除空格
            text.upper(),  # 转大写
            text.upper().rstrip('款型式号'),  # 大写+去后缀
            text.upper().rstrip('#'),  # 大写+去除#
            text.upper().rstrip('款'),  # 大写+去除款字
        ]

        # 5. 逐一尝试，找到第一个匹配的
        for candidate in cleaning_strategies:
            if candidate and candidate in style_db:
                return candidate

        # 6. 都不匹配，返回最干净的版本
        return text.rstrip('款型式号').strip()

    @staticmethod
    def normalize_supplier_name(extracted_name, supplier_db):
        """
        供应商名称标准化匹配（优化版：增强短词匹配和包含逻辑）
        """
        if not extracted_name or not supplier_db:
            return None

        extracted_name = str(extracted_name).strip()
        # 去除常见的无关后缀，提高纯净度（可根据实际情况添加）
        clean_extracted = extracted_name.replace("商行", "").replace("有限公司", "").replace("布行", "").strip()

        if not extracted_name:
            return None

        # 1. 精确匹配 (最快)
        if extracted_name in supplier_db:
            print(f"   [匹配] 精确命中: {extracted_name}")
            return extracted_name

        # 2. 核心包含逻辑 (解决 "罗卡" vs "罗卡家")
        # 遍历库里的标准名
        for db_name in supplier_db:
            # A. 库里的名字包含提取的名字 (如 库:"杭州罗卡", 提:"罗卡")
            if extracted_name in db_name:
                print(f"   [匹配] 包含命中(提取在库中): {extracted_name} -> {db_name}")
                return db_name

            # B. 提取的名字包含库里的名字 (如 库:"罗卡", 提:"罗卡家" / "罗卡纺织")
            # ⚠️关键：防止由"罗"匹配到"罗卡"，限制标准名长度至少为2
            if len(db_name) >= 2 and db_name in extracted_name:
                print(f"   [匹配] 包含命中(库在提取中): {extracted_name} -> {db_name}")
                return db_name

            # C. 针对清洗后的包含 (如 库:"罗卡", 提:"罗卡商行"->clean:"罗卡")
            if len(db_name) >= 2 and db_name == clean_extracted:
                print(f"   [匹配] 清洗后命中: {extracted_name} -> {db_name}")
                return db_name

        # 3. 编辑距离匹配 (容错)
        def levenshtein_distance(s1, s2):
            if len(s1) < len(s2):
                return levenshtein_distance(s2, s1)
            if len(s2) == 0:
                return len(s1)
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

            # 动态阈值计算
            # 如果标准名很短(<=3)，允许1个字符差异
            # 如果标准名较长，允许 40% 的差异
            length = max(len(extracted_name), len(standard_name))
            if length <= 3:
                threshold = 1.0  # 允许错1个字
            else:
                threshold = length * 0.4

            if dist <= threshold and dist < min_distance:
                min_distance = dist
                best_match = standard_name

        if best_match:
            print(f"   [匹配] 模糊命中(距离{min_distance}): {extracted_name} -> {best_match}")

        return best_match

    @staticmethod
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
                cleaned = DataProcessor.smart_clean_with_db(clean_text, style_db)
                print(f"✅ 命中规则1 (表格内且在库): {cleaned}")
                return cleaned

        # --- 规则 2: 识别红色字体 (若存在且在库中) ---
        red_candidates = [c for c in candidates if c.get('is_red') == True]
        for item in red_candidates:
            clean_text = item['text'].strip()
            if clean_text in style_db:
                cleaned = DataProcessor.smart_clean_with_db(clean_text, style_db)
                print(f"✅ 命中规则2 (红色字体且在库): {cleaned}")
                return cleaned

        # --- 规则 3: 兜底搜索 (符合 T/H/X/D 开头规律) ---
        for item in candidates:
            text = item['text'].strip().upper()
            # 这里逻辑可根据需求调整：如果不在库里，但长得像，要不要？
            # 下面代码假设：只要长得像就提取，作为最后的保底
            if text.startswith(('T', 'H', 'X', 'D', 'L', 'S', 'F')):  # 我看你csv里还有L/S/F开头
                cleaned = DataProcessor.smart_clean_with_db(text, style_db)
                print(f"✅ 命中规则3 (符合命名规律): {cleaned}")
                return cleaned

        return None

    @staticmethod
    def reconstruct_rpa_data(match_result, original_parsed_data, original_records):
        """
        根据 llm 匹配后的返回构建 RPA 传递数据
        """
        matched_tasks = []

        # 1. 建立快速查找表 (Hash Map)
        # 将 list 转换为 dict: { "uuid": record_obj }，方便 O(1) 查找
        record_map = {rec['Id']: rec for rec in original_records}
        ocr_items_list = original_parsed_data.get('items', [])

        # 辅助函数：安全获取 OCR Item
        def get_item_by_index(idx):
            if isinstance(idx, int) and 0 <= idx < len(ocr_items_list):
                return ocr_items_list[idx]
            return None

        # --- Type 1: Direct (1对1) ---
        for match in match_result.get('direct_matches', []):
            rid = match.get('record_id')
            idx = match.get('ocr_index')

            target_record = record_map.get(rid)
            target_item = get_item_by_index(idx)

            if target_record and target_item:
                matched_tasks.append({
                    "match_type": "DIRECT",
                    "record": target_record,  # 完整 Record
                    "items": [target_item],  # 完整 Item (统一放入列表)
                    "ocr_context": original_parsed_data
                })

        # --- Type 2: Merge (N对1) ---
        for match in match_result.get('merge_matches', []):
            rid = match.get('record_id')
            indices = match.get('ocr_indices', [])

            target_record = record_map.get(rid)
            # 获取所有对应的 items
            target_items = [get_item_by_index(i) for i in indices if get_item_by_index(i)]

            if target_record and target_items:
                matched_tasks.append({
                    "match_type": "MERGE",
                    "record": target_record,  # 完整 Record
                    "items": target_items,  # 多个 完整 Items
                    "ocr_context": original_parsed_data
                })

        # --- Type 3: Split (1对N) ---
        for match in match_result.get('split_matches', []):
            rid = match.get('record_id')
            idx = match.get('ocr_index')

            target_record = record_map.get(rid)
            target_item = get_item_by_index(idx)

            if target_record and target_item:
                matched_tasks.append({
                    "match_type": "SPLIT",
                    "record": target_record,  # 完整 Record
                    "items": [target_item],  # 完整 Item
                    "ocr_context": original_parsed_data
                })

        return matched_tasks
