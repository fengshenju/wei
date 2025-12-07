import json
import re
import datetime
from datetime import timedelta
import time
from app_config import CONFIG
from utils.util_llm import call_llm_text, call_dmxllm_text

class MatchService:
    def __init__(self):
        pass
    # 预处理后台记录
    def _preprocess_records(self, records):
        """
        预处理 records，把 /Date(xxx)/ 时间戳转成人类可读日期
        """
        cleaned_records = []
        for rec in records:
            new_rec = rec.copy()
            # 处理 OrderReqCheckDate
            c_time = rec.get('OrderReqCheckDate')
            if c_time and isinstance(c_time, str) and '/Date(' in c_time:
                try:
                    match = re.search(r'\/Date\((-?\d+)\)\/', c_time)
                    if match:
                        timestamp = int(match.group(1)) / 1000
                        if timestamp > 0:
                            dt = datetime.datetime.fromtimestamp(timestamp)
                            new_rec['CreateTime_Readable'] = dt.strftime('%Y-%m-%d')
                        else:
                            new_rec['CreateTime_Readable'] = "Unknown"
                except:
                    new_rec['CreateTime_Readable'] = "Invalid"
            else:
                new_rec['CreateTime_Readable'] = "Unknown"

            # 精简字段
            cleaned_records.append({
                "Id": new_rec.get("Id"),
                "DBSupplierSpName": new_rec.get("DBSupplierSpName"),
                "DBSupplierSpShortName": new_rec.get("DBSupplierSpShortName"),
                "CreateTime_Readable": new_rec.get("CreateTime_Readable"),
                "TotalAmount": new_rec.get("TotalAmount"),
                "MaterialMtName": new_rec.get("MaterialMtName"),
                "MaterialSpec": new_rec.get("MaterialSpec")
            })
        return cleaned_records


    # 智能匹配
    def execute_smart_match(self, parsed_data, records):
        """
        执行智能匹配核心逻辑
        返回: (match_result, match_prompt, retry_count) 元组
        """
        # 1. 准备上下文数据
        today = datetime.date.today()
        two_weeks_ago = today - timedelta(days=14)

        # 2. 清洗后台记录
        clean_records = self._preprocess_records(records)

        # 为 OCR 明细项注入显式索引
        ocr_items_with_index = []
        original_items = parsed_data.get('items', [])
        for idx, item in enumerate(original_items):
            item_copy = item.copy()
            item_copy['_index'] = idx  # 显式注入索引，0, 1, 2...
            ocr_items_with_index.append(item_copy)

        # 构造传给 LLM 的数据视图 (包含索引)
        llm_input_ocr = {
            **parsed_data,
            "items": ocr_items_with_index
        }

        # 3. 从配置加载提示词模板
        prompt_template = CONFIG.get('match_prompt_template')
        if not prompt_template:
            print("!!! 错误: 未在 app_config.py 中找到 'match_prompt_template'")
            return {"status": "error", "reason": "配置缺失"}, "", 1

        # 4. 填充模板数据
        final_prompt = prompt_template.format(
            current_date=today.strftime('%Y-%m-%d'),
            two_weeks_ago=two_weeks_ago.strftime('%Y-%m-%d'),
            parsed_data_json=json.dumps(parsed_data, ensure_ascii=False, indent=2),
            records_json=json.dumps(clean_records, ensure_ascii=False, indent=2)
        )

        # 5. 重试逻辑
        max_retries = CONFIG.get('llm_match_max_retries', 3)
        match_result = None
        retry_count = 0

        for retry_count in range(1, max_retries + 1):
            # 第一次使用阿里通义千问，重试时使用DMX接口
            if retry_count == 1:
                print(">>> 使用阿里通义千问进行首次匹配...")
                match_result = call_llm_text(final_prompt, retry_count - 1)  # 阿里通义千问
            else:
                print(">>> 使用DMX接口进行重试匹配...")
                match_result = call_dmxllm_text(final_prompt, retry_count - 1)  # DMX接口

            if match_result and match_result.get('status') == 'success':
                return match_result, final_prompt, retry_count

            print(f">>> LLM匹配第{retry_count}次尝试失败")
            if retry_count < max_retries:  # 不是最后一次重试才等待
                wait_seconds = 2
                print(f">>> 等待{wait_seconds}秒后重试...")
                time.sleep(wait_seconds)

        return match_result, final_prompt, retry_count
