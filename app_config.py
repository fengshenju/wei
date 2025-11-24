#!/usr/bin/env python3
"""
配置文件
"""
# 三个主要功能开关
CONFIG = {

    # LLM 配置
    "llm_provider": "openai",  # 模型平台：siliconflow, openai, anthropic, etc.
    "llm_model": "deepseek-ai/DeepSeek-V3",  # 模型名称
    "llm_base_url": "https://api.siliconflow.cn/v1",  # API 端点
    "llm_api_key": "sk-650aa3e14159421bbbbbf2e2f285dfc5",  # API Key
    "llm_temperature": 0.1,  # 模型温度 (0.0-1.0)
    "llm_max_tokens": 500,  # 最大 token 数

# 其他配置
    "search_keyword": "cto",
    "valid_boss_status": ["在线"],
    "max_job_count": 15,
    "chrome_address": "192.168.50.207:9333",
    "base_url": "https://a1.scm321.com/Admin/Dashboard",

    # 待处理图片的本地文件夹路径 (请修改为您实际的图片文件夹)
    'image_source_dir': r'/Users/fengshenju/Downloads/program/pythonProjects/wei/data/images',

    'data_storage_path': r'/Users/fengshenju/Downloads/program/pythonProjects/wei/data/processed_data',

    # Excel 文件的绝对路径
    'style_db_path': r'/Users/fengshenju/Downloads/program/pythonProjects/wei/data/style/kuanhaochi.xlsx',

    # Excel 中存储款号的那一列的表头名称
    'style_db_column': '款式编号',

    # 图片数据提取的提示词配置
    'prompt_instruction': """
        你是一个专业的单据数据提取助手。请分析这张采购单据图片，提取关键信息并严格按照下方的 JSON 格式返回。
        
        请注意：不要使用 Markdown 格式，只返回纯 JSON 字符串。
        
        ### 1. 基础信息提取
        请提取以下字段：
        - 交付日期 (delivery_date): 格式为 YYYY-MM-DD，若未找到则留空。
        - 采购商名称 (buyer_name)
        - 供应商名称 (supplier_name)
        
        ### 2. 商品明细 (items)
        请提取表格中的每一行明细数据，包含：
        - 数量 (qty): 数字格式
        - 单价 (price): 数字格式
        - 单位 (unit)
        - 原始款号/品名 (raw_style_text): 表格这一行中原本写的款号或品名文字。
        
        ### 3. 款号候选池 (style_candidates) —— 关键任务
        为了辅助后台系统精准识别款号，请将图片中**所有**可能是“款号”的文本提取到一个列表中。
        **提取范围**：
        1. 表格明细中的款号列内容。
        2. 图片中**红色字体**的手写或打印标注（重点关注）。
        3. 任何以字母 T、H、X、D 开头的字母数字组合。
        
        **对于每一个候选文本，请提取以下属性：**
        - text: 文本内容（去除空格）。
        - is_red: (Boolean) 该文本在图片中是否显示为红色字体？这是最高优先级的判断依据。
        - position: (String) 文本所在位置描述，例如："表格内"、"手写标注"、"右上角"、"页眉"。
        
        ### 4. 返回格式示例 (JSON)
        {
            "delivery_date": "2025-11-24",
            "buyer_name": "某某服饰有限公司",
            "supplier_name": "某某制衣厂",
            "style_candidates": [
                {
                    "text": "T8821", 
                    "is_red": false, 
                    "position": "表格内"
                },
                {
                    "text": "H2201", 
                    "is_red": true, 
                    "position": "图片中间手写标注"
                }
            ],
            "items": [
                {
                    "qty": 500, 
                    "price": 12.5, 
                    "unit": "件", 
                    "raw_style_text": "T8821"
                }
            ]
        }
    """,
}