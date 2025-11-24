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
}