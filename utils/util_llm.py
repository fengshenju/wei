# 文件路径: utils/llm_utils.py
from litellm import completion
import json
import os
import sys
import base64
import mimetypes

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app_config import CONFIG


def encode_image_to_base64(image_path):
    """
    将本地图片转换为 Base64 编码，以便传递给大模型
    :param image_path: 图片的绝对路径或相对路径
    :return: data URI 格式的字符串 (e.g., "data:image/jpeg;base64,xxxx")
    """
    if not os.path.exists(image_path):
        raise FileNotFoundError(f"图片文件不存在: {image_path}")

    # 获取原本的 mime type (如 image/png, image/jpeg)
    mime_type, _ = mimetypes.guess_type(image_path)

    if not mime_type:
        mime_type = 'image/jpeg'  # 默认回退

    with open(image_path, "rb") as image_file:
        encoded_string = base64.b64encode(image_file.read()).decode('utf-8')
        return f"data:{mime_type};base64,{encoded_string}"


def extract_data_from_image(image_path, prompt_instructions):
    """
    使用 Qwen-VL-Max 模型分析图片并提取数据
    :param image_path: 本地图片路径
    :param prompt_instructions: 指示模型提取什么数据的提示词
    :return: 解析后的 JSON 数据 (dict) 或 原始文本 (str)
    """

    # 强制使用 qwen-vl-max
    # 注意：在 litellm 中调用通义千问通常使用 "openai/qwen-vl-max" (走兼容协议)
    # 或者配置好环境变量后直接用 "qwen-vl-max"
    model_name = "openai/qwen-vl-max"

    print(f">>> 正在处理图片: {os.path.basename(image_path)}")
    print(f">>> 使用模型: {model_name}")

    try:
        # 1. 准备图片数据
        base64_image = encode_image_to_base64(image_path)

        # 2. 构造符合 OpenAI Vision 标准的消息格式 (Qwen-VL 兼容此格式)
        # 这是一个多模态输入的标准写法：content 是一个数组，包含 text 和 image_url
        messages = [
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": f"{prompt_instructions}\n\n请直接返回纯 JSON 格式的数据，不要包含 markdown 标记。"
                    },
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": base64_image
                        }
                    }
                ]
            }
        ]

        # 3. 构造请求参数
        request_params = {
            "model": model_name,
            "messages": messages,
            "temperature": 0.1,  # 提取数据需要精确，温度设低
            "max_tokens": 2000,  # 防止截断，视图片内容多少而定
        }

        # 配置 API Key (优先读取配置，否则读取环境变量)
        # 注意：如果是 Qwen，通常需要在环境变量设置 DASHSCOPE_API_KEY
        # 或者在 litellm 调用时显式传入 api_key
        api_key = CONFIG.get('llm_api_key')
        if api_key:
            request_params["api_key"] = api_key

        request_params["base_url"] = "https://dashscope.aliyuncs.com/compatible-mode/v1"

        print(f">>> 发起 LLM 请求...")
        response = completion(**request_params)

        content = response.choices[0].message.content.strip()

        # 4. 数据清洗与 JSON 解析
        cleaned_content = content
        if cleaned_content.startswith("```json"):
            cleaned_content = cleaned_content[7:]
        if cleaned_content.endswith("```"):
            cleaned_content = cleaned_content[:-3]

        try:
            result_json = json.loads(cleaned_content.strip())
            return result_json
        except json.JSONDecodeError:
            print(">>> 警告: 模型返回的不是标准 JSON，返回原始文本")
            return {"raw_text": content}

    except Exception as e:
        print(f">>> 图片分析失败: {e}")
        return {"error": str(e)}


# 测试代码
if __name__ == "__main__":
    # 假设你有一张测试图片放在同一目录下，或者输入绝对路径
    test_image = "/Users/fengshenju/Desktop/aaa.jpg"  # 请修改为实际存在的图片路径

    # 你的业务 Prompt：告诉它读取哪些数据
    my_prompt = """
    这张图片是一张采购单据的照片，包含多个字段的信息。
    请提取以下关键信息，并按 JSON 格式返回：
    1. 款号
    2. 交付日期
    3. 单价、数量、单位
    4. 采购商名称、供应商名称
    """

    # 只有当文件存在时才运行测试
    if os.path.exists(test_image):
        result = extract_data_from_image(test_image, my_prompt)
        print("\n>>> 最终提取结果:")
        print(json.dumps(result, indent=4, ensure_ascii=False))
    else:
        print(f"测试文件不存在: {test_image}，请修改 main 函数中的路径进行测试。")