# 文件路径: utils/llm_utils.py
from litellm import completion
import json
import os
import sys
import base64
import mimetypes
import requests

try:
    from google import genai
    from google.genai import types
    GENAI_AVAILABLE = True
except ImportError:
    GENAI_AVAILABLE = False
    print("Warning: google-genai not installed. DMX image processing will not be available.")

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app_config import CONFIG


def get_api_key(key_name, attempt=0):
    """
    获取API key，支持轮转重试
    :param key_name: key配置名称，如 'llm_api_key' 或 'llm_image_api_key'
    :param attempt: 重试次数，从0开始
    :return: 对应的API key，如果没有配置则返回None
    """
    keys = CONFIG.get(key_name, [])
    if not keys:
        return None
    
    # 如果配置的是字符串（向后兼容），直接返回
    if isinstance(keys, str):
        return keys
    
    # 如果配置的是数组，循环使用
    return keys[attempt % len(keys)]


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


def extract_data_from_image(image_path, prompt_instructions, attempt=0):
    """
    使用 Qwen-VL-Max 模型分析图片并提取数据
    :param image_path: 本地图片路径
    :param prompt_instructions: 指示模型提取什么数据的提示词
    :param attempt: 重试次数，用于选择不同的API key
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

        # 配置 API Key (支持轮转重试)
        api_key = get_api_key('llm_image_api_key', attempt)
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
            print(">>> 警告: 模型返回的不是标准 JSON")
            print(f">>> 原始返回内容: {content[:500]}...")
            print(f">>> 清理后内容: {cleaned_content[:500]}...")
            return {"raw_text": content}

    except Exception as e:
        print(f">>> 图片分析失败: {e}")
        return {"error": str(e)}


# 文件路径: utils/llm_utils.py
# (请确保文件头部已经有 import json, os, sys, app_config 等引入，与之前保持一致)

def call_llm_text(prompt_text, attempt=0):
    model_name = "openai/qwen3-max"

    print(f">>> 正在执行纯文本任务...")
    print(f">>> 使用模型: {model_name}")

    try:
        # 1. 构造符合 OpenAI 标准的消息格式
        # 纯文本任务 content 直接传字符串即可
        # 保持一致性：强制追加 JSON 格式要求，防止模型废话
        final_content = f"{prompt_text}\n\n请直接返回纯 JSON 格式的数据，不要包含 markdown 标记。"

        messages = [
            {
                "role": "user",
                "content": final_content
            }
        ]

        # 2. 构造请求参数
        request_params = {
            "model": model_name,
            "messages": messages,
            "temperature": 0.01,  # 逻辑匹配需要严谨，温度设极低
            "max_tokens": 5000,  # 防止截断
        }

        # 配置 API Key (支持轮转重试)
        api_key = get_api_key('llm_api_key', attempt)
        if api_key:
            request_params["api_key"] = api_key

        request_params["base_url"] = "https://dashscope.aliyuncs.com/compatible-mode/v1"

        print(f">>> 发起 LLM 请求 (Text)...")
        response = completion(**request_params)

        content = response.choices[0].message.content.strip()

        # 3. 数据清洗与 JSON 解析 (逻辑与 extract_data_from_image 完全一致)
        cleaned_content = content
        if cleaned_content.startswith("```json"):
            cleaned_content = cleaned_content[7:]
        elif cleaned_content.startswith("```"):  # 兼容部分模型只写 ``` 的情况
            cleaned_content = cleaned_content[3:]

        if cleaned_content.endswith("```"):
            cleaned_content = cleaned_content[:-3]

        try:
            result_json = json.loads(cleaned_content.strip())
            # 添加使用的API key信息
            result_json['used_api_key'] = api_key if api_key else "N/A"
            return result_json
        except json.JSONDecodeError:
            print(">>> 警告: 模型返回的不是标准 JSON")
            print(f">>> 原始返回内容: {content[:500]}...")
            print(f">>> 清理后内容: {cleaned_content[:500]}...")
            return {"raw_text": content, "used_api_key": api_key if api_key else "N/A"}

    except Exception as e:
        print(f">>> 文本分析失败: {e}")
        return {"error": str(e), "used_api_key": api_key if api_key else "N/A"}


def call_gjllm_text(prompt_text, attempt=0):
    model_name = "openai/Pro/deepseek-ai/DeepSeek-V3.1-Terminus"

    print(f">>> 正在执行纯文本任务...")
    print(f">>> 使用模型: {model_name}")

    try:
        # 1. 构造消息
        final_content = f"{prompt_text}\n\n请直接返回纯 JSON 格式的数据，不要包含 markdown 标记。"
        messages = [{"role": "user", "content": final_content}]

        # 2. 获取 API Key (支持轮转重试)
        api_key = get_api_key('gjllm_api_key', attempt)
        if not api_key:
            print("!!! 错误: 未配置 gjllm_api_key")
            return {"status": "fail", "reason": "API Key缺失"}

        # 3. 构造请求参数
        # litellm.completion 的标准参数是 api_base 和 api_key
        request_params = {
            "model": model_name,
            "messages": messages,
            "temperature": 0.01,
            "max_tokens": 4000,
            "api_key": api_key,
            "api_base": "https://api.siliconflow.cn/v1",
            # 可选: 显式指定不用缓存以测试连通性，正式用可去掉
            "drop_params": True
        }

        print(f">>> 发起 LLM 请求 (Text)...")

        # 4. 调用 litellm
        response = completion(**request_params)

        content = response.choices[0].message.content.strip()

        # 5. 数据清洗
        cleaned_content = content
        if cleaned_content.startswith("```json"):
            cleaned_content = cleaned_content[7:]
        elif cleaned_content.startswith("```"):
            cleaned_content = cleaned_content[3:]

        if cleaned_content.endswith("```"):
            cleaned_content = cleaned_content[:-3]

        try:
            result_json = json.loads(cleaned_content.strip())
            # 添加使用的API key信息
            result_json['used_api_key'] = api_key if api_key else "N/A"
            return result_json
        except json.JSONDecodeError:
            print(f">>> 警告: 模型返回的不是标准 JSON:\n{content}")
            return {"raw_text": content, "used_api_key": api_key if api_key else "N/A"}

    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f">>> 文本分析失败: {e}")
        # 返回失败结构，防止主程序 RPA 崩溃
        return {
            "status": "fail",
            "global_reason": f"LLM调用异常: {str(e)}",
            "matched_record_ids": [],
            "detail_analysis": [],
            "used_api_key": api_key if api_key else "N/A"
        }


def call_dmxllm_text(prompt_text, attempt=0):
    """
    调用DMX大模型接口进行文本处理
    :param prompt_text: 输入的提示词
    :param attempt: 重试次数，用于选择不同的API key
    :return: 解析后的 JSON 数据 (dict) 或 原始文本 (str)
    """
    print(f">>> 正在执行DMX文本任务...")
    
    try:
        # 1. 获取配置信息
        api_key = get_api_key('dmx_api_key', attempt)
        api_url = CONFIG.get('dmx_api_url', 'https://www.dmxapi.cn/v1/chat/completions')
        model = CONFIG.get('dmx_model', 'gemini-3-pro-preview')
        
        if not api_key:
            print("!!! 错误: 未配置 dmx_api_key")
            return {"status": "fail", "reason": "API Key缺失", "used_api_key": "N/A"}
        
        print(f">>> 使用模型: {model}")
        
        # 2. 构造请求数据
        final_content = f"{prompt_text}\n\n请直接返回纯 JSON 格式的数据，不要包含 markdown 标记。"
        
        request_data = {
            "model": model,
            "messages": [
                {
                    "role": "user", 
                    "content": final_content
                }
            ],
            "stream": False
        }
        
        # 3. 构造请求头
        headers = {
            "Accept": "application/json",
            "Authorization": api_key,
            "User-Agent": f"DMXAPI/1.0.0 ({api_url})",
            "Content-Type": "application/json"
        }
        
        # 4. 发起HTTP请求
        url = f"{api_url}"
        print(f">>> 发起 DMX 请求到: {url}")
        
        response = requests.post(
            url, 
            headers=headers,
            json=request_data,
            timeout=300  # 5分钟超时
        )
        
        # 5. 处理响应
        if response.status_code == 200:
            response_data = response.json()
            
            # 提取生成的内容
            if response_data.get('choices') and len(response_data['choices']) > 0:
                content = response_data['choices'][0]['message']['content'].strip()
                
                # 数据清洗与 JSON 解析
                cleaned_content = content
                if cleaned_content.startswith("```json"):
                    cleaned_content = cleaned_content[7:]
                elif cleaned_content.startswith("```"):
                    cleaned_content = cleaned_content[3:]
                
                if cleaned_content.endswith("```"):
                    cleaned_content = cleaned_content[:-3]
                
                try:
                    result_json = json.loads(cleaned_content.strip())
                    # 添加使用的API key信息
                    result_json['used_api_key'] = api_key if api_key else "N/A"
                    return result_json
                except json.JSONDecodeError:
                    print(">>> 警告: DMX模型返回的不是标准 JSON")
                    print(f">>> 原始返回内容: {content[:500]}...")
                    print(f">>> 清理后内容: {cleaned_content[:500]}...")
                    return {"raw_text": content, "used_api_key": api_key if api_key else "N/A"}
            else:
                print("!!! DMX响应中没有有效的choices数据")
                return {"error": "无效的响应格式", "used_api_key": api_key if api_key else "N/A"}
        else:
            print(f"!!! DMX API调用失败，状态码：{response.status_code}")
            print(f"!!! 响应内容：{response.text}")
            return {"error": f"HTTP {response.status_code}: {response.text}", "used_api_key": api_key if api_key else "N/A"}
            
    except requests.exceptions.Timeout:
        print("!!! DMX API调用超时")
        return {"error": "请求超时", "used_api_key": api_key if api_key else "N/A"}
    except requests.exceptions.RequestException as e:
        print(f"!!! DMX API请求异常: {e}")
        return {"error": f"请求异常: {str(e)}", "used_api_key": api_key if api_key else "N/A"}
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f">>> DMX文本分析失败: {e}")
        # 返回失败结构，防止主程序 RPA 崩溃
        return {
            "status": "fail",
            "global_reason": f"DMX LLM调用异常: {str(e)}",
            "matched_record_ids": [],
            "detail_analysis": [],
            "used_api_key": api_key if api_key else "N/A"
        }


def extract_data_from_image_dmx(image_path, prompt_instructions, attempt=0):
    """
    使用 DMX 接口 (OpenAI 兼容协议) 分析图片并提取数据
    完全仿照 DMX 文档的 Base64 写法，使用 requests 库直接请求
    """
    print(f">>> 正在处理图片: {os.path.basename(image_path)}")
    print(">>> 使用DMX Gemini模型 (Requests方式)...")

    try:
        # 1. 检查图片文件是否存在
        if not os.path.exists(image_path):
            raise FileNotFoundError(f"图片文件不存在: {image_path}")

        # 2. 获取配置信息
        api_key = get_api_key('dmx_image_api_key', attempt)
        # 注意：DMX 图片通常走 /v1/chat/completions
        api_url = CONFIG.get('dmx_image_api_url', 'https://www.dmxapi.cn/v1/chat/completions')
        model = CONFIG.get('dmx_image_model', 'gemini-1.5-pro')  # 建议确认 DMX 支持的模型名称

        if not api_key:
            print("!!! 错误: 未配置 dmx_image_api_key")
            return {"error": "API Key缺失", "used_api_key": "N/A"}

        # 3. 图片转 Base64 (复用你已有的函数)
        # 注意：encode_image_to_base64 返回的是 "data:image/jpeg;base64,xxxx" 格式
        # DMX/OpenAI 格式通常只需要在这个 URL 字段里填这个字符串即可
        base64_image_url = encode_image_to_base64(image_path)

        # 4. 构造请求头
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        }

        # 5. 构造请求体 (OpenAI Vision 标准格式)
        final_content = f"{prompt_instructions}\n\n请直接返回纯 JSON 格式的数据，不要包含 markdown 标记。"

        payload = {
            "model": model,
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": final_content
                        },
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": base64_image_url
                            }
                        }
                    ]
                }
            ],
            "max_tokens": 4000,
            "stream": False
        }

        print(f">>> 发起请求到: {api_url}")
        print(f">>> 模型: {model}")

        # 6. 发起 HTTP POST 请求
        response = requests.post(api_url, headers=headers, json=payload, timeout=120)

        # 7. 处理响应
        if response.status_code == 200:
            res_json = response.json()
            if "choices" in res_json and len(res_json["choices"]) > 0:
                content = res_json["choices"][0]["message"]["content"].strip()

                # 数据清洗
                cleaned_content = content
                if cleaned_content.startswith("```json"):
                    cleaned_content = cleaned_content[7:]
                elif cleaned_content.startswith("```"):
                    cleaned_content = cleaned_content[3:]
                if cleaned_content.endswith("```"):
                    cleaned_content = cleaned_content[:-3]

                try:
                    result_json = json.loads(cleaned_content.strip())
                    result_json['used_api_key'] = api_key if api_key else "N/A"
                    return result_json
                except json.JSONDecodeError:
                    print(">>> 警告: DMX 返回内容无法解析为JSON")
                    return {"raw_text": content, "used_api_key": api_key}
            else:
                print(f"!!! DMX 返回结构异常: {res_json}")
                return {"error": "Invalid response structure", "used_api_key": api_key}
        else:
            print(f"!!! DMX 请求失败: {response.status_code}")
            print(f"!!! 错误详情: {response.text}")
            return {"error": f"HTTP {response.status_code}: {response.text}", "used_api_key": api_key}

    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f">>> DMX图片分析系统错误: {e}")
        return {"error": str(e), "used_api_key": api_key if 'api_key' in locals() else "N/A"}


def extract_data_from_image_gemini(image_path, prompt_instructions, attempt=0):
    """
    使用 Google Gemini 原生 API 分析图片并提取数据
    直接调用 Google Gemini API（不走 DMX 代理）

    :param image_path: 本地图片路径
    :param prompt_instructions: 指示模型提取什么数据的提示词
    :param attempt: 重试次数，用于选择不同的API key
    :return: 解析后的 JSON 数据 (dict) 或 原始文本 (str)
    """
    print(f">>> 正在处理图片: {os.path.basename(image_path)}")
    print(">>> 使用 Google Gemini 原生 API...")

    # 检查 SDK 是否可用
    if not GENAI_AVAILABLE:
        print("!!! 错误: google-genai SDK 未安装")
        return {"error": "google-genai SDK 未安装", "used_api_key": "N/A"}

    try:
        # 1. 检查图片文件是否存在
        if not os.path.exists(image_path):
            raise FileNotFoundError(f"图片文件不存在: {image_path}")

        # 2. 获取配置信息
        api_key = get_api_key('gemini_api_key', attempt)
        model = CONFIG.get('gemini_model', 'gemini-2.0-flash')
        resolution = CONFIG.get('gemini_image_resolution', 'high')
        max_tokens = CONFIG.get('gemini_max_tokens', 4000)

        if not api_key:
            print("!!! 错误: 未配置 gemini_api_key")
            return {"error": "API Key缺失", "used_api_key": "N/A"}

        print(f">>> 使用模型: {model}")
        print(f">>> 图片分辨率级别: {resolution}")

        # 3. 读取图片文件（直接读取字节数据，不需要 base64 编码）
        with open(image_path, 'rb') as f:
            image_bytes = f.read()

        # 获取 MIME 类型
        mime_type, _ = mimetypes.guess_type(image_path)
        if not mime_type:
            mime_type = 'image/jpeg'  # 默认回退

        print(f">>> 图片 MIME 类型: {mime_type}")

        # 4. 初始化 Gemini 客户端
        try:
            client = genai.Client(api_key=api_key)
        except Exception as init_error:
            print(f"!!! Gemini 客户端初始化失败: {init_error}")
            return {"error": f"客户端初始化失败: {str(init_error)}", "used_api_key": api_key}

        # 5. 构造请求内容
        final_content = f"{prompt_instructions}\n\n请直接返回纯 JSON 格式的数据，不要包含 markdown 标记。"

        # 构造带 media_resolution 的图片 Part
        media_resolution_value = f"media_resolution_{resolution}"
        image_part = types.Part(
            inline_data=types.Blob(
                mime_type=mime_type,
                data=image_bytes
            ),
            media_resolution={"level": media_resolution_value}
        )

        # 6. 发起 API 请求
        print(f">>> 发起 Gemini API 请求...")

        response = client.models.generate_content(
            model=model,
            contents=[
                types.Content(
                    role="user",
                    parts=[
                        types.Part(text=final_content),
                        image_part
                    ]
                )
            ],
            config=types.GenerateContentConfig(
                max_output_tokens=max_tokens,
                temperature=0.1,  # 数据提取需要精确，温度设低
                response_modalities=["TEXT"]
            )
        )

        # 7. 提取响应内容
        content = response.text.strip()
        print(f">>> Gemini API 返回成功，内容长度: {len(content)} 字符")

        # 8. 数据清洗与 JSON 解析
        cleaned_content = content
        if cleaned_content.startswith("```json"):
            cleaned_content = cleaned_content[7:]
        elif cleaned_content.startswith("```"):
            cleaned_content = cleaned_content[3:]

        if cleaned_content.endswith("```"):
            cleaned_content = cleaned_content[:-3]

        try:
            result_json = json.loads(cleaned_content.strip())
            result_json['used_api_key'] = api_key if api_key else "N/A"
            return result_json
        except json.JSONDecodeError:
            print(">>> 警告: Gemini 返回内容无法解析为 JSON")
            print(f">>> 原始返回内容: {content[:500]}...")
            print(f">>> 清理后内容: {cleaned_content[:500]}...")
            return {"raw_text": content, "used_api_key": api_key}

    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f">>> Gemini 图片分析系统错误: {e}")
        return {"error": str(e), "used_api_key": api_key if 'api_key' in locals() else "N/A"}


# 测试代码
if __name__ == "__main__":
    # 假设你有一张测试图片放在同一目录下，或者输入绝对路径
    test_image = "/Users/fengshenju/Desktop/aaa.jpg"  # 请修改为实际存在的图片路径

    # 从配置文件读取提示词
    my_prompt = CONFIG.get('prompt_instruction', '')

    # 只有当文件存在时才运行测试
    if os.path.exists(test_image):
        result = extract_data_from_image(test_image, my_prompt)
        print("\n>>> 最终提取结果:")
        print(json.dumps(result, indent=4, ensure_ascii=False))
    else:
        print(f"测试文件不存在: {test_image}，请修改 main 函数中的路径进行测试。")