#!/usr/bin/env python3
"""
HTML报告生成模块
"""
import os
import json
import datetime
import base64
from typing import Dict, Any


def generate_html_template():
    """生成HTML报告模板"""
    return """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>图片处理报告</title>
    <style>
        body {
            font-family: 'Microsoft YaHei', Arial, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f5f5f5;
        }
        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px;
            border-radius: 10px;
            margin-bottom: 20px;
        }
        .stats {
            display: flex;
            gap: 20px;
            margin-bottom: 20px;
        }
        .stat-card {
            background: white;
            padding: 15px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            flex: 1;
        }
        .stat-number {
            font-size: 2em;
            font-weight: bold;
            color: #667eea;
        }
        .image-card {
            background: white;
            margin-bottom: 20px;
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 2px 15px rgba(0,0,0,0.1);
        }
        .card-header {
            padding: 15px 20px;
            background-color: #f8f9fa;
            border-bottom: 1px solid #dee2e6;
            cursor: pointer;
        }
        .card-header:hover {
            background-color: #e9ecef;
        }
        .card-content {
            padding: 20px;
            display: none;
        }
        .status-success {
            color: #28a745;
            font-weight: bold;
        }
        .status-fail {
            color: #dc3545;
            font-weight: bold;
        }
        .info-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 15px;
            margin-bottom: 20px;
        }
        .info-section {
            background: #f8f9fa;
            padding: 15px;
            border-radius: 8px;
            border-left: 4px solid #667eea;
        }
        .info-title {
            font-weight: bold;
            color: #495057;
            margin-bottom: 10px;
        }
        .key-value {
            margin: 5px 0;
            padding: 3px 0;
        }
        .key {
            font-weight: bold;
            color: #6c757d;
        }
        .value {
            color: #495057;
        }
        .prompt-section {
            background: #fff3cd;
            padding: 15px;
            border-radius: 8px;
            border-left: 4px solid #ffc107;
            margin-top: 15px;
        }
        .prompt-content {
            max-height: 200px;
            overflow-y: auto;
            font-family: monospace;
            font-size: 0.9em;
            background: white;
            padding: 10px;
            border-radius: 4px;
            white-space: pre-wrap;
        }
        .record-list {
            background: #d1ecf1;
            padding: 15px;
            border-radius: 8px;
            border-left: 4px solid #17a2b8;
            margin-top: 15px;
        }
        .record-item {
            background: white;
            margin: 5px 0;
            padding: 10px;
            border-radius: 4px;
            font-size: 0.9em;
        }
        .toggle-icon {
            float: right;
            transition: transform 0.3s;
        }
        .expanded .toggle-icon {
            transform: rotate(180deg);
        }
        .image-container {
            text-align: center;
            margin: 15px 0;
        }
        .thumbnail {
            max-width: 200px;
            max-height: 150px;
            cursor: pointer;
            border: 2px solid #dee2e6;
            border-radius: 8px;
            transition: all 0.3s;
        }
        .thumbnail:hover {
            border-color: #667eea;
            box-shadow: 0 4px 15px rgba(102, 126, 234, 0.3);
        }

        /* 模态框样式优化 - 开始 */
        .modal {
            display: none;
            position: fixed;
            z-index: 1000;
            left: 0;
            top: 0;
            width: 100%;
            height: 100%;
            background-color: rgba(0,0,0,0.85); /*稍微加深背景*/
            backdrop-filter: blur(5px); /* 添加背景模糊效果 */
        }
        .modal-content {
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%); /* 完美居中 */
            width: auto;
            max-width: 95%; /* 左右留一点边距 */
            background: transparent;
            text-align: center;
        }
        .modal img {
            display: block;
            max-width: 100%;
            max-height: 85vh; /* 限制高度为屏幕高度的85%，防止溢出 */
            width: auto;      /* 保持原始比例 */
            height: auto;     /* 保持原始比例 */
            border-radius: 8px;
            box-shadow: 0 5px 25px rgba(0,0,0,0.5); /* 添加阴影 */
            margin: 0 auto;
        }
        /* 模态框样式优化 - 结束 */

        .close {
            position: absolute;
            top: -40px;
            right: 0;
            color: white;
            font-size: 35px;
            font-weight: bold;
            cursor: pointer;
            text-shadow: 0 2px 4px rgba(0,0,0,0.5); /* 增加文字阴影，防止背景太白看不清 */
            transition: color 0.2s;
        }
        .close:hover {
            color: #ccc;
        }
        .collapsible-section {
            background: #f8f9fa;
            border: 1px solid #dee2e6;
            border-radius: 8px;
            margin: 15px 0;
        }
        .collapsible-header {
            padding: 12px 15px;
            background: #e9ecef;
            border-bottom: 1px solid #dee2e6;
            cursor: pointer;
            user-select: none;
        }
        .collapsible-header:hover {
            background: #dee2e6;
        }
        .collapsible-content {
            padding: 15px;
            display: none;
            max-height: 400px;
            overflow-y: auto;
        }
        .copy-btn {
            background: #667eea;
            color: white;
            border: none;
            padding: 5px 10px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 0.8em;
            margin-left: 10px;
        }
        .copy-btn:hover {
            background: #5a6fd8;
        }
        .code-block {
            background: #f8f9fa;
            border: 1px solid #e9ecef;
            border-radius: 4px;
            padding: 15px;
            font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
            white-space: pre-wrap;
            word-wrap: break-word;
            font-size: 0.85em;
            line-height: 1.5;
            overflow-x: auto;
        }
        .markdown-content {
            background: #fafafa;
            border: 1px solid #e1e4e8;
            border-radius: 6px;
            padding: 16px;
            font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
            font-size: 0.85em;
            line-height: 1.6;
            overflow-x: auto;
        }
        .markdown-content .md-header {
            color: #0366d6;
            font-weight: bold;
            margin: 8px 0 4px 0;
        }
        .markdown-content .md-header1 { font-size: 1.2em; }
        .markdown-content .md-header2 { font-size: 1.1em; }
        .markdown-content .md-header3 { font-size: 1.0em; }
        .markdown-content .md-list {
            color: #6f42c1;
            margin: 2px 0;
        }
        .markdown-content .md-bold {
            color: #d73a49;
            font-weight: bold;
        }
        .markdown-content .md-code {
            background: #f6f8fa;
            color: #e36209;
            padding: 2px 4px;
            border-radius: 3px;
            font-family: monospace;
        }
        .json-content {
            background: #2d3748;
            color: #e2e8f0;
            border-radius: 6px;
            padding: 16px;
            font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
            font-size: 0.85em;
            line-height: 1.6;
            overflow-x: auto;
        }
        .json-key { color: #81c784; }
        .json-string { color: #ffb74d; }
        .json-number { color: #64b5f6; }
        .json-boolean { color: #f06292; }
        .json-null { color: #90a4ae; }
        .json-bracket { color: #ffffff; }

        .code-container {
            display: flex;
            border-radius: 6px;
            overflow: hidden;
        }
        .formatted-content {
            flex: 1;
            overflow-x: auto;
        }
    </style>
</head>
<body>
    <div class="header">
        <h1>图片处理报告</h1>
        <p>生成时间: %s</p>
    </div>

    <div class="stats">
        <div class="stat-card">
            <div class="stat-number" id="total-count">0</div>
            <div>总处理数量</div>
        </div>
        <div class="stat-card">
            <div class="stat-number" id="success-count" style="color: #28a745;">0</div>
            <div>解析成功</div>
        </div>
        <div class="stat-card">
            <div class="stat-number" id="match-success-count" style="color: #17a2b8;">0</div>
            <div>匹配成功</div>
        </div>
        <div class="stat-card">
            <div class="stat-number" id="fail-count" style="color: #dc3545;">0</div>
            <div>处理失败</div>
        </div>
    </div>

    <div id="results-container">
        </div>

    <div id="imageModal" class="modal">
        <div class="modal-content">
            <span class="close" onclick="closeModal()">&times;</span>
            <img id="modalImg" src="">
        </div>
    </div>

    <script>
        // 卡片折叠功能
        function toggleCard(element) {
            const content = element.nextElementSibling;
            const icon = element.querySelector('.toggle-icon');
            if (content.style.display === 'none' || content.style.display === '') {
                content.style.display = 'block';
                element.classList.add('expanded');
            } else {
                content.style.display = 'none';
                element.classList.remove('expanded');
            }
        }

        function updateStats() {
            const cards = document.querySelectorAll('.image-card');
            let total = cards.length;
            let success = 0;
            let matchSuccess = 0;
            let fail = 0;

            cards.forEach(card => {
                const parseStatus = card.querySelector('.parse-status').textContent;
                const matchStatus = card.querySelector('.match-status').textContent;

                if (parseStatus.includes('成功')) {
                    success++;
                }
                if (matchStatus.includes('成功')) {
                    matchSuccess++;
                }
                if (parseStatus.includes('失败') || matchStatus.includes('失败')) {
                    fail++;
                }
            });

            document.getElementById('total-count').textContent = total;
            document.getElementById('success-count').textContent = success;
            document.getElementById('match-success-count').textContent = matchSuccess;
            document.getElementById('fail-count').textContent = fail;
        }


        // 图片放大功能
        function showModal(imgSrc) {
            document.getElementById('imageModal').style.display = 'block';
            document.getElementById('modalImg').src = imgSrc;
        }

        function closeModal() {
            document.getElementById('imageModal').style.display = 'none';
        }

        // 点击模态框外部关闭
        window.onclick = function(event) {
            const modal = document.getElementById('imageModal');
            if (event.target == modal) {
                modal.style.display = 'none';
            }
        }

        // 可折叠区域功能
        function toggleCollapsible(element) {
            const content = element.nextElementSibling;
            const icon = element.querySelector('.toggle-icon');
            if (content.style.display === 'none' || content.style.display === '') {
                content.style.display = 'block';
                if (icon) icon.textContent = '▲';
            } else {
                content.style.display = 'none';
                if (icon) icon.textContent = '▼';
            }
        }

        // 安全的复制函数
        function copyFromElement(elementId, event) {
            var element = document.getElementById(elementId);
            if (!element) {
                console.error('找不到元素:', elementId);
                return;
            }
            var text = element.value || element.textContent;

            navigator.clipboard.writeText(text).then(function() {
                var btn = event.target;
                var originalText = btn.textContent;
                btn.textContent = '✅ 已复制!';
                setTimeout(() => {
                    btn.textContent = originalText;
                }, 1000);
            }).catch(function(err) {
                console.error('复制失败: ', err);
                alert('复制失败，请手动选择文本复制');
            });
        }

        // Markdown格式化
        function formatMarkdown(text) {
            return text
                .replace(/^(#{1,3})\\s+(.+)$/gm, (match, hashes, content) => {
                    const level = hashes.length;
                    return `<div class="md-header md-header${level}">${hashes} ${content}</div>`;
                })
                .replace(/^\\s*[-*]\\s+(.+)$/gm, '<div class="md-list">• $1</div>')
                .replace(/\\*\\*(.+?)\\*\\*/g, '<span class="md-bold">$1</span>')
                .replace(/`(.+?)`/g, '<span class="md-code">$1</span>')
                .replace(/\\n/g, '<br>');
        }

        // JSON格式化和高亮
        function formatJSON(jsonStr) {
            try {
                const parsed = JSON.parse(jsonStr);
                const formatted = JSON.stringify(parsed, null, 2);

                return formatted
                    .replace(/(\"[\\w_-]+\"):/g, '<span class=\"json-key\">$1</span>:')
                    .replace(/:​\\s*\"([^\"]*)\"/g, ': <span class=\"json-string\">\"$1\"</span>')
                    .replace(/:​\\s*(\\d+\\.?​\\d*)/g, ': <span class=\"json-number\">$1</span>')
                    .replace(/:​\\s*(true|false)/g, ': <span class=\"json-boolean\">$1</span>')
                    .replace(/:​\\s*(null)/g, ': <span class=\"json-null\">$1</span>')
                    .replace(/([{}\\[\\]])/g, '<span class=\"json-bracket\">$1</span>');
            } catch (e) {
                return jsonStr;
            }
        }

        // 页面加载后执行初始化
        window.addEventListener('DOMContentLoaded', function() {
            // 更新统计
            updateStats();

            // 格式化所有Markdown内容
            document.querySelectorAll('.markdown-raw').forEach(function(element) {
                const rawText = element.textContent;
                const formatted = formatMarkdown(rawText);
                element.innerHTML = formatted;
                element.className = 'markdown-content';
            });

            // 格式化所有JSON内容
            document.querySelectorAll('.json-raw').forEach(function(element) {
                const rawText = element.textContent;
                const formatted = formatJSON(rawText);

                const container = document.createElement('div');
                container.className = 'code-container';

                const contentDiv = document.createElement('div');
                contentDiv.className = 'formatted-content json-content';
                contentDiv.innerHTML = formatted;

                container.appendChild(contentDiv);

                element.parentNode.replaceChild(container, element);
            });
        });
    </script>
</body>
</html>
"""

import uuid  # 记得在文件顶部导入 uuid

import uuid  # 确保文件头部导入了 uuid


def create_image_result_card(image_name: str, result_data: Dict[str, Any]) -> str:
    """创建单个图片结果卡片HTML"""

    # 生成唯一ID
    unique_id = str(uuid.uuid4())[:8]
    prompt_id = f"prompt-{unique_id}"
    json_id = f"json-{unique_id}"

    # 解析状态
    parse_status = "解析成功" if result_data.get('parse_success', False) else "解析失败"
    parse_status_class = "status-success" if result_data.get('parse_success', False) else "status-fail"

    # 匹配状态
    match_status = "匹配成功" if result_data.get('match_success', False) else "匹配失败"
    match_status_class = "status-success" if result_data.get('match_success', False) else "status-fail"

    # 图片展示
    image_base64 = result_data.get('image_base64', '')
    image_section = ""
    if image_base64:
        image_section = f"""
            <div class="image-container">
                <img src="{image_base64}" alt="{image_name}" class="thumbnail" onclick="showModal('{image_base64}')">
                <div style="margin-top: 5px; font-size: 0.9em; color: #6c757d;">点击图片放大查看</div>
            </div>
        """

    # 解析数据
    parsed_data = result_data.get('parsed_data', {})

    # 基本信息
    basic_info = f"""
        <div class="key-value">
            <span class="key">交付日期:</span> <span class="value">{parsed_data.get('delivery_date', 'N/A')}</span>
        </div>
        <div class="key-value">
            <span class="key">采购商:</span> <span class="value">{parsed_data.get('buyer_name', 'N/A')}</span>
        </div>
        <div class="key-value">
            <span class="key">供应商:</span> <span class="value">{parsed_data.get('supplier_name', 'N/A')}</span>
        </div>
        <div class="key-value">
            <span class="key">码单号:</span> <span class="value">{parsed_data.get('delivery_order_number', 'N/A')}</span>
        </div>
        <div class="key-value">
            <span class="key">最终款号:</span> <span class="value">{result_data.get('final_style', 'N/A')}</span>
        </div>
        <div class="key-value">
            <span class="key">LLM尝试次数:</span> <span class="value">{result_data.get('retry_count', 1)}</span>
        </div>
        <div class="key-value">
            <span class="key">匹配使用的API Key:</span> <span class="value">{result_data.get('used_api_key', 'N/A')}</span>
        </div>
        <div class="key-value">
            <span class="key">DMX重新识别:</span> <span class="value">{'是' if result_data.get('used_dmx_recheck', False) else '否'}</span>
        </div>
        <div class="key-value">
            <span class="key">失败原因:</span> <span class="value">{result_data.get('failure_reason', 'N/A')}</span>
        </div>
    """

    # 商品明细
    items_info = ""
    items = parsed_data.get('items', [])
    if items:
        for i, item in enumerate(items):
            items_info += f"""
                <div class="record-item">
                    <strong>商品 {i + 1}:</strong><br>
                    数量: {item.get('qty', 'N/A')} | 
                    单价: {item.get('price', 'N/A')} | 
                    单位: {item.get('unit', 'N/A')}<br>
                    商品编码: {item.get('raw_style_text', 'N/A')}<br>
                    商品描述: {item.get('product_description', 'N/A')}
                </div>
            """

    # 最终确定的查询款号
    final_style = result_data.get('final_style', '')
    candidates_info = ""
    if final_style:
        candidates_info = f"""
            <div class="record-item">
                <strong>RPA查询款号: {final_style}</strong>
            </div>
        """
    else:
        candidates_info = '<div class="record-item">未确定查询款号</div>'

    # 匹配提示词
    match_prompt = result_data.get('match_prompt', 'N/A')

    # 匹配结果处理
    match_result = result_data.get('match_result', {})
    original_records = result_data.get('original_records', [])
    match_records_info = ""

    if match_result:
        if result_data.get('match_success', False):
            matched_ids = match_result.get('matched_record_ids', [])

            # --- 【核心修改】构建 ID -> (索引, 记录数据) 的映射 ---
            # i+1 表示人类可读的行号（从1开始）
            records_map = {
                rec.get('Id'): (i + 1, rec)
                for i, rec in enumerate(original_records)
            }

            match_records_info = f"<div class='record-item'><strong>匹配成功记录:</strong></div>"

            for matched_id in matched_ids:
                # 获取 (行号, 数据) 元组
                found = records_map.get(matched_id)

                if found:
                    index, record = found  # 解包
                    match_records_info += f"""
                        <div class='record-item' style='border-left: 3px solid #28a745; margin: 5px 0; padding-left: 10px;'>
                            <div style="margin-bottom: 6px;">
                                <span style="background-color: #28a745; color: white; padding: 2px 6px; border-radius: 4px; font-size: 0.8em; margin-right: 5px;">
                                    第 {index} 条
                                </span>
                                <span style="font-family: monospace; color: #666; font-size: 0.9em;">ID: {matched_id}</span>
                            </div>
                            <strong>产品名称:</strong> {record.get('DBProductName', 'N/A')}<br>
                            <strong>物料规格:</strong> {record.get('MaterialSpec', 'N/A')}<br>
                            <strong>供应商:</strong> {record.get('DBSupplierSpName', 'N/A')}<br>
                            <strong>订单编号:</strong> {record.get('OrderReqCode', 'N/A')}<br>
                            <strong>需求数量:</strong> {record.get('ReqQty', 'N/A')}
                        </div>
                    """
                else:
                    match_records_info += f"<div class='record-item'>ID: {matched_id} (原始列表中未找到该ID，可能数据未保存完整)</div>"
        else:
            reason = match_result.get('global_reason', 'N/A')
            match_records_info = f"<div class='record-item' style='border-left: 3px solid #dc3545; padding-left: 10px;'><strong>失败原因:</strong> {reason}</div>"

            # 详细分析
            details = match_result.get('detail_analysis', [])
            if details:
                for detail in details:
                    match_records_info += f"""
                        <div class='record-item'>
                            <strong>项目 {detail.get('ocr_item_index', 'N/A') + 1}:</strong> {detail.get('ocr_desc', 'N/A')}<br>
                            <strong>匹配逻辑:</strong> {detail.get('match_logic', 'N/A')}<br>
                            <strong>备注:</strong> {detail.get('notes', 'N/A')}
                        </div>
                    """

    # 准备 JSON 字符串
    json_str = json.dumps(match_result, ensure_ascii=False, indent=2)

    card_html = f"""
    <div class="image-card">
        <div class="card-header" onclick="toggleCard(this)">
            <strong>{image_name}</strong>
            <span class="parse-status {parse_status_class}">{parse_status}</span> | 
            <span class="match-status {match_status_class}">{match_status}</span>
            <span class="toggle-icon">▼</span>
        </div>
        <div class="card-content">
            {image_section}
            <div class="info-grid">
                <div class="info-section">
                    <div class="info-title">基本信息</div>
                    {basic_info}
                </div>
                <div class="info-section">
                    <div class="info-title">RPA查询款号</div>
                    {candidates_info}
                </div>
            </div>

            <div class="record-list">
                <div class="info-title">商品明细</div>
                {items_info if items_info else '<div class="record-item">暂无商品明细</div>'}
            </div>

            <textarea id="{prompt_id}" style="display:none;">{match_prompt}</textarea>

            <div class="collapsible-section">
                <div class="collapsible-header" onclick="toggleCollapsible(this)">
                    <strong>匹配提示词</strong> (字符数: {len(match_prompt)})
                    <button class="copy-btn" onclick="event.stopPropagation(); copyFromElement('{prompt_id}', event)">复制</button>
                    <span class="toggle-icon">▼</span>
                </div>
                <div class="collapsible-content">
                    <div class="markdown-raw">{match_prompt}</div>
                </div>
            </div>

            <div class="record-list">
                <div class="info-title">匹配结果</div>
                {match_records_info if match_records_info else '<div class="record-item">暂无匹配结果</div>'}
            </div>

            <textarea id="{json_id}" style="display:none;">{json_str}</textarea>

            <div class="collapsible-section">
                <div class="collapsible-header" onclick="toggleCollapsible(this)">
                    <strong>LLM原始返回</strong>
                    <button class="copy-btn" onclick="event.stopPropagation(); copyFromElement('{json_id}', event)">复制JSON</button>
                    <span class="toggle-icon">▼</span>
                </div>
                <div class="collapsible-content">
                    <div class="json-raw">{json_str}</div>
                </div>
            </div>
        </div>
    </div>
    """

    return card_html

def update_html_report(report_path: str, image_name: str, result_data: Dict[str, Any]):
    """更新HTML报告"""
    
    # 如果传入的是目录路径，自动添加默认文件名
    if os.path.isdir(report_path):
        report_path = os.path.join(report_path, 'report.html')
    elif not report_path.endswith('.html'):
        # 如果路径不以.html结尾且不是目录，也添加默认文件名
        report_path = os.path.join(report_path, 'report.html')
    
    # 生成当前结果的卡片
    new_card = create_image_result_card(image_name, result_data)
    
    # 如果报告文件不存在，创建新文件
    if not os.path.exists(report_path):
        # 确保目录存在
        os.makedirs(os.path.dirname(report_path), exist_ok=True)
        
        # 生成初始HTML
        timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        html_content = generate_html_template().replace('%s', timestamp)
        
        # 插入第一个卡片
        html_content = html_content.replace(
            '<div id="results-container">',
            f'<div id="results-container">\n{new_card}'
        )
        
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
    else:
        # 读取现有文件
        with open(report_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
        
        # 在results-container中插入新卡片
        insert_pos = html_content.find('<div id="results-container">') + len('<div id="results-container">')
        html_content = html_content[:insert_pos] + f'\n{new_card}' + html_content[insert_pos:]
        
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(html_content)

def image_to_base64(image_path: str) -> str:
    """将图片转换为base64编码"""
    try:
        if not os.path.exists(image_path):
            return ""
        
        with open(image_path, 'rb') as img_file:
            img_data = img_file.read()
            base64_str = base64.b64encode(img_data).decode('utf-8')
            
        # 根据文件扩展名确定MIME类型
        ext = os.path.splitext(image_path)[1].lower()
        mime_type = {
            '.jpg': 'image/jpeg',
            '.jpeg': 'image/jpeg', 
            '.png': 'image/png',
            '.gif': 'image/gif'
        }.get(ext, 'image/jpeg')
        
        return f"data:{mime_type};base64,{base64_str}"
    except Exception as e:
        print(f"图片转换失败: {e}")
        return ""

def collect_result_data(image_name: str, parsed_data: Dict[str, Any], final_style: str, 
                       match_prompt: str = "", match_result: Dict[str, Any] = None, 
                       original_records: list = None, image_path: str = "", retry_count: int = 1, failure_reason: str = "", used_api_key: str = "") -> Dict[str, Any]:
    """收集处理结果数据"""
    
    parse_success = parsed_data and 'error' not in parsed_data
    match_success = match_result and match_result.get('status') == 'success' if match_result else False
    
    # 转换图片为base64
    image_base64 = image_to_base64(image_path) if image_path else ""
    
    # 优先从match_result中获取API key，如果没有则使用传入的参数
    actual_used_key = match_result.get('used_api_key', used_api_key) if match_result else used_api_key
    
    # 检查是否使用了DMX重新识别
    used_dmx_recheck = parsed_data.get('used_dmx_for_date_check', False) if parsed_data else False
    dmx_recheck_failed = parsed_data.get('dmx_recheck_failed', False) if parsed_data else False
    
    return {
        'image_name': image_name,
        'parse_success': parse_success,
        'parsed_data': parsed_data,
        'final_style': final_style,
        'match_prompt': match_prompt,
        'match_result': match_result,
        'match_success': match_success,
        'original_records': original_records or [],
        'image_base64': image_base64,
        'retry_count': retry_count,
        'failure_reason': failure_reason,
        'used_api_key': actual_used_key,
        'used_dmx_recheck': used_dmx_recheck,
        'dmx_recheck_failed': dmx_recheck_failed
    }