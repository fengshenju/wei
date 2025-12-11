# -*- coding: utf-8 -*-
"""
RPA工具类 - 提供通用的浏览器操作和数据处理方法
"""
import json
import re
import datetime
from datetime import timedelta
import time

# 导入配置和工具
try:
    from app_config import CONFIG
    from utils.util_llm import call_llm_text, call_dmxllm_text
except ImportError as e:
    print(f"!!! RPAUtils 导入模块失败: {e}")
    exit(1)


class RPAUtils:
    """RPA工具类 - 无状态设计，支持并发调用"""

    @staticmethod
    def fill_date_input(scope, selector, date_value,
                        remove_readonly=True,
                        trigger_events=True,
                        use_enter_key=True,
                        click_body_after=True,
                        scroll_to_see=True,
                        timeout=0.5):
        """
        [修复版] 兼容 Element 和 Tab/Frame 的暴力日期填充
        解决了 'ChromiumElement object has no attribute page' 的崩溃问题
        """
        try:
            if not date_value or not date_value.strip():
                return False

            # 1. 查找元素
            date_input = scope.ele(selector, timeout=timeout)
            if not date_input:
                # 尝试稍微等待后再次查找
                time.sleep(0.2)
                date_input = scope.ele(selector, timeout=0.2)
                if not date_input:
                    return False

            # 滚动到可见区域
            if scroll_to_see:
                date_input.scroll.to_see()
                time.sleep(0.1)

            print(f"   -> [RPAUtils] 正在填入日期: {date_value}")

            # ============================================================
            # [关键修复] 安全地获取操作驱动器 (Driver)
            # ============================================================
            driver = None
            # 1. 优先尝试从元素本身获取 page 对象 (新版 DrissionPage)
            if hasattr(date_input, 'page'):
                driver = date_input.page
            # 2. 旧版本可能是 owner
            elif hasattr(date_input, 'owner'):
                driver = date_input.owner

            # 3. 如果 scope 本身就是 Tab/Frame (具有 actions 属性)，则直接使用 scope
            if not driver and hasattr(scope, 'actions'):
                driver = scope

            # (A) 移除 readonly (直接在元素上操作，不需要 driver)
            if remove_readonly:
                date_input.run_js('this.removeAttribute("readonly");')

            # (B) 清空并输入
            date_input.clear()
            time.sleep(0.1)
            date_input.input(date_value)
            time.sleep(0.2)

            # (C) 模拟按下"回车"键 (触发日历关闭和验证)
            if use_enter_key:
                if driver:
                    # 使用找到的 driver 执行动作链
                    driver.actions.key_down('ENTER').key_up('ENTER')
                else:
                    # [兜底] 如果找不到 driver，使用 JS 模拟回车事件
                    # print("   ℹ️ 使用 JS 模拟回车")
                    date_input.run_js(
                        'this.dispatchEvent(new KeyboardEvent("keydown", {bubbles:true, keyCode:13, key:"Enter"}));')
                time.sleep(0.2)

            # (D) 模拟失焦 (触发 onchange)
            if click_body_after:
                if driver:
                    driver.run_js('document.body.click();')
                else:
                    # 如果没有 driver，就让输入框自己失去焦点
                    date_input.run_js('this.blur();')

                # 再次点击确保状态重置
                try:
                    date_input.click()
                except:
                    pass
                time.sleep(0.1)

            return True

        except Exception as e:
            print(f"   !!! 日期填充异常: {e}")
            # 最后的暴力兜底：直接修改 value 并触发 change
            try:
                if date_input:
                    js = f'this.removeAttribute("readonly"); this.value="{date_value}"; this.dispatchEvent(new Event("change"));'
                    date_input.run_js(js)
            except:
                pass
            return False

    @staticmethod
    def search_and_select_from_popup(scope, trigger_selector, search_input_selector, 
                                    table_id, search_value, item_name="项目", timeout=1):
        """
        通用的弹框搜索选择方法
        
        :param scope: 操作作用域对象（tab、frame等）
        :param trigger_selector: 触发按钮的选择器
        :param search_input_selector: 搜索框选择器
        :param table_id: 结果表格的ID
        :param search_value: 要搜索的值
        :param item_name: 项目类型名称（用于日志显示）
        :param timeout: 查找元素的超时时间
        :return: 是否成功选择
        """
        try:
            if not search_value or not search_value.strip():
                print(f"⚠️ {item_name}值为空，跳过选择")
                return False
                
            # 1. 点击触发按钮打开弹框
            trigger_btn = scope.ele(trigger_selector, timeout=timeout)
            if not trigger_btn or not trigger_btn.states.is_displayed:
                print(f"⚠️ 未找到{item_name}触发按钮: {trigger_selector}")
                return False
                
            trigger_btn.click()
            time.sleep(0.5)
            
            # 2. 定位搜索输入框
            search_box = scope.ele(search_input_selector, timeout=timeout)
            if not search_box or not search_box.states.is_displayed:
                print(f"⚠️ 未找到{item_name}搜索框: {search_input_selector}")
                # 如果找不到搜索框，关闭弹框
                trigger_btn.click()
                return False
                
            # 3. 清空并输入搜索值
            search_box.clear()
            search_box.input(search_value)
            time.sleep(0.2)
            
            # 4. 按Enter键触发搜索
            scope.actions.key_down('ENTER').key_up('ENTER')
            time.sleep(0.5)
            
            # 5. 在结果表格中查找匹配的单元格
            # 优先使用精确匹配，如果没找到再使用包含匹配
            target_td_xpath_exact = f'x://table[@id="{table_id}"]//tbody//tr//td[text()="{search_value}"]'
            target_td = scope.ele(target_td_xpath_exact, timeout=timeout)
            
            # 如果精确匹配失败，再尝试包含匹配
            if not target_td:
                target_td_xpath_contains = f'x://table[@id="{table_id}"]//tbody//tr//td[contains(text(), "{search_value}")]'
                target_td = scope.ele(target_td_xpath_contains, timeout=timeout)
            
            if target_td:
                print(f"  -> 找到{item_name} [{search_value}]，执行双击选择...")
                # 6. 双击匹配的单元格进行选择
                js_code = """this.click(); this.dispatchEvent(new MouseEvent('dblclick', {bubbles: true, cancelable: true, view: window}));"""
                target_td.run_js(js_code)
                time.sleep(0.5)
                return True
            else:
                print(f"  ⚠️ {item_name}列表中未搜索到: {search_value}")
                # 7. 搜索失败，关闭弹框
                trigger_btn.click()
                return False
                
        except Exception as e:
            print(f"!!! {item_name}选择异常: {e}")
            return False

    @staticmethod
    def find_element_in_iframes(tab, selector, max_retries=10, retry_interval=0.5, 
                               timeout=0.2, return_frame=False, validate_visibility=True,
                               fallback_selectors=None):
        """
        在所有可见iframe中查找元素的通用方法
        
        :param tab: 浏览器标签页对象
        :param selector: CSS/XPath选择器（主选择器）
        :param max_retries: 最大重试次数
        :param retry_interval: 重试间隔时间(秒)
        :param timeout: 元素查找超时时间(秒)
        :param return_frame: 是否同时返回找到元素的iframe
        :param validate_visibility: 是否验证元素可见性
        :param fallback_selectors: 备用选择器列表
        :return: 元素对象 或 (元素对象, iframe对象) 元组
        """
        target_element = None
        target_frame = None
        
        # 构建选择器列表
        selectors = [selector]
        if fallback_selectors:
            selectors.extend(fallback_selectors)
        
        for retry_count in range(max_retries):
            print(f">>> [调试] 第 {retry_count + 1}/{max_retries} 次尝试查找元素: {selector}")
            iframe_count = 0
            visible_iframe_count = 0
            
            for frame in tab.eles('tag:iframe'):
                iframe_count += 1
                # 跳过不可见的iframe
                if not frame.states.is_displayed:
                    continue
                visible_iframe_count += 1
                print(f">>> [调试] 检查第 {visible_iframe_count} 个可见iframe (总计 {iframe_count} 个iframe)")
                
                try:
                    frame_src = frame.attr('src') or 'unknown'
                    print(f">>> [调试] iframe src: {frame_src}")
                except:
                    print(f">>> [调试] 无法获取iframe src")
                    pass
                    
                # 尝试所有选择器
                for sel in selectors:
                    try:
                        print(f">>> [调试] 在iframe中尝试选择器: {sel}")
                        # 在当前iframe中查找元素
                        element = frame.ele(sel, timeout=timeout)
                        
                        # 验证元素存在且可见（可选）
                        if element:
                            print(f">>> [调试] 找到元素，验证可见性: validate_visibility={validate_visibility}")
                            if validate_visibility and not element.states.is_displayed:
                                print(f">>> [调试] 元素不可见，跳过")
                                continue
                            
                            print(f">>> [调试] ✅ 成功找到可用元素!")
                            target_element = element
                            target_frame = frame
                            break
                        else:
                            print(f">>> [调试] 未找到元素")
                            
                    except Exception as e:
                        # 忽略单个选择器的查找异常，继续下一个
                        print(f">>> [调试] 选择器查找异常: {e}")
                        continue
                
                # 如果找到元素，跳出frame循环
                if target_element:
                    break
                    
            # 找到元素则跳出重试循环
            if target_element:
                break
                
            # 未找到则等待后重试
            print(f">>> [调试] 第 {retry_count + 1} 次尝试失败，等待 {retry_interval} 秒后重试...")
            time.sleep(retry_interval)
        
        # 根据需求返回不同格式
        if target_element:
            print(f">>> [调试] 最终成功找到元素")
        else:
            print(f">>> [调试] ❌ 所有重试失败，未找到元素")
            
        if return_frame:
            return (target_element, target_frame)
        else:
            return target_element

    @staticmethod
    def input_text_char_by_char(input_element, text_value, char_interval=0.2, 
                               click_first=True, clear_first=True, pre_clear_sleep=0):
        """
        逐字符输入文本的通用方法
        
        :param input_element: 输入框元素
        :param text_value: 要输入的文本
        :param char_interval: 字符间隔时间(秒)
        :param click_first: 是否先点击输入框
        :param clear_first: 是否先清空输入框
        :param pre_clear_sleep: 清空后等待时间(秒)
        :return: 是否成功输入
        """
        try:
            if not text_value or not input_element:
                return False
                
            # 点击输入框
            if click_first:
                input_element.click()
                time.sleep(0.2)
            
            # 清空输入框
            if clear_first:
                input_element.clear()
                if pre_clear_sleep > 0:
                    time.sleep(pre_clear_sleep)
            
            # 逐字符输入
            for char in text_value:
                input_element.input(char, clear=False)
                time.sleep(char_interval)
            
            return True
            
        except Exception as e:
            print(f"!!! 逐字符输入失败: {e}")
            return False

    @staticmethod
    def handle_alert_confirmation(tab, timeout=3, double_confirm=False, 
                                 second_timeout=2, middle_sleep=1):
        """
        处理弹窗确认的通用方法
        
        :param tab: 浏览器标签页对象
        :param timeout: 第一次等待超时时间(秒)
        :param double_confirm: 是否进行双重确认
        :param second_timeout: 第二次等待超时时间(秒)
        :param middle_sleep: 双重确认中间等待时间(秒)
        :return: (first_handled, second_handled) 处理结果元组
        """
        first_handled = False
        second_handled = False
        
        try:
            # 第一次确认
            if tab.wait.alert(timeout=timeout):
                tab.alert.accept()
                first_handled = True
        except Exception:
            pass
        
        # 双重确认
        if double_confirm:
            time.sleep(middle_sleep)
            try:
                if tab.wait.alert(timeout=second_timeout):
                    tab.alert.accept()
                    second_handled = True
            except Exception:
                pass
        
        return (first_handled, second_handled)

    @staticmethod
    def navigate_to_menu(tab, main_menu_text, sub_menu_text, timeout=3):
        """
        统一的菜单导航方法
        :param tab: 浏览器标签页
        :param main_menu_text: 一级菜单文本（如"物料"、"财务"）
        :param sub_menu_text: 二级菜单文本（如"物料采购需求"）
        :param timeout: 超时时间
        :return: 是否导航成功
        """
        try:
            print(f'>>> 正在定位"{main_menu_text}"菜单...')
            main_btn = tab.ele(f'x://div[contains(@class, "title") and .//div[contains(text(), "{main_menu_text}")]]')
            
            if main_btn:
                print(f'>>> 找到"{main_menu_text}"菜单，正在点击...')
                main_btn.click()
                time.sleep(0.5)
            else:
                print(f'!!! 未找到{main_menu_text}菜单')
                return False

            print(f'>>> 正在定位二级菜单"{sub_menu_text}"...')
            sub_menu_btn = tab.wait.ele_displayed(f'x://a[contains(text(), "{sub_menu_text}")]', timeout=timeout)
            
            if sub_menu_btn:
                sub_menu_btn.click()
                print(f'✅ 成功点击"{sub_menu_text}"')
                time.sleep(0.5)
                return True
            else:
                print(f'⚠️ 未检测到二级菜单展开，尝试重新点击"{main_menu_text}"...')
                main_btn.click()
                time.sleep(1)
                sub_menu_btn = tab.wait.ele_displayed(f'x://a[contains(text(), "{sub_menu_text}")]', timeout=timeout)
                if sub_menu_btn:
                    sub_menu_btn.click()
                    print("✅ 重试后成功点击")
                    time.sleep(0.5)
                    return True
                else:
                    print("!!! 错误：无法展开二级菜单，请检查页面遮挡或网络卡顿。")
                    return False
                    
        except Exception as e:
            print(f"!!! 菜单导航异常: {e}")
            return False

    @staticmethod
    def search_with_network_listen(tab, input_element, target_url, timeout=10, 
                                 success_message="搜索成功", retry_on_concurrent=False, 
                                 retry_delay=5, auto_stop_listen=True):
        """
        网络监听搜索通用方法
        :param tab: 浏览器标签页  
        :param input_element: 输入框元素
        :param target_url: 监听的目标URL关键字
        :param timeout: 等待超时时间
        :param success_message: 成功时的日志消息
        :param retry_on_concurrent: 是否在并发限制时重试
        :param retry_delay: 重试延迟时间  
        :param auto_stop_listen: 是否自动停止监听
        :return: 网络响应包对象或None
        """
        try:
            tab.listen.start(targets=target_url)
            print(f">>> 已开启网络监听，目标: {target_url}")
            
            input_element.run_js("""
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                arguments[0].dispatchEvent(new KeyboardEvent("keydown", {bubbles:true, keyCode:13, key:"Enter"}));
                arguments[0].dispatchEvent(new KeyboardEvent("keyup", {bubbles:true, keyCode:13, key:"Enter"}));
            """, input_element)
            print("✅ 输入完毕并回车")
            
            res_packet = tab.listen.wait(timeout=timeout)
            
            if res_packet:
                print(f"✅ {success_message}: {res_packet.url}")
                
                if retry_on_concurrent:
                    response_body = res_packet.response.body
                    msg = response_body.get('msg', '')
                    if '上一个相同请求未结束' in msg or '请勿重复请求' in msg:
                        print(f"⚠️ 触发系统并发限制: {msg}")
                        print(f">>> 正在等待 {retry_delay} 秒后重试搜索...")
                        time.sleep(retry_delay)
                        
                        input_element.run_js("""
                            arguments[0].dispatchEvent(new KeyboardEvent("keydown", {bubbles:true, keyCode:13, key:"Enter"}));
                            arguments[0].dispatchEvent(new KeyboardEvent("keyup", {bubbles:true, keyCode:13, key:"Enter"}));
                        """, input_element)
                        res_packet = tab.listen.wait(timeout=timeout)
                        if res_packet:
                            print(f"✅ 重试后{success_message}")
                
                return res_packet
            else:
                print(f"⚠️ 网络监听超时，未捕获到响应")
                return None
                
        except Exception as e:
            print(f"!!! 网络监听搜索异常: {e}")
            return None
        finally:
            if auto_stop_listen:
                try:
                    tab.listen.stop()
                except:
                    pass

    @staticmethod
    def select_matched_checkboxes(tab, matched_ids):
        """勾选匹配的记录复选框"""
        print(f">>> 开始勾选匹配的记录: {len(matched_ids)} 条")
        for record_id in matched_ids:
            try:
                checkbox_selector = f'x://tr[.//a[contains(@data-sub-html, "{record_id}")]]//input[contains(@class, "ckbox")]'
                checkbox = RPAUtils.find_element_in_iframes(
                    tab=tab,
                    selector=checkbox_selector,
                    max_retries=1,
                    timeout=0.2,
                    validate_visibility=True
                )
                
                if checkbox:
                    if not checkbox.states.is_checked:
                        checkbox.click()
                        print(f"✅ 已勾选记录: {record_id}")
                else:
                    print(f"⚠️ 未找到记录 {record_id} 对应的checkbox")
            except Exception as e:
                print(f"!!! 勾选记录 {record_id} 时发生异常: {e}")
        print(f">>> 勾选操作完成")

    @staticmethod
    def find_and_use_select_all_button(frame, fallback_checkbox_selector='x://tr//input[contains(@class, "ckbox")]'):
        """
        查找并使用全选按钮，失败则回退到逐个勾选
        
        :param frame: 操作的frame对象
        :param fallback_checkbox_selector: 回退时使用的复选框选择器
        :return: (success, selected_count) - 是否成功和勾选数量
        """
        try:
            # 尝试查找全选按钮
            select_all_btn = frame.ele('x://input[@type="checkbox" and contains(@onclick, "selectAll")]', timeout=0.5)
            if not select_all_btn:
                select_all_btn = frame.ele('x://button[contains(text(), "全选") or contains(text(), "选择全部")]', timeout=0.5)
            
            if select_all_btn and select_all_btn.states.is_displayed:
                select_all_btn.click()
                return (True, -1)  # -1 表示使用了全选按钮，无法精确计数
            else:
                # 回退方案：逐个勾选所有复选框
                checkboxes = frame.eles(fallback_checkbox_selector, timeout=1)
                if checkboxes:
                    selected_count = 0
                    for checkbox in checkboxes:
                        if checkbox.states.is_displayed:
                            checkbox.scroll.to_see()
                            if not checkbox.states.is_checked:
                                checkbox.click()
                                selected_count += 1
                                time.sleep(0.1)
                    return (True, selected_count)
                else:
                    return (False, 0)
                    
        except Exception as e:
            print(f"!!! 全选操作异常: {e}")
            return (False, 0)

    @staticmethod  
    def select_checkboxes_in_table_rows(frame, table_selector, checkbox_selector='css:input.ckbox', timeout=2):
        """
        在表格行中批量勾选复选框
        
        :param frame: 操作的frame对象
        :param table_selector: 表格选择器
        :param checkbox_selector: 复选框选择器
        :param timeout: 超时时间
        :return: 勾选的数量
        """
        try:
            rows = frame.eles(table_selector, timeout=timeout)
            count_selected = 0
            if rows:
                for row in rows:
                    if not row.states.is_displayed: 
                        continue
                    try:
                        row.scroll.to_see()
                        checkbox = row.ele(checkbox_selector, timeout=0.5)
                        if checkbox:
                            if not checkbox.states.is_checked:
                                checkbox.click()
                                count_selected += 1
                        time.sleep(0.1)
                    except:
                        pass
            return count_selected
        except Exception as e:
            print(f"!!! 表格复选框勾选异常: {e}")
            return 0

    @staticmethod
    def select_all_checkboxes_in_frame(frame, table_selector=None, checkbox_selector='css:input.ckbox', label=""):
        """
        在指定frame的表格中勾选所有复选框
        
        :param frame: 操作的frame对象
        :param table_selector: 表格行选择器，如果为None则直接查找复选框
        :param checkbox_selector: 复选框选择器
        :param label: 日志标签
        :return: 勾选的数量
        """
        selected_count = 0
        try:
            if table_selector:
                # 在指定表格中查找行
                current_rows = frame.eles(table_selector, timeout=3)
                if not current_rows: 
                    return 0
                for row in current_rows:
                    if not row.states.is_displayed: 
                        continue
                    try:
                        row.scroll.to_see()
                        ck = row.ele(checkbox_selector, timeout=0.5)
                        if ck:
                            if not ck.states.is_checked:
                                ck.click()
                                selected_count += 1
                        time.sleep(0.05)
                    except:
                        pass
            else:
                # 直接在frame中查找所有复选框
                checkboxes = frame.eles(checkbox_selector, timeout=3)
                for checkbox in checkboxes:
                    if checkbox.states.is_displayed:
                        checkbox.scroll.to_see()
                        if not checkbox.states.is_checked:
                            checkbox.click()
                            selected_count += 1
                        time.sleep(0.05)
        except Exception as e:
            if label:
                print(f"   !!! {label} 勾选异常: {e}")
            else:
                print(f"!!! 复选框勾选异常: {e}")
        return selected_count

    @staticmethod
    def execute_smart_match(parsed_data, records):
        """
        执行智能匹配核心逻辑
        """
        today = datetime.date.today()
        two_weeks_ago = today - timedelta(days=14)
        clean_records = RPAUtils.preprocess_records(records)

        ocr_items_with_index = []
        original_items = parsed_data.get('items', [])
        for idx, item in enumerate(original_items):
            item_copy = item.copy()
            item_copy['_index'] = idx
            ocr_items_with_index.append(item_copy)

        llm_input_ocr = {**parsed_data, "items": ocr_items_with_index}

        prompt_template = CONFIG.get('match_prompt_template')
        if not prompt_template:
            print("!!! 错误: 配置缺失 'match_prompt_template'")
            return {"status": "error", "reason": "配置缺失"}, "", 1

        final_prompt = prompt_template.format(
            current_date=today.strftime('%Y-%m-%d'),
            two_weeks_ago=two_weeks_ago.strftime('%Y-%m-%d'),
            parsed_data_json=json.dumps(parsed_data, ensure_ascii=False, indent=2),
            records_json=json.dumps(clean_records, ensure_ascii=False, indent=2)
        )

        max_retries = CONFIG.get('llm_match_max_retries', 3)
        match_result = None
        retry_count = 0

        for retry_count in range(1, max_retries + 1):
            if retry_count == 1:
                print(">>> 使用阿里通义千问进行首次匹配...")
                match_result = call_llm_text(final_prompt, retry_count - 1)
            else:
                print(">>> 使用DMX接口进行重试匹配...")
                match_result = call_dmxllm_text(final_prompt, retry_count - 1)

            if match_result and match_result.get('status') == 'success':
                return match_result, final_prompt, retry_count

            print(f">>> LLM匹配第{retry_count}次尝试失败")
            if retry_count < max_retries:
                time.sleep(2)

        return match_result, final_prompt, retry_count

    @staticmethod
    def preprocess_records(records):
        """预处理记录数据"""
        cleaned_records = []
        for rec in records:
            new_rec = rec.copy()
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

    @staticmethod
    def reconstruct_rpa_data(match_result, original_parsed_data, original_records):
        """重组RPA数据"""
        matched_tasks = []
        record_map = {rec['Id']: rec for rec in original_records}
        ocr_items_list = original_parsed_data.get('items', [])

        def get_item_by_index(idx):
            if isinstance(idx, int) and 0 <= idx < len(ocr_items_list):
                return ocr_items_list[idx]
            return None

        for match in match_result.get('direct_matches', []):
            rid = match.get('record_id')
            idx = match.get('ocr_index')
            target_record = record_map.get(rid)
            target_item = get_item_by_index(idx)
            if target_record and target_item:
                matched_tasks.append({
                    "match_type": "DIRECT",
                    "record": target_record,
                    "items": [target_item],
                    "ocr_context": original_parsed_data
                })

        for match in match_result.get('merge_matches', []):
            rid = match.get('record_id')
            indices = match.get('ocr_indices', [])
            target_record = record_map.get(rid)
            target_items = [get_item_by_index(i) for i in indices if get_item_by_index(i)]
            if target_record and target_items:
                matched_tasks.append({
                    "match_type": "MERGE",
                    "record": target_record,
                    "items": target_items,
                    "ocr_context": original_parsed_data
                })

        for match in match_result.get('split_matches', []):
            rid = match.get('record_id')
            idx = match.get('ocr_index')
            target_record = record_map.get(rid)
            target_item = get_item_by_index(idx)
            if target_record and target_item:
                matched_tasks.append({
                    "match_type": "SPLIT",
                    "record": target_record,
                    "items": [target_item],
                    "ocr_context": original_parsed_data
                })

        return matched_tasks

    @staticmethod
    @staticmethod
    def handle_purchase_order_popup(tab, timeout=5):
        """
        [增强版] 处理 Layui 合并采购确认弹窗
        策略：轮询主窗口和所有可见iframe，查找包含特定文本的弹窗确认按钮
        """
        target_text = "您选择了不同的款"
        # XPath: 找包含特定文本的 layui-layer -> 找下面的确定按钮 (.layui-layer-btn0)
        confirm_btn_xpath = (
            f'x://div[contains(@class, "layui-layer") and .//div[contains(text(), "{target_text}")]]'
            f'//a[contains(@class, "layui-layer-btn0")]'
        )

        print(f">>> [监测] 开始全域扫描“{target_text}”弹窗 (超时: {timeout}s)...")
        start_time = time.time()

        while time.time() - start_time < timeout:
            # --- 1. 尝试在主 Tab 中查找 ---
            try:
                btn = tab.ele(confirm_btn_xpath, timeout=0.1)
                if btn and btn.states.is_displayed:
                    print(f"✅ 在 [主窗口] 发现弹窗，正在点击确定...")
                    btn.click()
                    time.sleep(1.5)  # 等待弹窗消失和蒙层褪去
                    return True
            except Exception:
                pass

            # --- 2. 尝试在所有可见 Iframe 中查找 ---
            try:
                # 获取所有 iframe
                frames = tab.eles('tag:iframe')
                for frame in frames:
                    if not frame.states.is_displayed:
                        continue
                    try:
                        # 在当前 frame 查找
                        btn = frame.ele(confirm_btn_xpath, timeout=0.1)
                        if btn and btn.states.is_displayed:
                            print(f"✅ 在 [Iframe] 发现弹窗，正在点击确定...")
                            # 尝试滚动到可见（防止弹窗在视口外）
                            try:
                                btn.scroll.to_see()
                            except:
                                pass
                            btn.click()
                            time.sleep(1.5)
                            return True
                    except Exception:
                        continue
            except Exception:
                pass

            time.sleep(0.5)  # 轮询间隔

        # --- 3. 兜底诊断信息 ---
        print("⚠️ 超时未检测到合并采购弹窗。")
        # 简单检查一下页面上到底有没有 layui-layer (不管文本对不对)，用于调试
        try:
            layer_in_tab = tab.ele('css:.layui-layer', timeout=0.1)
            if layer_in_tab:
                print(f"   -> 调试: 主窗口存在 .layui-layer，但文本不匹配。内容: {layer_in_tab.text[:20]}...")
            else:
                # 检查 iframe 里有没有
                for frame in tab.eles('tag:iframe'):
                    if frame.states.is_displayed:
                        l = frame.ele('css:.layui-layer', timeout=0.1)
                        if l:
                            print(f"   -> 调试: Iframe中存在 .layui-layer，但文本不匹配。内容: {l.text[:20]}...")
                            break
        except:
            pass

        return False

    @staticmethod
    def fill_details_into_table(scope, structured_tasks, parsed_data=None):
        """
        填充物料明细数据到表格
        [最终增强版]
        1. 针对前端懒加载问题，实现了【失败自动重试机制】：搜不到->关闭->等待->重开->再搜。
        2. 保持了强力输入事件触发和双重XPath匹配策略。
        """
        print(f">>> 开始填充物料明细数据，共 {len(structured_tasks)} 条任务...")
        count_success = 0

        for task in structured_tasks:
            try:
                record_id = task['record'].get('Id')
                match_type = task['match_type']
                items = task['items']

                if not record_id or not items:
                    continue

                # --- 1. 定位行 (TR) ---
                tr_xpath = f'x://tr[.//input[@name="materialReqId" and @value="{record_id}"]]'
                tr = scope.ele(tr_xpath, timeout=1)

                if not tr:
                    print(f"   ⚠️ 未找到 ID 为 {record_id} 的行，跳过")
                    continue

                # 确保行可见
                tr.scroll.to_see()

                # --- 2. 计算填入的数据 ---
                target_unit = ""
                target_price = 0.0
                target_qty = 0.0
                target_date = ""

                raw_unit = items[0].get('unit', '') if items else ''
                raw_price = items[0].get('price', 0) if items else 0
                raw_date = task['ocr_context'].get('delivery_date')

                if match_type == 'DIRECT':
                    target_unit = raw_unit
                    target_price = raw_price
                    target_qty = items[0].get('qty', 0) if items else 0
                    target_date = raw_date

                elif match_type == 'MERGE':
                    total_qty = sum([float(i.get('qty', 0)) for i in items])
                    target_unit = raw_unit
                    target_price = raw_price
                    target_qty = total_qty
                    target_date = raw_date
                    print(f"   ℹ️ [合并] 记录 {record_id} 聚合了 {len(items)} 条明细，总数: {target_qty}")

                elif match_type == 'SPLIT':
                    target_unit = raw_unit
                    target_price = raw_price
                    target_qty = items[0].get('qty', 0) if items else 0
                    target_date = raw_date
                    print(f"   ℹ️ [拆分] 记录 {record_id} 强制调整数量为: {target_qty}")

                # --- 3. 执行填入操作 ---

                # ============================================================
                # A. 填入采购单位 (双重尝试 + 懒加载对抗逻辑)
                # ============================================================
                if target_unit:
                    inp_unit = tr.ele('css:input[name="unitCalc"]', timeout=0.5)
                    if inp_unit:

                        # --- 定义内部搜索函数，方便重试调用 ---
                        def execute_unit_search(is_retry=False):
                            # 1. 点击弹出窗口
                            inp_unit.click()
                            # 如果是重试，多等一会(1秒)，给前端加载数据的时间
                            time.sleep(1.0 if is_retry else 0.5)

                            # 2. 等待搜索框出现
                            search_box = scope.wait.ele_displayed('#txtMeteringPlusKey', timeout=2)
                            if not search_box:
                                return False

                            # 3. 强力输入序列 (聚焦->清空->逐字输入->触发事件)
                            search_box.click()
                            search_box.clear()
                            time.sleep(0.1)
                            for char in target_unit:
                                search_box.input(char)
                                time.sleep(0.05)

                            search_box.run_js('this.dispatchEvent(new Event("input", {bubbles:true}));')
                            search_box.run_js('this.dispatchEvent(new Event("change", {bubbles:true}));')

                            # 4. 回车搜索
                            scope.actions.key_down('ENTER').key_up('ENTER')

                            # 5. 循环检测结果 (3秒)
                            for _ in range(15):
                                # 方案1: 精确匹配
                                target_td = scope.ele(
                                    f'x://div[@id="unitWin"]//table[@id="meteringPlusGrid"]//tbody//tr//td[normalize-space(text())="{target_unit}"]',
                                    timeout=0.1)
                                # 方案2: 包含匹配
                                if not target_td:
                                    target_td = scope.ele(
                                        f'x://div[@id="unitWin"]//table[@id="meteringPlusGrid"]//tbody//tr//td[contains(text(), "{target_unit}")]',
                                        timeout=0.1)

                                if target_td:
                                    print(f"   -> [成功] 刷出单位 [{target_unit}]，执行双击...")
                                    js_code = "this.click(); this.dispatchEvent(new MouseEvent('dblclick', {bubbles: true, cancelable: true, view: window}));"
                                    target_td.run_js(js_code)
                                    time.sleep(0.5)
                                    return True

                                time.sleep(0.2)

                            # 没搜到，关闭弹窗，准备可能的重试
                            inp_unit.click()
                            return False

                        # --- 执行逻辑 ---

                        # 第一次尝试
                        if execute_unit_search(is_retry=False):
                            pass  # 成功
                        else:
                            # 失败了，很有可能是因为第一次打开时数据在异步加载
                            print(f"   ⚠️ 第一次搜索未找到，尝试等待数据加载后重试...")
                            time.sleep(1.0)  # 这里是关键：关掉弹窗，等1秒让缓存生效

                            # 第二次尝试 (Retry)
                            if execute_unit_search(is_retry=True):
                                print(f"   ✅ 重试搜索成功！")
                            else:
                                print(f"❌ 最终失败: 两次尝试均未找到单位 [{target_unit}]")
                                if parsed_data:
                                    parsed_data['processing_failed'] = True
                                    parsed_data['failure_reason'] = f"系统无此单位: {target_unit}"
                                    parsed_data['failure_stage'] = 'unit_search_failed'
                                return  # 停止后续操作

                # ============================================================
                # B. 填入含税单价
                # ============================================================
                if target_price is not None:
                    inp_price = tr.ele('css:input[name="Price"]', timeout=0.2)
                    if inp_price:
                        val = str(target_price)
                        js = f'this.value = "{val}"; this.dispatchEvent(new Event("input")); this.dispatchEvent(new Event("change")); this.dispatchEvent(new Event("blur"));'
                        inp_price.run_js(js)
                        time.sleep(0.1)

                # ============================================================
                # C. 填入数量
                # ============================================================
                if target_qty is not None:
                    inp_qty = tr.ele('css:input[name="Qty"]', timeout=0.2)
                    if inp_qty:
                        val = str(target_qty)
                        js = f'this.value = "{val}"; this.dispatchEvent(new Event("input")); this.dispatchEvent(new Event("change")); this.dispatchEvent(new Event("blur"));'
                        inp_qty.run_js(js)
                        time.sleep(0.1)

                # ============================================================
                # D. 触发总价计算
                # ============================================================
                inp_total = tr.ele('css:input[name="totalAmount"]', timeout=0.2)
                if inp_total:
                    inp_total.click()
                    time.sleep(0.2)

                # ============================================================
                # E. 填入交付日期
                # ============================================================
                if target_date and target_date.strip():
                    inp_date = tr.ele('css:input.deliveryDate', timeout=0.5)

                    if inp_date:
                        try:
                            print(f"   -> 正在填入日期: {target_date}")
                            inp_date.run_js('this.removeAttribute("readonly");')
                            inp_date.clear()
                            time.sleep(0.1)
                            inp_date.input(target_date)
                            time.sleep(0.2)
                            scope.actions.key_down('ENTER').key_up('ENTER')
                            time.sleep(0.2)
                            scope.run_js('document.body.click();')
                            inp_date.click()
                            time.sleep(0.2)
                        except Exception as e:
                            print(f"   ⚠️ 日期输入异常: {e}")
                            try:
                                inp_date.run_js(f'this.removeAttribute("readonly"); this.value="{target_date}";')
                            except:
                                pass

                count_success += 1
                time.sleep(0.1)

            except Exception as e:
                print(f"   !!! 填充行数据失败 (Record: {task.get('record', {}).get('Id')}): {e}")

        print(f"✅ 数据填充完成: 成功处理 {count_success}/{len(structured_tasks)} 行")

    @staticmethod
    def extract_total_amount_from_table(scope):
        """
        从物料采购单表格中提取总金额
        查找带有 class="sumqty" 的 span 元素，并提取其中的总金额数值
        
        :param scope: 操作作用域对象（tab、frame等）
        :return: 清理后的总金额字符串，失败时返回None
        """
        try:
            # 查找带有 sumqty class 的 span 元素
            sumqty_element = scope.ele('css:span.sumqty', timeout=2)
            if sumqty_element and sumqty_element.states.is_displayed:
                # 获取元素的文本内容
                total_amount_text = sumqty_element.text.strip()
                
                if total_amount_text:
                    # 清理金额格式，移除常见的货币符号和格式字符
                    clean_amount = (total_amount_text
                                  .replace(',', '')      # 移除千位分隔符
                                  .replace('￥', '')      # 移除人民币符号
                                  .replace('¥', '')       # 移除日元符号
                                  .replace('$', '')       # 移除美元符号
                                  .replace('€', '')       # 移除欧元符号
                                  .strip())               # 移除首尾空格
                    
                    print(f"   -> 提取到总金额文本: '{total_amount_text}' -> 清理后: '{clean_amount}'")
                    return clean_amount
                else:
                    print("   -> sumqty 元素文本为空")
                    return None
            else:
                print("   -> 未找到可见的 span.sumqty 元素")
                return None
                
        except Exception as e:
            print(f"   !!! 提取总金额异常: {e}")
            return None