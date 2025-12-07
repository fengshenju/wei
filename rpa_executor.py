# -*- coding: utf-8 -*-
"""
RPA执行器 - 负责浏览器自动化操作与单据处理
"""
import json
import re
import datetime
from datetime import timedelta
import os
import time
import random
from DrissionPage import Chromium

# 导入配置和工具
try:
    from app_config import CONFIG
    from utils.util_llm import call_llm_text, call_dmxllm_text
except ImportError as e:
    print(f"!!! RPAExecutor 导入模块失败: {e}")
    exit(1)


class RPAExecutor:
    def __init__(self, browser: Chromium):
        self.browser = browser
        self.PROMPT_INSTRUCTION = CONFIG.get('prompt_instruction', '')

    def run_process(self, data_json, file_name, img_path):
        """
        RPA 处理入口 (对应原 process_single_bill_rpa)
        """
        print(f"\n--- [RPA阶段] 开始处理: {file_name} ---")

        match_prompt = ""
        match_result = None
        original_records = []
        retry_count = 1
        tab = None

        try:
            # 1. 创建新页签
            print(f"[{file_name}] 正在创建新页签...")
            tab = self.browser.new_tab()

            # 2. 激活浏览器窗口
            if CONFIG.get('rpa_browser_to_front', True):
                print(f"[{file_name}] 正在激活浏览器窗口...")
                tab.set.window.normal()
                time.sleep(0.5)
                tab.set.window.full()
            else:
                print(f"[{file_name}] 跳过浏览器窗口激活")

            tab.get(CONFIG['base_url'])

            if not CONFIG.get('rpa_browser_to_front', True):
                print(f"[{file_name}] 将浏览器窗口最小化")
                tab.set.window.mini()

            # 3. 确保左侧菜单栏已经加载出来
            if not tab.wait.ele_displayed('.fixed-left-menu', timeout=5):
                print("!!! 错误: 未检测到左侧菜单栏，请确认网页已加载完成。")
                return "", None, []

            print(">>> 正在定位“物料”菜单...")
            material_btn = tab.ele('x://div[contains(@class, "title") and .//div[contains(text(), "物料")]]')

            if material_btn:
                print(">>> 找到“物料”菜单，正在点击...")
                material_btn.click()
            else:
                print("!!! 未找到物料菜单")

            print("✅ 成功点击“物料”菜单")
            time.sleep(0.5)

            print(">>> 正在定位二级菜单“物料采购需求”...")
            sub_menu_btn = tab.wait.ele_displayed('x://a[contains(text(), "物料采购需求")]', timeout=3)

            if sub_menu_btn:
                sub_menu_btn.click()
                print("✅ 成功点击“物料采购需求”")
            else:
                print("⚠️ 未检测到二级菜单展开，尝试重新点击“物料”...")
                material_btn.click()
                time.sleep(1)
                sub_menu_btn = tab.wait.ele_displayed('x://a[contains(text(), "物料采购需求")]', timeout=3)
                if sub_menu_btn:
                    sub_menu_btn.click()
                    print("✅ 重试后成功点击")
                else:
                    print("!!! 错误：无法展开二级菜单，请检查页面遮挡或网络卡顿。")
                    return "", None, []

            print(">>> 正在 iframe 中查找可见的搜索框...")
            search_input = None

            for _ in range(20):
                for frame in tab.eles('tag:iframe'):
                    if not frame.states.is_displayed:
                        continue
                    ele = frame.ele('#txtSearchKey', timeout=0.2)
                    try:
                        if ele and ele.states.is_displayed:
                            search_input = ele
                            break
                    except:
                        pass
                if search_input:
                    break
                time.sleep(0.5)

            if search_input:
                input_value = data_json.get('final_selected_style', '')
                if not input_value:
                    print(f"⚠️ 警告: 款号为空或None，跳过RPA处理")
                    return "", None, []

                random_sleep = random.uniform(1, 3)
                print(f"[{file_name}] 为防止并发冲突，随机等待 {random_sleep:.2f} 秒...")
                time.sleep(random_sleep)

                print(f">>> 开始输入款号: {input_value}")
                search_input.click()
                time.sleep(0.2)
                search_input.clear()
                time.sleep(0.8)

                for char in input_value:
                    search_input.input(char, clear=False)
                    time.sleep(0.2)

                if search_input.value != input_value:
                    print(f"   -> 检测到输入框值不匹配，强制修正...")
                    search_input.run_js(f'this.value = "{input_value}"')

                target_url_substring = 'Admin/MtReq/NewGet'
                tab.listen.start(targets=target_url_substring)
                print(f">>> 已开启网络监听，目标: {target_url_substring}")

                search_input.run_js("""
                                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                                arguments[0].dispatchEvent(new KeyboardEvent("keydown", {bubbles:true, keyCode:13, key:"Enter"}));
                                arguments[0].dispatchEvent(new KeyboardEvent("keyup", {bubbles:true, keyCode:13, key:"Enter"}));
                            """, search_input)
                print("✅ 输入完毕并回车")

                res_packet = tab.listen.wait(timeout=10)

                if res_packet:
                    print(f"✅ 成功捕获接口数据: {res_packet.url}")
                    response_body = res_packet.response.body
                    msg = response_body.get('msg', '')
                    if '上一个相同请求未结束' in msg or '请勿重复请求' in msg:
                        print(f"⚠️ 触发系统并发限制: {msg}")
                        print(">>> 正在等待 5 秒后重试搜索...")
                        time.sleep(5)
                        search_input.run_js("""
                                arguments[0].dispatchEvent(new KeyboardEvent("keydown", {bubbles:true, keyCode:13, key:"Enter"}));
                                arguments[0].dispatchEvent(new KeyboardEvent("keyup", {bubbles:true, keyCode:13, key:"Enter"}));
                            """, search_input)
                        res_packet = tab.listen.wait(timeout=10)
                        if res_packet:
                            response_body = res_packet.response.body

                    if isinstance(response_body, dict):
                        records = response_body.get('data')
                        if records is None:
                            records = []
                        print(f"数据统计: 共找到 {len(records)} 条记录")

                        if not records:
                            print("⚠️ 警告: 搜索结果为空，无需匹配。")
                        else:
                            original_records = records
                            print(">>> 正在调用 LLM 进行智能匹配...")
                            match_result, match_prompt, retry_count = self.execute_smart_match(data_json, records)

                            print("\n" + "=" * 30)
                            print(f"🤖 智能匹配结果: {match_result.get('status', 'FAIL').upper()}")
                            print(f"匹配原因: {match_result.get('global_reason', match_result.get('reason'))}")
                            print("=" * 30 + "\n")

                            matched_ids = []
                            structured_tasks = []

                            if match_result.get('status') == 'success':
                                structured_tasks = self.reconstruct_rpa_data(match_result, data_json, original_records)
                                print(f">>> 数据组装完成，共生成 {len(structured_tasks)} 个任务包")

                                seen_ids = set()
                                for task in structured_tasks:
                                    rec_id = task['record'].get('Id')
                                    if rec_id and rec_id not in seen_ids:
                                        matched_ids.append(rec_id)
                                        seen_ids.add(rec_id)

                            if matched_ids:
                                self.select_matched_checkboxes(tab, matched_ids)

                                print(">>> 正在查找并点击“物料采购单”生成按钮...")
                                button_found = False
                                scopes = [tab] + [f for f in tab.eles('tag:iframe') if f.states.is_displayed]

                                for scope in scopes:
                                    btn = scope.ele('x://button[contains(text(), "物料采购单")]', timeout=0.5)
                                    if btn and btn.states.is_displayed:
                                        btn.scroll.to_see()
                                        time.sleep(0.5)
                                        btn.click()
                                        print("✅ 成功点击“物料采购单”按钮")
                                        button_found = True
                                        time.sleep(2)
                                        break

                                if not button_found:
                                    print("⚠️ 未找到“物料采购单”按钮")
                                else:
                                    print(">>> 正在等待页面加载并切换为“月结采购”...")
                                    time.sleep(2)
                                    type_selected = False
                                    current_scopes = [tab] + [f for f in tab.eles('tag:iframe') if
                                                              f.states.is_displayed]

                                    for scope in current_scopes:
                                        try:
                                            dropdown_btn = scope.ele('css:button[data-id="OrderTypeId"]', timeout=0.5)
                                            if dropdown_btn and dropdown_btn.states.is_displayed:
                                                print("   -> 找到采购类型下拉框")
                                                dropdown_btn.scroll.to_see()
                                                dropdown_btn.click()
                                                time.sleep(0.5)
                                                option = scope.ele('x://span[@class="text" and text()="月结采购"]',
                                                                   timeout=1)
                                                if option and option.states.is_displayed:
                                                    option.click()
                                                    print("✅ 成功选择“月结采购”")
                                                    type_selected = True
                                                    time.sleep(1)
                                                    break
                                        except Exception:
                                            continue

                                    if not type_selected:
                                        print("⚠️ 未能完成月结采购选择")

                                    supplier_name = data_json.get('supplier_name', '').strip()
                                    if supplier_name:
                                        print(f">>> 正在设置供应商: {supplier_name}")
                                        supplier_selected = False
                                        for scope in current_scopes:
                                            try:
                                                supplier_label = scope.ele('#lbSupplierInfo', timeout=0.5)
                                                if supplier_label and supplier_label.states.is_displayed:
                                                    supplier_label.click()
                                                    time.sleep(0.5)
                                                    search_box = scope.ele('#txtMpSupplierPlusContent', timeout=1)
                                                    if search_box and search_box.states.is_displayed:
                                                        search_box.clear()
                                                        search_box.input(supplier_name)
                                                        time.sleep(0.2)
                                                        scope.actions.key_down('ENTER').key_up('ENTER')
                                                        time.sleep(0.5)
                                                        target_td_xpath = f'x://table[@id="mtSupplierPlusGrid"]//tbody//tr//td[text()="{supplier_name}"]'
                                                        target_td = scope.ele(target_td_xpath, timeout=1)
                                                        if target_td:
                                                            print(f"  -> 找到供应商 [{supplier_name}]，执行双击选择...")
                                                            js_code = """this.click(); this.dispatchEvent(new MouseEvent('dblclick', {bubbles: true, cancelable: true, view: window}));"""
                                                            target_td.run_js(js_code)
                                                            time.sleep(0.5)
                                                            supplier_selected = True
                                                            break
                                                        else:
                                                            print(f"  ⚠️ 供应商列表中未搜索到: {supplier_name}")
                                                            supplier_label.click()
                                            except Exception:
                                                continue
                                        if not supplier_selected:
                                            print(f"⚠️ 未能完成供应商选择: {supplier_name}")
                                    else:
                                        print("⚠️ 未获取到供应商名称")

                                    style_code = data_json.get('final_selected_style', '').strip().upper()
                                    target_brand = None
                                    if style_code.startswith('T'):
                                        target_brand = "CHENXIHE"
                                    elif style_code.startswith('X'):
                                        target_brand = "CHENXIHE抖音"
                                    elif style_code.startswith('H'):
                                        target_brand = "SUNONEONE"
                                    elif style_code.startswith('D'):
                                        target_brand = "SUNONEONE抖音"

                                    if target_brand:
                                        print(f">>> 识别到款号[{style_code}]，准备选择品牌: [{target_brand}]...")
                                        brand_selected = False
                                        for scope in current_scopes:
                                            try:
                                                brand_btn = scope.ele('css:button[data-id="BrandId"]', timeout=0.3)
                                                if brand_btn and brand_btn.states.is_displayed:
                                                    brand_btn.scroll.to_see()
                                                    brand_btn.click()
                                                    time.sleep(0.5)
                                                    open_menu = scope.ele('css:div.btn-group.open', timeout=1)
                                                    if open_menu:
                                                        brand_opt = open_menu.ele(
                                                            f'x:.//span[contains(@class, "text") and contains(text(), "{target_brand}")]',
                                                            timeout=1)
                                                        if brand_opt:
                                                            brand_opt.scroll.to_see()
                                                            time.sleep(0.1)
                                                            brand_opt.click()
                                                            print(f"✅ 成功选择品牌: {target_brand}")
                                                            brand_selected = True
                                                            time.sleep(0.5)
                                                            break
                                                        else:
                                                            print(f"   ⚠️ 未找到选项 [{target_brand}]")
                                                            brand_btn.click()
                                            except Exception:
                                                continue
                                        if not brand_selected:
                                            print(f"⚠️ 品牌选择失败")

                                    ocr_date = data_json.get('delivery_date', '')
                                    if ocr_date:
                                        print(f">>> 正在查找并填写码单日期: {ocr_date} ...")
                                        att01_filled = False
                                        for scope in current_scopes:
                                            try:
                                                att01_input = scope.ele('#Att01', timeout=0.5)
                                                if att01_input and att01_input.states.is_displayed:
                                                    att01_input.scroll.to_see()
                                                    att01_input.clear()
                                                    att01_input.input(ocr_date)
                                                    att01_input.run_js(
                                                        'this.dispatchEvent(new Event("change", {bubbles: true})); this.dispatchEvent(new Event("blur"));')
                                                    print("✅ 成功填写码单日期")
                                                    att01_filled = True
                                                    break
                                            except Exception:
                                                continue
                                        if not att01_filled:
                                            print("⚠️ 未找到码单日期输入框 (#Att01)")

                                    self.fill_details_into_table(scope, structured_tasks)

                                    print(">>> 表格填写完毕，正在查找并点击“保存并审核”按钮...")
                                    try:
                                        save_btn = scope.ele('css:button[data-amid="btnSaveAndAudit"]', timeout=1)
                                        if not save_btn:
                                            save_btn = scope.ele('x://button[contains(text(), "保存并审核")]',
                                                                 timeout=1)

                                        if save_btn and save_btn.states.is_displayed:
                                            save_btn.scroll.to_see()
                                            time.sleep(0.5)
                                            save_btn.click()
                                            print("✅ 成功点击“保存并审核”")

                                            try:
                                                if tab.wait.alert(timeout=2):
                                                    tab.alert.accept()
                                            except:
                                                pass

                                            print(">>> 等待保存结果...")
                                            time.sleep(3)
                                            print(">>> 正在获取生成的订单编号...")
                                            try:
                                                code_input = scope.ele('#Code', timeout=2)
                                                if code_input:
                                                    order_code = code_input.value or code_input.attr(
                                                        'valuecontent') or code_input.attr('value')
                                                    if order_code:
                                                        print(f"✅ 成功获取订单编号: [{order_code}]")
                                                        data_json['rpa_order_code'] = order_code
                                                    else:
                                                        print("⚠️ 无法提取到编号值")
                                            except Exception as e:
                                                print(f"!!! 获取订单编号异常: {e}")

                                            print(">>> 准备跳转至“物料采购订单”列表...")
                                            time.sleep(0.5)

                                            try:
                                                material_btn_nav = tab.ele(
                                                    'x://div[contains(@class, "title") and .//div[contains(text(), "物料")]]')
                                                if material_btn_nav:
                                                    material_btn_nav.click()
                                                    time.sleep(0.5)

                                                target_menu_text = "物料采购订单"
                                                purchase_order_menu = tab.wait.ele_displayed(
                                                    f'x://a[contains(text(), "{target_menu_text}")]', timeout=3)
                                                if purchase_order_menu:
                                                    purchase_order_menu.click()
                                                else:
                                                    if material_btn_nav:
                                                        material_btn_nav.click()
                                                        time.sleep(0.5)
                                                    purchase_order_menu = tab.wait.ele_displayed(
                                                        f'x://a[contains(text(), "{target_menu_text}")]', timeout=3)
                                                    if purchase_order_menu:
                                                        purchase_order_menu.click()

                                                time.sleep(2)

                                                order_code = data_json.get('rpa_order_code')
                                                if order_code:
                                                    print(f">>> 准备在“物料采购订单”列表搜索单号: {order_code}")
                                                    search_input_order = None
                                                    for _ in range(10):
                                                        for frame in tab.eles('tag:iframe'):
                                                            if not frame.states.is_displayed: continue
                                                            ele = frame.ele(
                                                                'css:input#txtSearchKey[data-grid="POMtPurchaseGrid"]',
                                                                timeout=0.2)
                                                            if ele and ele.states.is_displayed:
                                                                search_input_order = ele
                                                                break
                                                        if search_input_order: break
                                                        time.sleep(0.5)

                                                    if search_input_order:
                                                        search_input_order.click()
                                                        time.sleep(0.2)
                                                        search_input_order.clear()
                                                        for char in order_code:
                                                            search_input_order.input(char, clear=False)
                                                            time.sleep(0.1)

                                                        tab.listen.start(targets='Admin/MtPurchase')
                                                        search_input_order.run_js("""
                                                            this.dispatchEvent(new Event('change', { bubbles: true }));
                                                            this.dispatchEvent(new KeyboardEvent("keydown", {bubbles:true, keyCode:13, key:"Enter"}));
                                                            this.dispatchEvent(new KeyboardEvent("keyup", {bubbles:true, keyCode:13, key:"Enter"}));
                                                        """)
                                                        print("✅ 输入完毕并回车")

                                                        res_packet_order = tab.listen.wait(timeout=10)
                                                        if res_packet_order:
                                                            print(f"✅ 搜索成功")
                                                            time.sleep(0.5)
                                                            all_selected = False
                                                            target_frame = None

                                                            for frame in tab.eles('tag:iframe'):
                                                                if not frame.states.is_displayed: continue
                                                                try:
                                                                    select_all_btn = frame.ele(
                                                                        'x://input[@type="checkbox" and contains(@onclick, "selectAll")]',
                                                                        timeout=0.5)
                                                                    if not select_all_btn:
                                                                        select_all_btn = frame.ele(
                                                                            'x://button[contains(text(), "全选") or contains(text(), "选择全部")]',
                                                                            timeout=0.5)
                                                                    if select_all_btn and select_all_btn.states.is_displayed:
                                                                        select_all_btn.click()
                                                                        all_selected = True
                                                                        target_frame = frame
                                                                        break

                                                                    checkboxes = frame.eles(
                                                                        'x://tr//input[contains(@class, "ckbox")]',
                                                                        timeout=1)
                                                                    if checkboxes:
                                                                        selected_count = 0
                                                                        for checkbox in checkboxes:
                                                                            if checkbox.states.is_displayed:
                                                                                checkbox.scroll.to_see()
                                                                                if not checkbox.states.is_checked:
                                                                                    checkbox.click()
                                                                                    selected_count += 1
                                                                                    time.sleep(0.1)
                                                                        if selected_count > 0:
                                                                            all_selected = True
                                                                            target_frame = frame
                                                                            break
                                                                except Exception:
                                                                    continue

                                                            if all_selected and target_frame:
                                                                target_frame.scroll.down(200)
                                                                time.sleep(0.5)
                                                                print(">>> 记录已选中，准备触发附件上传...")
                                                                try:
                                                                    adjunct_tab = target_frame.ele(
                                                                        'x://a[contains(text(), "附件") and contains(@href, "tb_Adjunct")]',
                                                                        timeout=2)
                                                                    if adjunct_tab:
                                                                        adjunct_tab.click()
                                                                        time.sleep(0.5)
                                                                        upload_label = target_frame.ele(
                                                                            'x://div[@id="tb_Adjunct"]//label[contains(@style, "opacity: 0")]',
                                                                            timeout=2)
                                                                        if upload_label:
                                                                            abs_img_path = os.path.abspath(img_path)
                                                                            upload_label.click.to_upload(abs_img_path)
                                                                            print(">>> 正在上传附件...")
                                                                            time.sleep(5)
                                                                            save_img_btn = target_frame.ele(
                                                                                'x://button[@onclick="AddImg()"]',
                                                                                timeout=2)
                                                                            if not save_img_btn:
                                                                                save_img_btn = target_frame.ele(
                                                                                    'x://button[contains(text(), "保存") and contains(@class, "btn-success")]',
                                                                                    timeout=2)
                                                                            if not save_img_btn:
                                                                                save_img_btn = target_frame.ele(
                                                                                    'css:button.btn.btn-success.btn-sm',
                                                                                    timeout=2)

                                                                            if save_img_btn:
                                                                                save_img_btn.scroll.to_see()
                                                                                time.sleep(0.5)
                                                                                save_img_btn.click()
                                                                                print("✅ 成功点击附件保存按钮")
                                                                                time.sleep(2)
                                                                except Exception as e:
                                                                    print(f"!!! 附件上传异常: {e}")

                                                                print(">>> 准备执行采购任务...")
                                                                time.sleep(1)
                                                                try:
                                                                    more_btn = target_frame.ele(
                                                                        'x://button[contains(text(), "更多")]',
                                                                        timeout=2)
                                                                    if more_btn:
                                                                        more_btn.click()
                                                                        time.sleep(0.5)
                                                                        exec_task_btn = target_frame.ele(
                                                                            'css:a[onclick="doMtPurTask()"]', timeout=1)
                                                                        if not exec_task_btn:
                                                                            exec_task_btn = target_frame.ele(
                                                                                'x://a[contains(text(), "执行采购任务")]',
                                                                                timeout=1)
                                                                        if exec_task_btn:
                                                                            exec_task_btn.click()
                                                                            try:
                                                                                if tab.wait.alert(
                                                                                    timeout=3): tab.alert.accept()
                                                                            except:
                                                                                pass
                                                                            print("✅ 成功点击“执行采购任务”")
                                                                            time.sleep(2)

                                                                            try:
                                                                                confirm_btn = tab.ele(
                                                                                    'css:a.layui-layer-btn0', timeout=3)
                                                                                if not confirm_btn: confirm_btn = target_frame.ele(
                                                                                    'css:a.layui-layer-btn0', timeout=2)
                                                                                if confirm_btn:
                                                                                    confirm_btn.click()
                                                                                    time.sleep(1)
                                                                                    self.navigate_and_search_purchase_task(
                                                                                        tab,
                                                                                        data_json.get('rpa_order_code'),
                                                                                        data_json)
                                                                                    self.navigate_to_bill_list(tab,
                                                                                                               data_json.get(
                                                                                                                   'rpa_order_code'))
                                                                            except Exception:
                                                                                print("⚠️ 未检测到结果弹窗")
                                                                except Exception as e:
                                                                    print(f"!!! 执行采购任务操作异常: {e}")

                                                        try:
                                                            if tab.listen: tab.listen.stop()
                                                        except:
                                                            pass
                                                else:
                                                    print("ℹ️ 无订单编号，跳过搜索")
                                            except Exception as e:
                                                print(f"!!! 菜单跳转异常: {e}")
                                    except Exception as e:
                                        print(f"!!! 点击保存按钮异常: {e}")
                    else:
                        print("响应内容不是 JSON 格式")
                else:
                    print(f"!!! 警告: 等待超时，未捕获到请求。")

                try:
                    if tab.listen: tab.listen.stop()
                except:
                    pass
            else:
                print("!!! 错误：没找到可见的搜索框")

        except Exception as e:
            error_msg = f"RPA执行异常: {str(e)}"
            print(f"!!! {error_msg}")
            if match_result is None:
                match_result = {
                    "status": "fail",
                    "reason": error_msg,
                    "global_reason": error_msg
                }
            else:
                match_result['reason'] = f"{match_result.get('reason', '')} | {error_msg}"
        finally:
            try:
                if tab and tab.listen: tab.listen.stop()
            except:
                pass
            if tab:
                try:
                    # tab.close()
                    print(f"[{file_name}] 页签已关闭 (模拟)")
                except:
                    pass

        return match_prompt, match_result, original_records, retry_count

    def handle_new_reconciliation_bill(self, tab):
        print("\n>>> [阶段: 新增对账单处理] 开始...")
        try:
            target_frame = None
            save_audit_btn = None
            print(">>> 正在查找“保存并审核”按钮所在的 iframe...")
            for _ in range(5):
                for frame in tab.eles('tag:iframe'):
                    if not frame.states.is_displayed: continue
                    btn = frame.ele('css:button[data-amid="btnPaySaveAndAduit"]', timeout=0.1)
                    if not btn: btn = frame.ele('css:button[onclick="saveRecord(1)"]', timeout=0.1)
                    if not btn: btn = frame.ele('x://button[contains(text(), "保存并审核")]', timeout=0.1)
                    if btn and btn.states.is_displayed:
                        save_audit_btn = btn
                        target_frame = frame
                        break
                if save_audit_btn: break
                time.sleep(1)

            if save_audit_btn:
                print("   -> 找到“保存并审核”按钮，准备点击...")
                save_audit_btn.scroll.to_see()
                time.sleep(0.5)
                save_audit_btn.click()
                print("✅ “新增对账单”审核流程完成")
            else:
                print("⚠️ 未在任何可见 iframe 中找到“保存并审核”按钮")
        except Exception as e:
            print(f"!!! 新增对账单处理异常: {e}")

    def navigate_to_bill_list(self, tab, order_code):
        print("\n>>> [阶段: 跳转账单列表] 开始处理...")
        try:
            print(">>> 正在定位“财务”菜单...")
            finance_btn = tab.ele('x://div[contains(@class, "title") and .//div[contains(text(), "财务")]]')
            if finance_btn:
                finance_btn.click()
                time.sleep(0.5)
            else:
                print("!!! 错误: 未找到“财务”菜单")
                return

            target_menu_text = "账单列表"
            bill_menu = tab.wait.ele_displayed(f'x://a[contains(text(), "{target_menu_text}")]', timeout=3)
            if bill_menu:
                bill_menu.click()
                print(f"✅ 成功点击左侧菜单“{target_menu_text}”")
                time.sleep(2)
            else:
                print(f"⚠️ 未检测到二级菜单，尝试重新展开一级菜单...")
                if finance_btn:
                    finance_btn.click()
                    time.sleep(0.5)
                bill_menu = tab.wait.ele_displayed(f'x://a[contains(text(), "{target_menu_text}")]', timeout=3)
                if bill_menu:
                    bill_menu.click()
                    print(f"✅ (重试) 成功点击左侧菜单“{target_menu_text}”")
                    time.sleep(2)
                else:
                    return

            print(f">>> 正在查找搜索框 (data-grid='FMAccountsReceivableGrid')...")
            if not order_code: return

            search_input_bill = None
            target_frame = None
            for _ in range(10):
                for frame in tab.eles('tag:iframe'):
                    if not frame.states.is_displayed: continue
                    ele = frame.ele('css:input#txtSearchKey[data-grid="FMAccountsReceivableGrid"]', timeout=0.2)
                    if ele and ele.states.is_displayed:
                        search_input_bill = ele
                        target_frame = frame
                        break
                if search_input_bill: break
                time.sleep(0.5)

            if search_input_bill:
                print(f">>> 找到账单列表搜索框，正在输入: {order_code}")
                search_input_bill.click()
                time.sleep(0.2)
                search_input_bill.clear()
                for char in order_code:
                    search_input_bill.input(char, clear=False)
                    time.sleep(0.2)

                tab.listen.start(targets='Admin/AccountsReceivable/NewGet')
                search_input_bill.run_js("""
                    this.dispatchEvent(new Event('change', { bubbles: true }));
                    this.dispatchEvent(new KeyboardEvent("keydown", {bubbles:true, keyCode:13, key:"Enter"}));
                    this.dispatchEvent(new KeyboardEvent("keyup", {bubbles:true, keyCode:13, key:"Enter"}));
                """)
                print("✅ 输入完毕并触发回车")

                res = None
                try:
                    res = tab.listen.wait(timeout=10)
                finally:
                    tab.listen.stop()

                if res:
                    print(f"✅ 账单列表搜索响应成功")
                    time.sleep(1)
                    if target_frame:
                        rows = target_frame.eles('css:table#FMAccountsReceivableGrid tbody tr', timeout=2)
                        count_selected = 0
                        if rows:
                            for row in rows:
                                if not row.states.is_displayed: continue
                                try:
                                    row.scroll.to_see()
                                    checkbox = row.ele('css:input.ckbox', timeout=0.5)
                                    if checkbox:
                                        if not checkbox.states.is_checked:
                                            checkbox.click()
                                            count_selected += 1
                                    time.sleep(0.1)
                                except:
                                    pass

                            if count_selected > 0:
                                print(f"✅ 已勾选 {count_selected} 条账单记录")
                                print(">>> 准备点击“发起对账”...")
                                try:
                                    btn_check = target_frame.ele('css:button[onclick="aReconciliation()"]', timeout=2)
                                    if not btn_check: btn_check = target_frame.ele(
                                        'x://button[contains(text(), "发起对账")]', timeout=1)
                                    if btn_check:
                                        btn_check.run_js('this.click()')
                                        time.sleep(2)
                                        print("✅ “发起对账”操作完成")
                                        print(">>> 等待“新增对账单”页面加载...")
                                        time.sleep(3)
                                        self.handle_new_reconciliation_bill(tab)
                                    else:
                                        print("⚠️ 未找到“发起对账”按钮")
                                except Exception as e:
                                    print(f"!!! 发起对账操作异常: {e}")
                        else:
                            print("⚠️ 未勾选任何记录")
                else:
                    print("⚠️ 搜索超时")
            else:
                print("!!! 错误: 未找到账单列表搜索框")
        except Exception as e:
            print(f"!!! 跳转账单列表时发生异常: {e}")

    def navigate_and_search_purchase_task(self, tab, order_code, parsed_data):
        print(f"\n>>> [阶段: 跳转物料采购任务] 开始处理，目标单号: {order_code}")
        if not order_code: return

        delivery_date = parsed_data.get('delivery_date', '')
        delivery_order_no = parsed_data.get('delivery_order_number', '')

        try:
            print(">>> 正在重新定位“物料”菜单...")
            material_btn_nav = tab.ele('x://div[contains(@class, "title") and .//div[contains(text(), "物料")]]')
            if material_btn_nav:
                material_btn_nav.click()
                time.sleep(0.5)

            target_menu_text = "物料采购任务"
            task_menu = tab.wait.ele_displayed(f'x://a[contains(text(), "{target_menu_text}")]', timeout=3)
            if task_menu:
                task_menu.click()
                print(f"✅ 成功点击左侧菜单“{target_menu_text}”")
            else:
                if material_btn_nav:
                    material_btn_nav.click()
                    time.sleep(0.5)
                task_menu = tab.wait.ele_displayed(f'x://a[contains(text(), "{target_menu_text}")]', timeout=3)
                if task_menu:
                    task_menu.click()
                else:
                    return

            time.sleep(2)
            print(f">>> 正在查找搜索框 (data-grid='poMtPurTaskGrid')...")
            search_input_task = None
            target_frame = None

            for _ in range(10):
                for frame in tab.eles('tag:iframe'):
                    if not frame.states.is_displayed: continue
                    ele = frame.ele('css:input#txtSearchKey[data-grid="poMtPurTaskGrid"]', timeout=0.2)
                    if ele and ele.states.is_displayed:
                        search_input_task = ele
                        target_frame = frame
                        break
                if search_input_task: break
                time.sleep(0.5)

            if search_input_task:
                print(f">>> 找到搜索框，正在输入: {order_code}")
                search_input_task.click()
                time.sleep(0.2)
                search_input_task.clear()
                for char in order_code:
                    search_input_task.input(char, clear=False)
                    time.sleep(0.2)

                tab.listen.start(targets='Admin/MtPurchase')
                search_input_task.run_js("""
                    this.dispatchEvent(new Event('change', { bubbles: true }));
                    this.dispatchEvent(new KeyboardEvent("keydown", {bubbles:true, keyCode:13, key:"Enter"}));
                    this.dispatchEvent(new KeyboardEvent("keyup", {bubbles:true, keyCode:13, key:"Enter"}));
                """)
                print("✅ 输入完毕并触发回车")
                try:
                    res = tab.listen.wait(timeout=10)
                finally:
                    tab.listen.stop()

                if res:
                    print(f"✅ 搜索响应成功")
                    time.sleep(1)
                    if target_frame:
                        def select_all_rows(frame_obj, label=""):
                            selected_count = 0
                            try:
                                current_rows = frame_obj.eles('css:table#poMtPurTaskGrid tbody tr', timeout=3)
                                if not current_rows: return 0
                                for row in current_rows:
                                    if not row.states.is_displayed: continue
                                    try:
                                        row.scroll.to_see()
                                        ck = row.ele('css:input.ckbox', timeout=0.5)
                                        if ck:
                                            if not ck.states.is_checked:
                                                ck.click()
                                                selected_count += 1
                                        time.sleep(0.05)
                                    except:
                                        pass
                            except Exception as e:
                                print(f"   !!! {label} 勾选异常: {e}")
                            return selected_count

                        select_all_rows(target_frame, "首次勾选")
                        print(">>> [1/3] 准备点击“一键绑定加工单”...")
                        try:
                            btn_bind = target_frame.ele('#btnOneKeyBindPM', timeout=2)
                            if btn_bind:
                                btn_bind.click()
                                try:
                                    if tab.wait.alert(timeout=3): tab.alert.accept()
                                except:
                                    pass
                                try:
                                    confirm_btn = tab.ele('css:a.layui-layer-btn0', timeout=3)
                                    if confirm_btn: confirm_btn.click()
                                except:
                                    pass
                                time.sleep(1)
                                try:
                                    if tab.wait.alert(timeout=2): tab.alert.accept()
                                except:
                                    pass
                                print("   ✅ 一键绑定操作结束")

                                print(">>> 等待系统处理一键绑定...")
                                binding_completed = False
                                for attempt in range(15):
                                    try:
                                        total_rows = target_frame.eles('css:table#poMtPurTaskGrid tbody tr', timeout=2)
                                        visible_rows = [r for r in total_rows if r.states.is_displayed]
                                        total_count = len(visible_rows)
                                        if total_count == 0:
                                            time.sleep(2)
                                            continue

                                        factory_cells = target_frame.eles('css:td[masking="SpName"]', timeout=1)
                                        completed_count = 0
                                        for cell in factory_cells:
                                            if cell.states.is_displayed and cell.text.strip():
                                                completed_count += 1

                                        if completed_count >= total_count and total_count > 0:
                                            binding_completed = True
                                            break
                                        time.sleep(2)
                                    except:
                                        time.sleep(2)

                                print(">>> 一键绑定完成，开始填写码单信息...")
                                try:
                                    rows = target_frame.eles('css:table#poMtPurTaskGrid tbody tr', timeout=3)
                                    for row in rows:
                                        if not row.states.is_displayed: continue
                                        try:
                                            row.scroll.to_see()
                                            if delivery_order_no:
                                                inp_no = row.ele('css:input.Att01', timeout=0.2)
                                                if inp_no:
                                                    js_no = f'this.value = "{delivery_order_no}"; this.dispatchEvent(new Event("input")); this.dispatchEvent(new Event("blur"));'
                                                    inp_no.run_js(js_no)
                                            if delivery_date:
                                                inp_date = row.ele('css:input.Att02', timeout=0.2)
                                                if inp_date:
                                                    inp_date.run_js('this.removeAttribute("readonly");')
                                                    inp_date.clear()
                                                    inp_date.input(delivery_date)
                                                    target_frame.run_js('document.body.click();')
                                        except:
                                            pass
                                    print(f"✅ 码单信息填写完成")
                                except Exception as e:
                                    print(f"!!! 填写信息异常: {e}")
                            else:
                                print("⚠️ 未找到一键绑定加工单按钮")
                        except Exception as e:
                            print(f"!!! 绑定操作异常: {e}")

                        print("\n>>> [重要] 准备提交，正在强制重新勾选所有记录...")
                        time.sleep(1)
                        reselect_count = select_all_rows(target_frame, "提交前重选")
                        print(f"✅ 已确认勾选 {reselect_count} 行")

                        print(">>> [2/3] 准备点击“提交”...")
                        try:
                            btn_submit = target_frame.ele('#btnSubmitTasks', timeout=2)
                            if btn_submit:
                                btn_submit.click()
                                try:
                                    if tab.wait.alert(timeout=3): tab.alert.accept()
                                except:
                                    pass
                                time.sleep(1)
                                try:
                                    if tab.wait.alert(timeout=2): tab.alert.accept()
                                except:
                                    pass
                                print("   ✅ “提交”操作结束")
                                time.sleep(2)
                            else:
                                print("⚠️ 未找到“提交”按钮")
                        except Exception as e:
                            print(f"!!! 提交操作异常: {e}")

                        print(">>> [3/3] 准备点击“确认”...")
                        try:
                            btn_confirm = target_frame.ele('#btnConfirmToDoTask', timeout=2)
                            if btn_confirm:
                                btn_confirm.click()
                                try:
                                    if tab.wait.alert(timeout=3): tab.alert.accept()
                                except:
                                    pass
                                time.sleep(2)
                                try:
                                    lay_confirm = tab.ele('css:a.layui-layer-btn0', timeout=3)
                                    if not lay_confirm: lay_confirm = target_frame.ele('css:a.layui-layer-btn0',
                                                                                       timeout=1)
                                    if lay_confirm: lay_confirm.click()
                                except:
                                    pass
                                print("✅ “确认”操作全部完成")
                            else:
                                print("⚠️ 未找到“确认”按钮")
                        except Exception as e:
                            print(f"!!! 确认操作异常: {e}")
                    else:
                        print("!!! 错误: 丢失了 iframe 上下文")
                else:
                    print("⚠️ 搜索超时")
            else:
                print("!!! 错误: 未找到搜索框")
        except Exception as e:
            print(f"!!! 跳转或搜索'物料采购任务'时发生异常: {e}")

    def select_matched_checkboxes(self, tab, matched_ids):
        print(f">>> 开始勾选匹配的记录: {len(matched_ids)} 条")
        for record_id in matched_ids:
            try:
                checkbox_selector = f'x://tr[.//a[contains(@data-sub-html, "{record_id}")]]//input[contains(@class, "ckbox")]'
                checkbox_found = False
                for frame in tab.eles('tag:iframe'):
                    if not frame.states.is_displayed: continue
                    checkbox = frame.ele(checkbox_selector, timeout=0.2)
                    if checkbox and checkbox.states.is_displayed:
                        if not checkbox.states.is_checked:
                            checkbox.click()
                            print(f"✅ 已勾选记录: {record_id}")
                        checkbox_found = True
                        break
                if not checkbox_found:
                    print(f"⚠️ 未找到记录 {record_id} 对应的checkbox")
            except Exception as e:
                print(f"!!! 勾选记录 {record_id} 时发生异常: {e}")
        print(f">>> 勾选操作完成")

    def execute_smart_match(self, parsed_data, records):
        """
        执行智能匹配核心逻辑
        """
        today = datetime.date.today()
        two_weeks_ago = today - timedelta(days=14)
        clean_records = self.preprocess_records(records)

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

    def preprocess_records(self, records):
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

    def reconstruct_rpa_data(self, match_result, original_parsed_data, original_records):
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

    def fill_details_into_table(self, scope, structured_tasks):
        print(f">>> 开始填充物料明细数据，共 {len(structured_tasks)} 条任务...")
        count_success = 0
        for task in structured_tasks:
            try:
                record_id = task['record'].get('Id')
                match_type = task['match_type']
                items = task['items']
                if not record_id or not items: continue

                tr_xpath = f'x://tr[.//input[@name="materialReqId" and @value="{record_id}"]]'
                tr = scope.ele(tr_xpath, timeout=1)
                if not tr:
                    print(f"   ⚠️ 未找到 ID 为 {record_id} 的行，跳过")
                    continue
                tr.scroll.to_see()

                target_unit = ""
                target_price = 0.0
                target_qty = 0.0
                target_date = ""
                first_item = items[0]
                raw_unit = first_item.get('unit', '')
                raw_price = first_item.get('price', 0)
                raw_date = task['ocr_context'].get('delivery_date')

                if match_type == 'DIRECT':
                    target_unit = raw_unit
                    target_price = raw_price
                    target_qty = first_item.get('qty', 0)
                    target_date = raw_date
                elif match_type == 'MERGE':
                    total_qty = sum([float(i.get('qty', 0)) for i in items])
                    target_unit = raw_unit
                    target_price = raw_price
                    target_qty = total_qty
                    target_date = raw_date
                    print(f"   ℹ️ [合并] 记录 {record_id} 聚合了 {len(items)} 条明细")
                elif match_type == 'SPLIT':
                    target_unit = raw_unit
                    target_price = raw_price
                    target_qty = first_item.get('qty', 0)
                    target_date = raw_date

                if target_unit:
                    inp_unit = tr.ele('css:input[name="unitCalc"]', timeout=0.5)
                    if inp_unit:
                        inp_unit.click()
                        time.sleep(0.5)
                        search_box = scope.ele('#txtMeteringPlusKey', timeout=1)
                        if search_box and search_box.states.is_displayed:
                            search_box.clear()
                            search_box.input(target_unit)
                            time.sleep(0.2)
                            scope.actions.key_down('ENTER').key_up('ENTER')
                            target_td_xpath = f'x://table[@id="meteringPlusGrid"]//tbody//tr//td[text()="{target_unit}"]'
                            target_td = scope.ele(target_td_xpath, timeout=1)
                            if target_td:
                                js_code = """this.click(); this.dispatchEvent(new MouseEvent('dblclick', {bubbles: true, cancelable: true, view: window}));"""
                                target_td.run_js(js_code)
                                time.sleep(0.5)
                            else:
                                inp_unit.click()
                        else:
                            inp_unit.input(target_unit, clear=True)

                if target_price is not None:
                    inp_price = tr.ele('css:input[name="Price"]', timeout=0.2)
                    if inp_price:
                        val = str(target_price)
                        js = f'this.value = "{val}"; this.dispatchEvent(new Event("input")); this.dispatchEvent(new Event("change")); this.dispatchEvent(new Event("blur"));'
                        inp_price.run_js(js)
                        time.sleep(0.1)

                if target_qty is not None:
                    inp_qty = tr.ele('css:input[name="Qty"]', timeout=0.2)
                    if inp_qty:
                        val = str(target_qty)
                        js = f'this.value = "{val}"; this.dispatchEvent(new Event("input")); this.dispatchEvent(new Event("change")); this.dispatchEvent(new Event("blur"));'
                        inp_qty.run_js(js)
                        time.sleep(0.1)

                inp_total = tr.ele('css:input[name="totalAmount"]', timeout=0.2)
                if inp_total:
                    inp_total.click()
                    time.sleep(0.2)

                if target_date and target_date.strip():
                    inp_date = tr.ele('css:input.deliveryDate', timeout=0.5)
                    if inp_date:
                        try:
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
                        except:
                            try:
                                inp_date.run_js(f'this.removeAttribute("readonly"); this.value="{target_date}";')
                            except:
                                pass

                count_success += 1
                time.sleep(0.1)
            except Exception as e:
                print(f"   !!! 填充行数据失败 (Record: {task.get('record', {}).get('Id')}): {e}")

        print(f"✅ 数据填充完成: 成功处理 {count_success}/{len(structured_tasks)} 行")