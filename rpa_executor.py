# -*- coding: utf-8 -*-
"""
RPA执行器 - 负责浏览器自动化操作与单据处理
"""
import os
import time
import random
from DrissionPage import Chromium

# 导入配置和工具
try:
    from app_config import CONFIG
    from utils.rpa_utils import RPAUtils
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

            if not RPAUtils.navigate_to_menu(tab, "物料", "物料采购需求"):
                print("!!! 错误：无法导航到物料采购需求页面")
                return "", None, []

            #查找搜索框
            print(">>> 正在 iframe 中查找可见的搜索框...")
            search_input = RPAUtils.find_element_in_iframes(
                tab=tab,
                selector='#txtSearchKey',
                max_retries=20,
                retry_interval=1,
                timeout=0.5
            )

            # 找到搜索框输入款号搜索
            if search_input:
                input_value = data_json.get('final_selected_style', '')
                if not input_value:
                    print(f"⚠️ 警告: 款号为空或None，跳过RPA处理")
                    return "", None, []

                random_sleep = random.uniform(1, 3)
                print(f"[{file_name}] 为防止并发冲突，随机等待 {random_sleep:.2f} 秒...")
                time.sleep(random_sleep)

                print(f">>> 开始输入款号: {input_value}")
                RPAUtils.input_text_char_by_char(
                    input_element=search_input,
                    text_value=input_value,
                    char_interval=0.2,
                    pre_clear_sleep=0.8
                )

                if search_input.value != input_value:
                    print(f"   -> 检测到输入框值不匹配，强制修正...")
                    search_input.run_js(f'this.value = "{input_value}"')

                # 输入好款号，进行搜索，并获取查询返回
                res_packet = RPAUtils.search_with_network_listen(
                    tab=tab,
                    input_element=search_input,
                    target_url='Admin/MtReq/NewGet',
                    success_message="成功捕获接口数据",
                    retry_on_concurrent=True,
                    auto_stop_listen=False
                )

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
                            # 搜索返回的值与 ocr 识别到的码单信息发给大模型进行匹配
                            print(">>> 正在调用 LLM 进行智能匹配...")
                            match_result, match_prompt, retry_count = RPAUtils.execute_smart_match(data_json, records)

                            print("\n" + "=" * 30)
                            print(f"🤖 智能匹配结果: {match_result.get('status', 'FAIL').upper()}")
                            print(f"匹配原因: {match_result.get('global_reason', match_result.get('reason'))}")
                            print("=" * 30 + "\n")

                            matched_ids = []
                            structured_tasks = []

                            if match_result.get('status') == 'success':
                                structured_tasks = RPAUtils.reconstruct_rpa_data(match_result, data_json, original_records)
                                print(f">>> 数据组装完成，共生成 {len(structured_tasks)} 个任务包")

                                seen_ids = set()
                                for task in structured_tasks:
                                    rec_id = task['record'].get('Id')
                                    if rec_id and rec_id not in seen_ids:
                                        matched_ids.append(rec_id)
                                        seen_ids.add(rec_id)

                            if matched_ids:
                                RPAUtils.select_matched_checkboxes(tab, matched_ids)

                                print(">>> 正在查找并点击\"物料采购单\"生成按钮...")
                                button_found = False
                                scopes = [tab] + [f for f in tab.eles('tag:iframe') if f.states.is_displayed]

                                for scope in scopes:
                                    btn = scope.ele('x://button[contains(text(), "物料采购单")]', timeout=0.5)
                                    if btn and btn.states.is_displayed:
                                        btn.scroll.to_see()
                                        time.sleep(0.5)
                                        btn.click()
                                        print("✅ 成功点击\"物料采购单\"按钮")
                                        button_found = True
                                        # time.sleep(1.5)
                                        # 处理可能出现的合并采购确认弹窗
                                        RPAUtils.handle_purchase_order_popup(tab)
                                        
                                        time.sleep(2)
                                        break

                                if not button_found:
                                    print("⚠️ 未找到\"物料采购单\"按钮")
                                else:
                                    print(">>> 正在等待页面加载并切换为\"月结采购\"...")
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
                                                    print("✅ 成功选择\"月结采购\"")
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
                                            if RPAUtils.search_and_select_from_popup(
                                                scope=scope,
                                                trigger_selector='#lbSupplierInfo', 
                                                search_input_selector='#txtMpSupplierPlusContent',
                                                table_id='mtSupplierPlusGrid',
                                                search_value=supplier_name,
                                                item_name='供应商',
                                                timeout=1
                                            ):
                                                supplier_selected = True
                                                break
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
                                                if RPAUtils.fill_date_input(scope, '#Att01', ocr_date, 
                                                                        remove_readonly=False, trigger_events=True,
                                                                        scroll_to_see=True, timeout=0.5):
                                                    print("✅ 成功填写码单日期")
                                                    att01_filled = True
                                                    break
                                            except Exception:
                                                continue
                                        if not att01_filled:
                                            print("⚠️ 未找到码单日期输入框 (#Att01)")


                                    RPAUtils.fill_details_into_table(scope, structured_tasks)

                                    # 提取物料采购单总金额
                                    print(">>> 表格填写完毕，正在提取物料采购单总金额...")
                                    total_amount = RPAUtils.extract_total_amount_from_table(scope)
                                    if total_amount:
                                        print(f"✅ 成功提取总金额: {total_amount}")
                                        data_json['total_amount'] = total_amount
                                    else:
                                        print("⚠️ 未能提取到总金额")

                                    print(">>> 表格填写完毕，正在查找并点击\"保存并审核\"按钮...")
                                    try:
                                        save_btn = scope.ele('css:button[data-amid="btnSaveAndAudit"]', timeout=1)
                                        if not save_btn:
                                            save_btn = scope.ele('x://button[contains(text(), "保存并审核")]',
                                                                 timeout=1)

                                        if save_btn and save_btn.states.is_displayed:
                                            save_btn.scroll.to_see()
                                            time.sleep(0.5)
                                            save_btn.click()
                                            print("✅ 成功点击\"保存并审核\"")

                                            RPAUtils.handle_alert_confirmation(tab, timeout=2)

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

                                            print(">>> 准备跳转至\"物料采购订单\"列表...")
                                            time.sleep(0.5)

                                            try:
                                                if not RPAUtils.navigate_to_menu(tab, "物料", "物料采购订单"):
                                                    print("!!! 错误：无法导航到物料采购订单页面")
                                                else:
                                                    time.sleep(2)

                                                order_code = data_json.get('rpa_order_code')
                                                if order_code:
                                                    print(f">>> 准备在\"物料采购订单\"列表搜索单号: {order_code}")
                                                    search_input_order = RPAUtils.find_element_in_iframes(
                                                        tab=tab,
                                                        selector='css:input#txtSearchKey[data-grid="POMtPurchaseGrid"]',
                                                        max_retries=10,
                                                        retry_interval=0.5,
                                                        timeout=0.2
                                                    )

                                                    if search_input_order:
                                                        RPAUtils.input_text_char_by_char(
                                                            input_element=search_input_order,
                                                            text_value=order_code,
                                                            char_interval=0.1
                                                        )

                                                        res_packet_order = RPAUtils.search_with_network_listen(
                                                            tab=tab,
                                                            input_element=search_input_order,
                                                            target_url='Admin/MtPurchase',
                                                            success_message="搜索成功",
                                                            auto_stop_listen=False
                                                        )
                                                        if res_packet_order:
                                                            time.sleep(0.5)
                                                            all_selected = False
                                                            target_frame = None

                                                            for frame in tab.eles('tag:iframe'):
                                                                if not frame.states.is_displayed: continue
                                                                try:
                                                                    # 尝试使用通用的全选方法
                                                                    success, selected_count = RPAUtils.find_and_use_select_all_button(frame)
                                                                    if success:
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
                                                                            print("✅ 成功点击\"执行采购任务\"")
                                                                            time.sleep(2)

                                                                            try:
                                                                                confirm_btn = tab.ele(
                                                                                    'css:a.layui-layer-btn0', timeout=3)
                                                                                if not confirm_btn: confirm_btn = target_frame.ele(
                                                                                    'css:a.layui-layer-btn0', timeout=2)
                                                                                if confirm_btn:
                                                                                    confirm_btn.click()
                                                                                    time.sleep(1)
                                                                                    task_result = self.process_purchase_task(
                                                                                        tab,
                                                                                        data_json.get('rpa_order_code'),
                                                                                        data_json)
                                                                                    
                                                                                    if task_result and not task_result.get('success', True):
                                                                                        print(f"❌ 物料采购任务处理失败: {task_result.get('error', '未知错误')}")
                                                                                        print(f"   失败阶段: {task_result.get('error_stage', '未知阶段')}")
                                                                                        print("   跳过后续账单处理步骤")
                                                                                        # 记录失败并跳过账单处理
                                                                                        data_json['processing_failed'] = True
                                                                                        data_json['failure_reason'] = task_result.get('error', '物料采购任务处理失败')
                                                                                        data_json['failure_stage'] = task_result.get('error_stage', 'unknown')
                                                                                    # else:
                                                                                    #     # 只有成功时才处理账单
                                                                                    #     self.process_bill_list(tab,data_json.get('rpa_order_code'))
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

    def process_reconciliation_bill(self, tab):
        print("\n>>> [阶段: 新增对账单处理] 开始...")
        try:
            print(">>> 正在查找\"保存并审核\"按钮所在的 iframe...")
            save_audit_btn, target_frame = RPAUtils.find_element_in_iframes(
                tab=tab,
                selector='css:button[data-amid="btnPaySaveAndAduit"]',
                fallback_selectors=[
                    'css:button[onclick="saveRecord(1)"]',
                    'x://button[contains(text(), "保存并审核")]'
                ],
                max_retries=5,
                retry_interval=1,
                timeout=0.1,
                return_frame=True
            )

            if save_audit_btn:
                print("   -> 找到\"保存并审核\"按钮，准备点击...")
                save_audit_btn.scroll.to_see()
                time.sleep(0.5)
                save_audit_btn.click()
                print("✅ \"新增对账单\"审核流程完成")
            else:
                print("⚠️ 未在任何可见 iframe 中找到\"保存并审核\"按钮")
        except Exception as e:
            print(f"!!! 新增对账单处理异常: {e}")

    def process_bill_list(self, tab, order_code):
        print("\n>>> [阶段: 跳转账单列表] 开始处理...")
        try:
            if not RPAUtils.navigate_to_menu(tab, "财务", "账单列表"):
                print("!!! 错误：无法导航到账单列表页面")
                return
            
            time.sleep(2)

            print(f">>> 正在查找搜索框 (data-grid='FMAccountsReceivableGrid')...")
            if not order_code: return

            search_input_bill, target_frame = RPAUtils.find_element_in_iframes(
                tab=tab,
                selector='css:input#txtSearchKey[data-grid="FMAccountsReceivableGrid"]',
                max_retries=10,
                retry_interval=0.5,
                timeout=0.2,
                return_frame=True
            )

            if search_input_bill:
                print(f">>> 找到账单列表搜索框，正在输入: {order_code}")
                RPAUtils.input_text_char_by_char(
                    input_element=search_input_bill,
                    text_value=order_code,
                    char_interval=0.2
                )

                res = RPAUtils.search_with_network_listen(
                    tab=tab,
                    input_element=search_input_bill,
                    target_url='Admin/AccountsReceivable/NewGet',
                    success_message="账单列表搜索响应成功"
                )

                if res:
                    time.sleep(1)
                    if target_frame:
                        count_selected = RPAUtils.select_checkboxes_in_table_rows(
                            frame=target_frame,
                            table_selector='css:table#FMAccountsReceivableGrid tbody tr'
                        )

                        if count_selected > 0:
                            print(f"✅ 已勾选 {count_selected} 条账单记录")
                            print(">>> 准备点击\"发起对账\"...")
                            try:
                                btn_check = target_frame.ele('css:button[onclick="aReconciliation()"]', timeout=2)
                                if not btn_check: btn_check = target_frame.ele(
                                    'x://button[contains(text(), "发起对账")]', timeout=1)
                                if btn_check:
                                    btn_check.run_js('this.click()')
                                    time.sleep(2)
                                    print("✅ \"发起对账\"操作完成")
                                    print(">>> 等待\"新增对账单\"页面加载...")
                                    time.sleep(3)
                                    self.process_reconciliation_bill(tab)
                                else:
                                    print("⚠️ 未找到\"发起对账\"按钮")
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

    def process_purchase_task(self, tab, order_code, parsed_data):
        print(f"\n>>> [阶段: 跳转物料采购任务] 开始处理，目标单号: {order_code}")
        if not order_code: 
            return {"success": False, "error": "缺少订单编号", "error_stage": "missing_order_code"}

        delivery_date = parsed_data.get('delivery_date', '')
        delivery_order_no = parsed_data.get('delivery_order_number', '')

        try:
            if not RPAUtils.navigate_to_menu(tab, "物料", "物料采购任务"):
                print("!!! 错误：无法导航到物料采购任务页面")
                return {"success": False, "error": "无法导航到物料采购任务页面", "error_stage": "navigation_failed"}

            time.sleep(2)
            print(f">>> 正在查找搜索框 (data-grid='poMtPurTaskGrid')...")
            
            search_input_task, target_frame = RPAUtils.find_element_in_iframes(
                tab=tab,
                selector='css:input#txtSearchKey[data-grid="poMtPurTaskGrid"]',
                max_retries=10,
                retry_interval=0.5,
                timeout=0.2,
                return_frame=True
            )

            if search_input_task:
                print(f">>> 找到搜索框，正在输入: {order_code}")
                RPAUtils.input_text_char_by_char(
                    input_element=search_input_task,
                    text_value=order_code,
                    char_interval=0.2
                )

                res = RPAUtils.search_with_network_listen(
                    tab=tab,
                    input_element=search_input_task,
                    target_url='Admin/MtPurchase',
                    success_message="搜索响应成功"
                )

                if res:
                    time.sleep(1)
                    if target_frame:
                        # 使用通用方法勾选所有复选框
                        select_count = RPAUtils.select_all_checkboxes_in_frame(
                            frame=target_frame,
                            table_selector='css:table#poMtPurTaskGrid tbody tr',
                            label="首次勾选"
                        )
                        print(">>> [1/3] 准备点击\"一键绑定加工单\"...")
                        try:
                            btn_bind = target_frame.ele('#btnOneKeyBindPM', timeout=2)
                            if btn_bind:
                                btn_bind.click()
                                RPAUtils.handle_alert_confirmation(tab, timeout=3)
                                try:
                                    confirm_btn = tab.ele('css:a.layui-layer-btn0', timeout=3)
                                    if confirm_btn: confirm_btn.click()
                                except:
                                    pass
                                time.sleep(1)
                                RPAUtils.handle_alert_confirmation(tab, timeout=2)
                                print("   ✅ \"一键绑定\"操作结束")

                                print(">>> 等待系统处理一键绑定(页面可能刷新)...")
                                binding_completed = False
                                max_wait_time = 30  # 最大等待30秒
                                check_interval = 2  # 每2秒检查一次
                                
                                # 先获取总的记录行数
                                try:
                                    total_rows = target_frame.eles('css:table#poMtPurTaskGrid tbody tr', timeout=2)
                                    visible_rows = [row for row in total_rows if row.states.is_displayed]
                                    total_count = len(visible_rows)
                                    print(f"   -> 检测到 {total_count} 行记录需要处理")
                                except:
                                    total_count = 1  # 兜底，至少有1行
                                    print("   -> 无法获取行数，默认为1行")
                                
                                for attempt in range(max_wait_time // check_interval):
                                    try:
                                        # 必须重新从 iframe 获取元素，因为页面刷新了
                                        total_rows = target_frame.eles('css:table#poMtPurTaskGrid tbody tr', timeout=2)
                                        visible_rows = [r for r in total_rows if r.states.is_displayed]
                                        total_count = len(visible_rows)
                                        if total_count == 0:
                                            print(f"   -> 第{attempt + 1}次检查: 未发现可见行，继续等待...")
                                            time.sleep(check_interval)
                                            continue

                                        factory_cells = target_frame.eles('css:td[masking="SpName"]', timeout=1)
                                        completed_count = 0
                                        for cell in factory_cells:
                                            if cell.states.is_displayed and cell.text.strip():
                                                completed_count += 1

                                        if completed_count >= total_count and total_count > 0:
                                            print(f"   ✅ 系统处理完成，{completed_count}/{total_count} 行已绑定工厂")
                                            binding_completed = True
                                            break
                                            
                                        print(f"   -> 第{attempt + 1}次检查: {completed_count}/{total_count} 行已完成，继续等待...")
                                        time.sleep(check_interval)
                                    except Exception as e:
                                        print(f"   ⚠️ 检查加工厂字段时出错: {e}")
                                        time.sleep(check_interval)

                                if not binding_completed:
                                    print("   ❌ 一键绑定失败：加工厂字段未填充，单据处理终止")
                                    # 设置失败标记，确保报告能正确记录失败状态
                                    parsed_data['processing_failed'] = True
                                    parsed_data['failure_reason'] = "一键绑定失败：系统无法为此单据自动分配加工厂"
                                    parsed_data['failure_stage'] = 'binding_failed'
                                    return {
                                        "success": False, 
                                        "error": "一键绑定失败：系统无法为此单据自动分配加工厂",
                                        "error_stage": "binding_failed"
                                    }

                                print(">>> [系统修复] 页面已刷新，正在重新定位 iframe 上下文...")
                                time.sleep(1)  # 等待渲染
                                frame_refreshed = False

                                # 重新遍历所有可见 iframe，找到包含特征元素的那个
                                for frame in tab.eles('tag:iframe'):
                                    if not frame.states.is_displayed: continue
                                    # 使用特有的搜索框作为特征来确认是不是目标 frame
                                    if frame.ele('css:input#txtSearchKey[data-grid="poMtPurTaskGrid"]', timeout=0.5):
                                        target_frame = frame
                                        frame_refreshed = True
                                        print("✅成功重新获取 iframe 对象")
                                        break

                                if not frame_refreshed:
                                    print("   ❌ 严重错误：页面刷新后无法找回 iframe，流程终止")
                                    return  # 找不到就直接停止，防止后面报 'NoneType' 错误

                                print(">>> 一键绑定完成，开始填写码单信息...")
                                try:
                                    # 重新验证并获取target_frame（防止页面刷新导致iframe失效）
                                    print(">>> 重新验证iframe上下文...")
                                    current_frame = RPAUtils.find_element_in_iframes(
                                        tab=tab,
                                        selector='css:input#txtSearchKey[data-grid="poMtPurTaskGrid"]',
                                        max_retries=3,
                                        retry_interval=1,
                                        timeout=0.5,
                                        return_frame=True
                                    )
                                    if current_frame and len(current_frame) == 2:
                                        _, verified_frame = current_frame
                                        target_frame = verified_frame
                                        print(">>> ✅ iframe上下文验证成功，继续使用更新后的frame")
                                    else:
                                        print(">>> ⚠️ iframe验证失败，继续使用原frame")
                                    
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
                                                RPAUtils.fill_date_input(
                                                    scope=row,  # 传入行元素
                                                    selector='css:input.Att02',
                                                    date_value=delivery_date,
                                                    remove_readonly=True,
                                                    use_enter_key=True,  # 必须为 True
                                                    click_body_after=True,  # 必须为 True
                                                    timeout=0.5
                                                )
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
                        
                        # 在提交前再次验证iframe有效性
                        print(">>> 提交前验证iframe上下文...")
                        submit_frame = RPAUtils.find_element_in_iframes(
                            tab=tab,
                            selector='css:table#poMtPurTaskGrid',
                            max_retries=3,
                            retry_interval=1,
                            timeout=0.5,
                            return_frame=True
                        )
                        if submit_frame and len(submit_frame) == 2:
                            _, verified_frame = submit_frame
                            target_frame = verified_frame
                            print(">>> ✅ 提交前iframe验证成功")
                        else:
                            print(">>> ⚠️ 提交前iframe验证失败，继续使用原frame")
                        
                        reselect_count = RPAUtils.select_all_checkboxes_in_frame(
                            frame=target_frame,
                            table_selector='css:table#poMtPurTaskGrid tbody tr',
                            label="提交前重选"
                        )
                        print(f"✅ 已确认勾选 {reselect_count} 行")

                        print(">>> [2/3] 准备点击\"提交\"...")
                        try:
                            btn_submit = target_frame.ele('#btnSubmitTasks', timeout=2)
                            if btn_submit:
                                btn_submit.click()
                                RPAUtils.handle_alert_confirmation(tab, timeout=3)
                                time.sleep(1)
                                RPAUtils.handle_alert_confirmation(tab, timeout=2)
                                print("   ✅ \"提交\"操作结束")
                                time.sleep(2)
                            else:
                                print("⚠️ 未找到\"提交\"按钮")
                        except Exception as e:
                            print(f"!!! 提交操作异常: {e}")

                        print(">>> [3/3] 准备点击\"确认\"...")
                        try:
                            # 确认前再次验证iframe上下文
                            print(">>> 确认前验证iframe上下文...")
                            confirm_frame = RPAUtils.find_element_in_iframes(
                                tab=tab,
                                selector='css:button#btnConfirmToDoTask',
                                max_retries=3,
                                retry_interval=1,
                                timeout=0.5,
                                return_frame=True
                            )
                            if confirm_frame and len(confirm_frame) == 2:
                                _, verified_frame = confirm_frame
                                target_frame = verified_frame
                                print(">>> ✅ 确认前iframe验证成功")
                            else:
                                print(">>> ⚠️ 确认前iframe验证失败，继续使用原frame")
                                
                            btn_confirm = target_frame.ele('#btnConfirmToDoTask', timeout=2)
                            if btn_confirm:
                                btn_confirm.click()
                                RPAUtils.handle_alert_confirmation(tab, timeout=3)
                                time.sleep(2)
                                try:
                                    lay_confirm = tab.ele('css:a.layui-layer-btn0', timeout=3)
                                    if not lay_confirm: lay_confirm = target_frame.ele('css:a.layui-layer-btn0',
                                                                                       timeout=1)
                                    if lay_confirm: lay_confirm.click()
                                except:
                                    pass
                                print("✅ \"确认\"操作全部完成")
                                return {"success": True}
                            else:
                                print("⚠️ 未找到\"确认\"按钮")
                                return {"success": False, "error": "未找到确认按钮", "error_stage": "confirm_failed"}
                        except Exception as e:
                            print(f"!!! 确认操作异常: {e}")
                            return {"success": False, "error": f"确认操作异常: {e}", "error_stage": "confirm_exception"}
                    else:
                        print("!!! 错误: 丢失了 iframe 上下文")
                        return {"success": False, "error": "丢失了iframe上下文", "error_stage": "iframe_lost"}
                else:
                    print("⚠️ 搜索超时")
                    return {"success": False, "error": "搜索超时", "error_stage": "search_timeout"}
            else:
                print("!!! 错误: 未找到搜索框")
                return {"success": False, "error": "未找到搜索框", "error_stage": "search_box_not_found"}
        except Exception as e:
            print(f"!!! 跳转或搜索'物料采购任务'时发生异常: {e}")
            return {"success": False, "error": f"跳转或搜索异常: {e}", "error_stage": "navigation_exception"}

