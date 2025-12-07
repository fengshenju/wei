import time
import os
import json
from app_config import CONFIG
from services.match_service import MatchService
from services.data_processor import DataProcessor


class RpaService:
    def __init__(self):
        self.match_service = MatchService()

    def handle_new_reconciliation_bill(self, tab):
        """处理新增对账单页面"""
        print("\n>>> [阶段: 新增对账单处理] 开始...")
        try:
            target_frame = None
            save_audit_btn = None
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
                print("   -> 找到\"保存并审核\"按钮，准备点击...")
                save_audit_btn.scroll.to_see()
                time.sleep(0.5)
                save_audit_btn.click()
                print("✅ \"新增对账单\"审核流程完成")
            else:
                print("⚠️ 未在任何可见 iframe 中找到\"保存并审核\"按钮")
        except Exception as e:
            print(f"!!! 新增对账单处理异常: {e}")

    def navigate_to_bill_list(self, tab, order_code):
        """跳转账单列表并发起对账"""
        print("\n>>> [阶段: 跳转账单列表] 开始处理...")
        try:
            finance_btn = tab.ele('x://div[contains(@class, "title") and .//div[contains(text(), "财务")]]')
            if finance_btn:
                finance_btn.click()
                time.sleep(0.5)
            else:
                print("!!! 错误: 未找到\"财务\"菜单")
                return

            target_menu_text = "账单列表"
            menu_xpath = f'x://a[contains(text(), "{target_menu_text}")]'
            bill_menu = tab.wait.ele_displayed(menu_xpath, timeout=3)
            if bill_menu:
                bill_menu.click()
                print(f"✅ 成功点击左侧菜单\"{target_menu_text}\"")
                time.sleep(2)
            else:
                print(f"⚠️ 未检测到二级菜单，尝试重新展开一级菜单...")
                if finance_btn:
                    finance_btn.click()
                    time.sleep(0.5)
                bill_menu = tab.wait.ele_displayed(menu_xpath, timeout=3)
                if bill_menu:
                    bill_menu.click()
                    print(f"✅ (重试) 成功点击左侧菜单\"{target_menu_text}\"")
                    time.sleep(2)
                else:
                    print(f"!!! 错误: 无法找到左侧菜单项\"{target_menu_text}\"")
                    return

            print(f">>> 正在查找搜索框 (data-grid='FMAccountsReceivableGrid')...")
            if not order_code:
                print("⚠️ 警告: 未获取到有效的订单编号，跳过搜索")
                return

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
                    print(">>> 正在遍历记录并勾选所有记录...")
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
                                            print(f"   ✅ 已勾选行")
                                        else:
                                            print(f"   ℹ️ 行已被勾选")
                                        count_selected += 1
                                    else:
                                        print("   ⚠️ 未找到复选框")
                                    time.sleep(0.1)
                                except Exception as inner_e:
                                    print(f"   !!! 勾选行出错: {inner_e}")

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
                                        print("✅ \"发起对账\"操作完成")
                                        print(">>> 等待\"新增对账单\"页面加载...")
                                        time.sleep(3)
                                        self.handle_new_reconciliation_bill(tab)
                                    else:
                                        print("⚠️ 未找到“发起对账”按钮")
                                except Exception as e:
                                    print(f"!!! 发起对账操作异常: {e}")
                            else:
                                print("⚠️ 未勾选任何记录，跳过“发起对账”")
                    else:
                        print("!!! 错误: 丢失了 iframe 上下文")
                else:
                    print("⚠️ 搜索超时，未收到响应")
            else:
                print("!!! 错误: 未找到账单列表搜索框")
        except Exception as e:
            print(f"!!! 跳转账单列表时发生异常: {e}")

    def navigate_and_search_purchase_task(self, tab, order_code, parsed_data):
        """跳转物料采购任务并处理"""
        print(f"\n>>> [阶段: 跳转物料采购任务] 开始处理，目标单号: {order_code}")
        if not order_code:
            print("⚠️ 错误: 未获取到有效的订单编号，无法执行搜索。")
            return
        delivery_date = parsed_data.get('delivery_date', '')
        delivery_order_no = parsed_data.get('delivery_order_number', '')

        try:
            print(">>> 正在重新定位“物料”菜单...")
            material_btn_nav = tab.ele('x://div[contains(@class, "title") and .//div[contains(text(), "物料")]]')
            if material_btn_nav:
                material_btn_nav.click()
                time.sleep(0.5)

            target_menu_text = "物料采购任务"
            menu_xpath = f'x://a[contains(text(), "{target_menu_text}")]'
            task_menu = tab.wait.ele_displayed(menu_xpath, timeout=3)
            if task_menu:
                task_menu.click()
                print(f"✅ 成功点击左侧菜单\"{target_menu_text}\"")
            else:
                print(f"⚠️ 未检测到菜单，尝试重新展开一级菜单...")
                if material_btn_nav:
                    material_btn_nav.click()
                    time.sleep(0.5)
                task_menu = tab.wait.ele_displayed(menu_xpath, timeout=3)
                if task_menu:
                    task_menu.click()
                    print(f"✅ (重试) 成功点击左侧菜单\"{target_menu_text}\"")
                else:
                    print(f"!!! 错误: 无法找到左侧菜单项“{target_menu_text}”")
                    return

            time.sleep(2)
            print(f">>> 正在查找搜索框 (data-grid='poMtPurTaskGrid')...")
            search_input_task = None
            target_frame = None

            for i in range(20):
                for frame in tab.eles('tag:iframe'):
                    if not frame.states.is_displayed: continue
                    ele = frame.ele('css:input#txtSearchKey[data-grid="poMtPurTaskGrid"]', timeout=0.5)
                    if ele and ele.states.is_displayed:
                        search_input_task = ele
                        target_frame = frame
                        break
                if search_input_task:
                    print(f"   -> 在第 {i + 1} 次尝试中找到搜索框")
                    break
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
                    print(">>> 正在遍历记录并勾选所有记录...")
                    time.sleep(1)
                    if target_frame:
                        rows = target_frame.eles('css:table#poMtPurTaskGrid tbody tr', timeout=2)
                        count_selected = 0
                        for row in rows:
                            if not row.states.is_displayed: continue
                            try:
                                row.scroll.to_see()
                                checkbox = row.ele('css:input.ckbox', timeout=0.5)
                                if checkbox:
                                    if not checkbox.states.is_checked:
                                        checkbox.click()
                                        print(f"   ✅ 已勾选行")
                                        count_selected += 1
                                    else:
                                        print(f"   ℹ️ 行已被勾选")
                                else:
                                    print("   ⚠️ 未找到复选框")
                                time.sleep(0.1)
                            except Exception as inner_e:
                                print(f"   !!! 勾选行时出错: {inner_e}")
                        print(f"✅ 记录勾选完成，共勾选 {count_selected} 行")

                        print(">>> [1/3] 准备点击\"一键绑定加工单\"...")
                        try:
                            btn_bind = target_frame.ele('#btnOneKeyBindPM', timeout=2)
                            if btn_bind:
                                btn_bind.click()
                                try:
                                    if tab.wait.alert(timeout=3):
                                        print(f"   ℹ️ 绑定确认弹窗: {tab.alert.text} -> 自动接受")
                                        tab.alert.accept()
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

                                print(">>> 等待系统处理一键绑定，检查所有行的加工厂字段...")
                                binding_completed = False
                                max_wait_time = 30
                                check_interval = 2
                                try:
                                    total_rows = target_frame.eles('css:table#poMtPurTaskGrid tbody tr', timeout=2)
                                    visible_rows = [row for row in total_rows if row.states.is_displayed]
                                    total_count = len(visible_rows)
                                    print(f"   -> 检测到 {total_count} 行记录需要处理")
                                except:
                                    total_count = 1
                                    print("   -> 无法获取行数，默认为1行")

                                for attempt in range(max_wait_time // check_interval):
                                    try:
                                        factory_cells = target_frame.eles('css:td[masking="SpName"]', timeout=1)
                                        completed_count = 0
                                        for cell in factory_cells:
                                            if cell.states.is_displayed:
                                                cell_text = cell.text.strip()
                                                if cell_text: completed_count += 1
                                        if completed_count == total_count and completed_count > 0:
                                            print(f"   ✅ 所有 {total_count} 行记录的加工厂信息都已填入，系统处理完成")
                                            binding_completed = True
                                            break
                                        print(
                                            f"   -> 第{attempt + 1}次检查: {completed_count}/{total_count} 行已完成，继续等待...")
                                        time.sleep(check_interval)
                                    except Exception as e:
                                        print(f"   ⚠️ 检查加工厂字段时出错: {e}")
                                        time.sleep(check_interval)

                                if not binding_completed:
                                    print("   ⚠️ 等待超时，但继续执行后续操作...")
                                    time.sleep(2)

                                print(">>> 一键绑定完成，开始填写码单信息...")
                                try:
                                    rows = target_frame.eles('css:table#poMtPurTaskGrid tbody tr', timeout=3)
                                    count_filled = 0
                                    for row in rows:
                                        if not row.states.is_displayed: continue
                                        try:
                                            row.scroll.to_see()
                                            if delivery_order_no:
                                                inp_no = row.ele('css:input.Att01', timeout=0.2)
                                                if inp_no:
                                                    js_no = f'this.value = "{delivery_order_no}"; this.dispatchEvent(new Event("input")); this.dispatchEvent(new Event("change")); this.dispatchEvent(new Event("blur"));'
                                                    inp_no.run_js(js_no)
                                                    print(f"   -> 已填入码单编号: {delivery_order_no}")

                                            if delivery_date:
                                                inp_date = row.ele('css:input.Att02', timeout=0.2)
                                                if inp_date:
                                                    try:
                                                        print(f"   -> 正在填入码单日期: {delivery_date}")
                                                        inp_date.run_js('this.removeAttribute("readonly");')
                                                        inp_date.clear()
                                                        time.sleep(0.1)
                                                        inp_date.input(delivery_date)
                                                        time.sleep(0.2)
                                                        target_frame.actions.key_down('ENTER').key_up('ENTER')
                                                        time.sleep(0.2)
                                                        target_frame.run_js('document.body.click();')
                                                        inp_date.click()
                                                        time.sleep(0.2)
                                                    except Exception as e:
                                                        print(f"   ⚠️ 日期输入异常: {e}")
                                                        try:
                                                            inp_date.run_js(
                                                                f'this.removeAttribute("readonly"); this.value="{delivery_date}";')
                                                        except:
                                                            pass
                                            count_filled += 1
                                            time.sleep(0.1)
                                        except Exception as inner_e:
                                            print(f"   !!! 填写行数据时出错: {inner_e}")
                                    print(f"✅ 码单信息填写完成，共处理 {count_filled} 行")
                                except Exception as e:
                                    print(f"!!! 填写码单信息时发生异常: {e}")
                            else:
                                print("⚠️ 未找到一键绑定加工单按钮")
                        except Exception as e:
                            print(f"!!! 绑定操作异常: {e}")

                        print(">>> [2/3] 准备点击\"提交\"...")
                        try:
                            btn_submit = target_frame.ele('#btnSubmitTasks', timeout=2)
                            if btn_submit:
                                btn_submit.click()
                                try:
                                    if tab.wait.alert(timeout=3):
                                        print(f"   ℹ️ 提交确认弹窗: {tab.alert.text} -> 自动接受")
                                        tab.alert.accept()
                                except:
                                    pass
                                time.sleep(1)
                                try:
                                    if tab.wait.alert(timeout=2): tab.alert.accept()
                                except:
                                    pass
                                print("   ✅ \"提交\"操作结束")
                                time.sleep(2)
                            else:
                                print("⚠️ 未找到\"提交\"按钮")
                        except Exception as e:
                            print(f"!!! 提交操作异常: {e}")

                        print(">>> [3/3] 准备点击\"确认\"...")
                        try:
                            btn_confirm = target_frame.ele('#btnConfirmToDoTask', timeout=2)
                            if btn_confirm:
                                btn_confirm.click()
                                try:
                                    if tab.wait.alert(timeout=3):
                                        print(f"   ℹ️ 确认操作弹窗: {tab.alert.text} -> 自动接受")
                                        tab.alert.accept()
                                except:
                                    pass
                                print("   -> 等待系统处理确认逻辑...")
                                time.sleep(2)
                                print("   -> 检查\"成功确认\"Layui弹窗...")
                                try:
                                    lay_confirm = tab.ele('css:a.layui-layer-btn0', timeout=3)
                                    if not lay_confirm: lay_confirm = target_frame.ele('css:a.layui-layer-btn0',
                                                                                       timeout=1)
                                    if lay_confirm:
                                        lay_confirm.click()
                                        print("   ✅ 检测到Layui成功弹窗，已点击\"确定\"关闭")
                                    else:
                                        print("   ℹ️ 未检测到Layui成功弹窗 (可能已自动关闭或无提示)")
                                except Exception as e:
                                    print(f"   ⚠️ 处理Layui弹窗时出错 (非阻断): {e}")
                                print("✅ \"确认\"操作全部完成")
                            else:
                                print("⚠️ 未找到\"确认\"按钮")
                        except Exception as e:
                            print(f"!!! 确认操作异常: {e}")
                    else:
                        print("!!! 错误: 丢失了 iframe 上下文")
                else:
                    print("⚠️ 搜索超时，未收到响应")
            else:
                print("!!! 错误: 未找到搜索框")
        except Exception as e:
            print(f"!!! 跳转或搜索'物料采购任务'时发生异常: {e}")

    def select_matched_checkboxes(self, tab, matched_ids):
        """根据匹配的记录ID勾选对应行的checkbox"""
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
                        else:
                            print(f"⚠️ 记录 {record_id} 已被勾选")
                        checkbox_found = True
                        break
                if not checkbox_found:
                    print(f"⚠️ 未找到记录 {record_id} 对应的checkbox")
            except Exception as e:
                print(f"!!! 勾选记录 {record_id} 时发生异常: {e}")
        print(f">>> 勾选操作完成")

    def fill_details_into_table(self, scope, structured_tasks):
        """根据匹配任务填充 tbody 中的物料数据"""
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
                    target_unit, target_price, target_qty, target_date = raw_unit, raw_price, first_item.get('qty',
                                                                                                             0), raw_date
                elif match_type == 'MERGE':
                    target_qty = sum([float(i.get('qty', 0)) for i in items])
                    target_unit, target_price, target_date = raw_unit, raw_price, raw_date
                    print(f"   ℹ️ [合并] 记录 {record_id} 聚合了 {len(items)} 条明细，总数: {target_qty}")
                elif match_type == 'SPLIT':
                    target_unit, target_price, target_qty, target_date = raw_unit, raw_price, first_item.get('qty',
                                                                                                             0), raw_date
                    print(f"   ℹ️ [拆分] 记录 {record_id} 强制调整数量为: {target_qty}")

                # 填充单位
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
                                js_code = "this.click(); this.dispatchEvent(new MouseEvent('dblclick', {bubbles: true, cancelable: true, view: window}));"
                                target_td.run_js(js_code)
                                time.sleep(0.5)
                            else:
                                inp_unit.click()
                        else:
                            inp_unit.input(target_unit, clear=True)

                # 填充价格
                if target_price is not None:
                    inp_price = tr.ele('css:input[name="Price"]', timeout=0.2)
                    if inp_price:
                        val = str(target_price)
                        inp_price.run_js(
                            f'this.value = "{val}"; this.dispatchEvent(new Event("input")); this.dispatchEvent(new Event("change")); this.dispatchEvent(new Event("blur"));')

                # 填充数量
                if target_qty is not None:
                    inp_qty = tr.ele('css:input[name="Qty"]', timeout=0.2)
                    if inp_qty:
                        val = str(target_qty)
                        inp_qty.run_js(
                            f'this.value = "{val}"; this.dispatchEvent(new Event("input")); this.dispatchEvent(new Event("change")); this.dispatchEvent(new Event("blur"));')

                # 触发总价计算
                inp_total = tr.ele('css:input[name="totalAmount"]', timeout=0.2)
                if inp_total:
                    inp_total.click()

                # 填充日期
                if target_date and target_date.strip():
                    inp_date = tr.ele('css:input.deliveryDate', timeout=0.5)
                    if inp_date:
                        try:
                            inp_date.run_js('this.removeAttribute("readonly");')
                            inp_date.clear()
                            inp_date.input(target_date)
                            scope.actions.key_down('ENTER').key_up('ENTER')
                            scope.run_js('document.body.click();')
                        except Exception as e:
                            try:
                                inp_date.run_js(f'this.removeAttribute("readonly"); this.value="{target_date}";')
                            except:
                                pass

                count_success += 1
                time.sleep(0.1)
            except Exception as e:
                print(f"   !!! 填充行数据失败: {e}")
        print(f"✅ 数据填充完成: 成功处理 {count_success}/{len(structured_tasks)} 行")

    def process_single_bill_rpa(self, browser, data_json, file_name, img_path):
        """
        完整RPA流程方法
        """
        print(f"\n--- [RPA阶段] 开始处理: {file_name} ---")

        match_prompt = ""
        match_result = None
        original_records = []
        retry_count = 1
        tab = None

        try:
            # 1. 创建新页签并初始化
            print(f"[{file_name}] 正在创建新页签...")
            tab = browser.new_tab()
            if CONFIG.get('rpa_browser_to_front', True):
                tab.set.window.normal()
                time.sleep(0.5)
                tab.set.window.full()

            tab.get(CONFIG['base_url'])

            if not CONFIG.get('rpa_browser_to_front', True):
                tab.set.window.mini()

            # 2. 菜单导航
            if not tab.wait.ele_displayed('.fixed-left-menu', timeout=5):
                print("!!! 错误: 未检测到左侧菜单栏，请确认网页已加载完成。")
                return "", None, [], 1

            # 定位“物料”菜单按钮
            material_btn = tab.ele('x://div[contains(@class, "title") and .//div[contains(text(), "物料")]]')
            if material_btn:
                material_btn.click()
            else:
                print("!!! 未找到物料菜单")

            # 定位“物料采购需求”子菜单
            sub_menu_btn = tab.wait.ele_displayed('x://a[contains(text(), "物料采购需求")]', timeout=3)
            if sub_menu_btn:
                sub_menu_btn.click()
            else:
                print("⚠️ 未检测到二级菜单展开，尝试重新点击“物料”...")
                if material_btn: material_btn.click()
                time.sleep(1)
                sub_menu_btn = tab.wait.ele_displayed('x://a[contains(text(), "物料采购需求")]', timeout=3)
                if sub_menu_btn:
                    sub_menu_btn.click()
                else:
                    print("!!! 错误：无法展开二级菜单")
                    return "", None, [], 1

            # 3. 搜索款号
            search_input = None
            for _ in range(20):
                for frame in tab.eles('tag:iframe'):
                    if not frame.states.is_displayed: continue
                    ele = frame.ele('#txtSearchKey', timeout=0.2)
                    if ele and ele.states.is_displayed:
                        search_input = ele
                        break
                if search_input: break
                time.sleep(0.5)

            if search_input:
                input_value = data_json.get('final_selected_style', '')
                if not input_value:
                    print(f"⚠️ 警告: 款号为空，跳过RPA处理")
                    return "", None, [], 1

                print(f">>> 开始输入款号: {input_value}")
                search_input.click()
                search_input.clear()
                for char in input_value:
                    search_input.input(char, clear=False)
                    time.sleep(0.1)

                target_url_substring = 'Admin/MtReq/NewGet'
                tab.listen.start(targets=target_url_substring)

                search_input.run_js("""
                    arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                    arguments[0].dispatchEvent(new KeyboardEvent("keydown", {bubbles:true, keyCode:13, key:"Enter"}));
                    arguments[0].dispatchEvent(new KeyboardEvent("keyup", {bubbles:true, keyCode:13, key:"Enter"}));
                """, search_input)

                res_packet = tab.listen.wait(timeout=10)

                if res_packet:
                    print(f"✅ 成功捕获接口数据: {res_packet.url}")
                    response_body = res_packet.response.body

                    if isinstance(response_body, dict):
                        records = response_body.get('data', [])
                        print(f"数据统计: 共找到 {len(records)} 条记录")

                        if records:
                            original_records = records
                            # 4. LLM 智能匹配
                            match_result, match_prompt, retry_count = self.match_service.execute_smart_match(data_json,
                                                                                                             records)

                            print(f"🤖 智能匹配结果: {match_result.get('status', 'FAIL')}")

                            matched_ids = []
                            structured_tasks = []
                            if match_result.get('status') == 'success':
                                structured_tasks = DataProcessor.reconstruct_rpa_data(match_result, data_json,
                                                                                      original_records)
                                seen_ids = set()
                                for task in structured_tasks:
                                    rec_id = task['record'].get('Id')
                                    if rec_id and rec_id not in seen_ids:
                                        matched_ids.append(rec_id)
                                        seen_ids.add(rec_id)

                            # 5. 执行 RPA 动作
                            if matched_ids:
                                self.select_matched_checkboxes(tab, matched_ids)

                                # 点击“物料采购单”生成按钮
                                button_found = False
                                scopes = [tab] + [f for f in tab.eles('tag:iframe') if f.states.is_displayed]

                                for scope in scopes:
                                    btn = scope.ele('x://button[contains(text(), "物料采购单")]', timeout=0.5)
                                    if btn and btn.states.is_displayed:
                                        btn.click()
                                        button_found = True
                                        time.sleep(2)
                                        break

                                if button_found:
                                    # 处理新页面弹窗逻辑
                                    current_scopes = [tab] + [f for f in tab.eles('tag:iframe') if
                                                              f.states.is_displayed]

                                    # 选择"月结采购"
                                    for scope in current_scopes:
                                        try:
                                            dropdown_btn = scope.ele('css:button[data-id="OrderTypeId"]', timeout=0.5)
                                            if dropdown_btn:
                                                dropdown_btn.click()
                                                time.sleep(0.5)
                                                scope.ele('x://span[@class="text" and text()="月结采购"]').click()
                                                break
                                        except:
                                            continue

                                    # 设置供应商
                                    supplier_name = data_json.get('supplier_name', '').strip()
                                    if supplier_name:
                                        for scope in current_scopes:
                                            try:
                                                slabel = scope.ele('#lbSupplierInfo', timeout=0.5)
                                                if slabel:
                                                    slabel.click()
                                                    time.sleep(0.5)
                                                    sbox = scope.ele('#txtMpSupplierPlusContent')
                                                    sbox.clear()
                                                    sbox.input(supplier_name)
                                                    time.sleep(0.2)
                                                    scope.actions.key_down('ENTER').key_up('ENTER')
                                                    time.sleep(0.5)
                                                    td = scope.ele(
                                                        f'x://table[@id="mtSupplierPlusGrid"]//tbody//tr//td[text()="{supplier_name}"]')
                                                    if td:
                                                        td.run_js(
                                                            "this.click(); this.dispatchEvent(new MouseEvent('dblclick', {bubbles: true, cancelable: true, view: window}));")
                                                        break
                                            except:
                                                continue

                                    # 选择品牌
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
                                        for scope in current_scopes:
                                            try:
                                                bbtn = scope.ele('css:button[data-id="BrandId"]', timeout=0.3)
                                                if bbtn:
                                                    bbtn.click()
                                                    time.sleep(0.5)
                                                    opt = scope.ele(
                                                        f'x:.//span[contains(@class, "text") and contains(text(), "{target_brand}")]')
                                                    if opt:
                                                        opt.click()
                                                        break
                                            except:
                                                continue

                                    # 填写日期和明细
                                    ocr_date = data_json.get('delivery_date', '')
                                    if ocr_date:
                                        for scope in current_scopes:
                                            try:
                                                att01 = scope.ele('#Att01', timeout=0.5)
                                                if att01:
                                                    att01.clear()
                                                    att01.input(ocr_date)
                                                    att01.run_js(
                                                        'this.dispatchEvent(new Event("change", {bubbles: true})); this.dispatchEvent(new Event("blur"));')
                                                    break
                                            except:
                                                continue

                                    self.fill_details_into_table(scope, structured_tasks)

                                    # 保存并审核
                                    print(">>> 正在保存并审核...")
                                    save_btn = scope.ele('css:button[data-amid="btnSaveAndAudit"]', timeout=1)
                                    if not save_btn: save_btn = scope.ele('x://button[contains(text(), "保存并审核")]',
                                                                          timeout=1)

                                    if save_btn:
                                        save_btn.click()
                                        try:
                                            if tab.wait.alert(timeout=2): tab.alert.accept()
                                        except:
                                            pass

                                        # 获取生成的单号
                                        order_code = None
                                        for attempt in range(10):
                                            code_input = scope.ele('#Code', timeout=2)
                                            if code_input:
                                                val = code_input.value or code_input.attr(
                                                    'valuecontent') or code_input.attr('value')
                                                if val and val.strip():
                                                    order_code = val.strip()
                                                    data_json['rpa_order_code'] = order_code
                                                    print(f"✅ 获取到订单编号: {order_code}")
                                                    break
                                            time.sleep(1)

                                        # 如果成功获取单号，跳转到采购订单列表
                                        if order_code:
                                            # 跳转菜单
                                            material_nav = tab.ele(
                                                'x://div[contains(@class, "title") and .//div[contains(text(), "物料")]]')
                                            if material_nav: material_nav.click()

                                            po_menu = tab.wait.ele_displayed('x://a[contains(text(), "物料采购订单")]',
                                                                             timeout=3)
                                            if po_menu: po_menu.click()
                                            time.sleep(2)

                                            # 搜索订单
                                            search_po = None
                                            for _ in range(10):
                                                for fr in tab.eles('tag:iframe'):
                                                    if not fr.states.is_displayed: continue
                                                    ele = fr.ele('css:input#txtSearchKey[data-grid="POMtPurchaseGrid"]',
                                                                 timeout=0.2)
                                                    if ele:
                                                        search_po = ele
                                                        break
                                                if search_po: break
                                                time.sleep(0.5)

                                            if search_po:
                                                search_po.click()
                                                search_po.clear()
                                                for char in order_code:
                                                    search_po.input(char, clear=False)
                                                    time.sleep(0.1)

                                                tab.listen.start(targets='Admin/MtPurchase')
                                                search_po.run_js("""
                                                    this.dispatchEvent(new Event('change', { bubbles: true }));
                                                    this.dispatchEvent(new KeyboardEvent("keydown", {bubbles:true, keyCode:13, key:"Enter"}));
                                                    this.dispatchEvent(new KeyboardEvent("keyup", {bubbles:true, keyCode:13, key:"Enter"}));
                                                """)

                                                res_po = tab.listen.wait(timeout=10)
                                                if res_po:
                                                    # 选中记录
                                                    time.sleep(0.5)
                                                    target_fr = None
                                                    for fr in tab.eles('tag:iframe'):
                                                        if not fr.states.is_displayed: continue
                                                        sel_btn = fr.ele('x://button[contains(text(), "全选")]',
                                                                         timeout=0.5)
                                                        if sel_btn:
                                                            sel_btn.click()
                                                            target_fr = fr
                                                            break
                                                        else:
                                                            cks = fr.eles('x://tr//input[contains(@class, "ckbox")]')
                                                            if cks:
                                                                for ck in cks:
                                                                    if ck.states.is_displayed: ck.click()
                                                                target_fr = fr
                                                                break

                                                    # 上传附件
                                                    if target_fr:
                                                        adj_tab = target_fr.ele('x://a[contains(text(), "附件")]',
                                                                                timeout=2)
                                                        if adj_tab:
                                                            adj_tab.click()
                                                            up_lbl = target_fr.ele(
                                                                'x://div[@id="tb_Adjunct"]//label[contains(@style, "opacity: 0")]',
                                                                timeout=2)
                                                            if up_lbl:
                                                                up_lbl.click.to_upload(os.path.abspath(img_path))
                                                                time.sleep(5)  # 等待上传
                                                                save_img = target_fr.ele(
                                                                    'x://button[@onclick="AddImg()"]', timeout=2)
                                                                if save_img: save_img.click()

                                                        # 执行采购任务
                                                        more_btn = target_fr.ele('x://button[contains(text(), "更多")]',
                                                                                 timeout=2)
                                                        if more_btn: more_btn.click()

                                                        do_task = target_fr.ele('css:a[onclick="doMtPurTask()"]',
                                                                                timeout=1)
                                                        if do_task:
                                                            do_task.click()
                                                            try:
                                                                if tab.wait.alert(timeout=3): tab.alert.accept()
                                                            except:
                                                                pass

                                                            # 确认弹窗
                                                            cfm = tab.ele('css:a.layui-layer-btn0', timeout=3)
                                                            if not cfm: cfm = target_fr.ele('css:a.layui-layer-btn0',
                                                                                            timeout=2)
                                                            if cfm: cfm.click()

                                                            # 后续流程
                                                            self.navigate_and_search_purchase_task(tab, order_code,
                                                                                                   data_json)
                                                            self.navigate_to_bill_list(tab, order_code)

                                                try:
                                                    tab.listen.stop()
                                                except:
                                                    pass
                                            else:
                                                print("⚠️ 未找到采购订单搜索框")
                                else:
                                    print("⚠️ 未找到物料采购单生成按钮")
                        else:
                            print("⚠️ 搜索结果为空")
                    else:
                        print("⚠️ 响应非JSON格式")
                else:
                    print("⚠️ 搜索接口超时")
            else:
                print("⚠️ 未找到搜索框")

        except Exception as e:
            error_msg = f"RPA执行异常: {str(e)}"
            print(f"!!! {error_msg}")
            if match_result is None:
                match_result = {"status": "fail", "reason": error_msg}
            else:
                match_result['reason'] = f"{match_result.get('reason', '')} | {error_msg}"
        finally:
            try:
                if tab and tab.listen: tab.listen.stop()
                if tab:
                    # tab.close()
                    print(f"[{file_name}] 处理结束")
            except:
                pass

        return match_prompt, match_result, original_records, retry_count