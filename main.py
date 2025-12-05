#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DrissionPage 爬虫项目主入口 - 单据自动化处理版本
"""
import json
import re
import datetime
from datetime import timedelta
import os
import time
import glob
import asyncio
from concurrent.futures import ThreadPoolExecutor
from DrissionPage import Chromium, ChromiumOptions

# ---------------------------------------------------------
# 导入配置和工具
# 请确保当前目录下有 app_config.py 和 utils 文件夹
# ---------------------------------------------------------
try:
    from app_config import CONFIG
    from utils.data_manager import DataManager, load_style_db_with_cache
    from utils.report_generator import collect_result_data, update_html_report
    from utils.util_llm import extract_data_from_image, call_llm_text, call_gjllm_text, call_dmxllm_text, \
        extract_data_from_image_dmx
except ImportError as e:
    print(f"!!! 导入模块失败: {e}")
    print("请检查项目结构是否包含 app_config.py 和 utils/ 目录")
    exit(1)

# 从配置文件获取提示词
PROMPT_INSTRUCTION = CONFIG.get('prompt_instruction', '')

# 假设这是你的本地款号库（从数据库或文件加载），也可以在 async_main 中动态加载覆盖
LOCAL_STYLE_DB = {"T8821", "H2005", "X3002", "D5001"}

#  处理新增对账单的函数
def handle_new_reconciliation_bill(tab):

    print("\n>>> [阶段: 新增对账单处理] 开始...")

    try:
        # 1. 查找新的 iframe (通常是新增的，或者当前激活的)
        # 策略：遍历所有可见 iframe，查找含有“保存并审核”按钮的那个
        target_frame = None
        save_audit_btn = None

        print(">>> 正在查找“保存并审核”按钮所在的 iframe...")

        for _ in range(5):  # 重试几次，防止页面未完全渲染
            for frame in tab.eles('tag:iframe'):
                if not frame.states.is_displayed: continue

                # 查找按钮
                # HTML: <button ... data-amid="btnPaySaveAndAduit" onclick="saveRecord(1)" ...>
                # 策略A: data-amid
                btn = frame.ele('css:button[data-amid="btnPaySaveAndAduit"]', timeout=0.1)

                # 策略B: onclick
                if not btn:
                    btn = frame.ele('css:button[onclick="saveRecord(1)"]', timeout=0.1)

                # 策略C: 文本内容
                if not btn:
                    btn = frame.ele('x://button[contains(text(), "保存并审核")]', timeout=0.1)

                if btn and btn.states.is_displayed:
                    save_audit_btn = btn
                    target_frame = frame
                    break

            if save_audit_btn: break
            time.sleep(1)

        # 2. 点击按钮
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

# 跳转至“财务” -> “账单列表”
# 跳转至“财务” -> “账单列表”
def navigate_to_bill_list(tab, order_code):
    print("\n>>> [阶段: 跳转账单列表] 开始处理...")

    try:
        # 1. 定位并点击一级菜单“财务”
        print(">>> 正在定位“财务”菜单...")
        finance_btn = tab.ele('x://div[contains(@class, "title") and .//div[contains(text(), "财务")]]')

        if finance_btn:
            finance_btn.click()
            time.sleep(0.5)  # 等待折叠/展开动画
        else:
            print("!!! 错误: 未找到“财务”菜单")
            return

        # 2. 定位并点击二级菜单“账单列表”
        target_menu_text = "账单列表"
        menu_xpath = f'x://a[contains(text(), "{target_menu_text}")]'

        bill_menu = tab.wait.ele_displayed(menu_xpath, timeout=3)

        if bill_menu:
            bill_menu.click()
            print(f"✅ 成功点击左侧菜单“{target_menu_text}”")
            time.sleep(2)
            print(">>> 页面跳转等待完成")

        else:
            # 重试机制：尝试重新点击一级菜单
            print(f"⚠️ 未检测到二级菜单，尝试重新展开一级菜单...")
            if finance_btn:
                finance_btn.click()
                time.sleep(0.5)

            bill_menu = tab.wait.ele_displayed(menu_xpath, timeout=3)
            if bill_menu:
                bill_menu.click()
                print(f"✅ (重试) 成功点击左侧菜单“{target_menu_text}”")
                time.sleep(2)
            else:
                print(f"!!! 错误: 无法找到左侧菜单项\"{target_menu_text}\"")
                return

        # 3. 等待页面加载完成后开始搜索
        print(f">>> 正在查找搜索框 (data-grid='FMAccountsReceivableGrid')...")

        if not order_code:
            print("⚠️ 警告: 未获取到有效的订单编号，跳过搜索")
            return

        search_input_bill = None
        target_frame = None

        # 循环遍历 iframe 查找 (防止加载延迟)
        for _ in range(10):
            for frame in tab.eles('tag:iframe'):
                if not frame.states.is_displayed:
                    continue

                # 精确查找账单列表搜索框
                ele = frame.ele('css:input#txtSearchKey[data-grid="FMAccountsReceivableGrid"]', timeout=0.2)

                if ele and ele.states.is_displayed:
                    search_input_bill = ele
                    target_frame = frame
                    break

            if search_input_bill:
                break
            time.sleep(0.5)

        # 4. 输入订单编号并触发搜索
        if search_input_bill:
            print(f">>> 找到账单列表搜索框，正在输入: {order_code}")

            search_input_bill.click()
            time.sleep(0.2)
            search_input_bill.clear()

            # 逐字输入，实现"打字机"间歇效果
            for char in order_code:
                search_input_bill.input(char, clear=False)
                time.sleep(0.2)

            # 开启网络监听 (模糊匹配账单相关接口)
            tab.listen.start(targets='Admin/AccountsReceivable/NewGet')

            # 触发回车
            search_input_bill.run_js("""
                this.dispatchEvent(new Event('change', { bubbles: true }));
                this.dispatchEvent(new KeyboardEvent("keydown", {bubbles:true, keyCode:13, key:"Enter"}));
                this.dispatchEvent(new KeyboardEvent("keyup", {bubbles:true, keyCode:13, key:"Enter"}));
            """)
            print("✅ 输入完毕并触发回车")

            # 等待搜索响应
            res = None
            try:
                res = tab.listen.wait(timeout=10)
                if res:
                    print(f"✅ 账单列表搜索响应成功")
                else:
                    print("⚠️ 搜索超时，未收到响应")
            finally:
                # 【修复】使用 finally 确保 stop 只调用一次，防止 NoneType 错误
                tab.listen.stop()

            if res:
                print(">>> 正在遍历记录并勾选所有记录...")
                time.sleep(1)

                if target_frame:
                    rows = target_frame.eles('css:table#FMAccountsReceivableGrid tbody tr', timeout=2)

                    count_selected = 0

                    if not rows:
                        print("⚠️ 搜索结果为空，未找到任何行")
                    else:
                        for row in rows:
                            if not row.states.is_displayed: continue
                            try:
                                row.scroll.to_see()

                                # 查找复选框 (input.ckbox)
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

                            # --- 动作: 发起对账 ---
                            print(">>> 准备点击“发起对账”...")
                            try:
                                # 策略A: onclick 属性
                                btn_check = target_frame.ele('css:button[onclick="aReconciliation()"]', timeout=2)

                                # 策略B: 文本内容
                                if not btn_check:
                                    btn_check = target_frame.ele('x://button[contains(text(), "发起对账")]', timeout=1)

                                if btn_check:
                                    # 使用 JS 点击防止遮挡
                                    btn_check.run_js('this.click()')
                                    # 等待处理完成 (可能有 Layui 成功提示)
                                    time.sleep(2)
                                    print("✅ “发起对账”操作完成")
                                else:
                                    print("⚠️ 未找到“发起对账”按钮")

                                print(">>> 等待“新增对账单”页面加载...")
                                time.sleep(3)  # 给新页面一点加载时间

                                # 这里调用之前写好的独立函数
                                handle_new_reconciliation_bill(tab)

                            except Exception as e:
                                print(f"!!! 发起对账操作异常: {e}")
                        else:
                            print("⚠️ 未勾选任何记录，跳过“发起对账”")

                else:
                    print("!!! 错误: 丢失了 iframe 上下文")

        else:
            print("!!! 错误: 未找到账单列表搜索框")

    except Exception as e:
        print(f"!!! 跳转账单列表时发生异常: {e}")

# 跳转至“物料采购任务”并搜索订单
# 跳转至“物料采购任务”并搜索订单
def navigate_and_search_purchase_task(tab, order_code, parsed_data):
    print(f"\n>>> [阶段: 跳转物料采购任务] 开始处理，目标单号: {order_code}")

    if not order_code:
        print("⚠️ 错误: 未获取到有效的订单编号，无法执行搜索。")
        return

    # 提取需要填写的字段
    delivery_date = parsed_data.get('delivery_date', '')
    delivery_order_no = parsed_data.get('delivery_order_number', '')

    try:
        # 1. 定位并点击一级菜单“物料”
        print(">>> 正在重新定位“物料”菜单...")
        material_btn_nav = tab.ele('x://div[contains(@class, "title") and .//div[contains(text(), "物料")]]')
        if material_btn_nav:
            material_btn_nav.click()
            time.sleep(0.5)  # 等待动画

        # 2. 定位并点击二级菜单“物料采购任务”
        target_menu_text = "物料采购任务"
        menu_xpath = f'x://a[contains(text(), "{target_menu_text}")]'

        task_menu = tab.wait.ele_displayed(menu_xpath, timeout=3)

        if task_menu:
            task_menu.click()
            print(f"✅ 成功点击左侧菜单“{target_menu_text}”")
        else:
            # 重试机制
            print(f"⚠️ 未检测到菜单，尝试重新展开一级菜单...")
            if material_btn_nav:
                material_btn_nav.click()
                time.sleep(0.5)

            task_menu = tab.wait.ele_displayed(menu_xpath, timeout=3)
            if task_menu:
                task_menu.click()
                print(f"✅ (重试) 成功点击左侧菜单“{target_menu_text}”")
            else:
                print(f"!!! 错误: 无法找到左侧菜单项“{target_menu_text}”")
                return

        # 3. 等待页面加载 (iframe 载入)
        time.sleep(2)

        # 4. 在新打开的 iframe 中查找搜索框
        print(f">>> 正在查找搜索框 (data-grid='poMtPurTaskGrid')...")

        search_input_task = None
        target_frame = None  # 记录当前操作的iframe

        # 循环遍历 iframe 查找 (防止加载延迟)
        for _ in range(10):
            for frame in tab.eles('tag:iframe'):
                if not frame.states.is_displayed:
                    continue

                # 精确查找
                ele = frame.ele('css:input#txtSearchKey[data-grid="poMtPurTaskGrid"]', timeout=0.2)

                if ele and ele.states.is_displayed:
                    search_input_task = ele
                    target_frame = frame  # 锁定iframe
                    break

            if search_input_task:
                break
            time.sleep(0.5)

        # 5. 输入单号并触发搜索
        if search_input_task:
            print(f">>> 找到搜索框，正在输入: {order_code}")

            search_input_task.click()
            time.sleep(0.2)
            search_input_task.clear()

            # 逐字输入，实现"打字机"间歇效果
            for char in order_code:
                search_input_task.input(char, clear=False)
                time.sleep(0.2)  # 这里的 0.2 就是间歇时间，可自己调

            # 开启网络监听 (模糊匹配 PurchaseTask 相关接口)
            tab.listen.start(targets='Admin/MtPurchase')

            # 触发回车
            search_input_task.run_js("""
                this.dispatchEvent(new Event('change', { bubbles: true }));
                this.dispatchEvent(new KeyboardEvent("keydown", {bubbles:true, keyCode:13, key:"Enter"}));
                this.dispatchEvent(new KeyboardEvent("keyup", {bubbles:true, keyCode:13, key:"Enter"}));
            """)
            print("✅ 输入完毕并触发回车")

            # 等待响应
            try:
                res = tab.listen.wait(timeout=10)
            finally:
                tab.listen.stop()  # 确保停止监听

            if res:
                print(f"✅ 搜索响应成功")

                # ==========================================
                # [新增功能] 遍历结果行并勾选所有记录
                # ==========================================
                print(">>> 正在遍历记录并勾选所有记录...")
                time.sleep(1)  # 等待表格DOM渲染

                if target_frame:
                    # 定位表格行
                    rows = target_frame.eles('css:table#poMtPurTaskGrid tbody tr', timeout=2)

                    count_selected = 0
                    for row in rows:
                        if not row.states.is_displayed: continue

                        try:
                            # 确保行可见
                            row.scroll.to_see()

                            # 勾选复选框
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

                            time.sleep(0.1)  # 稍微防抖

                        except Exception as inner_e:
                            print(f"   !!! 勾选行时出错: {inner_e}")

                    print(f"✅ 记录勾选完成，共勾选 {count_selected} 行")

                    print(">>> [1/3] 准备点击“一键绑定加工单”...")
                    try:
                        btn_bind = target_frame.ele('#btnOneKeyBindPM', timeout=2)
                        if btn_bind:
                            btn_bind.click()

                            # [处理弹窗] 绑定操作通常会弹窗确认
                            try:
                                if tab.wait.alert(timeout=3):
                                    print(f"   ℹ️ 绑定确认弹窗: {tab.alert.text} -> 自动接受")
                                    tab.alert.accept()
                            except:
                                pass

                            # 处理layui弹窗
                            try:
                                confirm_btn = tab.ele('css:a.layui-layer-btn0', timeout=3)
                                if confirm_btn:
                                    confirm_btn.click()
                            except:
                                pass

                            # 等待绑定完成 (可能还有成功提示)
                            time.sleep(1)
                            try:
                                if tab.wait.alert(timeout=2):
                                    tab.alert.accept()
                            except:
                                pass

                            print("   ✅ 一键绑定操作结束")

                            # 等待系统处理完成（检查所有行的加工厂字段都有值）
                            print(">>> 等待系统处理一键绑定，检查所有行的加工厂字段...")
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
                                    # 检查所有行的加工厂字段
                                    factory_cells = target_frame.eles('css:td[masking="SpName"]', timeout=1)
                                    completed_count = 0

                                    for cell in factory_cells:
                                        if cell.states.is_displayed:
                                            cell_text = cell.text.strip()
                                            if cell_text:  # 有内容的算完成
                                                completed_count += 1

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
                                time.sleep(2)  # 兜底等待

                            # 一键绑定完成后，重新填写码单信息
                            print(">>> 一键绑定完成，开始填写码单信息...")
                            try:
                                # 重新获取表格行（页面可能已刷新）
                                rows = target_frame.eles('css:table#poMtPurTaskGrid tbody tr', timeout=3)
                                count_filled = 0

                                for row in rows:
                                    if not row.states.is_displayed: continue

                                    try:
                                        # 确保行可见
                                        row.scroll.to_see()

                                        # 填写 码单编号 (Att01)
                                        if delivery_order_no:
                                            inp_no = row.ele('css:input.Att01', timeout=0.2)
                                            if inp_no:
                                                # 普通文本框直接赋值并触发事件
                                                js_no = f"""
                                                        this.value = "{delivery_order_no}";
                                                        this.dispatchEvent(new Event("input"));
                                                        this.dispatchEvent(new Event("change"));
                                                        this.dispatchEvent(new Event("blur"));
                                                    """
                                                inp_no.run_js(js_no)
                                                print(f"   -> 已填入码单编号: {delivery_order_no}")
                                            else:
                                                print("   ⚠️ 未找到 Att01 (编号) 输入框")

                                        # 填写 码单日期 (Att02)
                                        if delivery_date:
                                            inp_date = row.ele('css:input.Att02', timeout=0.2)
                                            if inp_date:
                                                try:
                                                    print(f"   -> 正在填入码单日期: {delivery_date}")

                                                    # 1. 【核心】完全模仿 Selenium：JS 移除 readonly 属性
                                                    inp_date.run_js('this.removeAttribute("readonly");')

                                                    # 2. 清空输入框
                                                    inp_date.clear()
                                                    time.sleep(0.1)

                                                    # 3. 模拟键盘输入
                                                    inp_date.input(delivery_date)
                                                    time.sleep(0.2)  # 等待输入反应

                                                    # 4. 模拟按下"回车"键
                                                    target_frame.actions.key_down('ENTER').key_up('ENTER')
                                                    time.sleep(0.2)

                                                    # 5. 点击空白处失焦
                                                    target_frame.run_js('document.body.click();')
                                                    inp_date.click()  # 再次点击输入框确保焦点在上面
                                                    time.sleep(0.2)
                                                except Exception as e:
                                                    print(f"   ⚠️ 日期输入异常: {e}")
                                                    # 兜底策略
                                                    try:
                                                        inp_date.run_js(
                                                            f'this.removeAttribute("readonly"); this.value="{delivery_date}";')
                                                    except:
                                                        pass
                                            else:
                                                print("   ⚠️ 未找到 Att02 (日期) 输入框")

                                        count_filled += 1
                                        time.sleep(0.1)  # 防抖

                                    except Exception as inner_e:
                                        print(f"   !!! 填写行数据时出错: {inner_e}")

                                print(f"✅ 码单信息填写完成，共处理 {count_filled} 行")

                            except Exception as e:
                                print(f"!!! 填写码单信息时发生异常: {e}")

                        else:
                            print("⚠️ 未找到一键绑定加工单按钮")
                    except Exception as e:
                        print(f"!!! 绑定操作异常: {e}")

                    # --- 动作 B: [2/3] 提交 ---
                    print(">>> [2/3] 准备点击“提交”...")
                    try:
                        btn_submit = target_frame.ele('#btnSubmitTasks', timeout=2)
                        if btn_submit:
                            btn_submit.click()

                            # [处理弹窗] 提交确认
                            try:
                                if tab.wait.alert(timeout=3):
                                    print(f"   ℹ️ 提交确认弹窗: {tab.alert.text} -> 自动接受")
                                    tab.alert.accept()
                            except:
                                pass

                            # 提交后的提示
                            time.sleep(1)
                            try:
                                if tab.wait.alert(timeout=2):
                                    tab.alert.accept()
                            except:
                                pass

                            print("   ✅ “提交”操作结束")
                            time.sleep(2)
                        else:
                            print("⚠️ 未找到“提交”按钮")
                    except Exception as e:
                        print(f"!!! 提交操作异常: {e}")

                    # --- 动作 C: [3/3] 确认 ---
                    print(">>> [3/3] 准备点击“确认”...")
                    try:
                        btn_confirm = target_frame.ele('#btnConfirmToDoTask', timeout=2)
                        if btn_confirm:
                            btn_confirm.click()

                            # [处理弹窗 1] 浏览器原生 Alert
                            try:
                                if tab.wait.alert(timeout=3):
                                    print(f"   ℹ️ 确认操作弹窗: {tab.alert.text} -> 自动接受")
                                    tab.alert.accept()
                            except:
                                pass

                            print("   -> 等待系统处理确认逻辑...")
                            time.sleep(2)

                            # [处理弹窗 2] Layui 成功提示弹窗
                            print("   -> 检查“成功确认”Layui弹窗...")
                            try:
                                lay_confirm = tab.ele('css:a.layui-layer-btn0', timeout=3)
                                if not lay_confirm:
                                    lay_confirm = target_frame.ele('css:a.layui-layer-btn0', timeout=1)

                                if lay_confirm:
                                    lay_confirm.click()
                                    print("   ✅ 检测到Layui成功弹窗，已点击“确定”关闭")
                                else:
                                    print("   ℹ️ 未检测到Layui成功弹窗 (可能已自动关闭或无提示)")

                            except Exception as e:
                                print(f"   ⚠️ 处理Layui弹窗时出错 (非阻断): {e}")

                            print("✅ “确认”操作全部完成")
                        else:
                            print("⚠️ 未找到“确认”按钮")
                    except Exception as e:
                        print(f"!!! 确认操作异常: {e}")

                    # --- 动作 D: [新增] 执行采购任务 (修复：使用JS点击) ---
                    print(">>> 准备执行采购任务...")
                    time.sleep(1)

                    try:
                        # 定位并点击“更多”按钮
                        more_btn = target_frame.ele(
                            'x://button[contains(text(), "更多")]', timeout=2)

                        if more_btn:
                            # 【修复】使用 JS 点击，避免下拉动画期间点击失败
                            more_btn.run_js('this.click()')
                            time.sleep(0.5)  # 等待下拉菜单展开动画

                            # 定位并点击“执行采购任务”
                            exec_task_btn = target_frame.ele(
                                'css:a[onclick="doMtPurTask()"]', timeout=1)

                            if not exec_task_btn:
                                exec_task_btn = target_frame.ele(
                                    'x://a[contains(text(), "执行采购任务")]',
                                    timeout=1)

                            if exec_task_btn:
                                print("   -> 找到“执行采购任务”按钮，正在点击...")
                                # 【修复】使用 JS 点击，忽略位置大小计算
                                exec_task_btn.run_js('this.click()')

                                # 处理可能出现的 Alert/Confirm 弹窗
                                try:
                                    if tab.wait.alert(timeout=3):
                                        alert_text = tab.alert.text
                                        print(f"   ℹ️ 检测到系统弹窗: [{alert_text}]，自动接受...")
                                        tab.alert.accept()
                                except:
                                    pass

                                print("   ✅ 成功点击“执行采购任务”")
                                time.sleep(2)  # 等待系统处理

                                print(">>> 等待并处理结果弹窗...")
                                try:
                                    confirm_btn = tab.ele('css:a.layui-layer-btn0', timeout=3)
                                    if not confirm_btn:
                                        confirm_btn = target_frame.ele('css:a.layui-layer-btn0', timeout=2)

                                    if confirm_btn:
                                        confirm_btn.click()
                                        print("   ✅ 成功点击弹窗“确定”按钮")
                                        time.sleep(1)
                                    else:
                                        print("   ⚠️ 未检测到 Layui 结果弹窗 (可能已自动关闭或无需确认)")

                                except Exception as e:
                                    print(f"   !!! 处理结果弹窗时发生异常: {e}")
                            else:
                                print("   ⚠️ 展开了“更多”菜单，但未找到“执行采购任务”选项")
                                more_btn.run_js('this.click()')  # 尝试关闭
                        else:
                            print("   ⚠️ 未找到“更多”按钮")

                    except Exception as e:
                        print(f"!!! 执行采购任务操作异常: {e}")

                else:
                    print("!!! 错误: 丢失了 iframe 上下文")

            else:
                print("⚠️ 搜索超时，未收到响应")

        else:
            print("!!! 错误: 未找到搜索框")
    except Exception as e:
        print(f"!!! 跳转或搜索'物料采购任务'时发生异常: {e}")

def should_use_dmx_for_date_check(delivery_date_str):
    """
    检查交付日期是否异常，如果与当前日期相差超过阈值天数则返回True
    :param delivery_date_str: 交付日期字符串，格式如 "2025-12-05"
    :return: bool, True表示需要用DMX重新识别
    """
    try:
        if not delivery_date_str:
            return False

        # 解析交付日期
        if len(delivery_date_str) == 10 and delivery_date_str.count('-') == 2:
            # YYYY-MM-DD 格式
            delivery_date = datetime.datetime.strptime(delivery_date_str, '%Y-%m-%d').date()
        else:
            # 尝试其他常见格式
            for fmt in ['%Y/%m/%d', '%m/%d/%Y', '%d/%m/%Y', '%Y.%m.%d']:
                try:
                    delivery_date = datetime.datetime.strptime(delivery_date_str, fmt).date()
                    break
                except ValueError:
                    continue
            else:
                # 所有格式都失败
                print(f">>> 无法解析交付日期格式: {delivery_date_str}")
                return False

        # 获取当前日期
        current_date = datetime.datetime.now().date()

        # 计算日期差值（绝对值）
        date_diff = abs((delivery_date - current_date).days)

        # 获取阈值
        threshold_days = CONFIG.get('delivery_date_threshold_days', 7)

        if date_diff > threshold_days:
            print(
                f">>> 交付日期异常检测: 当前日期 {current_date}, 交付日期 {delivery_date}, 差值 {date_diff} 天 > 阈值 {threshold_days} 天")
            return True
        else:
            print(f">>> 交付日期正常: 当前日期 {current_date}, 交付日期 {delivery_date}, 差值 {date_diff} 天")
            return False

    except Exception as e:
        print(f">>> 交付日期解析失败: {delivery_date_str}, 错误: {e}")
        return False


def select_matched_checkboxes(tab, matched_ids):
    """根据匹配的记录ID勾选对应行的checkbox"""
    print(f">>> 开始勾选匹配的记录: {len(matched_ids)} 条")

    for record_id in matched_ids:
        try:
            checkbox_selector = f'x://tr[.//a[contains(@data-sub-html, "{record_id}")]]//input[contains(@class, "ckbox")]'

            # 在所有iframe中查找checkbox
            checkbox_found = False
            for frame in tab.eles('tag:iframe'):
                if not frame.states.is_displayed:
                    continue

                # 使用新的 XPath selector 查找
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
                # 调试建议：如果还是找不到，打印一下所有可见的 data-sub-html 看看格式是否一致

        except Exception as e:
            print(f"!!! 勾选记录 {record_id} 时发生异常: {e}")

    print(f">>> 勾选操作完成")


# 智能款号清理函数
def smart_clean_with_db(text, style_db):
    """基于款号库的智能款号清理"""
    if not text:
        return text

    # 1. 直接匹配（原文就在库中）
    if text in style_db:
        return text

    # 2. 检测重复款号（如H1591A-AH1591A-B）
    # 先检查是否包含库中的款号
    for db_style in style_db:
        if db_style in text and len(db_style) > 3:  # 避免匹配过短的款号
            # 计算该款号在文本中出现的次数
            count = text.count(db_style)
            if count > 1:
                # 重复出现，返回单个款号
                print(f">>> 检测到重复款号，提取: {db_style}")
                return db_style
            elif count == 1:
                # 只出现一次，可能是正确的
                return db_style

    # 3. OCR字符纠错：处理常见的OCR识别错误
    ocr_char_map = {
        'J': '1',  # HJ643C → H1643C
        'O': '0',  # HO123A → H0123A
        'I': '1',  # HI456B → H1456B
        'S': '5',  # H5789C → H5789C (S→5)
        'Z': '2'  # HZ321A → H2321A
    }

    for wrong_char, correct_char in ocr_char_map.items():
        if wrong_char in text:
            corrected_text = text.replace(wrong_char, correct_char)
            if corrected_text in style_db:
                print(f">>> OCR纠错成功: {text} → {corrected_text}")
                return corrected_text

    # 4. 尝试多种清理策略
    # 先处理常见的末尾字符
    text_clean_hash = text.rstrip('#') if text.endswith('#') else text
    text_clean_kuan = text.rstrip('款') if text.endswith('款') else text

    cleaning_strategies = [
        text_clean_hash,  # 去除#结尾
        text_clean_kuan,  # 去除款字结尾
        text.rstrip('款型式号'),  # 移除常见后缀
        text.split('款')[0].strip(),  # 取"款"字前的部分
        text.split('型')[0].strip(),  # 取"型"字前的部分
        text.split('式')[0].strip(),  # 取"式"字前的部分
        text.replace(' ', ''),  # 去除空格
        text.upper(),  # 转大写
        text.upper().rstrip('款型式号'),  # 大写+去后缀
        text.upper().rstrip('#'),  # 大写+去除#
        text.upper().rstrip('款'),  # 大写+去除款字
    ]

    # 5. 逐一尝试，找到第一个匹配的
    for candidate in cleaning_strategies:
        if candidate and candidate in style_db:
            return candidate

    # 6. 都不匹配，返回最干净的版本
    return text.rstrip('款型式号').strip()


# 根据规则确定最终款号
def determine_final_style(json_data, style_db):
    """
    根据业务规则从 candidates 中选出最终款号
    :param json_data: LLM 返回的 JSON
    :param style_db: 本地款号库 (Set集合)
    """
    candidates = json_data.get("style_candidates", [])
    if not candidates:
        return None

    # --- 规则 1: 识别明细中存在的款号 (位置在表格内且在库中) ---
    for item in candidates:
        clean_text = item['text'].strip()
        # 这里直接用 style_db 进行 O(1) 复杂度的极速查找
        if "表格" in item.get('position', '') and clean_text in style_db:
            cleaned = smart_clean_with_db(clean_text, style_db)
            print(f"✅ 命中规则1 (表格内且在库): {cleaned}")
            return cleaned

    # --- 规则 2: 识别红色字体 (若存在且在库中) ---
    red_candidates = [c for c in candidates if c.get('is_red') == True]
    for item in red_candidates:
        clean_text = item['text'].strip()
        if clean_text in style_db:
            cleaned = smart_clean_with_db(clean_text, style_db)
            print(f"✅ 命中规则2 (红色字体且在库): {cleaned}")
            return cleaned

    # --- 规则 3: 兜底搜索 (符合 T/H/X/D 开头规律) ---
    for item in candidates:
        text = item['text'].strip().upper()
        # 这里逻辑可根据需求调整：如果不在库里，但长得像，要不要？
        # 下面代码假设：只要长得像就提取，作为最后的保底
        if text.startswith(('T', 'H', 'X', 'D', 'L', 'S', 'F')):  # 我看你csv里还有L/S/F开头
            cleaned = smart_clean_with_db(text, style_db)
            print(f"✅ 命中规则3 (符合命名规律): {cleaned}")
            return cleaned

    return None


# 完整RPA流程方法
def process_single_bill_rpa(browser, data_json, file_name, img_path):
    print(f"\n--- [RPA阶段] 开始处理: {file_name} ---")

    match_prompt = ""
    match_result = None
    original_records = []
    retry_count = 1
    tab = None

    try:
        # 1. 创建新页签
        print(f"[{file_name}] 正在创建新页签...")
        tab = browser.new_tab()

        # 2. 激活浏览器窗口（根据配置决定是否置前）
        if CONFIG.get('rpa_browser_to_front', True):
            print(f"[{file_name}] 正在激活浏览器窗口...")
            tab.set.window.normal()
            time.sleep(0.5)
            tab.set.window.full()
        else:
            print(f"[{file_name}] 跳过浏览器窗口激活")

        tab.get(CONFIG['base_url'])

        # 如果配置为不置前，则将浏览器最小化
        if not CONFIG.get('rpa_browser_to_front', True):
            print(f"[{file_name}] 将浏览器窗口最小化")
            tab.set.window.mini()

        # 3. 开始具体的业务操作 (和之前一样)
        # 2. 确保左侧菜单栏已经加载出来
        # 使用 provided HTML 中的顶层 class "fixed-left-menu" 作为定位基准
        if not tab.wait.ele_displayed('.fixed-left-menu', timeout=5):
            print("!!! 错误: 未检测到左侧菜单栏，请确认网页已加载完成。")
            return

        print(">>> 正在定位“物料”菜单...")

        # 定位“物料”菜单按钮
        # 查找 class为title 且 包含 '物料' 文字的 div
        material_btn = tab.ele('x://div[contains(@class, "title") and .//div[contains(text(), "物料")]]')

        if material_btn:
            print(">>> 找到“物料”菜单，正在点击...")
            material_btn.click()
        else:
            print("!!! 未找到物料菜单")

        print("✅ 成功点击“物料”菜单")

        # 等待子菜单展开 (给自己留点缓冲时间，0.5~1秒通常够了)
        time.sleep(0.5)

        print(">>> 正在定位二级菜单“物料采购需求”...")

        # 定位“物料采购需求”子菜单按钮
        # sub_menu_btn = tab.wait.ele_displayed('a@@text:物料采购需求', timeout=3)
        sub_menu_btn = tab.wait.ele_displayed('x://a[contains(text(), "物料采购需求")]', timeout=3)

        if sub_menu_btn:
            sub_menu_btn.click()
            print("✅ 成功点击“物料采购需求”")
        else:
            # 如果等了3秒还没显示，可能是以及菜单没点开，尝试补救（再次点击一级菜单）
            print("⚠️ 未检测到二级菜单展开，尝试重新点击“物料”...")
            material_btn.click()
            time.sleep(1)
            sub_menu_btn = tab.wait.ele_displayed('x://a[contains(text(), "物料采购需求")]', timeout=3)
            if sub_menu_btn:
                sub_menu_btn.click()
                print("✅ 重试后成功点击")
            else:
                print("!!! 错误：无法展开二级菜单，请检查页面遮挡或网络卡顿。")
                return

        print(">>> 正在 iframe 中查找可见的搜索框...")
        search_input = None

        # 循环尝试 10 次 (共等待 5 秒)，防止 iframe 还没加载完
        for _ in range(10):
            # 遍历所有 iframe
            for frame in tab.eles('tag:iframe'):

                # --- 优化点1: iframe 自己都不显示，就别进去找了，浪费时间 ---
                if not frame.states.is_displayed:
                    continue

                # --- 优化点2: timeout=0.2 (关键) ---
                # 进去找元素时，只给 0.2 秒。找不到就说明不在这个 iframe 里，立刻换下一个
                ele = frame.ele('#txtSearchKey', timeout=0.2)

                try:
                    if ele and ele.states.is_displayed:
                        search_input = ele
                        break
                except:
                    pass

            if search_input:
                break

            # 没找到就稍微歇一下再重试
            time.sleep(0.5)

        if search_input:
            input_value = data_json.get('final_selected_style', '')

            # 检查款号是否为空或None
            if not input_value:
                print(f"⚠️ 警告: 款号为空或None，跳过RPA处理")
                return "", None, []

            print(f">>> 开始输入款号: {input_value}")

            search_input.click()
            time.sleep(0.2)
            # 先清空输入框
            search_input.clear()

            # 逐字输入，实现"打字机"间歇效果
            for char in input_value:
                search_input.input(char, clear=False)
                time.sleep(0.2)  # 这里的 0.2 就是间歇时间，可自己调

            # 设置网络监听，捕获请求后的返回数据
            target_url_substring = 'Admin/MtReq/NewGet'
            tab.listen.start(targets=target_url_substring)
            print(f">>> 已开启网络监听，目标: {target_url_substring}")

            # 输入完毕后，触发回车事件
            search_input.run_js("""
                            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                            arguments[0].dispatchEvent(new KeyboardEvent("keydown", {bubbles:true, keyCode:13, key:"Enter"}));
                            arguments[0].dispatchEvent(new KeyboardEvent("keyup", {bubbles:true, keyCode:13, key:"Enter"}));
                        """, search_input)

            print("✅ 输入完毕并回车")

            res_packet = tab.listen.wait(timeout=10)

            # 处理捕获到的响应数据
            if res_packet:
                print(f"✅ 成功捕获接口数据: {res_packet.url}")

                # 获取响应内容 (自动解析 JSON)
                response_body = res_packet.response.body

                # --- 在这里处理你拿到的 JSON 数据 ---
                if isinstance(response_body, dict):
                    records = response_body.get('data', [])
                    print(f"数据统计: 共找到 {len(records)} 条记录")

                    if not records:
                        print("⚠️ 警告: 搜索结果为空，无需匹配。")
                    else:
                        # 保存原始records用于报告生成
                        original_records = records

                        # 4. 调用 LLM 进行智能匹配
                        print(">>> 正在调用 LLM 进行智能匹配...")
                        match_result, match_prompt, retry_count = execute_smart_match(data_json, records)

                        print("\n" + "=" * 30)
                        print(f"🤖 智能匹配结果: {match_result.get('status', 'FAIL').upper()}")
                        print(f"匹配原因: {match_result.get('global_reason', match_result.get('reason'))}")
                        print("=" * 30 + "\n")

                        # 5. 基于匹配结果的后续RPA操作
                        # structured_tasks = []
                        matched_ids = []

                        if match_result.get('status') == 'success':
                            # A. 组装数据结构体
                            structured_tasks = reconstruct_rpa_data(match_result, data_json, original_records)
                            print(f">>> 数据组装完成，共生成 {len(structured_tasks)} 个任务包")

                            # B. 提取所有需要勾选的 Record ID (去重，防止 Split 模式下重复勾选)
                            seen_ids = set()
                            for task in structured_tasks:
                                rec_id = task['record'].get('Id')
                                if rec_id and rec_id not in seen_ids:
                                    matched_ids.append(rec_id)
                                    seen_ids.add(rec_id)

                        # 6. 执行 RPA 动作 (勾选 + 点击)
                        if matched_ids:
                            # 步骤 A: 勾选匹配的记录
                            select_matched_checkboxes(tab, matched_ids)

                            # 步骤 B: 点击“物料采购单”按钮 (使用全域搜索)
                            print(">>> 正在查找并点击“物料采购单”生成按钮...")
                            button_found = False

                            # 搜索范围：主界面 + 所有可见iframe
                            scopes = [tab] + [f for f in tab.eles('tag:iframe') if f.states.is_displayed]

                            for scope in scopes:
                                # 直接使用验证过的 XPath 文本定位
                                btn = scope.ele('x://button[contains(text(), "物料采购单")]', timeout=0.5)

                                if btn and btn.states.is_displayed:
                                    btn.scroll.to_see()
                                    time.sleep(0.5)  # 等待滚动稳定
                                    btn.click()
                                    print("✅ 成功点击“物料采购单”按钮")
                                    button_found = True
                                    time.sleep(2)  # 等待页面响应
                                    break

                            if not button_found:
                                print("⚠️ 未找到“物料采购单”按钮")
                            else:
                                # 采购类型为“月结采购”
                                print(">>> 正在等待页面加载并切换为“月结采购”...")
                                time.sleep(2)  # 等待新页面/弹窗渲染

                                type_selected = False

                                # 页面变动了，重新获取当前所有可见的 iframe 和 tab
                                # 注意：新页面可能加载在新的 iframe 中
                                current_scopes = [tab] + [f for f in tab.eles('tag:iframe') if f.states.is_displayed]

                                for scope in current_scopes:
                                    try:
                                        # 1. 定位下拉框按钮
                                        dropdown_btn = scope.ele('css:button[data-id="OrderTypeId"]', timeout=0.5)

                                        if dropdown_btn and dropdown_btn.states.is_displayed:
                                            print("   -> 找到采购类型下拉框")
                                            dropdown_btn.scroll.to_see()
                                            dropdown_btn.click()
                                            time.sleep(0.5)  # 等待下拉菜单展开动画

                                            # 2. 定位“月结采购”选项
                                            option = scope.ele('x://span[@class="text" and text()="月结采购"]',
                                                               timeout=1)

                                            if option and option.states.is_displayed:
                                                option.click()
                                                print("✅ 成功选择“月结采购”")
                                                type_selected = True
                                                time.sleep(1)  # 等待选择生效
                                                break
                                            else:
                                                print("   ⚠️ 展开了下拉框，但在当前Scope未找到'月结采购'选项")
                                    except Exception as e:
                                        # 忽略非目标iframe的查找错误
                                        continue

                                if not type_selected:
                                    print("⚠️ 未能完成“月结采购”选择 (可能已默认选中或元素定位失败)")

                                # 根据款号选择品牌
                                style_code = data_json.get('final_selected_style', '').strip().upper()  # 获取款号，转大写并去空格
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
                                            # A. 定位品牌下拉框按钮
                                            brand_btn = scope.ele('css:button[data-id="BrandId"]', timeout=0.3)

                                            if brand_btn and brand_btn.states.is_displayed:
                                                brand_btn.scroll.to_see()
                                                brand_btn.click()
                                                time.sleep(0.5)  # 等待下拉展开

                                                open_menu = scope.ele('css:div.btn-group.open', timeout=1)

                                                if open_menu:
                                                    # 查找选项 XPath: //span[contains(@class, 'text') and contains(text(), '目标')]
                                                    brand_opt = open_menu.ele(
                                                        f'x:.//span[contains(@class, "text") and contains(text(), "{target_brand}")]',
                                                        timeout=1)

                                                    if brand_opt:
                                                        # 关键：先滚动到该元素（处理下拉框滚动条）
                                                        brand_opt.scroll.to_see()
                                                        time.sleep(0.1)
                                                        brand_opt.click()
                                                        print(f"✅ 成功选择品牌: {target_brand}")
                                                        brand_selected = True
                                                        time.sleep(0.5)
                                                        break
                                                    else:
                                                        print(f"   ⚠️ 下拉框已打开，但在列表中未找到 [{target_brand}]")
                                                        brand_btn.click()  # 关闭下拉框
                                                else:
                                                    print("   ⚠️ 点击了按钮，但未检测到下拉菜单展开")

                                        except Exception as e:
                                            print(f"debug: {e}")
                                            continue
                                    if not brand_selected:
                                        print(f"⚠️ 品牌选择失败: 未找到下拉框或选项 [{target_brand}]")
                                else:
                                    print(f"ℹ️ 款号[{style_code}]无特殊前缀规则，跳过品牌选择")

                                # 填写码单日期
                                ocr_date = data_json.get('delivery_date', '')
                                if ocr_date:
                                    print(f">>> 正在查找并填写码单日期: {ocr_date} ...")
                                    att01_filled = False

                                    # 在当前所有可见区域(iframe/tab)中查找
                                    for scope in current_scopes:
                                        try:
                                            # 定位输入框
                                            att01_input = scope.ele('#Att01', timeout=0.5)

                                            if att01_input and att01_input.states.is_displayed:
                                                att01_input.scroll.to_see()
                                                att01_input.clear()  # 先清空
                                                att01_input.input(ocr_date)  # 输入

                                                # 触发 change 事件以确保数据被表单记录 (防止假输入)
                                                att01_input.run_js(
                                                    'this.dispatchEvent(new Event("change", {bubbles: true})); this.dispatchEvent(new Event("blur"));')

                                                print("✅ 成功填写码单日期")
                                                att01_filled = True
                                                break
                                        except Exception as e:
                                            # 忽略当前 scope 找不到的错误，继续找下一个
                                            continue

                                    if not att01_filled:
                                        print("⚠️ 未找到码单日期输入框 (#Att01)")
                                else:
                                    print("ℹ️ OCR数据中无码单日期，跳过填写")

                                fill_details_into_table(scope, structured_tasks)

                                print(">>> 表格填写完毕，正在查找并点击“保存并审核”按钮...")
                                try:
                                    # 1. 定位按钮
                                    # 策略A (推荐): 使用 data-amid 属性，精准定位
                                    save_btn = scope.ele('css:button[data-amid="btnSaveAndAudit"]', timeout=1)

                                    # 策略B (兜底): 如果属性找不到，尝试用文字内容
                                    if not save_btn:
                                        save_btn = scope.ele('x://button[contains(text(), "保存并审核")]', timeout=1)

                                    if save_btn and save_btn.states.is_displayed:
                                        # 2. 确保按钮可见并点击
                                        save_btn.scroll.to_see()
                                        time.sleep(0.5)  # 稍微停顿，模拟人工确认
                                        save_btn.click()
                                        print("✅ 成功点击“保存并审核”")

                                        # 3. 处理可能出现的 Alert/Confirm 弹窗
                                        # (有些ERP系统点击保存后会弹窗询问"确定保存吗？")
                                        try:
                                            # 等待弹窗出现 (最多等 2 秒)
                                            if tab.wait.alert(timeout=2):
                                                alert_text = tab.alert.text
                                                print(f"ℹ️ 检测到系统弹窗: [{alert_text}]，正在自动接受...")
                                                tab.alert.accept()
                                        except Exception:
                                            # 如果没有弹窗，pass 即可
                                            pass

                                        # 4. 等待保存完成 (防止立即关闭页面导致保存失败)
                                        print(">>> 等待保存结果...")
                                        time.sleep(3)

                                        print(">>> 正在获取生成的订单编号...")
                                        try:
                                            # 定位存放单号的 Input (ID="Code")
                                            # 直接使用当前的 scope (即找到保存按钮的那个 iframe)
                                            code_input = scope.ele('#Code', timeout=2)

                                            if code_input:
                                                # 尝试多种方式获取值
                                                # 1. 优先获取标准 value 属性 (DOM property)
                                                order_code = code_input.value

                                                # 2. 如果为空，尝试获取 HTML 属性 'valuecontent' (根据你提供的 HTML)
                                                if not order_code:
                                                    order_code = code_input.attr('valuecontent')

                                                # 3. 如果还为空，尝试获取 HTML 属性 'value'
                                                if not order_code:
                                                    order_code = code_input.attr('value')

                                                if order_code:
                                                    print(f"✅ 成功获取订单编号: [{order_code}]")
                                                    # 存入上下文 (data_json 即外部传入的 parsed_data)
                                                    # 这样在生成报告时可以直接读取到这个字段
                                                    data_json['rpa_order_code'] = order_code
                                                else:
                                                    print("⚠️ 找到了 #Code 输入框，但无法提取到编号值")
                                            else:
                                                print("⚠️ 未找到 ID 为 #Code 的输入框")

                                        except Exception as e:
                                            print(f"!!! 获取订单编号时发生异常: {e}")
                                        print(">>> 准备跳转至“物料采购订单”列表...")
                                        time.sleep(0.5)  # 稍作停顿，等待上一操作完全结束

                                        # 打开“物料采购订单”页面
                                        try:
                                            # 1. 重新定位并点击一级菜单“物料”
                                            # (虽然通常是展开的，但为了保险，再次确认点击或确保焦点)
                                            material_btn_nav = tab.ele(
                                                'x://div[contains(@class, "title") and .//div[contains(text(), "物料")]]')
                                            if material_btn_nav:
                                                material_btn_nav.click()
                                                time.sleep(0.5)  # 等待折叠/展开动画

                                            # 2. 定位并点击二级菜单“物料采购订单”
                                            # HTML: <a ...>物料采购订单</a>
                                            target_menu_text = "物料采购订单"
                                            menu_xpath = f'x://a[contains(text(), "{target_menu_text}")]'

                                            # 等待菜单出现 (timeout=3)
                                            purchase_order_menu = tab.wait.ele_displayed(menu_xpath, timeout=3)

                                            if purchase_order_menu:
                                                purchase_order_menu.click()
                                                print(f"✅ 成功点击左侧菜单“{target_menu_text}”")
                                            else:
                                                # 重试机制：如果没出来，可能是刚才点击“物料”把它折叠了，或者页面刷新了
                                                print(f"⚠️ 未检测到二级菜单，尝试重新展开一级菜单...")
                                                if material_btn_nav:
                                                    material_btn_nav.click()  # 再点一次
                                                    time.sleep(0.5)

                                                # 再次尝试查找
                                                purchase_order_menu = tab.wait.ele_displayed(menu_xpath, timeout=3)
                                                if purchase_order_menu:
                                                    purchase_order_menu.click()
                                                    print(f"✅ 重试后成功点击“{target_menu_text}”")
                                                else:
                                                    print(f"!!! 错误: 无法找到左侧菜单项“{target_menu_text}”")

                                            # 3. 等待页面跳转/加载
                                            # (点击菜单后通常右侧 iframe 会刷新，建议等待几秒)
                                            time.sleep(2)

                                            # 4. 在“物料采购订单”页面搜索刚才生成的订单编号
                                            order_code = data_json.get('rpa_order_code')
                                            if order_code:
                                                print(f">>> 准备在“物料采购订单”列表搜索单号: {order_code}")

                                                search_input_order = None

                                                # 循环尝试查找搜索框 (防止 iframe 加载延迟)
                                                for _ in range(10):
                                                    # 遍历所有可见 iframe
                                                    for frame in tab.eles('tag:iframe'):
                                                        if not frame.states.is_displayed:
                                                            continue

                                                        # 使用特有的 data-grid 属性精确定位
                                                        # HTML: <input id="txtSearchKey" data-grid="POMtPurchaseGrid" ...>
                                                        ele = frame.ele(
                                                            'css:input#txtSearchKey[data-grid="POMtPurchaseGrid"]',
                                                            timeout=0.2)

                                                        if ele and ele.states.is_displayed:
                                                            search_input_order = ele
                                                            break

                                                    if search_input_order:
                                                        break
                                                    time.sleep(0.5)

                                                if search_input_order:
                                                    print(">>> 找到订单搜索框，正在输入...")

                                                    # 1. 激活并清空
                                                    search_input_order.click()
                                                    time.sleep(0.2)
                                                    search_input_order.clear()

                                                    # 2. 逐字输入 (模拟打字)
                                                    for char in order_code:
                                                        search_input_order.input(char, clear=False)
                                                        time.sleep(0.1)  # 打字间隔

                                                    # 3. 开启网络监听 (预测接口包含 MtPurchase)
                                                    tab.listen.start(targets='Admin/MtPurchase')
                                                    print(">>> 已开启网络监听，等待搜索结果...")

                                                    # 4. 触发回车搜索
                                                    search_input_order.run_js("""
                                                        this.dispatchEvent(new Event('change', { bubbles: true }));
                                                        this.dispatchEvent(new KeyboardEvent("keydown", {bubbles:true, keyCode:13, key:"Enter"}));
                                                        this.dispatchEvent(new KeyboardEvent("keyup", {bubbles:true, keyCode:13, key:"Enter"}));
                                                    """)
                                                    print("✅ 输入完毕并回车")

                                                    # 5. 等待搜索结果返回
                                                    res_packet_order = tab.listen.wait(timeout=10)
                                                    if res_packet_order:
                                                        print(f"✅ 搜索成功，捕获到接口响应: {res_packet_order.url}")
                                                        print(">>> 正在选中所有搜索结果记录...")

                                                        # 给一点时间让表格渲染
                                                        time.sleep(0.5)

                                                        all_selected = False
                                                        target_frame = None

                                                        # 遍历所有可见 iframe，选中所有记录
                                                        for frame in tab.eles('tag:iframe'):
                                                            if not frame.states.is_displayed:
                                                                continue

                                                            try:
                                                                # 策略A: 尝试点击"全选"按钮（如果存在）
                                                                select_all_btn = frame.ele(
                                                                    'x://input[@type="checkbox" and contains(@onclick, "selectAll")]',
                                                                    timeout=0.5)
                                                                if not select_all_btn:
                                                                    # 尝试其他全选按钮的定位方式
                                                                    select_all_btn = frame.ele(
                                                                        'x://button[contains(text(), "全选") or contains(text(), "选择全部")]',
                                                                        timeout=0.5)

                                                                if select_all_btn and select_all_btn.states.is_displayed:
                                                                    select_all_btn.click()
                                                                    print("✅ 点击全选按钮成功")
                                                                    all_selected = True
                                                                    target_frame = frame
                                                                    break

                                                                # 策略B: 如果没有全选按钮，逐个勾选所有可见的复选框
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
                                                                                time.sleep(0.1)  # 防止点击过快

                                                                    if selected_count > 0:
                                                                        print(f"✅ 已勾选 {selected_count} 条记录")
                                                                        all_selected = True
                                                                        target_frame = frame
                                                                        break

                                                            except Exception as e:
                                                                print(f"   -> 在当前iframe中选择记录时出错: {e}")
                                                                continue

                                                        if not all_selected:
                                                            print("⚠️ 未能选中任何记录，可能表格结构发生变化")
                                                        else:
                                                            # 滚动一下确保页面稳定
                                                            if target_frame:
                                                                target_frame.scroll.down(200)
                                                                time.sleep(0.5)
                                                                print(">>> 记录选择完成，页面已滚动调整")

                                                        # 选中了记录，继续执行后续操作
                                                        if all_selected and target_frame:
                                                            print(">>> 记录已选中，准备触发附件上传...")
                                                            try:
                                                                # 1. 定位并点击“附件”页签
                                                                adjunct_tab = target_frame.ele(
                                                                    'x://a[contains(text(), "附件") and contains(@href, "tb_Adjunct")]',
                                                                    timeout=2)

                                                                if adjunct_tab:
                                                                    adjunct_tab.click()
                                                                    time.sleep(0.5)  # 等待面板切换

                                                                    # 2. 定位并点击上传图标 (透明 label)
                                                                    upload_label = target_frame.ele(
                                                                        'x://div[@id="tb_Adjunct"]//label[contains(@style, "opacity: 0")]',
                                                                        timeout=2)

                                                                    if upload_label:
                                                                        try:
                                                                            # 获取当前系统下的绝对路径 (自动兼容 Windows/Mac)
                                                                            abs_img_path = os.path.abspath(img_path)
                                                                            print(f"   -> 预设上传路径: {abs_img_path}")

                                                                            # 设置上传文件
                                                                            upload_label.click.to_upload(abs_img_path)

                                                                            print("✅ 成功点击上传图标并拦截文件选择框")

                                                                            # 等待上传进度 (根据网络情况调整等待时间)
                                                                            print(">>> 正在上传附件，请稍候...")
                                                                            time.sleep(5)

                                                                            # 上传完成后点击保存按钮
                                                                            print(
                                                                                ">>> 附件上传完成，正在查找并点击保存按钮...")
                                                                            try:
                                                                                # 定位保存按钮 - 多种策略
                                                                                save_img_btn = None

                                                                                # 策略A: 通过onclick属性定位
                                                                                save_img_btn = target_frame.ele(
                                                                                    'x://button[@onclick="AddImg()"]',
                                                                                    timeout=2)

                                                                                # 策略B: 通过按钮文本定位 (兜底)
                                                                                if not save_img_btn:
                                                                                    save_img_btn = target_frame.ele(
                                                                                        'x://button[contains(text(), "保存") and contains(@class, "btn-success")]',
                                                                                        timeout=2)

                                                                                # 策略C: 通过CSS类名定位
                                                                                if not save_img_btn:
                                                                                    save_img_btn = target_frame.ele(
                                                                                        'css:button.btn.btn-success.btn-sm',
                                                                                        timeout=2)

                                                                                if save_img_btn and save_img_btn.states.is_displayed:
                                                                                    # 确保按钮可见并点击
                                                                                    save_img_btn.scroll.to_see()
                                                                                    time.sleep(0.5)  # 稍微停顿，模拟人工确认
                                                                                    save_img_btn.click()
                                                                                    print("✅ 成功点击附件保存按钮")
                                                                                    time.sleep(2)  # 等待保存完成
                                                                                else:
                                                                                    print("⚠️ 未找到附件保存按钮")

                                                                            except Exception as e:
                                                                                print(
                                                                                    f"!!! 点击保存按钮时发生异常: {e}")

                                                                        except Exception as e:
                                                                            print(
                                                                                f"!!! 文件路径处理或上传设置失败: {e}")
                                                                    else:
                                                                        print("⚠️ 未在附件面板中找到上传图标")
                                                                else:
                                                                    print("⚠️ 未找到“附件”页签")

                                                            except Exception as e:
                                                                print(f"!!! 附件操作过程中发生异常: {e}")

                                                            print(">>> 准备执行采购任务...")
                                                            time.sleep(1)  # 稍作缓冲

                                                            try:
                                                                # 定位并点击“更多”按钮
                                                                more_btn = target_frame.ele(
                                                                    'x://button[contains(text(), "更多")]', timeout=2)

                                                                if more_btn:
                                                                    more_btn.click()
                                                                    time.sleep(0.5)  # 等待下拉菜单展开动画

                                                                    # 定位并点击“执行采购任务”
                                                                    exec_task_btn = target_frame.ele(
                                                                        'css:a[onclick="doMtPurTask()"]', timeout=1)

                                                                    # 策略B: 如果属性找不到，使用文字定位兜底
                                                                    if not exec_task_btn:
                                                                        exec_task_btn = target_frame.ele(
                                                                            'x://a[contains(text(), "执行采购任务")]',
                                                                            timeout=1)

                                                                    if exec_task_btn:
                                                                        print(
                                                                            "   -> 找到“执行采购任务”按钮，正在点击...")
                                                                        exec_task_btn.click()

                                                                        # 处理可能出现的 Alert/Confirm 弹窗
                                                                        try:
                                                                            if tab.wait.alert(timeout=3):
                                                                                alert_text = tab.alert.text
                                                                                print(
                                                                                    f"ℹ️ 检测到系统弹窗: [{alert_text}]，自动接受...")
                                                                                tab.alert.accept()
                                                                        except:
                                                                            pass

                                                                        print("✅ 成功点击“执行采购任务”")
                                                                        time.sleep(2)  # 等待系统处理

                                                                        print(">>> 等待并处理结果弹窗...")
                                                                        try:
                                                                            # 优先在主页面 (tab) 查找
                                                                            confirm_btn = tab.ele(
                                                                                'css:a.layui-layer-btn0', timeout=3)

                                                                            # 2. 如果主页面没找到，尝试在当前 iframe (target_frame) 查找
                                                                            if not confirm_btn:
                                                                                confirm_btn = target_frame.ele(
                                                                                    'css:a.layui-layer-btn0', timeout=2)

                                                                            if confirm_btn:
                                                                                confirm_btn.click()
                                                                                print("✅ 成功点击弹窗“确定”按钮")
                                                                                time.sleep(1)  # 等待弹窗关闭动画

                                                                                # 调用外部独立函数
                                                                                navigate_and_search_purchase_task(tab,
                                                                                                                  data_json.get(
                                                                                                                      'rpa_order_code'),
                                                                                                                  data_json)

                                                                                navigate_to_bill_list(tab,
                                                                                                      data_json.get(
                                                                                                          'rpa_order_code'))
                                                                            else:
                                                                                print(
                                                                                    "⚠️ 未检测到 Layui 结果弹窗 (可能已自动关闭或无需确认)")

                                                                        except Exception as e:
                                                                            print(f"!!! 处理结果弹窗时发生异常: {e}")
                                                                    else:
                                                                        print(
                                                                            "⚠️ 展开了“更多”菜单，但未找到“执行采购任务”选项")
                                                                    # 尝试关闭下拉菜单，避免遮挡
                                                                    more_btn.click()
                                                                else:
                                                                    print("⚠️ 未找到“更多”按钮")

                                                            except Exception as e:
                                                                print(f"!!! 执行采购任务操作异常: {e}")
                                                    else:
                                                        print(
                                                            "⚠️ 搜索请求发送了，但未捕获到预期的网络响应 (可能是接口规则不同)")

                                                    tab.listen.stop()

                                                    # 截图留存证据 (可选)
                                                    # tab.get_screenshot(path=os.path.join(CONFIG.get('data_storage_path'), f"{file_name}_final.jpg"))

                                                else:
                                                    print(
                                                        "!!! 错误: 遍历了所有iframe，未能在“物料采购订单”页面找到搜索框 (#txtSearchKey[data-grid='POMtPurchaseGrid'])")
                                            else:
                                                print("ℹ️ 上下文中无订单编号 (rpa_order_code)，跳过搜索步骤")


                                        except Exception as e:
                                            print(f"!!! 菜单跳转过程中发生异常: {e}")
                                    else:
                                        print("⚠️ 未找到“保存并审核”按钮 (可能已自动保存或界面结构变化)")

                                except Exception as e:
                                    print(f"!!! 点击保存按钮时发生异常: {e}")
                else:
                    print("响应内容不是 JSON 格式")

            else:
                print(f"!!! 警告: 等待 {10} 秒后未捕获到 URL 包含 '{target_url_substring}' 的请求。")

            # 停止监听 (虽然 wait 抓到一个后通常不需要手动停，但为了保险可以重置)
            tab.listen.stop()


        else:
            print("!!! 错误：没找到可见的搜索框，请检查页面是否加载完成。")

    except Exception as e:
        print(f"!!! rpa执行过程中发生异常: {e}")
    finally:
        # 关闭页签
        if tab:
            try:
                # tab.close()
                print(f"[{file_name}] 页签已关闭")
            except Exception as e:
                print(f"[{file_name}] 页签关闭异常: {e}")

    return match_prompt, match_result, original_records, retry_count


def parse_single_image(img_path, db_manager, LOCAL_STYLE_DB):
    """解析单张图片"""
    file_name = os.path.basename(img_path)

    try:
        # 状态检查 (去重)
        if db_manager.is_processed(file_name):
            print(f"[跳过] 文件已处理过: {file_name}")
            return None

        print(f"[处理中] 正在解析: {file_name} ...")

        # 根据配置决定是否使用LLM解析图片
        if CONFIG.get('use_llm_image_parsing', True):
            parsed_data = extract_data_from_image(img_path, PROMPT_INSTRUCTION)
            # parsed_data = extract_data_from_image_dmx(img_path, PROMPT_INSTRUCTION)
        else:
            parsed_data = {
                'buyer_name': '素本服饰',
                'delivery_date': '2025-11-22',
                'final_selected_style': 'H1635A-B',
                'items': [{'price': 18.0, 'qty': 11.0, 'raw_style_text': '00101020105', 'unit': '并'}],
                'style_candidates': [{'is_red': False, 'position': '表格备注栏', 'text': 'H1635A-B'}],
                'supplier_name': '杭州楼国忠辅料(辅料城仓库店)'
            }

        final_style = determine_final_style(parsed_data, LOCAL_STYLE_DB)

        parsed_data['final_selected_style'] = final_style
        print(f"最终判定款号: {final_style}")

        # 检查交付日期异常，如果异常则用DMX重新识别
        delivery_date = parsed_data.get('delivery_date', '')
        if delivery_date and should_use_dmx_for_date_check(delivery_date):
            print(f">>> 交付日期异常: {delivery_date}，使用DMX重新识别...")
            dmx_parsed_data = extract_data_from_image_dmx(img_path, PROMPT_INSTRUCTION, 0)
            if dmx_parsed_data and 'error' not in dmx_parsed_data:
                print(">>> DMX重新识别成功，替换原始数据")
                parsed_data = dmx_parsed_data
                final_style = determine_final_style(parsed_data, LOCAL_STYLE_DB)
                parsed_data['final_selected_style'] = final_style
                parsed_data['used_dmx_for_date_check'] = True  # 标记使用了DMX重新识别
                print(f">>> DMX识别后的款号: {final_style}")
            else:
                print(">>> DMX重新识别失败，继续使用原始数据")
                parsed_data['dmx_recheck_failed'] = True

        # 重试识别逻辑
        valid_prefixes = tuple(CONFIG.get('valid_style_prefixes', ['T', 'H', 'X', 'D']))
        max_retries = CONFIG.get('image_recognition_max_retries', 3)
        retry_count = 1
        failure_reason = None

        if not final_style or (final_style and not final_style.upper().startswith(valid_prefixes)):
            if not final_style:
                print(">>> 未识别到款号，开始重试识别...")
            else:
                print(f">>> 识别的款号{final_style}不符合规律，开始重试识别...")

            for retry_attempt in range(1, max_retries + 1):
                print(f">>> 第{retry_attempt}次重试...")
                # retry_parsed_data = extract_data_from_image(img_path, PROMPT_INSTRUCTION)
                retry_parsed_data = extract_data_from_image_dmx(img_path, PROMPT_INSTRUCTION)
                retry_count = retry_attempt + 1

                if retry_parsed_data and 'error' not in retry_parsed_data:
                    retry_final_style = determine_final_style(retry_parsed_data, LOCAL_STYLE_DB)
                    if retry_final_style and retry_final_style.upper().startswith(valid_prefixes):
                        print(f">>> 重试成功: {retry_final_style}")
                        parsed_data = retry_parsed_data
                        final_style = retry_final_style
                        parsed_data['final_selected_style'] = final_style
                        failure_reason = None
                        break
                else:
                    print(f">>> 第{retry_attempt}次重试识别失败")
            else:
                failure_reason = "款号没有解析到"
                print(f">>> 所有重试均失败，{failure_reason}")

        # 如果最终仍然没有有效款号，直接返回失败结果，不进入RPA逻辑
        if not final_style or not final_style.upper().startswith(valid_prefixes):
            print(f"!!! 款号解析最终失败: {file_name}, 款号: {final_style or 'None'}")
            return {
                'file_name': file_name,
                'img_path': img_path,
                'success': False,
                'error': '款号没有解析到',
                'parsed_data': parsed_data,
                'final_style': final_style or '',
                'failure_reason': '款号没有解析到'
            }

        parsed_data['retry_count'] = retry_count
        if failure_reason:
            parsed_data['failure_reason'] = failure_reason

        # 错误检查
        if not parsed_data or 'error' in parsed_data:
            error_msg = parsed_data.get('error', '未知') if parsed_data else '解析结果为空'
            print(f"!!! 解析失败，跳过: {file_name}, 原因: {error_msg}")
            return {
                'file_name': file_name,
                'img_path': img_path,
                'success': False,
                'error': error_msg,
                'parsed_data': parsed_data,
                'final_style': None,
                'failure_reason': parsed_data.get('failure_reason', error_msg) if parsed_data else error_msg
            }

        # 数据持久化
        saved = db_manager.save_data(file_name, parsed_data)
        if not saved:
            print("!!! 数据保存失败，中断处理本条")
            return {
                'file_name': file_name,
                'img_path': img_path,
                'success': False,
                'error': '数据保存失败',
                'parsed_data': parsed_data,
                'final_style': final_style,
                'failure_reason': parsed_data.get('failure_reason', '数据保存失败')
            }

        return {
            'file_name': file_name,
            'img_path': img_path,
            'success': True,
            'parsed_data': parsed_data,
            'final_style': final_style,
            'failure_reason': parsed_data.get('failure_reason', '')
        }

    except Exception as e:
        print(f"!!! 解析图片异常: {file_name}, 错误: {e}")
        return {
            'file_name': file_name,
            'img_path': img_path,
            'success': False,
            'error': str(e),
            'parsed_data': None,
            'final_style': None,
            'failure_reason': str(e)
        }


def process_complete_rpa(browser, result):
    """处理完整的RPA+LLM匹配流程并收集报告数据"""
    try:
        if result['success']:
            file_name = result['file_name']
            parsed_data = result['parsed_data']
            img_path = result['img_path']

            # 执行完整RPA+LLM流程
            match_prompt, match_result, original_records, retry_count = process_single_bill_rpa(browser, parsed_data,
                                                                                                file_name, img_path)

            # 收集报告数据
            return collect_result_data(
                image_name=file_name,
                parsed_data=parsed_data,
                final_style=result['final_style'],
                match_prompt=match_prompt,
                match_result=match_result,
                original_records=original_records,
                image_path=result['img_path'],
                retry_count=retry_count,
                failure_reason=result.get('failure_reason', '')
            )
        else:
            # 解析失败，生成失败报告
            return collect_result_data(
                image_name=result['file_name'],
                parsed_data=result.get('parsed_data', {}),
                final_style=result.get('final_style', ''),
                match_prompt="",
                match_result=None,
                original_records=[],
                image_path=result['img_path'],
                retry_count=1,
                failure_reason=result.get('failure_reason', '')
            )
    except Exception as e:
        print(f"!!! 完整流程异常: {result['file_name']}, 错误: {e}")
        return collect_result_data(
            image_name=result['file_name'],
            parsed_data=result.get('parsed_data', {}),
            final_style=result.get('final_style', ''),
            match_prompt="",
            match_result=None,
            original_records=[],
            image_path=result['img_path'],
            retry_count=1,
            failure_reason=str(e)
        )


async def async_main():
    # 1. 初始化阶段：加载款号库 (Excel + 缓存加速)
    excel_path = CONFIG.get('style_db_path')
    col_name = CONFIG.get('style_db_column', '款式编号')
    LOCAL_STYLE_DB = load_style_db_with_cache(excel_path, col_name)

    # 2. 初始化数据管理器
    storage_path = CONFIG.get('data_storage_path')
    if not storage_path:
        print("错误: 请在 app_config.py 中配置 'data_storage_path'")
        return
    db_manager = DataManager(storage_path_str=storage_path)

    # 3. 扫描图片文件夹
    source_dir = CONFIG.get('image_source_dir')
    if not source_dir or not os.path.exists(source_dir):
        print(f"错误: 图片源目录不存在，请检查 app_config.py 配置: {source_dir}")
        return

    image_files = []
    for ext in ['*.jpg', '*.jpeg', '*.png', '*.JPG', '*.PNG']:
        image_files.extend(glob.glob(os.path.join(source_dir, ext)))

    if not image_files:
        print(f"目录 {source_dir} 下没有找到图片文件。")
        return

    print(f"扫描到 {len(image_files)} 个图片文件，准备开始处理...")

    # 4. 连接浏览器 (只连一次)
    co = ChromiumOptions().set_address(CONFIG['chrome_address'])
    browser = Chromium(addr_or_opts=co)

    # 5. 阶段1：并发图片解析
    print("\n=== 阶段1：并发图片解析 ===")
    parsing_concurrency = CONFIG.get('image_parsing_concurrency', 3)
    semaphore = asyncio.Semaphore(parsing_concurrency)

    async def parse_with_semaphore(img_path):
        async with semaphore:
            loop = asyncio.get_event_loop()
            with ThreadPoolExecutor() as executor:
                return await loop.run_in_executor(executor, parse_single_image, img_path, db_manager, LOCAL_STYLE_DB)

    parse_tasks = [parse_with_semaphore(img_path) for img_path in image_files]
    parse_results = await asyncio.gather(*parse_tasks, return_exceptions=True)

    # 过滤有效结果
    valid_results = [r for r in parse_results if r is not None and not isinstance(r, Exception)]
    success_results = [r for r in valid_results if r.get('success', False)]

    print(f"解析完成: 总计 {len(image_files)} 张，有效 {len(valid_results)} 张，成功 {len(success_results)} 张")

    # 分离成功和失败的结果
    failed_results = [r for r in valid_results if not r.get('success', False)]
    print(f"款号解析失败: {len(failed_results)} 张")

    # 6. 阶段2：并发RPA+LLM匹配 (只处理成功解析款号的结果)
    print("\n=== 阶段2：并发RPA+LLM匹配 ===")
    rpa_concurrency = CONFIG.get('rpa_concurrency', 3)
    rpa_semaphore = asyncio.Semaphore(rpa_concurrency)

    async def process_with_semaphore(result):
        async with rpa_semaphore:
            loop = asyncio.get_event_loop()
            with ThreadPoolExecutor() as executor:
                return await loop.run_in_executor(executor, process_complete_rpa, browser, result)

    # 只对成功解析款号的结果进行RPA处理
    process_tasks = [process_with_semaphore(result) for result in success_results]
    rpa_results = await asyncio.gather(*process_tasks, return_exceptions=True)

    # 为失败的结果生成报告数据
    failed_report_results = []
    for failed_result in failed_results:
        report_data = collect_result_data(
            image_name=failed_result['file_name'],
            parsed_data=failed_result.get('parsed_data', {}),
            final_style=failed_result.get('final_style', ''),
            match_prompt="",
            match_result=None,
            original_records=[],
            image_path=failed_result['img_path'],
            retry_count=failed_result.get('parsed_data', {}).get('retry_count', 1),
            failure_reason=failed_result.get('failure_reason', '款号没有解析到')
        )
        failed_report_results.append(report_data)

    # 合并所有结果
    final_results = rpa_results + failed_report_results

    print(f"RPA+LLM匹配完成: {len(final_results)} 张")

    # 7. 统一生成报告
    print("\n=== 生成最终报告 ===")
    report_path = CONFIG.get('report_output_path')
    if report_path:
        try:
            # 生成带时间戳的报告文件名
            timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            report_filename = f'report_{timestamp}.html'

            # 确保目录存在
            if os.path.isdir(report_path):
                report_file = os.path.join(report_path, report_filename)
            else:
                report_file = os.path.join(report_path, report_filename)

            # 逐个更新报告
            for result_data in final_results:
                if not isinstance(result_data, Exception) and result_data:
                    update_html_report(os.path.dirname(report_file), result_data['image_name'], result_data)

            # 重命名最终生成的report.html为带时间戳的文件名
            default_report = os.path.join(os.path.dirname(report_file), 'report.html')
            if os.path.exists(default_report):
                os.rename(default_report, report_file)

            print(f"✅ 报告已生成: {report_file}")
        except Exception as e:
            print(f"⚠️ 报告生成失败: {e}")

    print(">>> 所有任务处理完毕。")


# 预处理 records，把 /Date(xxx)/ 时间戳转成人类可读日期
def preprocess_records(records):
    """
    (保持不变) 预处理 records，把 /Date(xxx)/ 时间戳转成人类可读日期
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


def execute_smart_match(parsed_data, records):
    """
    执行智能匹配核心逻辑
    返回: (match_result, match_prompt, retry_count) 元组
    """
    # 1. 准备上下文数据
    today = datetime.date.today()
    two_weeks_ago = today - timedelta(days=14)

    # 2. 清洗后台记录
    clean_records = preprocess_records(records)

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


# 根据 llm 匹配后的返回构建 RPA 传递数据
def reconstruct_rpa_data(match_result, original_parsed_data, original_records):
    matched_tasks = []

    # 1. 建立快速查找表 (Hash Map)
    # 将 list 转换为 dict: { "uuid": record_obj }，方便 O(1) 查找
    record_map = {rec['Id']: rec for rec in original_records}
    ocr_items_list = original_parsed_data.get('items', [])

    # 辅助函数：安全获取 OCR Item
    def get_item_by_index(idx):
        if isinstance(idx, int) and 0 <= idx < len(ocr_items_list):
            return ocr_items_list[idx]
        return None

    # --- Type 1: Direct (1对1) ---
    for match in match_result.get('direct_matches', []):
        rid = match.get('record_id')
        idx = match.get('ocr_index')

        target_record = record_map.get(rid)
        target_item = get_item_by_index(idx)

        if target_record and target_item:
            matched_tasks.append({
                "match_type": "DIRECT",
                "record": target_record,  # 完整 Record
                "items": [target_item],  # 完整 Item (统一放入列表)
                "ocr_context": original_parsed_data
            })

    # --- Type 2: Merge (N对1) ---
    for match in match_result.get('merge_matches', []):
        rid = match.get('record_id')
        indices = match.get('ocr_indices', [])

        target_record = record_map.get(rid)
        # 获取所有对应的 items
        target_items = [get_item_by_index(i) for i in indices if get_item_by_index(i)]

        if target_record and target_items:
            matched_tasks.append({
                "match_type": "MERGE",
                "record": target_record,  # 完整 Record
                "items": target_items,  # 多个 完整 Items
                "ocr_context": original_parsed_data
            })

    # --- Type 3: Split (1对N) ---
    for match in match_result.get('split_matches', []):
        rid = match.get('record_id')
        idx = match.get('ocr_index')

        target_record = record_map.get(rid)
        target_item = get_item_by_index(idx)

        if target_record and target_item:
            matched_tasks.append({
                "match_type": "SPLIT",
                "record": target_record,  # 完整 Record
                "items": [target_item],  # 完整 Item
                "ocr_context": original_parsed_data
            })

    return matched_tasks


# 选择并勾选匹配的记录复选框
def fill_details_into_table(scope, structured_tasks):
    """
    根据匹配任务填充 tbody 中的物料数据
    修复点：
    1. JS 里的 arguments[0] 全部改为 this，解决 'Cannot read properties of undefined' 报错
    2. 保持原有业务逻辑结构不变
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
            # 使用 hidden input 的 materialReqId 来定位所在的 TR
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

            # 获取第一条 Item 的基础信息
            first_item = items[0]
            # 注意：这里字段名要和 parse_single_image 返回的 key 一致
            raw_unit = first_item.get('unit', '')
            raw_price = first_item.get('price', 0)
            # 交付日期从 OCR 解析的整体数据中获取
            raw_date = task['ocr_context'].get('delivery_date')

            if match_type == 'DIRECT':
                # 一对一：直接填入
                target_unit = raw_unit
                target_price = raw_price
                target_qty = first_item.get('qty', 0)
                target_date = raw_date

            elif match_type == 'MERGE':
                # 多对一：数量累加，其他取第一条
                total_qty = sum([float(i.get('qty', 0)) for i in items])
                target_unit = raw_unit
                target_price = raw_price
                target_qty = total_qty
                target_date = raw_date
                print(f"   ℹ️ [合并] 记录 {record_id} 聚合了 {len(items)} 条明细，总数: {target_qty}")

            elif match_type == 'SPLIT':
                # 一对多：
                target_unit = raw_unit
                target_price = raw_price
                target_qty = first_item.get('qty', 0)
                target_date = raw_date
                print(f"   ℹ️ [拆分] 记录 {record_id} 强制调整数量为: {target_qty}")

            # --- 3. 执行填入操作 ---

            # A. 填入采购单位 (模拟：点击->搜索->双击)
            if target_unit:
                inp_unit = tr.ele('css:input[name="unitCalc"]', timeout=0.5)
                if inp_unit:
                    # 1. 点击弹出窗口
                    inp_unit.click()
                    time.sleep(0.5)  # 等待弹窗动画

                    # 2. 定位弹窗中的搜索框
                    search_box = scope.ele('#txtMeteringPlusKey', timeout=1)

                    if search_box and search_box.states.is_displayed:
                        # 3. 输入单位并回车
                        search_box.clear()
                        search_box.input(target_unit)
                        time.sleep(0.2)
                        scope.actions.key_down('ENTER').key_up('ENTER')  # 模拟回车键搜索
                        # time.sleep(1.0)  # [重要] 等待表格刷新，给足时间

                        # 4. 在结果表格中找到完全匹配的那一行并双击
                        target_td_xpath = f'x://table[@id="meteringPlusGrid"]//tbody//tr//td[text()="{target_unit}"]'
                        target_td = scope.ele(target_td_xpath, timeout=1)

                        if target_td:
                            print(f"   -> 找到单位单元格 [{target_unit}]，执行JS双击...")
                            # >>> 关键修改：将 arguments[0] 改为 this <<<
                            js_code = """
                                this.click(); 
                                this.dispatchEvent(new MouseEvent('dblclick', {bubbles: true, cancelable: true, view: window}));
                            """
                            target_td.run_js(js_code)
                            time.sleep(0.5)
                        else:
                            print(f"   ⚠️ 单位弹窗中未搜索到: {target_unit}")
                            inp_unit.click()  # 尝试关闭弹窗
                    else:
                        # 如果弹窗没出来，尝试直接输入(兜底)
                        print("   ⚠️ 单位选择弹窗未出现，尝试直接输入")
                        inp_unit.input(target_unit, clear=True)

            # B. 填入含税单价
            if target_price is not None:
                inp_price = tr.ele('css:input[name="Price"]', timeout=0.2)
                if inp_price:
                    val = str(target_price)
                    js = f'this.value = "{val}"; this.dispatchEvent(new Event("input")); this.dispatchEvent(new Event("change")); this.dispatchEvent(new Event("blur"));'
                    inp_price.run_js(js)
                    time.sleep(0.1)  # 给页面一点反应时间

            # C. 填入数量
            if target_qty is not None:
                inp_qty = tr.ele('css:input[name="Qty"]', timeout=0.2)
                if inp_qty:
                    # 【修改】同上，使用 JS 赋值
                    val = str(target_qty)
                    js = f'this.value = "{val}"; this.dispatchEvent(new Event("input")); this.dispatchEvent(new Event("change")); this.dispatchEvent(new Event("blur"));'
                    inp_qty.run_js(js)
                    time.sleep(0.1)

            # D. 触发总价计算 (点击 totalAmount)
            inp_total = tr.ele('css:input[name="totalAmount"]', timeout=0.2)
            if inp_total:
                inp_total.click()
                time.sleep(0.2)  # 等待系统自动计算

            # E. 填入交付日期 (模拟键盘手动输入)
            # target_date="2024-11-11" # debug line, remove later
            if target_date and target_date.strip():
                inp_date = tr.ele('css:input.deliveryDate', timeout=0.5)

                if inp_date:
                    try:
                        print(f"   -> 正在填入日期: {target_date}")

                        # 1. 【核心】完全模仿 Selenium：JS 移除 readonly 属性
                        # DrissionPage 中，this 直接指代当前元素，比 arguments[0] 更简洁
                        inp_date.run_js('this.removeAttribute("readonly");')

                        # 2. 清空输入框
                        inp_date.clear()
                        time.sleep(0.1)

                        # 3. 模拟键盘输入 (相当于 Selenium 的 send_keys)
                        inp_date.input(target_date)
                        time.sleep(0.2)  # 等待输入反应

                        # 4. 模拟按下“回车”键
                        # 这步很关键：通常回车会触发控件的 change 事件并自动关闭弹窗
                        scope.actions.key_down('ENTER').key_up('ENTER')
                        time.sleep(0.2)

                        # 5. 【兜底】如果弹窗还在，点一下旁边空白处强制失焦
                        # (相当于点击页面背景)
                        scope.run_js('document.body.click();')
                        inp_date.click()  # 再次点击输入框确保焦点在上面
                        time.sleep(0.2)
                    except Exception as e:
                        print(f"   ⚠️ 日期输入异常: {e}")
                        # 只有出错时才尝试暴力赋值兜底
                        try:
                            inp_date.run_js(f'this.removeAttribute("readonly"); this.value="{target_date}";')
                        except:
                            pass

            count_success += 1
            time.sleep(0.1)  # 稍微防抖

        except Exception as e:
            print(f"   !!! 填充行数据失败 (Record: {task.get('record', {}).get('Id')}): {e}")

    print(f"✅ 数据填充完成: 成功处理 {count_success}/{len(structured_tasks)} 行")


def main():
    """主入口函数"""
    asyncio.run(async_main())


if __name__ == "__main__":
    main()