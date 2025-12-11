#!/usr/bin/env python3
"""
配置文件
"""
# 三个主要功能开关
CONFIG = {

    # LLM 配置
    # "llm_image_api_key":"sk-81df8063219a4b398af8ea707e293034",    # 微微的阿里API Key
    # "llm_api_key":"sk-81df8063219a4b398af8ea707e293034",

     "llm_image_api_key": [
         "sk-a58eda60805e41999596ef92fb385480bak",   # 微妈的阿里API Key
        "sk-81df8063219a4b398af8ea707e293034bak" # 微微的阿里API Key
     ],

     "llm_api_key": [
         "sk-a58eda60805e41999596ef92fb385480",   # 微妈的阿里API Key
        "sk-81df8063219a4b398af8ea707e293034" # 微微的阿里API Key
     ],

    # "llm_image_api_key": "sk-a58eda60805e41999596ef92fb385480",  # 微妈的阿里API Key

    # "llm_api_key": "sk-a58eda60805e41999596ef92fb385480",

    # "gjllm_api_key":"sk-pwrkkffhkwgtkmhziljvfcvgmxdsgvbwuzozywiwclwmiywf",    # 杨青的硅基流动API Key

    # DMX 大模型配置
    "dmx_api_key": [
        "sk-Rr9CXIyBuc4sq9mpEyaN8l3WhVYB8P1TpahojO1HT1bozO4B",  # 主要DMX API Key
    ],
    "dmx_api_url": "https://www.dmxapi.cn/v1/chat/completions",
    "dmx_model": "gemini-3-pro-preview",

    # DMX 图片识别配置
    "dmx_image_api_key": [
        "sk-Rr9CXIyBuc4sq9mpEyaN8l3WhVYB8P1TpahojO1HT1bozO4B",  # DMX图片识别API Key
    ],
    "dmx_image_api_url": "https://www.dmxapi.cn/v1/chat/completions",
    "dmx_image_model": "gemini-3-pro-preview",
    "dmx_image_resolution": "media_resolution_high",  # 图片分辨率级别

    # 图片解析开关：是否发图片到大模型解析
    "use_llm_image_parsing": True,  # False时使用硬编码数据，True时正常调用LLM解析
    
    # 测试模式：是否复用已有JSON文件数据
    "test_use_existing_json": True,  # True时跳过LLM解析，直接使用已有JSON继续RPA流程
    
    # 并发控制配置
    "image_parsing_concurrency": 10,  # 大模型图片识别同时并发数
    "rpa_concurrency": 1,  # RPA操作并发数
    "llm_match_max_retries": 3,  # 大模型文本匹配最大重试次数
    "rpa_browser_to_front": True,  # RPA执行时是否将浏览器窗口置前
    
    # 图片识别重试配置
    "valid_style_prefixes": ['T', 'H', 'X', 'D'],  # 有效的款号前缀
    "image_recognition_max_retries": 3,  # 大模型图片识别最大重试次数

    "delivery_date_threshold_days": 5,  # 交付日期异常阈值（天数），超过此天数差值将触发DMX重新识别

# 其他配置
    "chrome_address": "127.0.0.1:9333",
    "base_url": "https://a1.scm321.com/Admin/Dashboard",

    # 待处理图片的本地文件夹路径 (请修改为您实际的图片文件夹)
    'image_source_dir': r'E:\pyprojects\wei\data\images',

    'data_storage_path': r'E:\pyprojects\wei\data',

    # HTML报告输出路径
    'report_output_path': r'E:\pyprojects\wei\data\report',

    # 款号池的绝对路径
    'style_db_path': r'E:\pyprojects\wei\data\style\kuanhaochi.xlsx',

    # Excel 中存储款号的那一列的表头名称
    'style_db_column': '款式编号',

    # 月结供应商目录 Excel 文件路径
    'supplier_db_path': r'E:\pyprojects\wei\data\supplier\月结供应商名单.xlsx',
    
    # 供应商 Excel 文件中供应商名称的列名（可配置）
    'supplier_name_column': '供应商简称',  # 默认列名，可根据实际 Excel 表头修改
    
    # 供应商 Excel 文件中采购助理的列名（可配置）
    'supplier_purchaser_column': '采购助理',  # 默认列名，可根据实际 Excel 表头修改

    # 优惠厂商excel列表
    'material_deduction_db_path': r'E:\pyprojects\wei\data\material\罗卡-素本-单价优惠.xlsx',

    # 物料扣减列表 Excel 中存储物料名称的那一列的表头名称
    'material_deduction_name_column': '品名',
    
    # 物料扣减列表 Excel 中存储扣减金额的那一列的表头名称
    'material_deduction_amount_column': '销售均价',
    
    # 支持物料扣减的供应商列表（这些供应商的商品会进行价格调整）
    'material_deduction_suppliers': ['发现色彩', '罗卡'],

    #成功处理的图片归档目录
    'success_archive_path': r'E:\pyprojects\wei\data\success_images',

    # 图片数据提取的提示词配置
    'prompt_instruction': """
    你是一个专业的单据数据提取助手。请分析这张采购单据图片，提取关键信息并严格按照下方的 JSON 格式返回。

    请注意：不要使用 Markdown 格式，只返回纯 JSON 字符串。

    ### 0. 供应商匹配规则 (⚠️最高优先级⚠️)
    **已知有效供应商列表**：
    {known_suppliers}

    **匹配与清洗逻辑（请严格执行）**：
    1. **多区域识别**：图片左上角的供应商信息通常包含两部分：
       - **大号字体**的招牌名/简称
       - **小号字体**的公司全称
       请同时识别这两部分文本。
    2. **列表比对**：将识别到的“大字招牌”和“小字全称”分别与上方的【已知有效供应商列表】进行比对。
    3. **模糊匹配**：如果OCR识别结果与列表中的名称存在包含关系，请直接使用列表中的标准名称。
    4. **兜底策略**：如果以上均未在列表中找到匹配项，则优先返回图片中字号最大的那个供应商名称。

    ### 1. 基础信息提取
    请提取以下字段：
    - 交付日期 (delivery_date): 格式为 YYYY-MM-DD。
    - 采购商名称 (buyer_name): 例如“素本-郑辉杰”。
    - 供应商名称 (supplier_name): 根据上述“供应商匹配规则”确定的最终标准名称。
    - 码单号 (delivery_order_number): 提取图片中的单号（如：BF-20251203-0641）。

    ### 2. 商品明细 (items)
    提取表格每一行：
    - 数量 (qty)
    - 单价 (price)
    - 单位 (unit)
    - 商品编码 (raw_style_text): 提取原有编码。
    - 商品描述 (product_description): 合并品名、颜色等描述列。

    ### 3. 款号候选池 (style_candidates)
    提取所有疑似款号的文本（重点关注 T/H/X/D 开头或红色手写字）。
    **特别注意**：请仔细检查图片底部的“备注”或“说明”栏，往往包含关键款号信息（例如“注：... T1341W ...”）。

    ### 4. 返回格式示例 (JSON)
    {{
        "delivery_date": "2025-12-03",
        "buyer_name": "素本-郑辉杰",
        "supplier_name": "罗卡家",
        "delivery_order_number": "BF-20251203-0641",
        "style_candidates": [
            {{
                "text": "T1341W", 
                "is_red": false, 
                "position": "底部备注栏"
            }}
        ],
        "items": [
            {{
                "qty": 178, 
                "price": 3, 
                "unit": "米", 
                "raw_style_text": "1003#超柔舒美绸",
                "product_description": "3#经典黑"
            }}
        ]
    }}
""",

# 单据匹配提示词模板
    'match_prompt_template': """
        # Role
        你是一个供应链单据匹配助手。你的任务是将 OCR 识别出的 Items 与 System Records 进行匹配。
        
        # Context
        当前日期: {current_date}
        有效日期范围: {two_weeks_ago} 至 {current_date}
        
        # Input Data
        1. OCR单据 (Parsed Data):
        {parsed_data_json}
        
        2. 系统记录 (System Records):
        {records_json}
        
        # Matching Logic (Strict Execution Steps)
        
        请严格遵循以下步骤进行逻辑推理：
        
        ## Step 1: 预处理与构建基准
        - **构建 Key**: 对于每条 System Record，组合 `System_Key` = `MaterialMtName` (物料名称) + `MaterialSpec` (物料规格)。
        - **日期过滤**: 标记出那些 `CreateTime_Readable` 在有效日期范围内的 Records。
        
        ## Step 2: 核心匹配 (遍历 OCR Items)
        
        对于 Parsed Data 中的每一条 Item，在 System Records 中寻找最佳匹配项。
        
        ** 核心认知 (重要):**
        系统记录 (System Record) 代表**“宽泛的采购需求/SKU总数”**，而 OCR 单据可能是**“具体的实物交付”**或**“分码细则”**。
        1.  **结构关系 (Structure)**:
            * 系统记录 (System Record) 代表**“宽泛的采购需求/SKU总数”**。
            * OCR 单据代表**“具体的实物交付”**或**“分码细则”**。
            * **多对一匹配**: 允许 **多条 OCR Items** (如分尺寸明细) 对应 **一条 System Record**。
        
        2.  **语义关系 (Semantics)**:
            * **包含即匹配**: 只要 System 核心词涵盖 OCR，即视为语义匹配。
            * **核心词主导 (重点)**: 判定匹配时，**核心名词**（如“弹簧扣”）+ **关键规格**（如“古银”）的权重为 100%。而**修饰性形容词**（如“哑光”vs“珍珠”、“雾纱”vs“凹槽”）视为**噪音**。
            * **指令**: 如果供应商、核心名词、颜色均一致，**必须忽略**中间修饰词的差异，直接认定为语义一致。
            * **行业术语/缩写映射 (Industry Terminology - 🌟关键新增)**:
              - **必须执行以下等价逻辑**，解决缩写与全称不一致的问题：
              - **拉链类**: 
                - System "隐拉" == OCR "隐形闭口" / "隐形拉链" / "3#隐形"。
                - System "拉链" == OCR "开门襟" / "牙开口" / "闭口" / "开口"。
                - **判定**: 只要出现上述任一对应词汇，**必须**视为核心名词完全一致。
                
        3.  **身份关系 (Identity)**:
            * **供应商模糊匹配**: "金致辅料" == "金致"。只要核心字号包含即可。
        
        **[数量校验规则]**
        1.  **数学红线 (General Math Guardrail)**: 
            * 通常情况下，语义匹配是前提，**数量一致是硬性指标**。
            * **计算公式**: `Diff_Rate = Abs(OCR_Qty - System_Amount) / OCR_Qty`。
            * **关键指令**: 计算差异率时，**必须以 OCR 数量作为分母**。
            * **阈值**: 如果 `Diff_Rate` > 50%，判定为 Fail。
        
        2.  **纯数值比对 (Pure Numerical Comparison)**:
            * **核心指令**: 在进行数量比对时，**完全忽略单位及其语义**。
            * **消除歧义**: 无论 System 字段名是 `TotalAmount` 还是其他，**严禁**将其推断为“金额/元”。
            * **操作**: 将 System 数值与 OCR `qty` 视为**纯数字**进行比对。
            * *例如*: System=200 (模型认为是金额) vs OCR=200 (米)，**必须**视为数值一致，**禁止**因为单位含义不同而拒绝。
        
        3.  **金致“包”单位特例 (JinZhi 'Bag' Exemption) - [最高优先级豁免]**:
            * **触发条件**: OCR `supplier_name` 包含 "金致" **且** 当前 Item 的 `unit` 为 "包"。
            * **执行动作**: **强制关闭数学红线**。
            * **判定逻辑**: 只要语义匹配（MaterialMtName/Spec 吻合），**无论数量差异多少**（例如 OCR=1 vs System=1000），**均视为匹配成功**。禁止因为数量不对而拒绝。
        
        
         
        必须遵循以下优先级顺序：
        
        **优先级 A: 单条完美匹配 (1-to-1 Perfect Match)**
        - **条件**: 供应商匹配 + 单条数量一致 (差异<=50%) + 语义一致。
        - **动作**: 直接绑定。
        - **特例执行**:
            1.  **兰搏拉链特例**: 如果供应商是“兰搏”，且数量对得上，**强制忽略** product_description 中的尺寸/长度规格（如 "50CM"），直接视为语义一致。
                ***注意**: 此特例仅豁免“语义”，**绝不豁免“数量”**。如果数量对不上，仍需拒绝。
            2.  **修饰词容错特例**: 如果供应商和核心词（如“弹簧扣”）及颜色一致，**强制忽略**修饰形容词（如“哑光”vs“珍珠”）的差异，直接视为语义一致。
            
        **优先级 B: 拆分/合并匹配 (N-to-N Aggregation Match - 🌟重点升级)**
        - **适用场景**: 单条 OCR 数量太小，但多条 OCR Items 的**数量之和**与 System Records 的数量吻合。
        - **核心逻辑**: **"智能分组求和 (Smart Grouping Sum)"**。
        - **动作**: 
          1. **识别同类**: 找到所有属于同一大类（如都有“拉链”）的 OCR Items。
          2. **强制累加 (Mandatory Summation)**:
             - **无论供应商是否一致**，只要 OCR 中存在多条描述相似的 Items，**必须**先计算它们的数量总和 (Total_OCR_Qty)。
             - **禁止**用单条数量去碰运气，必须用 **Total_OCR_Qty** 与 System Record 的 `TotalAmount` 进行比对。
          3. **尝试分组**: 
             - 如果 OCR 总数 (如 140) 远大于单条 System Record (如 70)，**严禁**直接用总数去匹配。
             - **必须**根据 OCR 描述中的**区分性关键词**（如颜色、镀色、规格）将 Items 分成子组。
             - *例如*: 将所有 "古银" 分为一组 (Sum=70)，将所有 "黑镍/枪色" 分为另一组 (Sum=70)。
          4. **兰搏拉链特例 (Lanbo Special Rule)**: 
             - 如果供应商含 "兰搏"，**严禁**因为尺寸/长度（如 "52.5CM" vs "53.5CM"）不同而将它们分为不同组。
             - **指令**: 不同长度的拉链，只要颜色/镀色一致，**必须合并**为同一个 Group 进行求和。
             - ⚠️ **警示**: 合并后的 **Group Sum** 必须与 System Record 的 TotalAmount 进行比对。
               - 如果 `Sum=24` vs `System=500` (差异 > 90%) -> **REJECT (拒绝匹配)**。
               - 只有差异 `<= 5%` 时，才允许绑定。
          5. **跨供应商聚合判定 (Cross-Supplier Aggregation - 关键)**:
             - **规则**: 如果 `Total_OCR_Qty` 与某条 System Record 的数量差异 <= 50% (且该 Record 的核心语义如“8212”被包含)，**即使供应商完全不同**，也**必须**判定为匹配。
             - **案例**: OCR(89.1+86.2+86.7 = 262) vs System(208)。差异率 |262-208|/262 ≈ 20.6% (小于50%) -> **MATCH**。严禁因为供应商不同而拒绝这种强数量关联。
          6. **分别匹配**: 
             - 将 **Group A 的总和** 与最匹配的 System Record A (如 "古银配色") 对比。若差异 <= 5%，绑定匹配。
             - 将 **Group B 的总和** 与最匹配的 System Record B (如 "枪色配色") 对比。若差异 <= 5%，绑定匹配。
          7. **语义宽容度**: 具体规格（如 "52.5CM"）是 System（如 "拉链"）的子集，允许匹配。
        
        **优先级 C: 兜底/跨品类匹配 (防幻觉重点)**
        - **适用场景**: 供应商完全不匹配，且无法进行聚合匹配。
        - **核心逻辑**: **"强行忽略供应商 + 数量近似 + 文本包含"**。
        - **动作**: 
          1. **数量计算**: `Abs(OCR_Qty - System_Amount) / OCR_Qty`。
             - 若差异 > 50%，直接拒绝。
             - 若差异 <= 50% (如 50 vs 50.4, 差异0.8%)，**必须进入下一步**。
          2. **文本强包含 (反幻觉检查)**: 
             - 检查 System `MaterialMtName` 是否被 OCR 描述 **包含**。
             - **禁止幻觉**: 如果字符串存在包含关系，**必须**判定为语义一致，**严禁**因为供应商不同而声称“物料不相关”。
          3. **短配长语义检查 (Short-match-Long)**: 
             - **核心规则**: 检查 System `MaterialMtName` (短) 是否是 OCR 描述 (长) 的**子字符串**或**核心词集合**。
             - **忽略缺失**: 如果 System 记录是 "100D四面弹"，而 OCR 是 "8820#100D平纹四面弹"，虽然 System 缺少 "8820#" 和 "平纹"，但 **System 核心词完整存在于 OCR 中**。
             - **强制匹配**: 只要满足上述“文本包含”且“数量一致”，**严禁**因为供应商不同或系统记录太短而拒绝匹配。
        
        ## Step 3: 判定案例 (Case Study - 必读)
        * **案例 1 (多对一拆分匹配)**:
          - System: "五号拉链" (Qty: 100)
          - OCR: "49CM 拉链"(40) + "50CM 拉链"(40) + "51CM 拉链"(20)
          - **结果**: **全部 MATCH 到同一 ID**。 (40+40+20=100)
        
        * **案例 2 (分组拆分匹配)**:
          - System A: "五号拉链" (Qty: 70, Spec: 古银)
          - System B: "五号拉链" (Qty: 70, Spec: 枪色)
          - OCR Items: 
            - "古银拉链 52CM"(40) + "古银拉链 53CM"(30) -> Group A (Sum=70)
            - "黑镍拉链 52CM"(40) + "黑镍拉链 53CM"(30) -> Group B (Sum=70)
          - **结果**: 
            - Group A -> Match System A
            - Group B -> Match System B
          - **原因**: 智能根据颜色关键词分组，每组总和分别匹配对应的 System Record。
        
        * **案例 3 (供应商冲突+数量微差)**:
          - System: "8819#50D四面弹" (Qty: 50.4, Supplier: 罗卡)
          - OCR: "8819#50D四面弹-270#裘皮咖" (Qty: 50, Supplier: 祖春)
          - **结果**: **MATCH**。 (数量差异 0.8% <= 50%，且文本包含，忽略供应商)
        
        ## Step 4: 规则校验 (关键!)
        1.  **一票否决**: 只要有任何一条 Item 没有找到匹配的 Record，本次任务直接判定为 Fail。
        
        ## Step 5: 判定案例
        * **案例 1 (修饰词容错匹配 - 本题关键)**:
          - System: "长条凹槽珍珠力弹簧扣 [H02]古银" (Qty: 320, Supp: 金致)
          - OCR: "哑光雾纱双孔弹簧扣 古银" (Qty: 400, Supp: 金致辅料)
          - **结果**: **MATCH** (优先级 A)。
          - **原因**: 供应商模糊匹配成功，核心词(弹簧扣)与颜色(古银)一致。"哑光" vs "珍珠" 视为修饰词噪音忽略。数量差异 25% 在溢短装允许范围内。
        
        
        # Output Format (JSON Only)
        请严格返回如下 JSON 结构，不要包含任何原始数据的文本细节，仅返回 ID 和 索引：
        
        {{
            "status": "success" | "fail",
            "global_reason": "如果失败，请明确指出是哪一条 Item 没找到",
            // 1. 一对一匹配 (Direct)
            "direct_matches": [
                {{
                    "record_id": "System_Record_Id",
                    "ocr_index": 0      //对应 OCR Items 列表中的下标
                }}
            ],
        
            // 2. 多对一合并 (Merge / N:1)
            "merge_matches": [
                {{
                    "record_id": "System_Record_Id",
                    "ocr_indices": [1, 2] //多个对应的 OCR Items 列表中的下标
                }}
            ],
        
            // 3. 一对多拆分 (Split / 1:N)
            "split_matches": [
                {{
                    "record_id": "System_Record_Id",
                    "ocr_index": 3       //对应 OCR Items 列表中的下标
                }},
                {{
                    "record_id": "System_Record_Id",
                    "ocr_index": 3       //对应 OCR Items 列表中的下标
                }}
            ]
        }}
    """
}