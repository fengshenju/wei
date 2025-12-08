import os
import json
import pathlib
import time
import pandas as pd

class DataManager:
    def __init__(self, storage_path_str):
        """
        初始化数据管理器
        :param storage_path_str: 数据存储的完整路径 (字符串)
        """
        # 直接根据配置路径创建 Path 对象
        # 如果配置的是相对路径，会相对于运行目录；如果是绝对路径，就用绝对路径
        self.data_dir = pathlib.Path(storage_path_str)

        # 如果目录不存在，自动创建 (parents=True 允许创建多级目录)
        if not self.data_dir.exists():
            try:
                self.data_dir.mkdir(parents=True, exist_ok=True)
                print(f">>> 数据存储目录已创建: {self.data_dir}")
            except Exception as e:
                print(f"!!! 无法创建数据目录，请检查路径权限: {e}")
                raise e
        else:
            print(f">>> 连接到数据存储目录: {self.data_dir}")


    def get_json_path(self, image_filename):
        """根据图片文件名，生成对应的 json 文件路径"""
        base_name = os.path.splitext(os.path.basename(image_filename))[0]
        return self.data_dir / f"{base_name}.json"

    def is_processed(self, image_filename):
        """检查该图片是否已经处理过"""
        json_path = self.get_json_path(image_filename)
        return json_path.exists()

    def save_data(self, image_filename, data):
        """保存数据"""
        json_path = self.get_json_path(image_filename)
        try:
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            print(f">>> 数据已保存: {json_path.name}")
            return True
        except Exception as e:
            print(f"!!! 数据保存失败: {e}")
            return False

    def load_data(self, image_filename):
        """读取数据"""
        json_path = self.get_json_path(image_filename)
        if json_path.exists():
            with open(json_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        return None

def load_style_db_from_excel(file_path, column_name='款式编号'):
        """
        读取 Excel 文件中的指定列，加载为 set 集合
        :param file_path: Excel 文件路径
        :param column_name: 包含款号的列名 (默认: 款式编号)
        :return: 包含所有款号的 set 集合
        """
        if not os.path.exists(file_path):
            print(f"!!! 错误: 款号库文件未找到: {file_path}")
            return set()

        print(f">>> 正在读取款号库 Excel: {os.path.basename(file_path)} ...")

        try:
            # 使用 pandas 读取 excel
            # engine='openpyxl' 专门用于读取 .xlsx 文件
            df = pd.read_excel(file_path, engine='openpyxl')

            # 检查列名是否存在
            if column_name not in df.columns:
                print(f"!!! 错误: Excel中未找到列名 '{column_name}'，现有列: {list(df.columns)}")
                return set()

            # 1. 提取该列
            # 2. dropna() 去除空行
            # 3. astype(str) 强制转为字符串 (防止纯数字款号被当成数字)
            # 4. str.strip() 去除首尾空格
            clean_series = df[column_name].dropna().astype(str).str.strip()

            # 转换为集合 (自动去重)
            style_set = set(clean_series)

            print(f">>> 款号库加载成功! 共加载 {len(style_set)} 个唯一款号。")
            # 打印前5个看看样子，确保没读错
            print(f"    示例数据: {list(style_set)[:5]}")

            return style_set

        except Exception as e:
            print(f"!!! 读取 Excel 失败: {e}")
            return set()


class StyleDBCache:
    """
    款号库缓存管理器
    功能：检查 Excel 文件修改时间，决定是否从缓存读取还是重新解析
    """
    
    def __init__(self, excel_path, cache_dir=None):
        """
        初始化缓存管理器
        :param excel_path: Excel 文件路径
        :param cache_dir: 缓存目录，默认与 Excel 同目录
        """
        self.excel_path = excel_path
        
        # 如果未指定缓存目录，使用 Excel 文件同目录
        if cache_dir is None:
            cache_dir = os.path.dirname(excel_path)
        
        # 生成缓存文件路径
        excel_filename = os.path.splitext(os.path.basename(excel_path))[0]
        self.cache_path = os.path.join(cache_dir, f"{excel_filename}_cache.json")
        
    def _get_file_mtime(self, file_path):
        """获取文件修改时间"""
        try:
            return os.path.getmtime(file_path)
        except OSError:
            return 0
            
    def _is_cache_valid(self):
        """检查缓存是否有效（缓存文件存在且比 Excel 新）"""
        if not os.path.exists(self.cache_path):
            return False
            
        excel_mtime = self._get_file_mtime(self.excel_path)
        cache_mtime = self._get_file_mtime(self.cache_path)
        
        return cache_mtime > excel_mtime
        
    def _save_cache(self, style_set):
        """保存缓存到 JSON 文件"""
        try:
            cache_data = {
                'style_list': list(style_set),
                'cache_time': time.time(),
                'source_excel': self.excel_path
            }
            with open(self.cache_path, 'w', encoding='utf-8') as f:
                json.dump(cache_data, f, ensure_ascii=False, indent=2)
            print(f">>> 缓存已保存: {os.path.basename(self.cache_path)}")
            return True
        except Exception as e:
            print(f"!!! 保存缓存失败: {e}")
            return False
            
    def _load_cache(self):
        """从缓存文件读取数据"""
        try:
            with open(self.cache_path, 'r', encoding='utf-8') as f:
                cache_data = json.load(f)
            style_set = set(cache_data['style_list'])
            print(f">>> 从缓存加载款号库: {len(style_set)} 个款号")
            return style_set
        except Exception as e:
            print(f"!!! 读取缓存失败: {e}")
            return None
            
    def get_style_db(self, column_name='款式编号'):
        """
        获取款号库（优先从缓存，缓存无效时从 Excel 重新加载）
        :param column_name: Excel 中存储款号的列名
        :return: 包含所有款号的 set 集合
        """
        # 检查缓存是否有效
        if self._is_cache_valid():
            print(f">>> 检测到有效缓存，快速加载中...")
            cached_data = self._load_cache()
            if cached_data is not None:
                return cached_data
                
        # 缓存无效，从 Excel 重新加载
        print(f">>> 缓存无效或不存在，从 Excel 重新加载...")
        style_set = load_style_db_from_excel(self.excel_path, column_name)
        
        # 保存新的缓存
        if style_set:
            self._save_cache(style_set)
            
        return style_set


def load_style_db_with_cache(file_path, column_name='款式编号', cache_dir=None):
    """
    使用缓存机制加载款号库（推荐使用此函数替代 load_style_db_from_excel）
    
    :param file_path: Excel 文件路径
    :param column_name: 包含款号的列名 (默认: 款式编号)
    :param cache_dir: 缓存目录，默认与 Excel 同目录
    :return: 包含所有款号的 set 集合
    """
    cache_manager = StyleDBCache(file_path, cache_dir)
    return cache_manager.get_style_db(column_name)


def load_supplier_db_from_excel(file_path):
    """
    读取 Excel 文件的第一列，加载月结供应商目录为 set 集合
    :param file_path: Excel 文件路径
    :return: 包含所有供应商名称的 set 集合
    """
    if not os.path.exists(file_path):
        print(f"!!! 错误: 供应商目录文件未找到: {file_path}")
        return set()

    print(f">>> 正在读取供应商目录 Excel: {os.path.basename(file_path)} ...")

    try:
        # 使用 pandas 读取 excel
        df = pd.read_excel(file_path, engine='openpyxl')

        if df.empty:
            print(f"!!! 错误: Excel文件为空")
            return set()

        # 读取第一列（索引为0），自动获取列名
        first_column = df.iloc[:, 0]
        
        # 去除空值，转为字符串，去除首尾空格
        clean_series = first_column.dropna().astype(str).str.strip()
        
        # 转换为集合 (自动去重)
        supplier_set = set(clean_series)

        print(f">>> 供应商目录加载成功! 共加载 {len(supplier_set)} 个唯一供应商。")
        # 打印前5个看看样子，确保没读错
        print(f"    示例数据: {list(supplier_set)[:5]}")

        return supplier_set

    except Exception as e:
        print(f"!!! 读取 Excel 失败: {e}")
        return set()


def load_material_deduction_db_from_excel(file_path, name_column='物料名称', amount_column='扣减金额'):
    """
    读取 Excel 文件中的物料名称和扣减金额两列，加载物料扣减列表为字典
    :param file_path: Excel 文件路径
    :param name_column: 包含物料名称的列名 (默认: 物料名称)
    :param amount_column: 包含扣减金额的列名 (默认: 扣减金额)
    :return: 包含物料名称和扣减金额的字典 {物料名称: 扣减金额}
    """
    if not os.path.exists(file_path):
        print(f"!!! 错误: 物料扣减列表文件未找到: {file_path}")
        return {}

    print(f">>> 正在读取物料扣减列表 Excel: {os.path.basename(file_path)} ...")

    try:
        # 使用 pandas 读取 excel
        df = pd.read_excel(file_path, engine='openpyxl')

        if df.empty:
            print(f"!!! 错误: Excel文件为空")
            return {}

        # 检查列名是否存在
        missing_columns = []
        if name_column not in df.columns:
            missing_columns.append(name_column)
        if amount_column not in df.columns:
            missing_columns.append(amount_column)
            
        if missing_columns:
            print(f"!!! 错误: Excel中未找到列名 {missing_columns}，现有列: {list(df.columns)}")
            return {}

        # 提取两列数据
        name_series = df[name_column].dropna().astype(str).str.strip()
        amount_series = df[amount_column].dropna()
        
        # 确保两列长度一致，取最小长度
        min_length = min(len(name_series), len(amount_series))
        name_series = name_series.iloc[:min_length]
        amount_series = amount_series.iloc[:min_length]
        
        # 转换为字典 {物料名称: 扣减金额}
        material_dict = {}
        for name, amount in zip(name_series, amount_series):
            if name and str(name).strip():  # 确保物料名称不为空
                try:
                    # 尝试转换金额为数字
                    amount_value = float(amount) if pd.notna(amount) else 0.0
                    material_dict[str(name).strip()] = amount_value
                except (ValueError, TypeError):
                    print(f"!!! 警告: 物料 '{name}' 的扣减金额 '{amount}' 无法转换为数字，设为0")
                    material_dict[str(name).strip()] = 0.0

        print(f">>> 物料扣减列表加载成功! 共加载 {len(material_dict)} 个物料。")
        # 打印前5个看看样子，确保没读错
        sample_items = list(material_dict.items())[:5]
        print(f"    示例数据: {sample_items}")

        return material_dict

    except Exception as e:
        print(f"!!! 读取物料扣减列表 Excel 失败: {e}")
        return {}


class SupplierDBCache:
    """
    供应商目录缓存管理器
    功能：检查 Excel 文件修改时间，决定是否从缓存读取还是重新解析
    """
    
    def __init__(self, excel_path, cache_dir=None):
        """
        初始化缓存管理器
        :param excel_path: Excel 文件路径
        :param cache_dir: 缓存目录，默认与 Excel 同目录
        """
        self.excel_path = excel_path
        
        # 如果未指定缓存目录，使用 Excel 文件同目录
        if cache_dir is None:
            cache_dir = os.path.dirname(excel_path)
        
        # 生成缓存文件路径
        excel_filename = os.path.splitext(os.path.basename(excel_path))[0]
        self.cache_path = os.path.join(cache_dir, f"{excel_filename}_supplier_cache.json")
        
    def _get_file_mtime(self, file_path):
        """获取文件修改时间"""
        try:
            return os.path.getmtime(file_path)
        except OSError:
            return 0
            
    def _is_cache_valid(self):
        """检查缓存是否有效（缓存文件存在且比 Excel 新）"""
        if not os.path.exists(self.cache_path):
            return False
            
        excel_mtime = self._get_file_mtime(self.excel_path)
        cache_mtime = self._get_file_mtime(self.cache_path)
        
        return cache_mtime > excel_mtime
        
    def _save_cache(self, supplier_set):
        """保存缓存到 JSON 文件"""
        try:
            cache_data = {
                'supplier_list': list(supplier_set),
                'cache_time': time.time(),
                'source_excel': self.excel_path
            }
            with open(self.cache_path, 'w', encoding='utf-8') as f:
                json.dump(cache_data, f, ensure_ascii=False, indent=2)
            print(f">>> 供应商缓存已保存: {os.path.basename(self.cache_path)}")
            return True
        except Exception as e:
            print(f"!!! 保存供应商缓存失败: {e}")
            return False
            
    def _load_cache(self):
        """从缓存文件读取数据"""
        try:
            with open(self.cache_path, 'r', encoding='utf-8') as f:
                cache_data = json.load(f)
            supplier_set = set(cache_data['supplier_list'])
            print(f">>> 从缓存加载供应商目录: {len(supplier_set)} 个供应商")
            return supplier_set
        except Exception as e:
            print(f"!!! 读取供应商缓存失败: {e}")
            return None
            
    def get_supplier_db(self):
        """
        获取供应商目录（优先从缓存，缓存无效时从 Excel 重新加载）
        :return: 包含所有供应商名称的 set 集合
        """
        # 检查缓存是否有效
        if self._is_cache_valid():
            print(f">>> 检测到有效供应商缓存，快速加载中...")
            cached_data = self._load_cache()
            if cached_data is not None:
                return cached_data
                
        # 缓存无效，从 Excel 重新加载
        print(f">>> 供应商缓存无效或不存在，从 Excel 重新加载...")
        supplier_set = load_supplier_db_from_excel(self.excel_path)
        
        # 保存新的缓存
        if supplier_set:
            self._save_cache(supplier_set)
            
        return supplier_set


def load_supplier_db_with_cache(file_path, cache_dir=None):
    """
    使用缓存机制加载供应商目录（推荐使用此函数替代 load_supplier_db_from_excel）
    
    :param file_path: Excel 文件路径
    :param cache_dir: 缓存目录，默认与 Excel 同目录
    :return: 包含所有供应商名称的 set 集合
    """
    cache_manager = SupplierDBCache(file_path, cache_dir)
    return cache_manager.get_supplier_db()


class MaterialDeductionDBCache:
    """
    物料扣减列表缓存管理器
    功能：检查 Excel 文件修改时间，决定是否从缓存读取还是重新解析
    """
    
    def __init__(self, excel_path, cache_dir=None):
        """
        初始化缓存管理器
        :param excel_path: Excel 文件路径
        :param cache_dir: 缓存目录，默认与 Excel 同目录
        """
        self.excel_path = excel_path
        
        # 如果未指定缓存目录，使用 Excel 文件同目录
        if cache_dir is None:
            cache_dir = os.path.dirname(excel_path)
        
        # 生成缓存文件路径
        excel_filename = os.path.splitext(os.path.basename(excel_path))[0]
        self.cache_path = os.path.join(cache_dir, f"{excel_filename}_material_deduction_cache.json")
        
    def _get_file_mtime(self, file_path):
        """获取文件修改时间"""
        try:
            return os.path.getmtime(file_path)
        except OSError:
            return 0
            
    def _is_cache_valid(self):
        """检查缓存是否有效（缓存文件存在且比 Excel 新）"""
        if not os.path.exists(self.cache_path):
            return False
            
        excel_mtime = self._get_file_mtime(self.excel_path)
        cache_mtime = self._get_file_mtime(self.cache_path)
        
        return cache_mtime > excel_mtime
        
    def _save_cache(self, material_dict):
        """保存缓存到 JSON 文件"""
        try:
            cache_data = {
                'material_dict': material_dict,
                'cache_time': time.time(),
                'source_excel': self.excel_path
            }
            with open(self.cache_path, 'w', encoding='utf-8') as f:
                json.dump(cache_data, f, ensure_ascii=False, indent=2)
            print(f">>> 物料扣减缓存已保存: {os.path.basename(self.cache_path)}")
            return True
        except Exception as e:
            print(f"!!! 保存物料扣减缓存失败: {e}")
            return False
            
    def _load_cache(self):
        """从缓存文件读取数据"""
        try:
            with open(self.cache_path, 'r', encoding='utf-8') as f:
                cache_data = json.load(f)
            material_dict = cache_data['material_dict']
            print(f">>> 从缓存加载物料扣减列表: {len(material_dict)} 个物料")
            return material_dict
        except Exception as e:
            print(f"!!! 读取物料扣减缓存失败: {e}")
            return None
            
    def get_material_deduction_db(self, name_column='物料名称', amount_column='扣减金额'):
        """
        获取物料扣减列表（优先从缓存，缓存无效时从 Excel 重新加载）
        :param name_column: Excel 中存储物料名称的列名
        :param amount_column: Excel 中存储扣减金额的列名
        :return: 包含物料名称和扣减金额的字典 {物料名称: 扣减金额}
        """
        # 检查缓存是否有效
        if self._is_cache_valid():
            print(f">>> 检测到有效物料扣减缓存，快速加载中...")
            cached_data = self._load_cache()
            if cached_data is not None:
                return cached_data
                
        # 缓存无效，从 Excel 重新加载
        print(f">>> 物料扣减缓存无效或不存在，从 Excel 重新加载...")
        material_dict = load_material_deduction_db_from_excel(self.excel_path, name_column, amount_column)
        
        # 保存新的缓存
        if material_dict:
            self._save_cache(material_dict)
            
        return material_dict


def load_material_deduction_db_with_cache(file_path, name_column='物料名称', amount_column='扣减金额', cache_dir=None):
    """
    使用缓存机制加载物料扣减列表（推荐使用此函数替代 load_material_deduction_db_from_excel）
    
    :param file_path: Excel 文件路径
    :param name_column: 包含物料名称的列名 (默认: 物料名称)
    :param amount_column: 包含扣减金额的列名 (默认: 扣减金额)
    :param cache_dir: 缓存目录，默认与 Excel 同目录
    :return: 包含物料名称和扣减金额的字典 {物料名称: 扣减金额}
    """
    cache_manager = MaterialDeductionDBCache(file_path, cache_dir)
    return cache_manager.get_material_deduction_db(name_column, amount_column)


def apply_material_deduction(parsed_data, material_deduction_db, deduction_suppliers):
    """
    对特定供应商的商品应用物料扣减价格调整
    
    :param parsed_data: 图片识别解析后的数据字典
    :param material_deduction_db: 物料扣减数据库字典 {物料名称: 扣减金额}
    :param deduction_suppliers: 支持扣减的供应商列表
    :return: 调整后的parsed_data，如果有调整返回True，否则返回False
    """
    if not parsed_data or not material_deduction_db or not deduction_suppliers:
        return parsed_data, False
    
    # 检查供应商是否在扣减列表中
    supplier_name = parsed_data.get('supplier_name', '')
    if not any(supplier in supplier_name for supplier in deduction_suppliers):
        return parsed_data, False
    
    print(f">>> 检测到扣减供应商: {supplier_name}，开始进行价格调整...")
    
    # 处理商品列表
    items = parsed_data.get('items', [])
    total_adjusted = 0
    
    for item in items:
        raw_style_text = item.get('raw_style_text', '')
        original_price = item.get('price', 0)
        
        if not raw_style_text or not original_price:
            continue
        
        # 精确前缀匹配
        matched_material = None
        deduction_amount = 0
        
        # 先尝试 raw_style_text 匹配
        for material_name, amount in material_deduction_db.items():
            if raw_style_text.startswith(material_name):
                matched_material = material_name
                deduction_amount = amount
                break
        
        # 如果没匹配上，再用 product_description 匹配
        if not matched_material:
            product_description = item.get('product_description', '')
            if product_description:
                for material_name, amount in material_deduction_db.items():
                    if material_name in product_description:
                        matched_material = material_name
                        deduction_amount = amount
                        break
        
        # 如果还没匹配上，尝试组合匹配
        if not matched_material:
            product_description = item.get('product_description', '')
            if raw_style_text and product_description:
                # 组合1: raw_style_text + product_description
                combination1 = raw_style_text + product_description
                # 组合2: product_description + raw_style_text  
                combination2 = product_description + raw_style_text
                
                for material_name, amount in material_deduction_db.items():
                    if material_name in combination1:
                        matched_material = material_name
                        deduction_amount = amount
                        item['match_method'] = 'combination1'
                        item['match_source'] = combination1
                        break
                    elif material_name in combination2:
                        matched_material = material_name
                        deduction_amount = amount
                        item['match_method'] = 'combination2'
                        item['match_source'] = combination2
                        break
        
        if matched_material:
            # 计算调整后价格
            try:
                original_price_float = float(original_price)
                deduction_amount_float = float(deduction_amount)
                new_price = original_price_float - deduction_amount_float
                
                # 确保价格不为负数
                new_price = max(0, new_price)
                
                # 更新价格
                item['price'] = new_price
                item['original_price'] = original_price_float  # 保存原价
                item['deduction_amount'] = deduction_amount_float  # 保存扣减金额
                item['matched_material'] = matched_material  # 保存匹配的物料
                
                total_adjusted += 1
                
                # 判断匹配方式和来源
                if item.get('match_method') in ['combination1', 'combination2']:
                    match_method = item.get('match_method')
                    match_source = item.get('match_source', '')
                elif raw_style_text.startswith(matched_material):
                    match_method = "raw_style_text"
                    match_source = raw_style_text
                else:
                    match_method = "product_description"
                    match_source = item.get('product_description', '')
                
                print(f"    ✅ 物料匹配成功 [{match_method}]: '{match_source}' → '{matched_material}'")
                print(f"       价格调整: {original_price_float} - {deduction_amount_float} = {new_price} (扣减前: {original_price_float}, 扣减后: {new_price})")
                
            except (ValueError, TypeError) as e:
                print(f"    ❌ 价格调整失败: {e}")
                continue
        else:
            print(f"    ⚠️ 未匹配物料: '{raw_style_text}'")
    
    if total_adjusted > 0:
        print(f">>> 价格调整完成，共调整 {total_adjusted} 个商品")
        return parsed_data, True
    else:
        print(">>> 未找到匹配的物料，无价格调整")
        return parsed_data, False
