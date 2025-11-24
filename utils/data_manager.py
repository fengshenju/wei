import os
import json
import pathlib


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