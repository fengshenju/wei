import json
import csv
import pandas as pd
from pathlib import Path
from typing import List, Dict, Any, Union
import pickle

from config import settings
from .logger import logger


class DataHandler:
    def __init__(self):
        self.data_dir = settings.DATA_PATH
        self._ensure_data_dir()

    def _ensure_data_dir(self):
        self.data_dir.mkdir(parents=True, exist_ok=True)

    def save_data(self, data: List[Dict[str, Any]], filename: str, format_type: str = "json"):
        filepath = self.data_dir / f"{filename}.{format_type}"
        
        try:
            if format_type.lower() == "json":
                self._save_json(data, filepath)
            elif format_type.lower() == "csv":
                self._save_csv(data, filepath)
            elif format_type.lower() == "xlsx":
                self._save_xlsx(data, filepath)
            elif format_type.lower() == "pickle":
                self._save_pickle(data, filepath)
            else:
                raise ValueError(f"不支持的文件格式: {format_type}")
            
            logger.info(f"数据已保存到: {filepath}")
        except Exception as e:
            logger.error(f"保存数据失败: {e}")
            raise

    def load_data(self, filename: str, format_type: str = "json") -> List[Dict[str, Any]]:
        if not filename.endswith(f".{format_type}"):
            filename = f"{filename}.{format_type}"
        
        filepath = self.data_dir / filename
        
        if not filepath.exists():
            raise FileNotFoundError(f"文件不存在: {filepath}")
        
        try:
            if format_type.lower() == "json":
                return self._load_json(filepath)
            elif format_type.lower() == "csv":
                return self._load_csv(filepath)
            elif format_type.lower() == "xlsx":
                return self._load_xlsx(filepath)
            elif format_type.lower() == "pickle":
                return self._load_pickle(filepath)
            else:
                raise ValueError(f"不支持的文件格式: {format_type}")
        except Exception as e:
            logger.error(f"加载数据失败: {e}")
            raise

    def _save_json(self, data: List[Dict[str, Any]], filepath: Path):
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def _load_json(self, filepath: Path) -> List[Dict[str, Any]]:
        with open(filepath, 'r', encoding='utf-8') as f:
            return json.load(f)

    def _save_csv(self, data: List[Dict[str, Any]], filepath: Path):
        if not data:
            return
        
        fieldnames = set()
        for item in data:
            fieldnames.update(item.keys())
        
        with open(filepath, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=list(fieldnames))
            writer.writeheader()
            writer.writerows(data)

    def _load_csv(self, filepath: Path) -> List[Dict[str, Any]]:
        with open(filepath, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            return list(reader)

    def _save_xlsx(self, data: List[Dict[str, Any]], filepath: Path):
        df = pd.DataFrame(data)
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Data')

    def _load_xlsx(self, filepath: Path) -> List[Dict[str, Any]]:
        df = pd.read_excel(filepath)
        return df.to_dict('records')

    def _save_pickle(self, data: List[Dict[str, Any]], filepath: Path):
        with open(filepath, 'wb') as f:
            pickle.dump(data, f)

    def _load_pickle(self, filepath: Path) -> List[Dict[str, Any]]:
        with open(filepath, 'rb') as f:
            return pickle.load(f)

    def merge_data_files(self, filenames: List[str], output_filename: str, format_type: str = "json"):
        merged_data = []
        
        for filename in filenames:
            try:
                data = self.load_data(filename, format_type)
                merged_data.extend(data)
                logger.info(f"已合并文件: {filename}")
            except Exception as e:
                logger.warning(f"合并文件 {filename} 失败: {e}")
        
        if merged_data:
            self.save_data(merged_data, output_filename, format_type)
            logger.info(f"合并完成，共 {len(merged_data)} 条数据")
        else:
            logger.warning("没有数据需要合并")

    def get_data_summary(self, filename: str, format_type: str = "json") -> Dict[str, Any]:
        try:
            data = self.load_data(filename, format_type)
            if not data:
                return {"count": 0}
            
            summary = {
                "count": len(data),
                "fields": list(data[0].keys()) if data else [],
                "sample": data[0] if data else None
            }
            
            return summary
        except Exception as e:
            logger.error(f"获取数据摘要失败: {e}")
            return {"error": str(e)}
    
    def _get_timestamp(self) -> str:
        from datetime import datetime
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")