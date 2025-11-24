import pytest
import tempfile
import shutil
from pathlib import Path
from utils.data_handler import DataHandler


class TestDataHandler:
    def setup_method(self):
        self.temp_dir = Path(tempfile.mkdtemp())
        self.data_handler = DataHandler()
        self.data_handler.data_dir = self.temp_dir
        
        self.test_data = [
            {"name": "张三", "age": 25, "city": "北京"},
            {"name": "李四", "age": 30, "city": "上海"},
            {"name": "王五", "age": 35, "city": "广州"}
        ]
    
    def teardown_method(self):
        shutil.rmtree(self.temp_dir)
    
    def test_save_and_load_json(self):
        filename = "test_data"
        self.data_handler.save_data(self.test_data, filename, "json")
        
        loaded_data = self.data_handler.load_data(filename, "json")
        assert loaded_data == self.test_data
    
    def test_save_and_load_csv(self):
        filename = "test_data"
        self.data_handler.save_data(self.test_data, filename, "csv")
        
        loaded_data = self.data_handler.load_data(filename, "csv")
        assert len(loaded_data) == len(self.test_data)
        assert loaded_data[0]["name"] == "张三"
    
    def test_get_data_summary(self):
        filename = "test_data"
        self.data_handler.save_data(self.test_data, filename, "json")
        
        summary = self.data_handler.get_data_summary(filename, "json")
        assert summary["count"] == 3
        assert "name" in summary["fields"]
        assert "age" in summary["fields"]