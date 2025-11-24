from abc import ABC, abstractmethod
from typing import Dict, List, Any, Optional
import time
import random

from .browser_manager import BrowserManager
from utils.logger import logger
from utils.data_handler import DataHandler


class BaseScraper(ABC):
    def __init__(self, headless: bool = None, user_agent: str = None):
        self.browser = BrowserManager(headless=headless, user_agent=user_agent)
        self.data_handler = DataHandler()
        self.results = []

    @abstractmethod
    def scrape(self, *args, **kwargs) -> List[Dict[str, Any]]:
        pass

    @abstractmethod
    def parse_page(self, *args, **kwargs) -> List[Dict[str, Any]]:
        pass

    def random_delay(self, min_seconds: float = 1, max_seconds: float = 3):
        delay = random.uniform(min_seconds, max_seconds)
        time.sleep(delay)
        logger.debug(f"随机延时: {delay:.2f}秒")

    def save_results(self, filename: str = None, format_type: str = "json"):
        if not self.results:
            logger.warning("没有数据需要保存")
            return

        if not filename:
            timestamp = int(time.time())
            filename = f"scraped_data_{timestamp}"

        try:
            self.data_handler.save_data(self.results, filename, format_type)
            logger.info(f"数据已保存: {filename}.{format_type}")
        except Exception as e:
            logger.error(f"保存数据失败: {e}")

    def load_results(self, filename: str, format_type: str = "json"):
        try:
            self.results = self.data_handler.load_data(filename, format_type)
            logger.info(f"数据已加载: {filename}")
        except Exception as e:
            logger.error(f"加载数据失败: {e}")

    def add_result(self, data: Dict[str, Any]):
        self.results.append(data)

    def get_results(self) -> List[Dict[str, Any]]:
        return self.results

    def clear_results(self):
        self.results = []
        logger.info("结果已清空")

    def close(self):
        self.browser.close()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()