from DrissionPage import ChromiumPage, ChromiumOptions
from typing import Optional, Dict, Any, List
import time
import os
from pathlib import Path

from config import settings
from utils.logger import logger


class BrowserManager:
    def __init__(self, headless: bool = None, user_agent: str = None):
        self.page: Optional[ChromiumPage] = None
        self.headless = headless if headless is not None else settings.HEADLESS
        self.user_agent = user_agent or settings.USER_AGENT
        self.options = self._setup_options()

    def _setup_options(self) -> ChromiumOptions:
        options = ChromiumOptions()
        
        if self.headless:
            options.headless()
        
        options.set_user_agent(self.user_agent)
        
        for option in settings.CHROME_OPTIONS:
            options.add_argument(option)
        
        if settings.DOWNLOAD_PATH:
            os.makedirs(settings.DOWNLOAD_PATH, exist_ok=True)
            options.set_download_path(str(settings.DOWNLOAD_PATH))
        
        window_size = settings.WINDOW_SIZE
        options.set_window_size(window_size[0], window_size[1])
        
        if settings.PROXY_CONFIG.get("enabled"):
            proxy_config = settings.PROXY_CONFIG
            proxy_str = f"{proxy_config['proxy_host']}:{proxy_config['proxy_port']}"
            if proxy_config.get('proxy_user'):
                proxy_str = f"{proxy_config['proxy_user']}:{proxy_config['proxy_pass']}@{proxy_str}"
            options.set_proxy(proxy_str)
        
        return options

    def create_page(self) -> ChromiumPage:
        try:
            self.page = ChromiumPage(addr_or_opts=self.options)
            self.page.set.timeouts(
                base=settings.TIMEOUT,
                page_load=settings.TIMEOUT,
                script=settings.TIMEOUT
            )
            self.page.set.timeouts(implicit=settings.IMPLICIT_WAIT)
            
            logger.info("浏览器页面创建成功")
            return self.page
        except Exception as e:
            logger.error(f"创建浏览器页面失败: {e}")
            raise

    def get_page(self) -> ChromiumPage:
        if self.page is None:
            self.create_page()
        return self.page

    def navigate_to(self, url: str, wait_time: int = 2) -> bool:
        try:
            page = self.get_page()
            page.get(url)
            time.sleep(wait_time)
            logger.info(f"成功导航到: {url}")
            return True
        except Exception as e:
            logger.error(f"导航到 {url} 失败: {e}")
            return False

    def wait_for_element(self, locator: str, timeout: int = None) -> bool:
        try:
            timeout = timeout or settings.TIMEOUT
            page = self.get_page()
            element = page.wait.ele_displayed(locator, timeout=timeout)
            return element is not None
        except Exception as e:
            logger.warning(f"等待元素 {locator} 超时: {e}")
            return False

    def click_element(self, locator: str, wait_time: int = 1) -> bool:
        try:
            page = self.get_page()
            element = page.ele(locator)
            if element:
                element.click()
                time.sleep(wait_time)
                logger.debug(f"成功点击元素: {locator}")
                return True
            else:
                logger.warning(f"未找到元素: {locator}")
                return False
        except Exception as e:
            logger.error(f"点击元素 {locator} 失败: {e}")
            return False

    def input_text(self, locator: str, text: str, clear: bool = True) -> bool:
        try:
            page = self.get_page()
            element = page.ele(locator)
            if element:
                if clear:
                    element.clear()
                element.input(text)
                logger.debug(f"成功输入文本到 {locator}")
                return True
            else:
                logger.warning(f"未找到输入框: {locator}")
                return False
        except Exception as e:
            logger.error(f"输入文本到 {locator} 失败: {e}")
            return False

    def get_text(self, locator: str) -> str:
        try:
            page = self.get_page()
            element = page.ele(locator)
            if element:
                return element.text
            else:
                logger.warning(f"未找到元素: {locator}")
                return ""
        except Exception as e:
            logger.error(f"获取元素 {locator} 文本失败: {e}")
            return ""

    def get_attribute(self, locator: str, attribute: str) -> str:
        try:
            page = self.get_page()
            element = page.ele(locator)
            if element:
                return element.attr(attribute) or ""
            else:
                logger.warning(f"未找到元素: {locator}")
                return ""
        except Exception as e:
            logger.error(f"获取元素 {locator} 属性 {attribute} 失败: {e}")
            return ""

    def scroll_to_bottom(self, pause_time: int = 2):
        try:
            page = self.get_page()
            page.scroll.to_bottom()
            time.sleep(pause_time)
            logger.debug("页面滚动到底部")
        except Exception as e:
            logger.error(f"滚动到底部失败: {e}")

    def take_screenshot(self, filename: str = None) -> str:
        try:
            page = self.get_page()
            if not filename:
                timestamp = int(time.time())
                filename = f"screenshot_{timestamp}.png"
            
            screenshot_path = settings.BASE_DIR / "logs" / filename
            os.makedirs(screenshot_path.parent, exist_ok=True)
            
            page.get_screenshot(str(screenshot_path))
            logger.info(f"截图保存至: {screenshot_path}")
            return str(screenshot_path)
        except Exception as e:
            logger.error(f"截图失败: {e}")
            return ""

    def close(self):
        try:
            if self.page:
                self.page.quit()
                self.page = None
                logger.info("浏览器已关闭")
        except Exception as e:
            logger.error(f"关闭浏览器失败: {e}")

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()