import pytest
import time
from src.browser_manager import BrowserManager


class TestBrowserManager:
    def test_create_page(self):
        with BrowserManager(headless=True) as browser:
            page = browser.create_page()
            assert page is not None
            
    def test_navigate_to(self):
        with BrowserManager(headless=True) as browser:
            result = browser.navigate_to("https://www.baidu.com")
            assert result is True
            
    def test_get_text(self):
        with BrowserManager(headless=True) as browser:
            browser.navigate_to("https://www.baidu.com")
            title = browser.get_text("title")
            assert "百度" in title
            
    def test_wait_for_element(self):
        with BrowserManager(headless=True) as browser:
            browser.navigate_to("https://www.baidu.com")
            result = browser.wait_for_element("#kw", timeout=10)
            assert result is True