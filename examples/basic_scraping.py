import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.browser_manager import BrowserManager
from utils.logger import logger


def basic_example():
    logger.info("基础爬取示例开始")
    
    with BrowserManager(headless=False) as browser:
        browser.navigate_to("https://www.baidu.com")
        
        browser.input_text("#kw", "DrissionPage")
        
        browser.click_element("#su")
        
        browser.wait_for_element(".result", timeout=10)
        
        results = browser.get_page().eles(".result")
        
        for i, result in enumerate(results[:5]):
            try:
                title_ele = result.ele("h3 a")
                title = title_ele.text if title_ele else "无标题"
                
                url_ele = result.ele("h3 a")
                url = url_ele.attr("href") if url_ele else "无链接"
                
                print(f"结果 {i+1}:")
                print(f"标题: {title}")
                print(f"链接: {url}")
                print("-" * 50)
                
            except Exception as e:
                logger.warning(f"解析第 {i+1} 个结果失败: {e}")


if __name__ == "__main__":
    basic_example()