import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.base_scraper import BaseScraper
from utils.logger import logger
from typing import List, Dict, Any


class NewsScraper(BaseScraper):
    def __init__(self, headless: bool = True):
        super().__init__(headless=headless)
        
    def scrape(self, keywords: List[str]) -> List[Dict[str, Any]]:
        all_results = []
        
        for keyword in keywords:
            logger.info(f"搜索关键词: {keyword}")
            results = self.search_news(keyword)
            all_results.extend(results)
            self.random_delay(2, 4)
        
        self.results = all_results
        return all_results
    
    def search_news(self, keyword: str) -> List[Dict[str, Any]]:
        search_url = f"https://www.baidu.com/s?tn=news&word={keyword}"
        
        if not self.browser.navigate_to(search_url):
            return []
        
        return self.parse_page(keyword)
    
    def parse_page(self, keyword: str = "") -> List[Dict[str, Any]]:
        results = []
        page = self.browser.get_page()
        
        try:
            news_items = page.eles(".result")
            
            for item in news_items[:10]:
                try:
                    title_ele = item.ele("h3 a")
                    title = title_ele.text if title_ele else ""
                    
                    link = title_ele.attr("href") if title_ele else ""
                    
                    snippet_ele = item.ele(".c-abstract")
                    snippet = snippet_ele.text if snippet_ele else ""
                    
                    source_ele = item.ele(".c-color-gray")
                    source = source_ele.text if source_ele else ""
                    
                    if title:
                        news_data = {
                            "keyword": keyword,
                            "title": title.strip(),
                            "link": link,
                            "snippet": snippet.strip(),
                            "source": source.strip(),
                            "scraped_at": self.data_handler._get_timestamp() if hasattr(self.data_handler, '_get_timestamp') else ""
                        }
                        results.append(news_data)
                        
                except Exception as e:
                    logger.warning(f"解析新闻项失败: {e}")
                    continue
            
            logger.info(f"关键词 '{keyword}' 解析完成，获取 {len(results)} 条新闻")
            
        except Exception as e:
            logger.error(f"解析页面失败: {e}")
        
        return results


def main():
    keywords = ["人工智能", "Python编程", "网络爬虫"]
    
    with NewsScraper(headless=False) as scraper:
        try:
            results = scraper.scrape(keywords)
            
            if results:
                scraper.save_results("news_data", "json")
                scraper.save_results("news_data", "xlsx")
                
                logger.info(f"新闻爬取完成，共获取 {len(results)} 条数据")
                
                for result in results[:3]:
                    print(f"标题: {result['title']}")
                    print(f"来源: {result['source']}")
                    print(f"摘要: {result['snippet'][:100]}...")
                    print("-" * 80)
            else:
                logger.warning("未获取到任何新闻数据")
                
        except Exception as e:
            logger.error(f"新闻爬取过程出错: {e}")


if __name__ == "__main__":
    main()