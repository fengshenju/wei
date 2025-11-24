# DrissionPage 爬虫项目

基于 DrissionPage 的标准 Python 爬虫项目骨架。

## 项目结构

```
wei/
├── src/                    # 核心代码
│   ├── __init__.py
│   ├── browser_manager.py  # 浏览器管理器
│   └── base_scraper.py     # 基础爬虫类
├── config/                 # 配置文件git init
│   ├── __init__.py
│   └── settings.py         # 项目设置
├── utils/                  # 工具类
│   ├── __init__.py
│   ├── logger.py           # 日志配置
│   └── data_handler.py     # 数据处理
├── tests/                  # 测试用例
│   ├── __init__.py
│   ├── test_browser_manager.py
│   └── test_data_handler.py
├── examples/               # 示例代码
│   ├── basic_scraping.py   # 基础爬取示例
│   └── news_scraper.py     # 新闻爬取示例
├── logs/                   # 日志文件
├── data/                   # 数据存储
├── downloads/              # 下载文件
├── docs/                   # 文档
├── main.py                 # 主程序入口
├── requirements.txt        # 依赖包
├── .env                    # 环境变量
├── .gitignore             # Git忽略文件
└── README.md              # 项目说明
```

## 安装依赖

```bash
pip install -r requirements.txt
```

## 配置说明

### 环境变量配置 (.env)

```env
# 浏览器配置
BROWSER_TYPE=chrome
HEADLESS=false
WINDOW_WIDTH=1920
WINDOW_HEIGHT=1080

# 超时设置
TIMEOUT=30
IMPLICIT_WAIT=10

# 日志配置
LOG_LEVEL=INFO

# 数据库配置
MONGODB_HOST=localhost
MONGODB_PORT=27017
MONGODB_DB=scraper_db

REDIS_HOST=localhost
REDIS_PORT=6379
REDIS_DB=0
```

### 项目设置 (config/settings.py)

主要配置项包括：
- 浏览器类型和选项
- 超时时间设置
- 日志配置
- 数据库连接
- 代理设置

## 快速开始

### 1. 运行主程序

```bash
python main.py
```

### 2. 基础爬取示例

```bash
python examples/basic_scraping.py
```

### 3. 新闻爬取示例

```bash
python examples/news_scraper.py
```

## 核心组件

### BrowserManager (src/browser_manager.py)

浏览器管理器，提供以下功能：
- 浏览器初始化和配置
- 页面导航和等待
- 元素查找和操作
- 截图功能
- 自动关闭浏览器

### BaseScraper (src/base_scraper.py)

基础爬虫类，提供以下功能：
- 抽象爬虫接口
- 数据保存和加载
- 随机延时
- 结果管理

### DataHandler (utils/data_handler.py)

数据处理器，支持多种格式：
- JSON
- CSV
- Excel (xlsx)
- Pickle

### Logger (utils/logger.py)

基于 loguru 的日志系统：
- 彩色控制台输出
- 文件日志轮转
- 多级别日志

## 自定义爬虫

继承 BaseScraper 类创建自定义爬虫：

```python
from src.base_scraper import BaseScraper
from typing import List, Dict, Any

class CustomScraper(BaseScraper):
    def scrape(self, url: str) -> List[Dict[str, Any]]:
        # 实现具体爬取逻辑
        if not self.browser.navigate_to(url):
            return []
        return self.parse_page()
    
    def parse_page(self) -> List[Dict[str, Any]]:
        # 实现页面解析逻辑
        results = []
        # ... 解析逻辑
        return results
```

## 运行测试

```bash
pytest tests/
```

## 注意事项

1. 首次运行前请确保已安装 Chrome 浏览器
2. 修改配置文件以适应具体的爬取需求
3. 遵守目标网站的 robots.txt 和使用条款
4. 添加适当的延时避免对服务器造成压力

## 扩展功能

- [ ] 支持更多浏览器类型
- [ ] 添加分布式爬取支持
- [ ] 集成更多数据存储方案
- [ ] 添加监控和报警功能
- [ ] 支持自动化测试框架