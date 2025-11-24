"""
项目配置
"""

from pydantic_settings import BaseSettings, SettingsConfigDict
from typing import Optional
from pathlib import Path


class Settings(BaseSettings):
    """项目设置"""
    
    # 基础配置
    PROJECT_NAME: str = "DrissionPage爬虫测试项目"
    DEBUG: bool = True
    
    # 浏览器配置
    BROWSER_TYPE: str = "chrome"
    HEADLESS: bool = False
    HEADLESS_MODE: bool = False
    DISABLE_IMAGES: bool = True
    WINDOW_WIDTH: int = 1920
    WINDOW_HEIGHT: int = 1080
    TIMEOUT: float = 30.0
    PAGE_TIMEOUT: float = 30.0
    IMPLICIT_WAIT: float = 10.0
    
    # 请求配置
    REQUEST_TIMEOUT: float = 30.0
    REQUEST_RETRY: int = 3
    REQUEST_DELAY_MIN: float = 1.0
    REQUEST_DELAY_MAX: float = 3.0
    
    # 并发配置
    MAX_WORKERS: int = 4
    MAX_CONCURRENT: int = 10
    
    # 数据存储配置
    DATA_DIR: str = "data"
    LOGS_DIR: str = "logs"
    
    # 日志配置
    LOG_LEVEL: str = "INFO"
    LOG_ROTATION: str = "100 MB"
    LOG_RETENTION: str = "30 days"
    
    # 代理配置（可选）
    USE_PROXY: bool = False
    PROXY_ENABLED: bool = False
    PROXY_TYPE: str = "http"
    PROXY_HOST: Optional[str] = None
    PROXY_PORT: Optional[int] = None
    PROXY_USER: Optional[str] = None
    PROXY_PASS: Optional[str] = None
    PROXY_USERNAME: Optional[str] = None
    PROXY_PASSWORD: Optional[str] = None
    
    # 数据库配置（可选）
    DATABASE_URL: Optional[str] = None
    
    # MongoDB 配置
    MONGODB_HOST: str = "localhost"
    MONGODB_PORT: int = 27017
    MONGODB_DB: str = "scraper_db"
    
    # Redis 配置
    REDIS_HOST: str = "localhost"
    REDIS_PORT: int = 6379
    REDIS_DB: int = 0
    
    # 邮件通知配置（可选）
    EMAIL_ENABLED: bool = False
    EMAIL_HOST: Optional[str] = None
    EMAIL_PORT: Optional[int] = None
    EMAIL_USERNAME: Optional[str] = None
    EMAIL_PASSWORD: Optional[str] = None
    EMAIL_TO: Optional[str] = None
    
    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        case_sensitive=True,
        extra="ignore"
    )


# 创建全局设置实例
settings = Settings()


# 确保必要的目录存在
def init_directories():
    """初始化必要的目录"""
    data_dir = Path(settings.DATA_DIR)
    logs_dir = Path(settings.LOGS_DIR)
    
    data_dir.mkdir(exist_ok=True)
    logs_dir.mkdir(exist_ok=True)


# 自动初始化目录
init_directories()