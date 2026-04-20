"""
日志模块
"""
import os
import logging
from logging.handlers import RotatingFileHandler
from typing import Optional


def setup_logger(
    name: str = 'email_scheduler',
    log_level: str = 'INFO',
    log_file: Optional[str] = None,
    log_max_size: str = '10MB',
    log_backup_count: int = 5,
    console_output: bool = True
) -> logging.Logger:
    """
    设置日志记录器
    
    Args:
        name: 日志记录器名称
        log_level: 日志级别
        log_file: 日志文件路径
        log_max_size: 日志文件最大大小
        log_backup_count: 日志文件备份数量
        console_output: 是否输出到控制台
    
    Returns:
        配置好的日志记录器
    """
    logger = logging.getLogger(name)
    logger.setLevel(getattr(logging, log_level.upper()))
    
    logger.handlers.clear()
    
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    if log_file:
        log_dir = os.path.dirname(log_file)
        if log_dir and not os.path.exists(log_dir):
            os.makedirs(log_dir, exist_ok=True)
        
        max_bytes = _parse_size(log_max_size)
        file_handler = RotatingFileHandler(
            log_file,
            maxBytes=max_bytes,
            backupCount=log_backup_count,
            encoding='utf-8'
        )
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    
    if console_output:
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)
    
    return logger


def _parse_size(size_str: str) -> int:
    """
    解析大小字符串为字节数
    
    Args:
        size_str: 大小字符串，如 '10MB', '1GB', '10M', '1G'
    
    Returns:
        字节数
    """
    size_str = size_str.upper().strip()
    
    units = {
        'TB': 1024 ** 4,
        'GB': 1024 ** 3,
        'MB': 1024 ** 2,
        'KB': 1024,
        'B': 1,
        'T': 1024 ** 4,
        'G': 1024 ** 3,
        'M': 1024 ** 2,
        'K': 1024,
    }
    
    for unit, multiplier in units.items():
        if size_str.endswith(unit):
            number = float(size_str[:-len(unit)].strip())
            return int(number * multiplier)
    
    return int(size_str)


def get_logger(name: str = 'email_scheduler') -> logging.Logger:
    """
    获取日志记录器
    
    Args:
        name: 日志记录器名称
    
    Returns:
        日志记录器
    """
    return logging.getLogger(name)
