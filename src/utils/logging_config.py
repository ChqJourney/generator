"""
日志配置模块
为整个项目提供统一的日志配置和获取logger的接口
"""
import logging
import sys
from typing import Optional


# 默认日志格式
DEFAULT_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
# 简洁日志格式（用于控制台输出）
SIMPLE_FORMAT = '%(levelname)s: %(message)s'


def setup_logging(
    level: int = logging.INFO,
    format_str: str = DEFAULT_FORMAT,
    stream: Optional[sys.__stdout__] = None,
    log_file: Optional[str] = None
) -> logging.Logger:
    """
    设置根日志配置
    
    Args:
        level: 日志级别，默认为 INFO
        format_str: 日志格式字符串
        stream: 输出流，默认为 sys.stdout
        log_file: 日志文件路径，如果指定则同时输出到文件
        
    Returns:
        logging.Logger: 根logger
    """
    # 清除现有的handlers
    root_logger = logging.getLogger()
    root_logger.handlers = []
    
    # 设置日志级别
    root_logger.setLevel(level)
    
    # 创建formatter
    formatter = logging.Formatter(format_str)
    
    # 控制台handler
    console_handler = logging.StreamHandler(stream or sys.stdout)
    console_handler.setLevel(level)
    console_handler.setFormatter(formatter)
    root_logger.addHandler(console_handler)
    
    # 文件handler（如果指定）
    if log_file:
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(level)
        file_handler.setFormatter(formatter)
        root_logger.addHandler(file_handler)
    
    return root_logger


def get_logger(name: str) -> logging.Logger:
    """
    获取指定名称的logger
    
    Args:
        name: logger名称，建议使用 __name__
        
    Returns:
        logging.Logger: logger实例
    """
    return logging.getLogger(name)


# 兼容旧代码的快捷函数
def debug(msg: str, *args, **kwargs):
    """输出debug级别日志"""
    logging.getLogger().debug(msg, *args, **kwargs)


def info(msg: str, *args, **kwargs):
    """输出info级别日志"""
    logging.getLogger().info(msg, *args, **kwargs)


def warning(msg: str, *args, **kwargs):
    """输出warning级别日志"""
    logging.getLogger().warning(msg, *args, **kwargs)


def error(msg: str, *args, **kwargs):
    """输出error级别日志"""
    logging.getLogger().error(msg, *args, **kwargs)


def exception(msg: str, *args, **kwargs):
    """输出exception级别日志（包含堆栈信息）"""
    logging.getLogger().exception(msg, *args, **kwargs)
