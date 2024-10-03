import sys

from loguru import logger


def setup_logger():
    # 移除默认的处理器
    logger.remove()
    # 添加控制台处理器
    logger.add(sys.stderr, format="{time} {level} {message}", level="INFO")
    # 添加文件处理器
    logger.add("app.log", rotation="500 MB", level="INFO")

    return logger


# 配置并返回logger实例
logger = setup_logger()
