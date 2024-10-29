import os
from functools import lru_cache

import pandas as pd

from config import use_unordered_set
from logger_config import logger


@lru_cache(maxsize=None)
def read_excel(excel_path, usecols="A") -> pd.DataFrame:
    """
    读取Excel文件中指定的列数据
    :param excel_path: Excel文件路径
    :param usecols: 要读取的列（默认读取第A列）
    :return: 包含指定列数据的DataFrame
    """
    df = pd.read_excel(excel_path, usecols=usecols, dtype=str)
    return df


@lru_cache(maxsize=None)
def read_csv(excel_path) -> pd.DataFrame:
    """
    读取CSV文件中的所有数据
    :param excel_path: CSV文件路径
    :return: 包含所有数据的DataFrame
    """
    df = pd.read_csv(excel_path, dtype=str)
    return df


def save_to_csv(excel_path, data):
    """
    将数据增量保存到CSV文件中
    :param excel_path: 输出CSV文件路径
    :param data: 要保存的数据列表
    """
    df = pd.DataFrame(data, columns=['订单号', '客户名称', '状态'])
    # 使用 'a' 模式追加数据，并在文件存在时不写入表头
    df.to_csv(excel_path, mode='a', header=not os.path.exists(excel_path), index=False)
    logger.info(f"数据已增量保存到 {excel_path}")


def load_processed_clients(excel_path):
    """
    从现有的Excel文件中加载已处理的客户
    :param excel_path: Excel文件路径
    :return: 包含已处理客户的集合
    """
    processed_clients = set()
    if os.path.exists(excel_path):
        df = pd.read_csv(excel_path, dtype=str)
        # 只加载状态为"发送"或"跳过"的客户
        processed_clients = set(df[df['状态'].isin(['发送', '跳过'])]['订单号'])
        logger.info(f"已从 {excel_path} 加载 {len(processed_clients)} 个已处理的客户")
    else:
        logger.info(f"未找到 {excel_path}，将创建新文件")
    return processed_clients


def filter_order_nums(excel_order_nums, client_set):
    """
    根据配置项选择使用有序或无序方式进行差集运算
    :param excel_order_nums: 原始订单号列表
    :param client_set: 已处理客户集合
    :return: 过滤后的订单号列表
    """
    if use_unordered_set:
        # 使用无序集合差集运算，提高效率但不保留顺序
        return list(set(excel_order_nums) - client_set)
    else:
        # 使用有序列表推导式，保留顺序
        return [num for num in excel_order_nums if num not in client_set]
