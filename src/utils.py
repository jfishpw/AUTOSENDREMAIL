"""
工具函数模块
"""
import os
import re
import glob
import shutil
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Tuple


def ensure_dir(path: str) -> str:
    """
    确保目录存在，不存在则创建
    
    Args:
        path: 目录路径
    
    Returns:
        目录路径
    """
    if not os.path.exists(path):
        os.makedirs(path, exist_ok=True)
    return path


def expand_path(path: str, variables: dict = None) -> str:
    """
    展开路径中的变量
    
    支持的变量:
        {date} - 当前日期 YYYY-MM-DD
        {time} - 当前时间 HH-MM-SS
        {datetime} - 当前日期时间 YYYY-MM-DD_HH-MM-SS
        {year} - 年份
        {month} - 月份
        {day} - 日期
        {sender} - 发件人邮箱
    
    Args:
        path: 包含变量的路径
        variables: 额外的变量字典
    
    Returns:
        展开后的路径
    """
    if variables is None:
        variables = {}
    
    now = datetime.now()
    
    default_vars = {
        'date': now.strftime('%Y-%m-%d'),
        'time': now.strftime('%H-%M-%S'),
        'datetime': now.strftime('%Y-%m-%d_%H-%M-%S'),
        'year': str(now.year),
        'month': f'{now.month:02d}',
        'day': f'{now.day:02d}',
    }
    
    all_vars = {**default_vars, **variables}
    
    result = path
    for key, value in all_vars.items():
        result = result.replace(f'{{{key}}}', str(value))
    
    return result


def find_files(pattern: str, base_dir: str = '.') -> List[str]:
    """
    查找匹配通配符的文件
    
    Args:
        pattern: 文件模式（支持通配符）
        base_dir: 基础目录
    
    Returns:
        匹配的文件列表
    """
    if not os.path.isabs(pattern):
        pattern = os.path.join(base_dir, pattern)
    
    files = glob.glob(pattern)
    return [f for f in files if os.path.isfile(f)]


def get_unique_filename(filepath: str) -> str:
    """
    获取唯一的文件名，如果文件已存在则添加序号
    
    Args:
        filepath: 原始文件路径
    
    Returns:
        唯一的文件路径
    """
    if not os.path.exists(filepath):
        return filepath
    
    directory = os.path.dirname(filepath)
    filename = os.path.basename(filepath)
    name, ext = os.path.splitext(filename)
    
    counter = 1
    while True:
        new_filename = f"{name}_{counter}{ext}"
        new_filepath = os.path.join(directory, new_filename)
        if not os.path.exists(new_filepath):
            return new_filepath
        counter += 1


def handle_file_conflict(
    filepath: str,
    strategy: str = 'rename'
) -> Tuple[str, bool]:
    """
    处理文件名冲突
    
    Args:
        filepath: 文件路径
        strategy: 处理策略 ('overwrite', 'rename', 'skip')
    
    Returns:
        (最终文件路径, 是否继续处理)
    """
    if not os.path.exists(filepath):
        return filepath, True
    
    if strategy == 'overwrite':
        return filepath, True
    elif strategy == 'rename':
        return get_unique_filename(filepath), True
    elif strategy == 'skip':
        return filepath, False
    else:
        return filepath, True


def validate_email(email: str) -> bool:
    """
    验证邮箱地址格式
    
    Args:
        email: 邮箱地址
    
    Returns:
        是否有效
    """
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(pattern, email))


def get_file_size(filepath: str) -> int:
    """
    获取文件大小（字节）
    
    Args:
        filepath: 文件路径
    
    Returns:
        文件大小
    """
    if not os.path.exists(filepath):
        return 0
    return os.path.getsize(filepath)


def format_size(size_bytes: int) -> str:
    """
    格式化文件大小
    
    Args:
        size_bytes: 字节数
    
    Returns:
        格式化后的大小字符串
    """
    for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
        if size_bytes < 1024.0:
            return f"{size_bytes:.2f} {unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:.2f} PB"


def parse_size_to_bytes(size_str: str) -> int:
    """
    将大小字符串转换为字节数
    
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


def sanitize_filename(filename: str) -> str:
    """
    清理文件名，移除非法字符
    
    Args:
        filename: 原始文件名
    
    Returns:
        清理后的文件名
    """
    illegal_chars = r'[<>:"/\\|?*]'
    return re.sub(illegal_chars, '_', filename)


def copy_file(src: str, dst: str) -> bool:
    """
    复制文件
    
    Args:
        src: 源文件路径
        dst: 目标文件路径
    
    Returns:
        是否成功
    """
    try:
        ensure_dir(os.path.dirname(dst))
        shutil.copy2(src, dst)
        return True
    except Exception as e:
        return False


def move_file(src: str, dst: str) -> bool:
    """
    移动文件
    
    Args:
        src: 源文件路径
        dst: 目标文件路径
    
    Returns:
        是否成功
    """
    try:
        ensure_dir(os.path.dirname(dst))
        shutil.move(src, dst)
        return True
    except Exception as e:
        return False


def delete_file(filepath: str) -> bool:
    """
    删除文件
    
    Args:
        filepath: 文件路径
    
    Returns:
        是否成功
    """
    try:
        if os.path.exists(filepath):
            os.remove(filepath)
        return True
    except Exception as e:
        return False
