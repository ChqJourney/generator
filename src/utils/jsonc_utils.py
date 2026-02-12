"""
JSONC (JSON with Comments) 支持工具
支持解析带注释的 JSON 文件（包括 // 行注释和 /* */ 块注释）
"""

import json
import re
from pathlib import Path
from typing import Any, Union


def remove_jsonc_comments(content: str) -> str:
    """
    移除 JSONC 中的注释
    
    支持的注释格式:
    - 单行注释: // 注释内容
    - 块注释: /* 注释内容 */
    
    Args:
        content: 带注释的 JSON 内容
        
    Returns:
        移除注释后的纯 JSON 内容
    """
    # 首先处理块注释 /* */
    content = re.sub(r'/\*.*?\*/', '', content, flags=re.DOTALL)
    
    # 处理单行注释 //
    # 注意：避免匹配 URL 中的 //
    lines = []
    for line in content.split('\n'):
        # 找到 // 的位置，但要排除在字符串中的情况
        in_string = False
        string_char = None
        comment_start = -1
        
        for i, char in enumerate(line):
            if not in_string:
                if char in '"\'':
                    in_string = True
                    string_char = char
                elif char == '/' and i + 1 < len(line) and line[i + 1] == '/':
                    comment_start = i
                    break
            else:
                if char == string_char:
                    # 检查是否是转义的
                    if i > 0 and line[i - 1] != '\\':
                        in_string = False
                        string_char = None
        
        if comment_start >= 0:
            line = line[:comment_start]
        
        lines.append(line)
    
    return '\n'.join(lines)


def load_jsonc(file_path: Union[str, Path]) -> Any:
    """
    加载 JSONC 文件（带注释的 JSON）
    
    Args:
        file_path: JSONC 文件路径
        
    Returns:
        解析后的 Python 对象
        
    Raises:
        FileNotFoundError: 文件不存在
        json.JSONDecodeError: JSON 格式错误
        UnicodeDecodeError: 文件编码错误
    """
    path = Path(file_path)
    
    if not path.exists():
        raise FileNotFoundError(f"找不到文件: {path}")
    
    # 读取文件内容
    content = path.read_text(encoding='utf-8')
    
    # 移除注释
    cleaned_content = remove_jsonc_comments(content)
    
    # 解析 JSON
    return json.loads(cleaned_content)


def load_json(file_path: Union[str, Path]) -> Any:
    """
    加载 JSON 或 JSONC 文件（自动识别）
    
    如果文件扩展名是 .jsonc 或文件内容包含注释，则使用 JSONC 解析
    否则使用标准 JSON 解析
    
    Args:
        file_path: JSON/JSONC 文件路径
        
    Returns:
        解析后的 Python 对象
    """
    path = Path(file_path)
    
    # 如果扩展名是 .jsonc，直接使用 JSONC 解析
    if path.suffix.lower() == '.jsonc':
        return load_jsonc(path)
    
    # 否则，尝试使用标准 JSON 解析
    # 如果失败，再尝试 JSONC 解析
    try:
        content = path.read_text(encoding='utf-8')
        return json.loads(content)
    except json.JSONDecodeError:
        # 可能是带注释的 JSON，尝试 JSONC 解析
        cleaned_content = remove_jsonc_comments(content)
        return json.loads(cleaned_content)


# 向后兼容的别名
parse_jsonc = load_jsonc
