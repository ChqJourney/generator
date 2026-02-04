"""
路径导航工具 - 支持点号路径访问分层数据
"""
from typing import Dict, Any, Optional


class PathNavigator:
    """路径导航器 - 支持点号路径访问分层数据"""
    
    @staticmethod
    def get_value(data: Dict, path: str) -> Any:
        """
        通过点号路径获取值
        
        Args:
            data: 分层数据字典
            path: 点号分隔的路径，如 'extracted_data.rated_wattage'
            
        Returns:
            路径对应的值，如果路径不存在返回None
        """
        if not path:
            return None
        
        parts = path.split('.')
        current = data
        
        for part in parts:
            if isinstance(current, dict) and part in current:
                current = current[part]
            else:
                return None
        
        return current
    
    @staticmethod
    def set_value(data: Dict, path: str, value: Any):
        """
        通过点号路径设置值
        
        Args:
            data: 分层数据字典
            path: 点号分隔的路径
            value: 要设置的值
        """
        parts = path.split('.')
        current = data
        
        for part in parts[:-1]:
            if part not in current:
                current[part] = {}
            current = current[part]
        
        current[parts[-1]] = value


# 向后兼容的别名
DataNavigator = PathNavigator
