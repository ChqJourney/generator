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
            支持字段名中包含点号，如 'calculated_data.v.1.a' 会先尝试查找 
            data['calculated_data']['v.1.a']，如果不存在则尝试 
            data['calculated_data']['v']['1']['a']
            
        Returns:
            路径对应的值，如果路径不存在返回None
        """
        if not path:
            return None
        
        # 策略：尝试最大匹配（支持字段名中包含点号）
        # 例如 'calculated_data.v.1.a' 会先尝试 'calculated_data' + 'v.1.a'
        parts = path.split('.')
        
        for split_idx in range(1, len(parts)):
            first_key = '.'.join(parts[:split_idx])
            rest_key = '.'.join(parts[split_idx:])
            
            if first_key in data:
                current = data[first_key]
                # 尝试将剩余部分作为完整键
                if isinstance(current, dict) and rest_key in current:
                    return current[rest_key]
                # 或者递归处理剩余部分
                result = PathNavigator.get_value(current, rest_key)
                if result is not None:
                    return result
        
        # 默认：按标准点号分割处理
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
