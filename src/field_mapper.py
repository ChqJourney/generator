"""
将字段映射配置转换为processor.py操作队列
新架构：从calculated_report.json读取数据，支持点号路径访问
"""
import json
from typing import Dict, Any, List, Optional
from utils.path_navigator import DataNavigator
from utils.logging_config import get_logger

logger = get_logger(__name__)


def parse_checkbox_value(value: Any) -> bool:
    """
    解析 checkbox 值，转换为布尔值
    
    Args:
        value: checkbox 值，可能是 {"type": "checkbox", "value": "true/false"} 格式
               或直接的字符串/布尔值
    
    Returns:
        bool: True 表示勾选，False 表示未勾选
    """
    if value is None:
        return False
    
    # 检查是否是标准 checkbox 格式 {"type": "checkbox", "value": "..."}
    if isinstance(value, dict) and value.get('type') == 'checkbox':
        value_str = value.get('value', 'false')
        if isinstance(value_str, str):
            return value_str.lower() == 'true'
        return bool(value_str)
    
    # 如果是字符串
    if isinstance(value, str):
        return value.lower() == 'true'
    
    # 其他情况直接转布尔值
    return bool(value)


def load_json(path: str) -> Dict:
    """加载JSON/JSONC文件（支持带注释的JSON）"""
    from utils.jsonc_utils import load_json as load_jsonc
    return load_jsonc(path)


def get_value_by_path(data: Dict, path: str) -> Any:
    """
    通过点号路径获取值
    
    Args:
        data: 分层数据字典
        path: 点号分隔的路径，如 'extracted_data.rated_wattage'
        
    Returns:
        路径对应的值
    """
    return DataNavigator.get_value(data, path)


def is_direct_table_data(value: Any) -> bool:
    """
    判断值是否为直接的表格数据（列表的列表或标准表格格式）
    
    Args:
        value: 字段值
        
    Returns:
        bool: 是否是直接表格数据
    """
    # 检查是否是标准表格格式 {"type": "table", "value": [...]}
    if isinstance(value, dict) and value.get('type') == 'table':
        table_value = value.get('value', [])
        return (
            isinstance(table_value, list) and 
            len(table_value) > 0 and 
            all(isinstance(row, list) for row in table_value)
        )
    
    # 检查是否是直接的列表的列表格式
    return (
        isinstance(value, list) and 
        len(value) > 0 and 
        all(isinstance(row, list) for row in value)
    )


def extract_table_data(value: Any) -> Optional[List[List]]:
    """
    从表格数据中提取实际的表格内容
    
    Args:
        value: 字段值（可能是列表的列表或标准表格格式）
        
    Returns:
        List[List]: 表格数据（列表的列表）
    """
    if isinstance(value, dict) and value.get('type') == 'table':
        return value.get('value')
    if isinstance(value, list):
        return value
    return None


def generate_operations(config: Dict, report_data: Dict) -> Dict:
    """
    生成操作队列
    
    Args:
        config: 配置字典
        report_data: 计算后的报告数据（calculated_report.json）
        
    Returns:
        Dict: 操作队列
    """
    operations = []
    
    for mapping in config.get('field_mappings', []):
        source_field = mapping.get('source_field')
        field_type = mapping.get('type')
        placeholder = mapping.get('template_field')
        
        if not source_field:
            logger.warning(f"Missing source_field for {placeholder}, skipping")
            continue
        
        # 通过点号路径获取值
        value = get_value_by_path(report_data, source_field)
        
        if value is None:
            # 对于 checkbox 类型，如果没有值则默认为 false
            if field_type == 'checkbox':
                value = False
                logger.info(f"Checkbox value not found for path '{source_field}', defaulting to False for {placeholder}")
            else:
                logger.warning(f"Value not found for path '{source_field}', skipping {placeholder}")
                continue
        
        if field_type == 'text':
            operations.append({
                'type': 'text',
                'placeholder': placeholder,
                'value': str(value),
                'location': mapping.get('location', 'body')
            })
        
        elif field_type == 'image':
            image_paths = value if isinstance(value, list) else [value]
            if image_paths and isinstance(image_paths[0], str):
                try:
                    parsed = json.loads(image_paths[0])
                    image_paths = parsed
                except:
                    pass
            operations.append({
                'type': 'image',
                'placeholder': placeholder,
                'image_paths': image_paths,
                'width': mapping.get('width'),
                'height': mapping.get('height'),
                'alignment': mapping.get('alignment'),
                'location': mapping.get('location', 'body')
            })
        
        elif field_type == 'table':
            # 提取表格数据（支持标准格式和直接列表格式）
            table_data = extract_table_data(value)
            
            if table_data is None:
                logger.warning(f"Unrecognized table data format for {placeholder}")
                continue
            
            logger.info(f"Using embedded table data for {placeholder}")
            
            operation = {
                'type': 'table',
                'placeholder': placeholder,
                'table_template_path': mapping['table_template_path'],
                'table_data': table_data
            }
            
            # 添加可选参数
            for key in ['transformations', 'row_strategy', 'skip_columns', 'header_rows','text_insert']:
                if key in mapping:
                    operation[key] = mapping[key]
            
            operations.append(operation)
        
        elif field_type == 'checkbox':
            # 解析 checkbox 值
            checkbox_value = parse_checkbox_value(value)
            
            # checkbox 使用 template_field 作为 checkbox 名称
            operations.append({
                'type': 'checkbox',
                'checkbox_mapping': {
                    placeholder: checkbox_value
                }
            })
    
    return {'operations': operations}


def main():
    """命令行入口"""
    import argparse
    
    parser = argparse.ArgumentParser(
        description='Generate processor operations from field mappings (new architecture)'
    )
    parser.add_argument(
        '--config', 
        required=True, 
        help='Path to report_config.json'
    )
    parser.add_argument(
        '--report', 
        required=True, 
        help='Path to calculated_report.json'
    )
    parser.add_argument(
        '--output', 
        required=True, 
        help='Path to output operations.json'
    )
    
    args = parser.parse_args()
    
    try:
        # 加载数据
        config = load_json(args.config)
        report_data = load_json(args.report)
        
        # 生成操作队列
        operations = generate_operations(config, report_data)
        
        # 保存结果
        with open(args.output, 'w', encoding='utf-8') as f:
            json.dump(operations, f, indent=2, ensure_ascii=False)
        
        op_count = len(operations.get('operations', []))
        logger.info(f"Generated {op_count} operations: {args.output}")
        return 0
    
    except FileNotFoundError as e:
        logger.error(f"File not found - {e}")
        return 1
    except json.JSONDecodeError as e:
        logger.error(f"Invalid JSON - {e}")
        return 1
    except Exception as e:
        logger.error(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == '__main__':
    import sys
    sys.exit(main())
