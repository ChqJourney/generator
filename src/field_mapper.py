"""
将字段映射配置转换为processor.py操作队列
"""
import json
import sys
import re
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, List
from openpyxl import load_workbook

def load_json(path: str) -> Dict:
    """加载JSON文件"""
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)

def extract_field_value(data: Dict, field_name: str) -> Any:
    """从extracted_data中提取字段值"""
    content = data.get('targets') or data.get('data', {})
    if isinstance(content, dict):
        field = content.get(field_name)
        if field and isinstance(field, dict):
            return field.get('value')
        return field
    elif isinstance(content, list):
        for item in content:
            if item['name'] == field_name:
                return item.get('value')
    return None

def generate_operations(config: Dict, metadata: Dict, extracted_data: Dict) -> Dict:
    """生成操作队列"""
    operations = []

    metadata_fields= {field['name']: field.get('value') for field in metadata.get('fields', [])}
    for mapping in config.get('field_mappings', []):
        source = mapping['source']
        source_field = mapping['source_field']
        field_type = mapping['type']
        placeholder = mapping['template_field']
        
        if source == 'metadata':
            value = metadata_fields.get(source_field)
        elif source == 'extracted_data':
            value = extract_field_value(extracted_data, source_field)
        else:
            continue
        
        if value is None:
            continue
        
        if field_type == 'text':
            operations.append({
                'type': 'text',
                'placeholder': placeholder,
                'value': str(value)
            })
        
        elif field_type == 'image':
            image_paths = value if isinstance(value, list) else [value]
            if image_paths and isinstance(image_paths[0], str):
                try:
                    import json
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
                'alignment': mapping.get('alignment')
            })
        
        elif field_type == 'table':
            table_data = None
            if check_type(value):
                table_data = value
            elif 'source_id' in value:
                template_headers = mapping.get('headers', [])
                table_data = build_table_data(value,template_headers)
            operations.append({
                'type': 'table',
                'placeholder': placeholder,
                'table_template_path': mapping['table_template_path'],
                'table_data': table_data
            })
    
    return {'operations': operations}

def build_table_data(value: Dict, headers: List[str]) -> List[List[str]]:
    # 假设value包含'source_id'，并且我们需要根据该ID构建表格数据
    source_id = value['source_id']
    sources=source_id.split('|')
    if len(sources)!=2:
        return []
    file_path=sources[0]
    sheet_name=sources[1]
    start_row=value.get('start_row',0)
    mapping=value.get('mapping',{})
    actual_path='data_files/'+file_path
    # 读取Excel文件
    table_data = get_xlsx_to_list(actual_path, sheet_name, start_row,mapping,headers)
    return table_data


def get_xlsx_to_list(file_path, sheet_name, start_row, mapping, headers) -> List[List[str]]:
    """
    读取 Excel 文件指定工作表中，从 start_row 下一行开始到末尾的数据。
    根据headers和mapping提取指定列，并按headers顺序排列。
    
    Args:
        file_path (str): xlsx 文件路径
        sheet_name (str): 工作表名称
        start_row (int): 表头所在行索引（0基）。数据从该行的下一行开始。
        mapping (List[Dict]): 列映射配置 [{'name': 'TargetCol', 'mapColumn': 'SourceCol'}, ...]
        headers (List[str]): 目标表头顺序
        
    Returns:
        List[List[str]]: 包含字符串的二维列表
    """
    # 使用 read_only=True 模式
    # 增加 data_only=True 以读取公式计算后的值
    wb = load_workbook(filename=file_path, read_only=True, data_only=True)
    
    try:
        # 获取指定的工作表
        ws = wb[sheet_name]
    except KeyError:
        wb.close()
        raise ValueError(f"工作表 '{sheet_name}' 不存在")

    # start_row 是0基索引，对应Excel行号是 start_row + 1
    header_row_num = start_row + 1
    
    # 使用 iter_rows 获取表头和数据
    # 从表头行开始读取
    row_iterator = ws.iter_rows(min_row=header_row_num, values_only=True)
    
    try:
        # 获取第一行作为表头
        header_values = next(row_iterator)
    except StopIteration:
        wb.close()
        return []

    # 辅助函数：规范化表头字符串（去除首尾空格，将连续空白字符替换为单个空格）
    def normalize_header(h):
        if h is None:
            return ""
        return " ".join(str(h).split())

    # 构建 Excel 表头到索引的映射 { "ColumnName": index }
    excel_headers_map = {}
    if header_values:
        for idx, val in enumerate(header_values):
            if val is not None:
                excel_headers_map[normalize_header(val)] = idx

    # 构建目标列索引列表
    target_col_indices = []
    
    # 解析 mapping 配置，转换为字典方便查找: { "TargetName": "SourceColumnName" }
    mapping_dict = {}
    if isinstance(mapping, list):
        for m in mapping:
            if isinstance(m, dict):
                mapping_dict[m.get('name')] = m.get('mapColumn')
    
    # 确定每列对应的 Excel 索引
    for header in headers:
        source_col_name = mapping_dict.get(header["name"])
        col_idx = None
        if source_col_name:
            normalized_source = normalize_header(source_col_name)
            if normalized_source in excel_headers_map:
                if header["type"]=='copy':
                    col_idx = excel_headers_map[normalized_source]
                else:
                    col_idx=excel_headers_map[normalized_source]+200
            else:
                # 可以在这里添加日志或打印警告
                print(f"警告: 未找到映射列 '{source_col_name}' (规范化: '{normalized_source}') 对应的 Excel 列，目标列 '{header}' 将填充空字符串。", file=sys.stderr)
        
        target_col_indices.append(col_idx)
            
    data = []
    
    # 继续遍历剩余的行（数据行）
    for row in row_iterator:
        row_data = []
        # 按 target_col_indices 的顺序提取数据
        for col_idx in target_col_indices:
            if col_idx is not None and col_idx < len(row):
                cell_val = row[col_idx]
                row_data.append(str(cell_val) if cell_val is not None else "")
            else:
                row_data.append("") # 缺失列填空字符串
                
        data.append(row_data)
        
    wb.close()
    return data

def check_type(data):
    # 逻辑：必须是一个list，且所有子元素是list，且所有子子元素是str
    return (
        isinstance(data, list) and 
        all(isinstance(sub, list) for sub in data) and 
        all(isinstance(item, str) for sub in data for item in sub)
    )

def main():
    import argparse
    parser = argparse.ArgumentParser(description='Generate processor operations from field mappings')
    parser.add_argument('--config', required=True, help='Path to report_config.json')
    parser.add_argument('--metadata', required=True, help='Path to metadata.json')
    parser.add_argument('--extracted_data', required=True, help='Path to extracted_data.json')
    parser.add_argument('--output', required=True, help='Path to output operations.json')
    args = parser.parse_args()
    
    try:
        config = load_json(args.config)
        metadata = load_json(args.metadata)
        extracted_data = load_json(args.extracted_data)
        
        operations = generate_operations(config, metadata, extracted_data)
        
        with open(args.output, 'w', encoding='utf-8') as f:
            json.dump(operations, f, indent=2, ensure_ascii=False)
        
        op_count = len(operations.get('operations', []))
        print(f"Generated {op_count} operations: {args.output}")
        return 0
    
    except FileNotFoundError as e:
        print(f"Error: File not found - {e}", file=sys.stderr)
        return 1
    except json.JSONDecodeError as e:
        print(f"Error: Invalid JSON - {e}", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1

if __name__ == '__main__':
    sys.exit(main())
