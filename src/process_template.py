"""
使用processor.py处理Word模板
通过calculated_report.json获取所有数据
"""
import sys
import json
from pathlib import Path
from docx.shared import Inches

# 将 src 目录添加到路径
sys.path.insert(0, str(Path(__file__).parent))

from utils.logging_config import get_logger

logger = get_logger(__name__)

try:
    from processor import DocxTemplateProcessor, DocxTemplateError
except ImportError as e:
    logger.error(f"processor.py import error: {e}")
    sys.exit(1)


def extract_template_checkboxes(template_path: Path) -> set:
    """
    从Word模板中提取所有checkbox的名称
    
    Args:
        template_path: Word模板文件路径
        
    Returns:
        set: checkbox名称集合
    """
    try:
        from docx import Document
        doc = Document(str(template_path))
        root = doc.part.element
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        w_ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        
        checkboxes = set()
        checkbox_elements = root.findall('.//w:checkBox', namespaces=ns)
        
        for checkbox in checkbox_elements:
            ffdata = checkbox.getparent()
            if ffdata is not None:
                name_elem = ffdata.find('w:name', namespaces=ns)
                if name_elem is not None:
                    field_name = name_elem.get(w_ns + 'val')
                    if field_name:
                        checkboxes.add(field_name)
        
        return checkboxes
    except Exception as e:
        logger.warning(f"Failed to extract checkboxes from template: {e}")
        return set()


def load_calculated_report(report_path: Path) -> dict:
    """
    加载calculated_report.json文件
    
    返回结构:
    {
        "metadata": {...},
        "extracted_data": {...},
        "calculated_data": {...}
    }
    """
    with open(report_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    # 确保基本结构存在
    if 'metadata' not in data:
        data['metadata'] = {}
    if 'extracted_data' not in data:
        data['extracted_data'] = {}
    if 'calculated_data' not in data:
        data['calculated_data'] = {}
    
    return data


def get_nested_value(data: dict, path: str, default=None):
    """
    获取嵌套字典的值，支持点号路径如 'metadata.report_no'
    也支持字段名中包含点号，如 'calculated_data.v.II.1.a.1'
    """
    # 策略：尝试最大匹配（支持字段名中包含点号）
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
            result = get_nested_value(current, rest_key, default)
            if result is not default:
                return result
    
    # 默认：按标准点号分割处理
    value = data
    for key in parts:
        if isinstance(value, dict) and key in value:
            value = value[key]
        else:
            return default
    return value





def resolve_text_value(value_ref: str, calculated_report: dict) -> str:
    """
    解析文本值引用
    
    支持格式:
    - 直接值: "some text"
    - 路径引用: "metadata.report_no"
    - 路径引用: "extracted_data.model_identifier"
    - 路径引用: "calculated_data.energy_class"
    """
    # 尝试作为路径解析
    value = get_nested_value(calculated_report, value_ref)
    if value is not None:
        return str(value)
    
    # 如果找不到，返回原值
    return value_ref


def resolve_table_data(data_ref: str, calculated_report: dict) -> list:
    """
    解析表格数据源
    
    支持格式:
    - 直接数据: 传入列表
    - 路径引用: "extracted_data.photometric_data"
    """
    if isinstance(data_ref, list):
        return data_ref
    
    value = get_nested_value(calculated_report, data_ref)
    if isinstance(value, list):
        return value
    
    return []


def main():
    import argparse
    parser = argparse.ArgumentParser(
        description='Process Word template using calculated report data',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python process_template.py 
    --template report_templates/template.docx 
    --operations config/operations.json 
    --calculated-report output/calculated_report.json 
    --output output/final_report.docx
        """
    )
    parser.add_argument('--template', required=True, help='Path to Word template file')
    parser.add_argument('--operations', required=True, help='Path to operations.json')
    parser.add_argument('--calculated-report', required=True, 
                        help='Path to calculated_report.json (contains metadata, extracted_data, calculated_data)')
    parser.add_argument('--output', required=True, help='Path to output file')
    args = parser.parse_args()
    
    try:
        template_path = Path(args.template)
        operations_path = Path(args.operations)
        report_path = Path(args.calculated_report)
        output_path = Path(args.output)

        if not template_path.exists():
            logger.error(f"Template file not found: {template_path}")
            return 1
        
        if not operations_path.exists():
            logger.error(f"Operations file not found: {operations_path}")
            return 1
        
        if not report_path.exists():
            logger.error(f"Calculated report file not found: {report_path}")
            return 1
        
        # 加载数据
        with open(operations_path, 'r', encoding='utf-8') as f:
            operations_data = json.load(f)
        
        calculated_report = load_calculated_report(report_path)
        
        logger.info(f"Loaded calculated report: {report_path}")
        logger.info(f"  Metadata fields: {len(calculated_report.get('metadata', {}))}")
        logger.info(f"  Extracted data fields: {len(calculated_report.get('extracted_data', {}))}")
        logger.info(f"  Calculated data fields: {len(calculated_report.get('calculated_data', {}))}")
        
        processor = DocxTemplateProcessor(str(template_path), str(output_path))
        
        # 处理未设置的 checkbox：模板中有但 operations 中没有的 checkbox 自动设为 false
        template_checkboxes = extract_template_checkboxes(template_path)
        if template_checkboxes:
            logger.info(f"Found {len(template_checkboxes)} checkboxes in template")
            
            # 收集 operations 中已设置的 checkbox
            operations_checkboxes = set()
            for op in operations_data.get('operations', []):
                if op.get('type') == 'checkbox':
                    checkbox_mapping = op.get('checkbox_mapping', {})
                    operations_checkboxes.update(checkbox_mapping.keys())
            
            # 找出未设置的 checkbox
            unchecked_checkboxes = template_checkboxes - operations_checkboxes
            if unchecked_checkboxes:
                logger.info(f"Auto-setting {len(unchecked_checkboxes)} checkboxes to false: {unchecked_checkboxes}")
                # 为未设置的 checkbox 添加 false 的 operation
                auto_checkbox_mapping = {name: False for name in unchecked_checkboxes}
                processor.add_checkboxes(auto_checkbox_mapping)
        
        op_count = 0
        for op in operations_data.get('operations', []):
            op_type = op.get('type')
            
            if op_type == 'text':
                placeholder = op['placeholder']
                # 支持直接值或路径引用
                value_ref = op.get('value', op.get('source_field', ''))
                value = resolve_text_value(value_ref, calculated_report)
                location = op.get('location', 'body')
                
                processor.add_text(placeholder, value, location)
                op_count += 1
            
            elif op_type == 'image':
                width = op.get('width')
                height = op.get('height')
                
                if width is not None and isinstance(width, (int, float)):
                    width = Inches(width)
                if height is not None and isinstance(height, (int, float)):
                    height = Inches(height)
                
                processor.add_image(
                    op['placeholder'],
                    op['image_paths'],
                    width,
                    height,
                    op.get('alignment'),
                    op.get('location', 'body')
                )
                op_count += 1
            
            elif op_type == 'table':
                placeholder = op['placeholder']
                table_template_path = op['table_template_path']
                
                # 解析表格数据
                raw_data = op.get('table_data', [])
                if isinstance(raw_data, str):
                    # 如果是字符串路径，从calculated_report解析
                    raw_data = resolve_table_data(raw_data, calculated_report)
                
                transformations = op.get('transformations', [])
                row_strategy = op.get('row_strategy', 'fixed_rows')
                skip_columns = op.get('skip_columns')
                header_rows = op.get('header_rows', 1)
                text_insert = op.get('text_insert')
                
                processor.add_table(
                    placeholder,
                    table_template_path,
                    raw_data,
                    transformations,
                    calculated_report,
                    row_strategy,
                    skip_columns,
                    header_rows,
                    text_insert
                )
                op_count += 1
            
            elif op_type == 'checkbox':
                checkbox_mapping = op.get('checkbox_mapping', {})
                if checkbox_mapping:
                    processor.add_checkboxes(checkbox_mapping)
                    op_count += 1
                    logger.debug(f"Added checkbox operation: {checkbox_mapping}")
        
        logger.info(f"Executing {op_count} operations...")
        result = processor.process()
        
        logger.info(f"Report generated successfully: {result}")
        
        # Auto open the output file in Windows
        if sys.platform == "win32":
            import os
            os.startfile(output_path)
        
        return 0
    
    except FileNotFoundError as e:
        logger.error(f"File not found - {e}")
        return 1
    except json.JSONDecodeError as e:
        logger.error(f"Invalid JSON - {e}")
        return 1
    except DocxTemplateError as e:
        logger.error(f"DocxTemplateError: {e}")
        return 1
    except Exception as e:
        logger.error(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == '__main__':
    sys.exit(main())
