"""
测试表格处理模块
演示如何使用 TableDataTransformer 和 EnhancedTableInserter
"""

from docx import Document
import sys
import os
import json
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from src.table_processor import TableDataTransformer, EnhancedTableInserter

def test_data_transformer():
    """测试数据转换器"""
    print("=" * 60)
    print("测试 1: 数据转换器 (TableDataTransformer)")
    print("=" * 60)
    
    raw_data = [
        ["0.1500", "12.5", "0.95", "1500", "120.5", "2700", "95", "85", "0.4000", "0.3500", "2.5", "0.0125"],
        ["0.2500", "25.0", "0.96", "2500", "100.0", "3000", "95", "85", "0.4100", "0.3600", "2.8", "0.0250"],
        ["0.3500", "37.5", "0.97", "3500", "107.1", "3300", "95", "85", "0.4200", "0.3700", "3.0", "0.0375"],
        ["", "", "", "", "", "", "", "", "", "", "", ""]
    ]
    
    metadata = {"model_name": "LED-001"}
    
    transformations = [
        {
            "type": "skip_columns",
            "columns": [0]
        },
        {
            "type": "add_column",
            "position": 0,
            "source": "metadata:model_name"
        },
        {
            "type": "calculate",
            "column": 3,
            "operation": "average",
            "decimal": 1
        },
        {
            "type": "format_column",
            "column": 2,
            "function": "lambda x: f'{x:.4f}'"
        },
        {
            "type": "format_column",
            "column": 4,
            "function": "lambda x: f'{x:.1f}' if x < 10 else f'{x:.0f}'"
        }
    ]
    
    transformer = TableDataTransformer()
    transformed_data = transformer.transform(raw_data, transformations, metadata)
    
    print("原始数据:")
    for row in raw_data:
        print(f"  {row}")
    
    print("\n转换后数据:")
    for idx, row in enumerate(transformed_data):
        print(f"  行{idx}: {row}")
    
    return transformed_data

def test_enhanced_inserter():
    """测试增强版表格插入器"""
    print("\n" + "=" * 60)
    print("测试 2: 增强版表格插入器 (EnhancedTableInserter)")
    print("=" * 60)
    
    template_path = "report_templates/report_template1.docx"
    table_template_path = "report_templates/photometric_table_template.docx"
    output_path = "output/test_table_processor.docx"
    
    if not os.path.exists(table_template_path):
        print(f"警告: 表格模板不存在: {table_template_path}")
        print("跳过插入器测试")
        return
    
    if not os.path.exists(template_path):
        print(f"警告: 主模板不存在: {template_path}")
        print("跳过插入器测试")
        return
    
    os.makedirs("output", exist_ok=True)
    
    import shutil
    shutil.copy(template_path, output_path)
    
    doc = Document(output_path)
    
    raw_data = [
        ["0.1500", "12.5", "0.95", "1500", "120.5", "2700", "95", "85", "0.4000", "0.3500", "2.5", "0.0125"],
        ["0.2500", "25.0", "0.96", "2500", "100.0", "3000", "95", "85", "0.4100", "0.3600", "2.8", "0.0250"],
        ["0.3500", "37.5", "0.97", "3500", "107.1", "3300", "95", "85", "0.4200", "0.3700", "3.0", "0.0375"],
        ["", "", "", "", "", "", "", "", "", "", "", ""]
    ]
    
    metadata = {"model_name": "LED-TEST-001"}
    
    transformations = [
        {
            "type": "skip_columns",
            "columns": [0]
        },
        {
            "type": "format_column",
            "column": 0,
            "decimal": 1
        },
        {
            "type": "format_column",
            "column": 1,
            "decimal": 1
        },
        {
            "type": "format_column",
            "column": 2,
            "function": "lambda x: f'{x:.4f}'"
        },
        {
            "type": "format_column",
            "column": 3,
            "decimal": 1
        },
        {
            "type": "format_column",
            "column": 4,
            "function": "lambda x: f'{x:.4f}' if x < 1 else f'{x:.2f}'"
        },
        {
            "type": "format_column",
            "column": 5,
            "function": "lambda x: f'{x:.1f}' if x < 10 else f'{x:.0f}'"
        },
        {
            "type": "format_column",
            "column": 8,
            "function": "lambda x: f'{x:.4f}'"
        },
        {
            "type": "format_column",
            "column": 9,
            "function": "lambda x: f'{x:.4f}'"
        },
        {
            "type": "format_column",
            "column": 10,
            "decimal": 1
        },
        {
            "type": "format_column",
            "column": 11,
            "decimal": 4
        }
    ]
    
    inserter = EnhancedTableInserter(doc)
    
    try:
        inserter.insert(
            placeholder="photometric_data",
            table_template_path=table_template_path,
            raw_data=raw_data,
            transformations=transformations,
            metadata=metadata,
            row_strategy='fixed_rows',
            skip_columns=[0, 1],
            header_rows=2,
            location='body'
        )
        
        doc.save(output_path)
        print(f"✓ 表格插入成功！")
        print(f"  输出文件: {output_path}")
        
    except Exception as e:
        print(f"✗ 表格插入失败: {e}")
        import traceback
        traceback.print_exc()

def test_complex_formatting():
    """测试复杂的小数点格式化"""
    print("\n" + "=" * 60)
    print("测试 3: 复杂小数点格式化")
    print("=" * 60)
    
    test_cases = [
        {
            "name": "CCT格式化 (<10保留1位，>=10保留0位)",
            "data": [
                ["5.5", "15.3", "2500", "9999"]
            ],
            "transformations": [
                {
                    "type": "format_column",
                    "column": 0,
                    "function": "lambda x: f'{x:.1f}' if x < 10 else f'{x:.0f}'"
                },
                {
                    "type": "format_column",
                    "column": 1,
                    "function": "lambda x: f'{x:.1f}' if x < 10 else f'{x:.0f}'"
                },
                {
                    "type": "format_column",
                    "column": 2,
                    "function": "lambda x: f'{x:.1f}' if x < 10 else f'{x:.0f}'"
                },
                {
                    "type": "format_column",
                    "column": 3,
                    "function": "lambda x: f'{x:.1f}' if x < 10 else f'{x:.0f}'"
                }
            ]
        },
        {
            "name": "光效格式化 (<1保留4位，>=1保留2位)",
            "data": [
                ["0.85", "95.5", "120.3"]
            ],
            "transformations": [
                {
                    "type": "format_column",
                    "column": 0,
                    "function": "lambda x: f'{x:.4f}' if x < 1 else f'{x:.2f}'"
                },
                {
                    "type": "format_column",
                    "column": 1,
                    "function": "lambda x: f'{x:.4f}' if x < 1 else f'{x:.2f}'"
                },
                {
                    "type": "format_column",
                    "column": 2,
                    "function": "lambda x: f'{x:.4f}' if x < 1 else f'{x:.2f}'"
                }
            ]
        },
        {
            "name": "XY坐标格式化 (<0.1保留6位，>=0.1保留4位)",
            "data": [
                ["0.04", "0.35", "0.15", "0.4"]
            ],
            "transformations": [
                {
                    "type": "format_column",
                    "column": 0,
                    "function": "lambda x: f'{x:.6f}' if abs(x) < 0.1 else f'{x:.4f}'"
                },
                {
                    "type": "format_column",
                    "column": 1,
                    "function": "lambda x: f'{x:.6f}' if abs(x) < 0.1 else f'{x:.4f}'"
                },
                {
                    "type": "format_column",
                    "column": 2,
                    "function": "lambda x: f'{x:.6f}' if abs(x) < 0.1 else f'{x:.4f}'"
                },
                {
                    "type": "format_column",
                    "column": 3,
                    "function": "lambda x: f'{x:.6f}' if abs(x) < 0.1 else f'{x:.4f}'"
                }
            ]
        }
    ]
    
    transformer = TableDataTransformer()
    
    for test_case in test_cases:
        print(f"\n{test_case['name']}")
        print(f"原始数据: {test_case['data'][0]}")
        
        result = transformer.transform(test_case['data'], test_case['transformations'])
        print(f"格式化后: {result[0]}")

def main():
    print("\n表格处理模块测试")
    print("=" * 60)
    
    try:
        test_data_transformer()
        test_complex_formatting()
        test_enhanced_inserter()
        
        print("\n" + "=" * 60)
        print("所有测试完成！")
        print("=" * 60)
        
    except Exception as e:
        print(f"\n测试失败: {e}")
        import traceback
        traceback.print_exc()
        return 1
    
    return 0

if __name__ == '__main__':
    sys.exit(main())
