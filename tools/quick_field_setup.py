#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
快速字段配置工具 - 批量设置同类字段

用法:
    # 批量设置metadata字段
    python tools/quick_field_setup.py config/extracted.json --source metadata \
        --fields report_no,issue_date,applicant_name,product_name
    
    # 批量设置extracted_data字段
    python tools/quick_field_setup.py config/extracted.json --source extracted_data \
        --fields model_identifier,rated_wattage,useful_luminous_flux
    
    # 批量设置计算字段
    python tools/quick_field_setup.py config/extracted.json --calculated \
        --function calculate_energy_class_rating \
        --args extracted_data.rated_wattage,extracted_data.useful_luminous_flux \
        --fields energy_class_rating
"""

import json
import argparse
import sys
import os
from typing import List, Dict


def load_mappings(input_file: str) -> List[Dict]:
    """加载现有的mappings"""
    with open(input_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    mappings = data.get("field_mappings", [])
    if not mappings:
        from extract_template_elements import generate_field_mappings
        mappings = generate_field_mappings(
            data.get("placeholders", {}),
            data.get("checkboxes", [])
        )
    
    return mappings, data


def save_mappings(mappings: List[Dict], original_data: Dict, output_file: str):
    """保存mappings到文件"""
    config = {
        "template_path": original_data.get("template_path", ""),
        "output_dir": "output/",
        "field_mappings": mappings,
        "merge_strategy": "prefer_first"
    }
    
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
    
    print(f"配置已保存: {output_file}")


def batch_set_source(mappings: List[Dict], fields: List[str], source: str):
    """批量设置数据源"""
    updated = 0
    for mapping in mappings:
        if mapping["template_field"] in fields:
            mapping["source_field"] = f"{source}.{mapping['template_field']}"
            mapping["inferred_source"] = source
            if source == "calculated_data":
                mapping["is_calculated"] = True
            else:
                mapping["is_calculated"] = False
            updated += 1
    
    print(f"已更新 {updated} 个字段的source为 {source}")
    return mappings


def batch_set_calculated(mappings: List[Dict], fields: List[str], 
                         function: str, args: List[str]):
    """批量设置计算字段"""
    updated = 0
    for mapping in mappings:
        if mapping["template_field"] in fields:
            mapping["source_field"] = f"calculated_data.{mapping['template_field']}"
            mapping["inferred_source"] = "calculated_data"
            mapping["is_calculated"] = True
            mapping["function"] = function
            mapping["args"] = args
            updated += 1
    
    print(f"已更新 {updated} 个字段为计算字段")
    print(f"  函数: {function}")
    print(f"  参数: {', '.join(args)}")
    return mappings


def batch_set_table_config(mappings: List[Dict], fields: List[str],
                           template_path: str = None, row_strategy: str = None,
                           header_rows: int = None):
    """批量设置table字段配置"""
    updated = 0
    for mapping in mappings:
        if mapping["template_field"] in fields:
            if template_path:
                mapping["table_template_path"] = template_path
            if row_strategy:
                mapping["row_strategy"] = row_strategy
            if header_rows is not None:
                mapping["header_rows"] = header_rows
            updated += 1
    
    print(f"已更新 {updated} 个table字段配置")
    return mappings


def batch_set_image_config(mappings: List[Dict], fields: List[str],
                           width: float = None, alignment: str = None):
    """批量设置image字段配置"""
    updated = 0
    for mapping in mappings:
        if mapping["template_field"] in fields:
            if width is not None:
                mapping["width"] = width
            if alignment:
                mapping["alignment"] = alignment
            updated += 1
    
    print(f"已更新 {updated} 个image字段配置")
    return mappings


def batch_set_type(mappings: List[Dict], fields: List[str], field_type: str):
    """批量设置字段类型"""
    updated = 0
    for mapping in mappings:
        if mapping["template_field"] in fields:
            mapping["type"] = field_type
            updated += 1
    
    print(f"已更新 {updated} 个字段的type为 {field_type}")
    return mappings


def show_statistics(mappings: List[Dict]):
    """显示配置统计信息"""
    print("\n=== 当前配置统计 ===")
    
    by_type = {}
    by_source = {}
    calculated = []
    
    for m in mappings:
        t = m.get("type", "unknown")
        by_type[t] = by_type.get(t, 0) + 1
        
        source = m.get("inferred_source", "unknown")
        by_source[source] = by_source.get(source, 0) + 1
        
        if m.get("is_calculated"):
            calculated.append(m["template_field"])
    
    print(f"\n按类型统计:")
    for t, count in sorted(by_type.items()):
        print(f"  {t}: {count}")
    
    print(f"\n按数据源统计:")
    for s, count in sorted(by_source.items()):
        print(f"  {s}: {count}")
    
    if calculated:
        print(f"\n计算字段 ({len(calculated)}):")
        for name in calculated:
            print(f"  - {name}")


def show_unconfigured(mappings: List[Dict]):
    """显示未配置的字段"""
    unconfigured = []
    
    for m in mappings:
        needs_config = False
        
        # 检查基本配置
        if not m.get("source_field"):
            needs_config = True
        elif m.get("type") == "table" and not m.get("table_template_path"):
            needs_config = True
        elif m.get("is_calculated") and not m.get("function"):
            needs_config = True
        
        if needs_config:
            unconfigured.append(m["template_field"])
    
    if unconfigured:
        print(f"\n需要配置的字段 ({len(unconfigured)}):")
        for name in unconfigured:
            print(f"  - {name}")
    else:
        print("\n✓ 所有字段已配置完成")


def main():
    parser = argparse.ArgumentParser(
        description="快速字段配置工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
    # 批量设置字段的source
    python tools/quick_field_setup.py extracted.json --source metadata \\
        --fields report_no,issue_date,applicant_name
    
    # 批量设置计算字段
    python tools/quick_field_setup.py extracted.json --calculated \\
        --function calculate_energy_class_rating \\
        --args extracted_data.wattage,extracted_data.flux \\
        --fields energy_class,efficiency_rating
    
    # 批量设置table配置
    python tools/quick_field_setup.py extracted.json \\
        --table-config --table-template report_templates/data_table.docx \\
        --row-strategy fixed_rows --fields photometric_data,test_data
    
    # 批量设置image配置
    python tools/quick_field_setup.py extracted.json \\
        --image-config --width 4.0 --alignment center \\
        --fields product_image,signature
    
    # 查看统计信息
    python tools/quick_field_setup.py extracted.json --stats
    
    # 查看未配置字段
    python tools/quick_field_setup.py extracted.json --unconfigured
        """
    )
    
    parser.add_argument("input", help="输入JSON文件")
    parser.add_argument("--output", "-o", help="输出JSON文件 (默认覆盖输入)")
    
    # Source设置
    parser.add_argument("--source", choices=["metadata", "extracted_data", "calculated_data"],
                       help="设置数据源")
    
    # 计算字段设置
    parser.add_argument("--calculated", action="store_true",
                       help="标记为计算字段")
    parser.add_argument("--function", help="计算函数名")
    parser.add_argument("--args", help="参数列表 (逗号分隔)")
    
    # Table配置
    parser.add_argument("--table-config", action="store_true",
                       help="设置table配置")
    parser.add_argument("--table-template", help="Table模板路径")
    parser.add_argument("--row-strategy", choices=["fixed_rows", "dynamic_rows"],
                       help="Row策略")
    parser.add_argument("--header-rows", type=int, help="Header行数")
    
    # Image配置
    parser.add_argument("--image-config", action="store_true",
                       help="设置image配置")
    parser.add_argument("--width", type=float, help="图片宽度(英寸)")
    parser.add_argument("--alignment", choices=["left", "center", "right"],
                       help="图片对齐方式")
    
    # 字段列表
    parser.add_argument("--fields", required=True,
                       help="要设置的字段名 (逗号分隔)")
    
    # 类型设置
    parser.add_argument("--type", choices=["text", "checkbox", "table", "image"],
                       help="设置字段类型")
    
    # 统计和检查
    parser.add_argument("--stats", action="store_true",
                       help="显示配置统计")
    parser.add_argument("--unconfigured", action="store_true",
                       help="显示未配置的字段")
    
    args = parser.parse_args()
    
    if not os.path.exists(args.input):
        print(f"错误: 文件不存在: {args.input}")
        sys.exit(1)
    
    # 加载配置
    mappings, original_data = load_mappings(args.input)
    
    # 处理统计请求
    if args.stats:
        show_statistics(mappings)
        return
    
    if args.unconfigured:
        show_unconfigured(mappings)
        return
    
    # 解析字段列表
    fields = [f.strip() for f in args.fields.split(",")]
    
    # 执行配置操作
    if args.source:
        mappings = batch_set_source(mappings, fields, args.source)
    
    if args.calculated or args.function:
        if not args.function:
            print("错误: 计算字段必须指定 --function")
            sys.exit(1)
        args_list = []
        if args.args:
            args_list = [a.strip() for a in args.args.split(",")]
        mappings = batch_set_calculated(mappings, fields, args.function, args_list)
    
    if args.table_config:
        mappings = batch_set_table_config(
            mappings, fields,
            template_path=args.table_template,
            row_strategy=args.row_strategy,
            header_rows=args.header_rows
        )
    
    if args.image_config:
        mappings = batch_set_image_config(
            mappings, fields,
            width=args.width,
            alignment=args.alignment
        )
    
    if args.type:
        mappings = batch_set_type(mappings, fields, args.type)
    
    # 保存结果
    output_file = args.output or args.input.replace('.json', '_configured.json')
    save_mappings(mappings, original_data, output_file)
    
    # 显示统计
    show_statistics(mappings)


if __name__ == "__main__":
    main()
