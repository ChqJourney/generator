#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
交互式配置向导 - 快速完成report_config.json的创建

用法:
    python tools/config_wizard.py config/extracted_elements.json
    
功能:
    1. 批量确认字段类型和数据源
    2. 快速设置计算字段的函数和参数
    3. 验证配置完整性
"""

import json
import sys
import os
from typing import Dict, List, Any


class Colors:
    """终端颜色"""
    HEADER = '\033[95m'
    BLUE = '\033[94m'
    CYAN = '\033[96m'
    GREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    END = '\033[0m'
    BOLD = '\033[1m'


def print_header(text: str):
    print(f"\n{Colors.HEADER}{Colors.BOLD}{'='*60}")
    print(f"  {text}")
    print(f"{'='*60}{Colors.END}\n")


def print_info(text: str):
    print(f"{Colors.BLUE}ℹ {text}{Colors.END}")


def print_success(text: str):
    print(f"{Colors.GREEN}✓ {text}{Colors.END}")


def print_warning(text: str):
    print(f"{Colors.WARNING}⚠ {text}{Colors.END}")


def print_error(text: str):
    print(f"{Colors.FAIL}✗ {text}{Colors.END}")


def input_default(prompt: str, default: str = "") -> str:
    """带默认值的输入"""
    if default:
        result = input(f"{prompt} [{default}]: ").strip()
        return result if result else default
    return input(f"{prompt}: ").strip()


def input_choice(prompt: str, choices: List[str], default: str = None) -> str:
    """选择题输入"""
    choices_str = "/".join([
        f"[{c[0].upper()}]{c[1:]}" if c == default else c
        for c in choices
    ])
    while True:
        result = input(f"{prompt} ({choices_str}): ").strip().lower()
        if not result and default:
            return default
        for choice in choices:
            if result == choice.lower() or result == choice[0].lower():
                return choice
        print_error(f"无效输入，请从 {choices} 中选择")


def batch_edit_fields(mappings: List[Dict]) -> List[Dict]:
    """
    批量编辑字段配置
    """
    print_header("批量字段配置")
    print_info(f"共有 {len(mappings)} 个字段需要配置\n")
    
    # 按类型分组
    text_fields = [m for m in mappings if m.get("type") == "text"]
    checkbox_fields = [m for m in mappings if m.get("type") == "checkbox"]
    table_fields = [m for m in mappings if m.get("type") == "table"]
    image_fields = [m for m in mappings if m.get("type") == "image"]
    
    print(f"字段统计:")
    print(f"  - Text字段: {len(text_fields)}")
    print(f"  - Checkbox字段: {len(checkbox_fields)}")
    print(f"  - Table字段: {len(table_fields)}")
    print(f"  - Image字段: {len(image_fields)}")
    
    # 快速模式 - 批量设置同类型字段
    print_header("快速批量设置")
    
    # 批量设置text字段的数据源
    if text_fields:
        print(f"\n{Colors.CYAN}Text字段批量设置:{Colors.END}")
        print(f"检测到 {len(text_fields)} 个text字段")
        
        # 显示前5个作为示例
        for i, field in enumerate(text_fields[:5]):
            print(f"  {i+1}. {field['template_field']}")
        if len(text_fields) > 5:
            print(f"  ... 还有 {len(text_fields)-5} 个")
        
        source = input_choice(
            "\n这些字段主要来源于",
            ["metadata", "extracted_data", "calculated_data"],
            "extracted_data"
        )
        
        if input_choice(f"将所有text字段的source设为 {source}?", ["yes", "no"], "yes") == "yes":
            for field in text_fields:
                field["source_field"] = f"{source}.{field['template_field']}"
                field["inferred_source"] = source
            print_success(f"已批量设置 {len(text_fields)} 个字段的source为 {source}")
    
    # 逐个确认计算字段
    calculated_fields = [m for m in mappings if m.get("is_calculated") or 
                        m.get("inferred_source") == "calculated_data"]
    
    if calculated_fields:
        print_header("计算字段配置")
        print_info(f"检测到 {len(calculated_fields)} 个可能的计算字段\n")
        
        for field in calculated_fields:
            name = field['template_field']
            print(f"\n{Colors.CYAN}字段: {name}{Colors.END}")
            print(f"  上下文: {field.get('context', 'N/A')[:60]}")
            
            if input_choice(f"  是否为计算字段?", ["yes", "no"], 
                           "yes" if field.get("is_calculated") else "no") == "yes":
                field["is_calculated"] = True
                field["source_field"] = f"calculated_data.{name}"
                
                # 选择或输入函数名
                suggested = field.get("suggested_function", "")
                func = input_default("  计算函数名", suggested)
                if func:
                    field["function"] = func
                    
                    # 输入参数
                    args = field.get("suggested_args", [])
                    args_str = ", ".join(args) if args else ""
                    args_input = input_default("  参数 (逗号分隔)", args_str)
                    if args_input:
                        field["args"] = [a.strip() for a in args_input.split(",")]
            else:
                field["is_calculated"] = False
                field["source_field"] = field["source_field"].replace("calculated_data.", "extracted_data.")
    
    # Table字段配置
    if table_fields:
        print_header("Table字段配置")
        for field in table_fields:
            name = field['template_field']
            print(f"\n{Colors.CYAN}Table: {name}{Colors.END}")
            
            # Table template路径
            current_path = field.get("table_template_path", "report_templates/table_template.docx")
            new_path = input_default("  Table模板路径", current_path)
            field["table_template_path"] = new_path
            
            # Row strategy
            strategy = input_choice("  Row策略", ["fixed_rows", "dynamic_rows"], 
                                   field.get("row_strategy", "fixed_rows"))
            field["row_strategy"] = strategy
            
            # Header rows
            header_rows = input_default("  Header行数", str(field.get("header_rows", 1)))
            field["header_rows"] = int(header_rows)
    
    # Image字段配置
    if image_fields:
        print_header("Image字段配置")
        for field in image_fields:
            name = field['template_field']
            print(f"\n{Colors.CYAN}Image: {name}{Colors.END}")
            
            width = input_default("  宽度(英寸)", str(field.get("width", 4.0)))
            field["width"] = float(width)
            
            alignment = input_choice("  对齐方式", ["left", "center", "right"],
                                    field.get("alignment", "center"))
            field["alignment"] = alignment
    
    return mappings


def review_and_fix(mappings: List[Dict]) -> List[Dict]:
    """
    审查并修复配置
    """
    print_header("配置审查")
    
    issues = []
    
    for i, field in enumerate(mappings):
        # 检查必填字段
        if not field.get("template_field"):
            issues.append((i, "缺少template_field", "error"))
        if not field.get("source_field"):
            issues.append((i, f"字段 {field.get('template_field', 'N/A')} 缺少source_field", "error"))
        if not field.get("type"):
            issues.append((i, f"字段 {field.get('template_field', 'N/A')} 缺少type", "error"))
        
        # 检查计算字段
        if field.get("is_calculated"):
            if not field.get("function"):
                issues.append((i, f"计算字段 {field['template_field']} 缺少function", "warning"))
            if not field.get("args"):
                issues.append((i, f"计算字段 {field['template_field']} 缺少args", "warning"))
        
        # 检查table字段
        if field.get("type") == "table":
            if not field.get("table_template_path"):
                issues.append((i, f"Table字段 {field['template_field']} 缺少table_template_path", "warning"))
    
    if not issues:
        print_success("配置检查通过，没有发现明显问题")
        return mappings
    
    print_warning(f"发现 {len(issues)} 个问题:\n")
    
    for idx, message, level in issues:
        prefix = "[错误]" if level == "error" else "[警告]"
        color = Colors.FAIL if level == "error" else Colors.WARNING
        print(f"{color}{prefix} {message}{Colors.END}")
    
    if input_choice("\n是否修复这些问题?", ["yes", "no"], "yes") == "yes":
        for idx, message, level in issues:
            field = mappings[idx]
            name = field.get("template_field", "N/A")
            print(f"\n{Colors.CYAN}修复: {name}{Colors.END}")
            print(f"  问题: {message}")
            
            if "缺少source_field" in message:
                source = input_default("  请输入source_field", f"extracted_data.{name}")
                field["source_field"] = source
            elif "缺少function" in message:
                func = input("  请输入function名称: ")
                field["function"] = func
            elif "缺少args" in message:
                args_input = input("  请输入参数 (逗号分隔): ")
                field["args"] = [a.strip() for a in args_input.split(",") if a.strip()]
            elif "缺少table_template_path" in message:
                path = input_default("  Table模板路径", "report_templates/table_template.docx")
                field["table_template_path"] = path
    
    return mappings


def clean_mappings(mappings: List[Dict]) -> List[Dict]:
    """
    清理mappings，移除辅助字段，保留最终配置
    """
    cleaned = []
    
    for field in mappings:
        clean_field = {
            "template_field": field["template_field"],
            "source_field": field["source_field"],
            "type": field["type"]
        }
        
        # 添加计算字段特有的属性
        if field.get("is_calculated"):
            clean_field["function"] = field.get("function", "")
            if field.get("args"):
                clean_field["args"] = field["args"]
        
        # 添加table特有的属性
        if field["type"] == "table":
            if field.get("table_template_path"):
                clean_field["table_template_path"] = field["table_template_path"]
            if field.get("row_strategy"):
                clean_field["row_strategy"] = field["row_strategy"]
            if field.get("header_rows") is not None:
                clean_field["header_rows"] = field["header_rows"]
            if field.get("skip_columns"):
                clean_field["skip_columns"] = field["skip_columns"]
            if field.get("transformations"):
                clean_field["transformations"] = field["transformations"]
        
        # 添加image特有的属性
        if field["type"] == "image":
            if field.get("width"):
                clean_field["width"] = field["width"]
            if field.get("height"):
                clean_field["height"] = field["height"]
            if field.get("alignment"):
                clean_field["alignment"] = field["alignment"]
        
        cleaned.append(clean_field)
    
    return cleaned


def generate_calculator_functions(mappings: List[Dict]) -> str:
    """
    生成calculator.py中需要的函数模板
    """
    calculated = [m for m in mappings if m.get("is_calculated") and m.get("function")]
    
    if not calculated:
        return ""
    
    lines = ["# 在 src/calculator.py 或 src/custom_calculations.py 中添加以下函数:\n"]
    
    for field in calculated:
        func_name = field["function"]
        args = field.get("args", [])
        arg_names = [a.split(".")[-1] for a in args]
        
        lines.append(f"""
@CalculationRegistry.register("{func_name}")
def {func_name}({', '.join(arg_names)}):
    \"\"\"
    计算: {field['template_field']}
    参数: {', '.join(args)}
    \"\"\"
    # TODO: 实现计算逻辑
    result = 0  # 替换为实际计算
    return f"{{result:.2f}}"
""")
    
    return "\n".join(lines)


def main():
    if len(sys.argv) < 2:
        print("用法: python tools/config_wizard.py <extracted_elements.json>")
        print("      python tools/config_wizard.py config/extracted_elements.json")
        sys.exit(1)
    
    input_file = sys.argv[1]
    
    if not os.path.exists(input_file):
        print_error(f"文件不存在: {input_file}")
        sys.exit(1)
    
    # 读取提取的元素
    with open(input_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    mappings = data.get("field_mappings", [])
    if not mappings:
        # 如果没有现成的field_mappings，从placeholders和checkboxes生成
        from extract_template_elements import generate_field_mappings
        mappings = generate_field_mappings(
            data.get("placeholders", {}),
            data.get("checkboxes", [])
        )
    
    print_header("Report Config 配置向导")
    print_info("此向导将帮助你快速完成report_config.json的配置\n")
    
    # 批量编辑
    mappings = batch_edit_fields(mappings)
    
    # 审查
    mappings = review_and_fix(mappings)
    
    # 清理配置
    cleaned_mappings = clean_mappings(mappings)
    
    # 生成最终配置
    config = {
        "template_path": data.get("template_path", ""),
        "output_dir": "output/",
        "field_mappings": cleaned_mappings,
        "merge_strategy": "prefer_first"
    }
    
    # 保存配置
    output_file = input_file.replace('.json', '_final.json')
    output_file = input_default("\n保存配置文件为", output_file)
    
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
    
    print_success(f"配置已保存: {output_file}")
    
    # 生成calculator函数模板
    calc_code = generate_calculator_functions(mappings)
    if calc_code:
        calc_file = output_file.replace('.json', '_calculator_functions.py')
        with open(calc_file, 'w', encoding='utf-8') as f:
            f.write(calc_code)
        print_success(f"计算函数模板已生成: {calc_file}")
    
    print_header("完成")
    print_info("下一步:")
    print("  1. 检查生成的配置文件")
    print("  2. 将计算函数模板复制到 calculator.py 或 custom_calculations.py")
    print("  3. 实现计算函数的具体逻辑")
    print("  4. 运行测试验证配置")


if __name__ == "__main__":
    # 检查Windows终端颜色支持
    if sys.platform == "win32":
        os.system("color")
    main()
