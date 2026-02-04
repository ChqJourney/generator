#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
计算函数生成器 - 根据report_config自动生成calculator函数

用法:
    # 生成所有计算函数
    python tools/generate_calculator_functions.py config/report_config.json
    
    # 生成并自动添加到calculator.py
    python tools/generate_calculator_functions.py config/report_config.json --append-to src/calculator.py
    
    # 生成到指定文件
    python tools/generate_calculator_functions.py config/report_config.json --output src/custom_calculations.py
"""

import json
import argparse
import sys
import os
import re
from typing import List, Dict, Tuple


# 常见计算函数的模板库
FUNCTION_TEMPLATES = {
    "calculate_energy_class_rating": """
@CalculationRegistry.register("calculate_energy_class_rating")
def calculate_energy_class_rating(wattage, flux):
    \"\"\"
    计算能效等级 (A++ 到 E)
    
    根据欧盟能效标准:
    η = 光通量(lm) / 功率(W)
    
    Args:
        wattage: 额定功率 (W)
        flux: 有用光通量 (lm)
    
    Returns:
        str: 能效等级 A++, A+, A, B, C, D, E
    \"\"\"
    try:
        w = float(wattage) if wattage else 0
        f = float(flux) if flux else 0
        
        if w <= 0:
            return "N/A"
        
        efficacy = f / w  # lm/W
        
        # 根据欧盟能效标签标准 (光源)
        if efficacy >= 210:
            return "A++"
        elif efficacy >= 185:
            return "A+"
        elif efficacy >= 160:
            return "A"
        elif efficacy >= 135:
            return "B"
        elif efficacy >= 110:
            return "C"
        elif efficacy >= 85:
            return "D"
        else:
            return "E"
    except (ValueError, TypeError):
        return "N/A"
""",
    
    "calculate_energy_efficacy": """
@CalculationRegistry.register("calculate_energy_efficacy")
def calculate_energy_efficacy(wattage, flux):
    \"\"\"
    计算能效值 (lm/W)
    
    Args:
        wattage: 额定功率 (W)
        flux: 有用光通量 (lm)
    
    Returns:
        str: 格式化的能效值，保留2位小数
    \"\"\"
    try:
        w = float(wattage) if wattage else 0
        f = float(flux) if flux else 0
        
        if w <= 0:
            return "0.00"
        
        efficacy = f / w
        return f"{efficacy:.2f}"
    except (ValueError, TypeError):
        return "0.00"
""",
    
    "calculate_percentage": """
@CalculationRegistry.register("calculate_percentage")
def calculate_percentage(value, total, decimal=2):
    \"\"\"
    计算百分比
    
    Args:
        value: 部分值
        total: 总值
        decimal: 小数位数 (默认2)
    
    Returns:
        str: 百分比值
    \"\"\"
    try:
        v = float(value) if value else 0
        t = float(total) if total else 0
        
        if t <= 0:
            return "0.00"
        
        percentage = (v / t) * 100
        return f"{percentage:.{decimal}f}"
    except (ValueError, TypeError):
        return "0.00"
""",
    
    "calculate_ratio": """
@CalculationRegistry.register("calculate_ratio")
def calculate_ratio(numerator, denominator, decimal=2):
    \"\"\"
    计算比率
    
    Args:
        numerator: 分子
        denominator: 分母
        decimal: 小数位数 (默认2)
    
    Returns:
        str: 比率值
    \"\"\"
    try:
        n = float(numerator) if numerator else 0
        d = float(denominator) if denominator else 0
        
        if d <= 0:
            return "0.00"
        
        ratio = n / d
        return f"{ratio:.{decimal}f}"
    except (ValueError, TypeError):
        return "0.00"
""",
    
    "format_number": """
@CalculationRegistry.register("format_number")
def format_number(value, decimal=2):
    \"\"\"
    格式化数字
    
    Args:
        value: 数值
        decimal: 小数位数 (默认2)
    
    Returns:
        str: 格式化后的数字
    \"\"\"
    try:
        v = float(value) if value else 0
        return f"{v:.{decimal}f}"
    except (ValueError, TypeError):
        return f"0.{ '0' * decimal }"
""",
    
    "concat": """
@CalculationRegistry.register("concat")
def concat(*values, separator=" "):
    \"\"\"
    连接多个字符串
    
    Args:
        *values: 要连接的字符串
        separator: 分隔符 (默认空格)
    
    Returns:
        str: 连接后的字符串
    \"\"\"
    str_values = [str(v) for v in values if v is not None]
    return separator.join(str_values)
""",
    
    "calculate_sum": """
@CalculationRegistry.register("calculate_sum")
def calculate_sum(*values):
    \"\"\"
    计算多个值的和
    
    Args:
        *values: 数值
    
    Returns:
        str: 和，保留2位小数
    \"\"\"
    try:
        total = sum(float(v) if v else 0 for v in values)
        return f"{total:.2f}"
    except (ValueError, TypeError):
        return "0.00"
""",
    
    "calculate_average": """
@CalculationRegistry.register("calculate_average")
def calculate_average(*values):
    \"\"\"
    计算多个值的平均值
    
    Args:
        *values: 数值
    
    Returns:
        str: 平均值，保留2位小数
    \"\"\"
    try:
        valid_values = [float(v) for v in values if v is not None]
        if not valid_values:
            return "0.00"
        avg = sum(valid_values) / len(valid_values)
        return f"{avg:.2f}"
    except (ValueError, TypeError):
        return "0.00"
""",
}


def generate_generic_function(func_name: str, args: List[str]) -> str:
    """
    生成通用计算函数模板
    """
    arg_names = [a.split(".")[-1] for a in args]
    arg_list = ", ".join(arg_names)
    
    # 根据函数名推断可能的逻辑
    hint = ""
    if "sum" in func_name.lower():
        hint = "# 实现求和逻辑: result = sum of args"
    elif "avg" in func_name.lower() or "mean" in func_name.lower():
        hint = "# 实现平均逻辑: result = sum / count"
    elif "diff" in func_name.lower():
        hint = "# 实现差值逻辑: result = arg1 - arg2"
    elif "mult" in func_name.lower() or "product" in func_name.lower():
        hint = "# 实现乘积逻辑: result = arg1 * arg2"
    elif "div" in func_name.lower() or "ratio" in func_name.lower():
        hint = "# 实现除法逻辑: result = arg1 / arg2 (注意检查除零)"
    elif "percent" in func_name.lower():
        hint = "# 实现百分比逻辑: result = (arg1 / arg2) * 100"
    else:
        hint = "# TODO: 实现计算逻辑"
    
    return f'''
@CalculationRegistry.register("{func_name}")
def {func_name}({arg_list}):
    """
    计算函数: {func_name}
    
    Args:
        {chr(10).join([f"        {name}: 参数{i+1}" for i, name in enumerate(arg_names)])}
    
    Returns:
        str: 计算结果
    """
    try:
        # 转换为数值
        {chr(10).join([f"        {name} = float({name}) if {name} else 0" for name in arg_names])}
        
        {hint}
        result = 0  # 替换为实际计算
        
        return f"{{result:.2f}}"
    except (ValueError, TypeError):
        return "0.00"
'''


def extract_calculated_fields(config_file: str) -> List[Dict]:
    """从配置中提取计算字段"""
    with open(config_file, 'r', encoding='utf-8') as f:
        config = json.load(f)
    
    calculated = []
    for mapping in config.get("field_mappings", []):
        if mapping.get("is_calculated") or mapping.get("function"):
            calculated.append({
                "field": mapping["template_field"],
                "function": mapping.get("function", ""),
                "args": mapping.get("args", [])
            })
    
    return calculated


def generate_functions(calculated_fields: List[Dict], 
                       use_templates: bool = True) -> Tuple[str, List[str]]:
    """
    生成计算函数代码
    
    Returns:
        (代码, 需要手动实现的函数列表)
    """
    code_lines = ["# 自动生成的计算函数", "# 请将这些函数添加到 calculator.py 或 custom_calculations.py", ""]
    code_lines.append("from src.calculator import CalculationRegistry")
    code_lines.append("")
    
    manual_needed = []
    generated_funcs = set()
    
    for field in calculated_fields:
        func_name = field["function"]
        args = field["args"]
        
        if not func_name or func_name in generated_funcs:
            continue
        
        generated_funcs.add(func_name)
        
        # 尝试使用模板
        if use_templates and func_name in FUNCTION_TEMPLATES:
            code_lines.append(FUNCTION_TEMPLATES[func_name])
        else:
            # 生成通用模板
            code_lines.append(generate_generic_function(func_name, args))
            manual_needed.append(func_name)
    
    return "\n".join(code_lines), manual_needed


def append_to_calculator(code: str, calculator_file: str):
    """将代码追加到calculator.py"""
    # 检查文件是否存在
    if not os.path.exists(calculator_file):
        print(f"错误: 文件不存在: {calculator_file}")
        return False
    
    # 读取原文件
    with open(calculator_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 检查是否已有相同的函数
    existing_funcs = set(re.findall(r'@CalculationRegistry\.register\(["\']([^"\']+)', content))
    new_funcs = set(re.findall(r'@CalculationRegistry\.register\(["\']([^"\']+)', code))
    
    duplicates = existing_funcs & new_funcs
    if duplicates:
        print(f"警告: 以下函数已存在，将被跳过: {', '.join(duplicates)}")
        # 移除重复的函数
        for func_name in duplicates:
            pattern = rf'@CalculationRegistry\.register\(["\']{func_name}["\']\).*?(?=\n@CalculationRegistry\.register|\Z)'
            code = re.sub(pattern, '', code, flags=re.DOTALL)
    
    # 追加代码
    with open(calculator_file, 'a', encoding='utf-8') as f:
        f.write("\n\n# === Auto-generated calculation functions ===\n")
        f.write(code)
    
    print(f"函数已追加到: {calculator_file}")
    return True


def main():
    parser = argparse.ArgumentParser(
        description="计算函数生成器",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
    # 生成函数代码并打印
    python tools/generate_calculator_functions.py config/report_config.json
    
    # 生成到文件
    python tools/generate_calculator_functions.py config/report_config.json \\
        --output src/custom_calculations.py
    
    # 追加到calculator.py
    python tools/generate_calculator_functions.py config/report_config.json \\
        --append-to src/calculator.py
        """
    )
    
    parser.add_argument("config", help="report_config.json文件路径")
    parser.add_argument("--output", "-o", help="输出Python文件")
    parser.add_argument("--append-to", help="追加到现有calculator.py文件")
    parser.add_argument("--no-templates", action="store_true",
                       help="不使用内置模板，全部生成通用函数")
    
    args = parser.parse_args()
    
    if not os.path.exists(args.config):
        print(f"错误: 配置文件不存在: {args.config}")
        sys.exit(1)
    
    # 提取计算字段
    calculated_fields = extract_calculated_fields(args.config)
    
    if not calculated_fields:
        print("配置中没有找到计算字段")
        return
    
    print(f"找到 {len(calculated_fields)} 个计算字段:")
    for field in calculated_fields:
        args_str = ", ".join(field["args"])
        print(f"  - {field['field']}: {field['function']}({args_str})")
    
    # 生成函数代码
    code, manual_needed = generate_functions(
        calculated_fields,
        use_templates=not args.no_templates
    )
    
    # 输出或保存
    if args.append_to:
        append_to_calculator(code, args.append_to)
    elif args.output:
        with open(args.output, 'w', encoding='utf-8') as f:
            f.write(code)
        print(f"\n函数已保存到: {args.output}")
    else:
        print("\n" + "="*60)
        print("生成的代码:")
        print("="*60)
        print(code)
    
    # 提示需要手动实现的函数
    if manual_needed:
        print("\n以下函数使用通用模板，请手动实现具体逻辑:")
        for func_name in manual_needed:
            print(f"  - {func_name}")


if __name__ == "__main__":
    main()
