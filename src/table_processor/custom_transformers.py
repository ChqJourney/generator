"""
专用表格数据转换器
为特定表格提供定制的数据转换逻辑
"""

import statistics
import re
from typing import List, Dict, Optional, Any, Callable
from dataclasses import dataclass
import sys
import os

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from utils.logging_config import get_logger

logger = get_logger(__name__)


# 格式化规则类
@dataclass
class FormatRule:
    """条件格式化规则"""
    condition: Callable[[float], bool]  # 条件函数
    format_str: str                    # 格式字符串，如 "{:.1f}"
    default_format: str = "{:.2f}"     # 默认格式
    
    def format(self, value: float) -> str:
        if self.condition(value):
            return self.format_str.format(value)
        return self.default_format.format(value)


# 专用转换器注册表
class CustomTransformerRegistry:
    """专用转换器注册表"""
    _transformers: Dict[str, Callable] = {}
    
    @classmethod
    def register(cls, name: str):
        """注册转换器装饰器"""
        def decorator(func: Callable):
            cls._transformers[name] = func
            return func
        return decorator
    
    @classmethod
    def transform(cls, name: str, data: Any, params: Dict, 
                  extracted_data: Optional[Dict] = None) -> List[List[Any]]:
        """执行转换"""
        if name not in cls._transformers:
            raise ValueError(f"Unknown transformer: {name}")
        return cls._transformers[name](data, params, extracted_data)


def format_number(value: Any, decimal: Optional[int] = None, 
                  format_rules: Optional[List[FormatRule]] = None) -> str:
    """
    格式化数字
    
    Args:
        value: 要格式化的值
        decimal: 固定小数位数
        format_rules: 条件格式化规则列表
    
    Returns:
        格式化后的字符串
    """
    try:
        num = float(value)
    except (ValueError, TypeError):
        return str(value) if value is not None else ''
    
    # 优先应用条件格式化规则
    if format_rules:
        for rule in format_rules:
            if rule.condition(num):
                return rule.format(num)
        # 没有匹配的规则，使用默认格式
        return format_rules[0].default_format.format(num)
    
    # 固定小数位
    if decimal is not None:
        return f"{num:.{decimal}f}"
    
    return str(num)


def parse_format_rules(rules_config: List[Dict]) -> List[FormatRule]:
    """解析格式化规则配置"""
    rules = []
    default_format = "{:.2f}"
    
    for rule in rules_config:
        condition_str = rule.get('condition', '')  # e.g., "x >= 100"
        format_str = rule.get('format', '{:.1f}')  # e.g., "{:.1f}"
        
        # 创建条件函数
        if condition_str:
            # 解析条件如 "x >= 100"
            match = re.match(r'x\s*([><=!]+)\s*([\d.]+)', condition_str)
            if match:
                op = match.group(1)
                threshold = float(match.group(2))
                
                def make_cond(o, t):
                    if o in ('>=', '=>'):
                        return lambda x: x >= t
                    elif o in ('<=', '=<'):
                        return lambda x: x <= t
                    elif o == '>':
                        return lambda x: x > t
                    elif o == '<':
                        return lambda x: x < t
                    elif o == '==':
                        return lambda x: x == t
                    else:
                        return lambda x: True
                
                cond = make_cond(op, threshold)
                rules.append(FormatRule(cond, format_str, default_format))
    
    return rules


# =============================================================================
# 1. Photometric Data 转换器
# =============================================================================

@CustomTransformerRegistry.register("photometric_data_transformer")
def photometric_data_transformer(data: List[List[Any]], params: Dict,
                                  extracted_data: Optional[Dict] = None) -> List[List[Any]]:
    """
    光度数据表格转换器
    
    - 数据源是二维数组
    - 指定列需要计算（公式）
    - 每列求平均数
    - 支持条件格式化
    
    配置示例:
    {
        "type": "custom_transform",
        "transformer": "photometric_data_transformer",
        "calculate_columns": [5, 6],  # 需要计算的列索引
        "formulas": {
            "5": "D{row}/C{row}*100",  # {row} 表示当前行号(1-based)
            "6": "E{row}/F{row}"
        },
        "average_columns": [2, 3, 4, 5, 6, 7, 8, 9, 10],  # 需要求平均的列
        "format_rules": {
            "4": [{"condition": "x >= 100", "format": "{:.1f}"}, {"condition": "x < 100", "format": "{:.2f}"}],
            "5": [{"condition": "x >= 100", "format": "{:.1f}"}, {"condition": "x < 100", "format": "{:.2f}"}]
        },
        "average_format_rules": {  # 平均行的格式化规则
            "4": [{"condition": "x >= 100", "format": "{:.1f}"}],
            "5": [{"condition": "x < 100", "format": "{:.2f}"}]
        }
    }
    """
    if not data:
        return []
    
    result = [row[:] for row in data]
    
    # 1. 执行列计算（公式）
    calculate_columns = params.get('calculate_columns', [])
    formulas = params.get('formulas', {})
    
    for col_idx in calculate_columns:
        formula_key = str(col_idx)
        if formula_key in formulas:
            formula = formulas[formula_key]
            for row_idx, row in enumerate(result):
                # 替换 {row} 为实际行号(1-based，因为Excel公式从1开始)
                actual_formula = formula.replace('{row}', str(row_idx + 1))
                
                # 替换列引用 A, B, C... 为实际值
                # 支持 A{row}, B{row} 格式
                def replace_col_ref(match):
                    col_letter = match.group(1)
                    col_num = ord(col_letter.upper()) - ord('A')  # A=0, B=1, ...
                    if col_num < len(row):
                        val = row[col_num]
                        try:
                            return str(float(val))
                        except (ValueError, TypeError):
                            return '0'
                    return '0'
                
                # 替换形如 A1, B2 的引用
                eval_formula = re.sub(r'([A-Z])\d+', replace_col_ref, actual_formula)
                
                try:
                    # 安全计算
                    value = eval(eval_formula, {"__builtins__": {}}, {})
                    # 确保行有足够列
                    while len(row) <= col_idx:
                        row.append('')
                    row[col_idx] = value
                except Exception as e:
                    logger.warning(f"Formula calculation error: {eval_formula}, {e}")
    
    # 2. 求平均数并添加平均行
    average_columns = params.get('average_columns', [])
    format_rules_config = params.get('format_rules', {})
    avg_format_rules_config = params.get('average_format_rules', {})
    
    if average_columns:
        avg_row = [''] * len(result[0]) if result else []
        avg_row[0] = 'Average'  # 第一列标记为 Average
        
        for col_idx in average_columns:
            values = []
            for row in result:
                if col_idx < len(row):
                    try:
                        val = float(row[col_idx])
                        values.append(val)
                    except (ValueError, TypeError):
                        pass
            
            if values:
                avg_value = statistics.mean(values)
                
                # 应用平均行的格式化规则
                col_key = str(col_idx)
                if col_key in avg_format_rules_config:
                    rules = parse_format_rules(avg_format_rules_config[col_key])
                    avg_row[col_idx] = format_number(avg_value, format_rules=rules)
                elif col_key in format_rules_config:
                    rules = parse_format_rules(format_rules_config[col_key])
                    avg_row[col_idx] = format_number(avg_value, format_rules=rules)
                else:
                    avg_row[col_idx] = format_number(avg_value, decimal=2)
        
        result.append(avg_row)
    
    # 3. 应用数据行的格式化规则
    for col_idx_str, rules_config in format_rules_config.items():
        col_idx = int(col_idx_str)
        rules = parse_format_rules(rules_config)
        
        for row in result[:-1] if average_columns else result:  # 不格式化平均行
            if col_idx < len(row):
                row[col_idx] = format_number(row[col_idx], format_rules=rules)
    
    return result


# =============================================================================
# 2. Beam Table 转换器
# =============================================================================

@CustomTransformerRegistry.register("beam_table_transformer")
def beam_table_transformer(data: Any, params: Dict,
                            extracted_data: Optional[Dict] = None) -> List[List[Any]]:
    """
    光束角表格转换器
    
    - 数据源是分散的字段
    - 填入固定位置：第二行第二列是 beam_angle，第三行第二列是 peak_intensity
    
    配置示例:
    {
        "type": "custom_transform",
        "transformer": "beam_table_transformer",
        "beam_angle_field": "beam_angle_0",
        "peak_intensity_field": "peak_intensity_0",
        "beam_angle_format": {:.1f},
        "peak_intensity_format": {:.0f}
    }
    """
    if not extracted_data:
        return []
    
    # 从 extracted_data 中获取字段值
    beam_angle_field = params.get('beam_angle_field', 'beam_angle')
    peak_intensity_field = params.get('peak_intensity_field', 'peak_intensity')
    
    beam_angle = extracted_data.get(beam_angle_field, '')
    peak_intensity = extracted_data.get(peak_intensity_field, '')
    
    # 格式化
    beam_format = params.get('beam_angle_format', '{:.1f}')
    intensity_format = params.get('peak_intensity_format', '{:.0f}')
    
    try:
        beam_angle_str = beam_format.format(float(beam_angle))
    except (ValueError, TypeError):
        beam_angle_str = str(beam_angle)
    
    try:
        intensity_str = intensity_format.format(float(peak_intensity))
    except (ValueError, TypeError):
        intensity_str = str(peak_intensity)
    
    # 构建表格数据：假设表格有3行，第一行是表头
    # 第二行第二列填 beam_angle，第三行第二列填 peak_intensity
    result = [
        ['', ''],  # 第一行（表头，会被保留）
        ['', beam_angle_str],  # 第二行第二列
        ['', intensity_str]    # 第三行第二列
    ]
    
    return result


# =============================================================================
# 3. EEI Table 转换器
# =============================================================================

@CustomTransformerRegistry.register("eei_table_transformer")
def eei_table_transformer(data: Any, params: Dict,
                          extracted_data: Optional[Dict] = None) -> List[List[Any]]:
    """
    EEI 能效表格转换器
    
    - 多行（多型号）
    - 每行：型号名、efficacy平均值、eei class
    - efficacy 从 photometric_data 的 efficacy 列计算平均值
    - eei class 由 efficacy 计算所得
    - 最后两列需要行合并标记
    
    配置示例:
    {
        "type": "custom_transform",
        "transformer": "eei_table_transformer",
        "model_fields": ["model_1", "model_2", "model_3"],  # 型号字段列表
        "photometric_data_ref": "photometric_data",  # 引用光度数据
        "efficacy_column": 5,  # efficacy 在 photometric_data 中的列索引
        "eei_thresholds": {
            "A++": 130,
            "A+": 110,
            "A": 90,
            "B": 70,
            "C": 50,
            "D": 30
        },
        "merge_columns": [3, 4],  # 需要合并的最后两列索引
        "format_rules": {
            "1": [{"condition": "x >= 100", "format": "{:.1f}"}]
        }
    }
    """
    if not extracted_data:
        return []
    
    model_fields = params.get('model_fields', [])
    photometric_ref = params.get('photometric_data_ref', 'photometric_data')
    efficacy_col = params.get('efficacy_column', 5)
    eei_thresholds = params.get('eei_thresholds', {
        "A++": 130, "A+": 110, "A": 90, "B": 70, "C": 50, "D": 30
    })
    merge_columns = params.get('merge_columns', [3, 4])
    format_rules_config = params.get('format_rules', {})
    
    # 获取 photometric_data 计算 efficacy 平均值
    photometric_data = extracted_data.get(photometric_ref, [])
    efficacy_values = []
    for row in photometric_data:
        if isinstance(row, list) and len(row) > efficacy_col:
            try:
                val = float(row[efficacy_col])
                efficacy_values.append(val)
            except (ValueError, TypeError):
                pass
    
    avg_efficacy = statistics.mean(efficacy_values) if efficacy_values else 0
    
    # 计算 EEI Class
    def calculate_eei_class(efficacy: float) -> str:
        for cls, threshold in sorted(eei_thresholds.items(), key=lambda x: x[1], reverse=True):
            if efficacy >= threshold:
                return cls
        return "E"
    
    eei_class = calculate_eei_class(avg_efficacy)
    
    # 构建结果表格
    result = []
    
    # 格式化 efficacy
    col_key = "1"  # 第二列（efficacy）
    if col_key in format_rules_config:
        rules = parse_format_rules(format_rules_config[col_key])
        efficacy_str = format_number(avg_efficacy, format_rules=rules)
    else:
        efficacy_str = format_number(avg_efficacy, decimal=1)
    
    # 为每个型号创建一行
    for model_field in model_fields:
        model_name = extracted_data.get(model_field, '')
        if model_name:  # 只添加有型号名的行
            row = [
                model_name,      # 第一列：型号名
                efficacy_str,    # 第二列：efficacy 平均值
                eei_class        # 第三列：eei class
            ]
            # 添加合并标记列（最后两列）
            for col_idx in merge_columns:
                while len(row) <= col_idx:
                    row.append('')
                # 使用特殊标记表示需要合并
                row[col_idx] = "__MERGE__"
            
            result.append(row)
    
    # 如果没有型号数据，至少创建一行
    if not result:
        row = ['', efficacy_str, eei_class]
        for col_idx in merge_columns:
            while len(row) <= col_idx:
                row.append('')
            row[col_idx] = "__MERGE__"
        result.append(row)
    
    return result


# =============================================================================
# 4. Zone Table 转换器
# =============================================================================

@CustomTransformerRegistry.register("zone_table_transformer")
def zone_table_transformer(data: Any, params: Dict,
                           extracted_data: Optional[Dict] = None) -> List[List[Any]]:
    """
    区域光强表格转换器
    
    - 数据源是分散的字段：0-30, 0-60, 0-90, 0-180 等
    - 字段数量不定
    - 最大角度必须比 beam_table 中的 beam_angle 大
    - 动态增减行
    
    配置示例:
    {
        "type": "custom_transform",
        "transformer": "zone_table_transformer",
        "zone_fields_pattern": "zone_{angle}",  # 字段名模式
        "zone_angles": [30, 60, 90, 120, 150, 180],  # 可选的角度列表
        "beam_angle_field": "beam_angle_0",  # beam_angle 字段引用
        "min_angle": 30,  # 最小角度
        "max_angle_override": 180,  # 强制最大角度
        "format": "{:.1f}"
    }
    """
    if not extracted_data:
        return []
    
    zone_angles = params.get('zone_angles', [30, 60, 90, 120, 150, 180])
    beam_angle_field = params.get('beam_angle_field', 'beam_angle')
    format_str = params.get('format', '{:.1f}')
    min_angle = params.get('min_angle', 30)
    max_angle_override = params.get('max_angle_override')
    
    # 获取 beam_angle
    beam_angle = 0
    try:
        beam_angle = float(extracted_data.get(beam_angle_field, 0))
    except (ValueError, TypeError):
        pass
    
    # 确定最大角度
    if max_angle_override:
        max_angle = max_angle_override
    else:
        # 最大角度必须比 beam_angle 大
        max_angle = max(beam_angle * 1.2, 180)  # 至少比 beam_angle 大 20% 或 180
    
    # 筛选有效的 zone 角度
    valid_zones = []
    for angle in zone_angles:
        if min_angle <= angle <= max_angle:
            field_name = f"zone_{angle}"  # 例如 zone_30, zone_60
            value = extracted_data.get(field_name)
            if value is not None and value != '':
                try:
                    formatted_value = format_str.format(float(value))
                except (ValueError, TypeError):
                    formatted_value = str(value)
                valid_zones.append((angle, formatted_value))
    
    # 构建表格数据：每行一个 zone
    result = []
    for angle, value in valid_zones:
        row = [
            f"0-{angle}°",  # 第一列：角度范围
            value           # 第二列：光强值
        ]
        result.append(row)
    
    return result


# =============================================================================
# 5. Life Table 转换器
# =============================================================================

@CustomTransformerRegistry.register("life_table_transformer")
def life_table_transformer(data: List[List[Any]], params: Dict,
                           extracted_data: Optional[Dict] = None) -> List[List[Any]]:
    """
    寿命测试表格转换器
    
    - 类似 photometric_data
    - 数据源是二维数组
    - 部分列需要计算
    - 每列需要平均数
    
    配置示例:
    {
        "type": "custom_transform",
        "transformer": "life_table_transformer",
        "calculate_columns": [3],
        "formulas": {
            "3": "B{row}*2"
        },
        "average_columns": [1, 2, 3, 4],
        "format_rules": {
            "1": [{"condition": "x >= 100", "format": "{:.1f}"}],
            "2": [{"condition": "x >= 1000", "format": "{:.0f}"}, {"condition": "x < 1000", "format": "{:.1f}"}]
        }
    }
    """
    # 直接复用 photometric_data_transformer 的逻辑
    return photometric_data_transformer(data, params, extracted_data)
