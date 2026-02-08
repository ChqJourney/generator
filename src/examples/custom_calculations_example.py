"""
自定义计算函数示例模块

展示如何创建自定义计算函数并通过装饰器注册到计算器
使用方式:
    1. 复制此文件并重命名为 custom_calculations.py
    2. 添加你的自定义计算函数
    3. 在运行计算器时指定: --functions-module custom_calculations

注意：此模块被导入时会自动注册所有带 @CalculationRegistry.register 的函数
"""

from src.calculator import CalculationRegistry


@CalculationRegistry.register("calculate_lumen_per_watt")
def calculate_lumen_per_watt(lumen: float, watt: float) -> str:
    """
    计算流明每瓦（光效）
    
    Args:
        lumen: 光通量 (lm)
        watt: 功率 (W)
        
    Returns:
        str: 光效值 (lm/W)，保留1位小数
    """
    if not watt or watt == 0:
        return "0.0"
    
    try:
        efficacy = float(lumen) / float(watt)
        return f"{efficacy:.1f}"
    except (ValueError, TypeError):
        return "0.0"


@CalculationRegistry.register("calculate_power_factor")
def calculate_power_factor(active_power: float, apparent_power: float) -> str:
    """
    计算功率因数
    
    Args:
        active_power: 有功功率 (W)
        apparent_power: 视在功率 (VA)
        
    Returns:
        str: 功率因数，保留2位小数
    """
    if not apparent_power or apparent_power == 0:
        return "0.00"
    
    try:
        pf = float(active_power) / float(apparent_power)
        return f"{pf:.2f}"
    except (ValueError, TypeError):
        return "0.00"


@CalculationRegistry.register("format_with_unit")
def format_with_unit(value: float, unit: str, decimal_places: int = 2) -> str:
    """
    格式化带单位的值
    
    Args:
        value: 数值
        unit: 单位（如 'W', 'lm', 'K' 等）
        decimal_places: 小数位数
        
    Returns:
        str: 带单位的格式化字符串
    """
    try:
        formatted = f"{float(value):.{int(decimal_places)}f}"
        return f"{formatted} {unit}"
    except (ValueError, TypeError):
        return f"{value} {unit}"


@CalculationRegistry.register("calculate_cct_deviation")
def calculate_cct_deviation(measured_cct: float, rated_cct: float) -> str:
    """
    计算色温偏差百分比
    
    Args:
        measured_cct: 实测色温 (K)
        rated_cct: 额定色温 (K)
        
    Returns:
        str: 偏差百分比，带符号，保留1位小数
    """
    if not rated_cct or rated_cct == 0:
        return "0.0%"
    
    try:
        deviation = ((float(measured_cct) - float(rated_cct)) / float(rated_cct)) * 100
        sign = "+" if deviation > 0 else ""
        return f"{sign}{deviation:.1f}%"
    except (ValueError, TypeError):
        return "0.0%"


@CalculationRegistry.register("calculate_average")
def calculate_average(*values) -> str:
    """
    计算多个值的平均值
    
    Args:
        *values: 可变数量的数值
        
    Returns:
        str: 平均值，保留2位小数
    """
    if not values:
        return "0.00"
    
    try:
        numeric_values = [float(v) for v in values if v is not None]
        if not numeric_values:
            return "0.00"
        avg = sum(numeric_values) / len(numeric_values)
        return f"{avg:.2f}"
    except (ValueError, TypeError):
        return "0.00"


@CalculationRegistry.register("check_pass_fail")
def check_pass_fail(measured_value: float, min_limit: float, max_limit: float) -> str:
    """
    根据上下限判断测试结果
    
    Args:
        measured_value: 实测值
        min_limit: 下限
        max_limit: 上限
        
    Returns:
        str: "Pass" 或 "Fail"
    """
    try:
        val = float(measured_value)
        min_val = float(min_limit)
        max_val = float(max_limit)
        
        if min_val <= val <= max_val:
            return "Pass"
        else:
            return "Fail"
    except (ValueError, TypeError):
        return "N/A"


@CalculationRegistry.register("calculate_luminous_intensity")
def calculate_luminous_intensity(luminous_flux: float, angle_degrees: float) -> str:
    """
    计算发光强度（坎德拉）
    
    公式：I = Φ / Ω
    其中Ω是立体角：Ω = 2π(1 - cos(θ/2))
    
    Args:
        luminous_flux: 光通量 (lm)
        angle_degrees: 光束角 (度)
        
    Returns:
        str: 发光强度 (cd)，保留1位小数
    """
    import math
    
    try:
        flux = float(luminous_flux)
        angle = float(angle_degrees)
        
        if angle <= 0 or angle > 360:
            return "0.0"
        
        # 计算立体角（球面度）
        theta_half = math.radians(angle / 2)
        solid_angle = 2 * math.pi * (1 - math.cos(theta_half))
        
        if solid_angle == 0:
            return "0.0"
        
        intensity = flux / solid_angle
        return f"{intensity:.1f}"
    except (ValueError, TypeError):
        return "0.0"


# 以下是如何使用这些函数的示例配置:
EXAMPLE_CONFIG = """
{
  "field_mappings": [
    {
      "template_field": "efficacy_formatted",
      "source": "calculated_data",
      "source_field": "efficacy_formatted",
      "args": [
        "extracted_data|total_luminous_flux",
        "extracted_data|rated_wattage"
      ],
      "function": "calculate_lumen_per_watt",
      "type": "text"
    },
    {
      "template_field": "power_factor_display",
      "source": "calculated_data",
      "source_field": "power_factor_display",
      "args": [
        "extracted_data|active_power",
        "extracted_data|apparent_power"
      ],
      "function": "calculate_power_factor",
      "type": "text"
    },
    {
      "template_field": "wattage_with_unit",
      "source": "calculated_data",
      "source_field": "wattage_with_unit",
      "args": [
        "extracted_data|rated_wattage",
        "W"
      ],
      "function": "format_with_unit",
      "type": "text"
    },
    {
      "template_field": "test_result",
      "source": "calculated_data",
      "source_field": "test_result",
      "args": [
        "extracted_data|measured_value",
        "extracted_data|min_limit",
        "extracted_data|max_limit"
      ],
      "function": "check_pass_fail",
      "type": "text"
    }
  ]
}
"""
