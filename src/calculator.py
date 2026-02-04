"""
根据配置文件计算字段值的计算器模块
支持从report.json分层数据结构获取数据，并根据配置进行计算
"""

import json
import re
from typing import Dict, Any, List, Optional, Callable, Union
from dataclasses import dataclass
import importlib
import sys

# 导入公共工具
from utils.path_navigator import DataNavigator


@dataclass
class FieldValue:
    """字段值数据类"""
    value: Any
    source: str
    field_name: str
    
    def __post_init__(self):
        """初始化后转换数值类型"""
        if isinstance(self.value, str):
            self.value = self._convert_value(self.value)
    
    @staticmethod
    def _convert_value(value: str) -> Union[str, float, int]:
        """尝试将字符串转换为数值"""
        if not value or not isinstance(value, str):
            return value
        
        # 尝试转换为整数
        try:
            return int(value)
        except ValueError:
            pass
        
        # 尝试转换为浮点数
        try:
            return float(value)
        except ValueError:
            pass
        
        return value


class CalculatorError(Exception):
    """计算器错误"""
    pass


class FieldNotFoundError(CalculatorError):
    """字段未找到错误"""
    def __init__(self, field_path: str, available_fields: Dict = None):
        self.field_path = field_path
        self.available_fields = available_fields
        msg = f"Field not found: {field_path}"
        if available_fields:
            msg += f"\nAvailable fields: {list(available_fields.keys())}"
        super().__init__(msg)


class FunctionNotFoundError(CalculatorError):
    """函数未找到错误"""
    def __init__(self, function_name: str):
        self.function_name = function_name
        super().__init__(f"Function not found: {function_name}")


class CalculationRegistry:
    """计算函数注册表"""
    
    _functions: Dict[str, Callable] = {}
    
    @classmethod
    def register(cls, name: str, func: Callable = None):
        """注册计算函数
        
        可以作为装饰器使用:
            @CalculationRegistry.register("my_function")
            def my_function(args):
                ...
        
        或直接调用:
            CalculationRegistry.register("my_function", my_function)
        """
        if func is None:
            # 作为装饰器使用
            def decorator(f):
                cls._functions[name] = f
                return f
            return decorator
        else:
            cls._functions[name] = func
    
    @classmethod
    def get(cls, name: str) -> Optional[Callable]:
        """获取计算函数"""
        return cls._functions.get(name)
    
    @classmethod
    def list_functions(cls) -> List[str]:
        """列出所有已注册的函数"""
        return list(cls._functions.keys())
    
    @classmethod
    def load_from_module(cls, module_path: str):
        """从模块加载计算函数"""
        try:
            module = importlib.import_module(module_path)
            # 模块中的函数会自动通过装饰器注册
        except ImportError as e:
            raise CalculatorError(f"Failed to load module {module_path}: {e}")


class FieldCalculator:
    """字段计算器 - 适应新架构"""
    
    def __init__(self, 
                 report_data: Dict,
                 config: Optional[Dict] = None):
        """
        初始化字段计算器
        
        Args:
            report_data: 分层报告数据，包含metadata/extracted_data/calculated_data
            config: 计算配置（可选）
        """
        self.report_data = report_data
        self.config = config or {}
        self.calculated_values: Dict[str, FieldValue] = {}
        self.navigator = DataNavigator()
    
    def get_value(self, field_path: str) -> FieldValue:
        """
        获取字段值 - 支持点号路径
        
        Args:
            field_path: 字段路径，如 "extracted_data.rated_wattage"
            
        Returns:
            FieldValue: 字段值对象
        """
        value = self.navigator.get_value(self.report_data, field_path)
        
        if value is None:
            # 构建可用字段信息
            available = self._get_available_fields()
            raise FieldNotFoundError(field_path, available)
        
        return FieldValue(value, 'report', field_path)
    
    def _get_available_fields(self) -> Dict:
        """获取可用的字段列表"""
        available = {}
        for section in ['metadata', 'extracted_data', 'calculated_data']:
            section_data = self.report_data.get(section, {})
            if isinstance(section_data, dict):
                available[section] = list(section_data.keys())
        return available
    
    def calculate_field(self, mapping: Dict) -> Optional[FieldValue]:
        """
        根据映射配置计算单个字段
        
        Args:
            mapping: 字段映射配置
            
        Returns:
            FieldValue: 计算后的字段值，如果无法计算则返回None
        """
        args_config = mapping.get('args', [])
        function_name = mapping.get('function')
        source_field = mapping.get('source_field')
        
        # 收集参数值
        args = []
        for arg_path in args_config:
            try:
                field_value = self.get_value(arg_path)
                args.append(field_value.value)
            except FieldNotFoundError as e:
                # 参数缺失时，根据配置决定如何处理
                if self.config.get('strict_mode', False):
                    raise
                args.append(None)
        
        # 执行计算
        if function_name:
            # 使用注册的函数
            func = CalculationRegistry.get(function_name)
            if func is None:
                raise FunctionNotFoundError(function_name)
            
            try:
                result = func(*args)
            except Exception as e:
                raise CalculatorError(
                    f"Error executing function '{function_name}' with args {args}: {e}"
                )
        else:
            # 无函数名时，返回第一个参数（简单透传）
            result = args[0] if args else None
        
        # 保存计算结果到report_data
        self.navigator.set_value(self.report_data, source_field, result)
        
        field_value = FieldValue(result, 'calculated_data', source_field)
        self.calculated_values[source_field] = field_value
        
        return field_value
    
    def process_config(self, config: Dict) -> Dict[str, FieldValue]:
        """
        处理整个配置，计算所有需要计算的字段
        
        Args:
            config: 配置字典，包含field_mappings
            
        Returns:
            Dict[str, FieldValue]: 计算结果字典
        """
        results = {}
        field_mappings = config.get('field_mappings', [])
        
        for mapping in field_mappings:
            if mapping.get('function'):  # 有函数才需要计算
                try:
                    field_value = self.calculate_field(mapping)
                    if field_value:
                        template_field = mapping.get('template_field')
                        results[template_field] = field_value
                except CalculatorError as e:
                    if self.config.get('raise_on_error', False):
                        raise
                    # 记录错误但继续处理
                    template_field = mapping.get('template_field', 'unknown')
                    print(f"Warning: Failed to calculate field '{template_field}': {e}")
        
        return results
    
    def get_calculated_report(self) -> Dict:
        """
        获取计算后的完整报告数据
        
        Returns:
            Dict: 包含所有计算字段的完整report数据
        """
        return self.report_data


# ============================================
# 内置计算函数
# ============================================

@CalculationRegistry.register("calculate_energy_class_rating")
def calculate_energy_class_rating(rated_wattage: float, useful_luminous_flux: float) -> str:
    """
    计算能源等级评级
    
    计算公式：η = Φ_use / P
    其中：
        η: 能源效率 (lm/W)
        Φ_use: 有用光通量 (lm)
        P: 额定功率 (W)
    
    Args:
        rated_wattage: 额定功率 (W)
        useful_luminous_flux: 有用光通量 (lm)
        
    Returns:
        str: 能源等级 (A++到E)
    """
    if not rated_wattage or rated_wattage == 0:
        return "N/A"
    
    if not useful_luminous_flux:
        return "N/A"
    
    # 确保是数值
    try:
        wattage = float(rated_wattage)
        flux = float(useful_luminous_flux)
    except (ValueError, TypeError):
        return "N/A"
    
    # 计算能效
    efficacy = flux / wattage
    
    # 根据能效确定能源等级（LED灯具标准）
    # 这些阈值需要根据实际标准调整
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


@CalculationRegistry.register("calculate_energy_efficacy")
def calculate_energy_efficacy(rated_wattage: float, useful_luminous_flux: float) -> str:
    """
    计算能源效率
    
    Args:
        rated_wattage: 额定功率 (W)
        useful_luminous_flux: 有用光通量 (lm)
        
    Returns:
        str: 能源效率值 (lm/W)，保留2位小数
    """
    if not rated_wattage or rated_wattage == 0:
        return "N/A"
    
    if not useful_luminous_flux:
        return "N/A"
    
    try:
        wattage = float(rated_wattage)
        flux = float(useful_luminous_flux)
        efficacy = flux / wattage
        return f"{efficacy:.2f}"
    except (ValueError, TypeError):
        return "N/A"


@CalculationRegistry.register("calculate_percentage")
def calculate_percentage(value: float, total: float) -> str:
    """
    计算百分比
    
    Args:
        value: 部分值
        total: 总值
        
    Returns:
        str: 百分比字符串，保留2位小数
    """
    if not total or total == 0:
        return "0.00%"
    
    try:
        percentage = (float(value) / float(total)) * 100
        return f"{percentage:.2f}%"
    except (ValueError, TypeError):
        return "0.00%"


@CalculationRegistry.register("format_number")
def format_number(value: float, decimal_places: int = 2) -> str:
    """
    格式化数字
    
    Args:
        value: 数值
        decimal_places: 小数位数
        
    Returns:
        str: 格式化后的数字字符串
    """
    try:
        return f"{float(value):.{int(decimal_places)}f}"
    except (ValueError, TypeError):
        return str(value)


@CalculationRegistry.register("concat")
def concat(*args, separator: str = " ") -> str:
    """
    连接字符串
    
    Args:
        *args: 要连接的字符串
        separator: 分隔符
        
    Returns:
        str: 连接后的字符串
    """
    return separator.join(str(arg) for arg in args if arg is not None)


@CalculationRegistry.register("multiply")
def multiply(a: float, b: float) -> float:
    """
    乘法计算
    
    Args:
        a: 第一个数值
        b: 第二个数值
        
    Returns:
        float: 乘积
    """
    try:
        return float(a) * float(b)
    except (ValueError, TypeError):
        return 0.0


@CalculationRegistry.register("divide")
def divide(a: float, b: float, default: float = 0.0) -> float:
    """
    除法计算
    
    Args:
        a: 被除数
        b: 除数
        default: 除数为0时的默认值
        
    Returns:
        float: 商
    """
    try:
        divisor = float(b)
        if divisor == 0:
            return default
        return float(a) / divisor
    except (ValueError, TypeError):
        return default


# ============================================
# 命令行接口
# ============================================

def load_json(path: str) -> Dict:
    """加载JSON文件"""
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def save_json(path: str, data: Dict):
    """保存JSON文件"""
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def main():
    """命令行入口"""
    import argparse
    
    parser = argparse.ArgumentParser(
        description='Calculate field values and save calculated report'
    )
    parser.add_argument(
        '--config', 
        required=True, 
        help='Path to report_config.json'
    )
    parser.add_argument(
        '--report', 
        required=True, 
        help='Path to report.json (input)'
    )
    parser.add_argument(
        '--output', 
        required=True, 
        help='Path to calculated_report.json (output)'
    )
    parser.add_argument(
        '--strict-mode',
        action='store_true',
        help='Raise error when field not found (default: False)'
    )
    parser.add_argument(
        '--functions-module',
        help='Path to custom functions module (e.g., custom_calculations)'
    )
    
    args = parser.parse_args()
    
    try:
        # 加载数据
        config = load_json(args.config)
        report_data = load_json(args.report)
        
        # 加载自定义函数模块
        if args.functions_module:
            CalculationRegistry.load_from_module(args.functions_module)
        
        # 创建计算器
        calculator_config = {
            'strict_mode': args.strict_mode,
            'raise_on_error': args.strict_mode
        }
        calculator = FieldCalculator(
            report_data=report_data,
            config=calculator_config
        )
        
        # 处理配置，执行计算
        results = calculator.process_config(config)
        
        # 获取计算后的完整报告
        calculated_report = calculator.get_calculated_report()
        
        # 保存结果
        save_json(args.output, calculated_report)
        
        print(f"Successfully calculated {len(results)} fields: {args.output}")
        for field, fv in results.items():
            print(f"  - {field}: {fv.value}")
        
        return 0
        
    except FileNotFoundError as e:
        print(f"Error: File not found - {e}", file=sys.stderr)
        return 1
    except json.JSONDecodeError as e:
        print(f"Error: Invalid JSON - {e}", file=sys.stderr)
        return 1
    except CalculatorError as e:
        print(f"Calculation error: {e}", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        return 1


if __name__ == '__main__':
    sys.exit(main())
