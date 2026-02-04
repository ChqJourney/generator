"""
单元测试：calculator.py
为每个函数/方法提供独立的单元测试
"""

import pytest
import json
import sys
import importlib
from typing import Dict, Any
from unittest.mock import patch, MagicMock, mock_open
from dataclasses import dataclass

# 导入被测试的模块
sys.path.insert(0, 'src')
from calculator import (
    FieldValue,
    CalculatorError,
    FieldNotFoundError,
    FunctionNotFoundError,
    DataNavigator,
    CalculationRegistry,
    FieldCalculator,
    calculate_energy_class_rating,
    calculate_energy_efficacy,
    calculate_percentage,
    format_number,
    concat,
    multiply,
    divide,
    load_json,
    save_json,
    main,
)


# =============================================================================
# FieldValue 类测试
# =============================================================================

class TestFieldValue:
    """FieldValue 数据类的单元测试"""
    
    def test_field_value_initialization(self):
        """测试 FieldValue 基本初始化"""
        fv = FieldValue(value="test", source="metadata", field_name="test_field")
        assert fv.value == "test"
        assert fv.source == "metadata"
        assert fv.field_name == "test_field"
    
    def test_field_value_string_conversion_to_int(self):
        """测试字符串自动转换为整数"""
        fv = FieldValue(value="42", source="test", field_name="num")
        assert fv.value == 42
        assert isinstance(fv.value, int)
    
    def test_field_value_string_conversion_to_float(self):
        """测试字符串自动转换为浮点数"""
        fv = FieldValue(value="3.14", source="test", field_name="pi")
        assert fv.value == 3.14
        assert isinstance(fv.value, float)
    
    def test_field_value_non_numeric_string_unchanged(self):
        """测试非数字字符串保持不变"""
        fv = FieldValue(value="hello", source="test", field_name="text")
        assert fv.value == "hello"
        assert isinstance(fv.value, str)
    
    def test_field_value_numeric_with_whitespace(self):
        """测试带空格的数值字符串"""
        fv = FieldValue(value="  100  ", source="test", field_name="num")
        assert fv.value == 100
    
    def test_field_value_empty_string(self):
        """测试空字符串处理"""
        fv = FieldValue(value="", source="test", field_name="empty")
        assert fv.value == ""
    
    def test_field_value_none_value(self):
        """测试 None 值处理"""
        fv = FieldValue(value=None, source="test", field_name="null")
        assert fv.value is None
    
    def test_field_value_float_string_scientific_notation(self):
        """测试科学计数法字符串转换"""
        fv = FieldValue(value="1.5e3", source="test", field_name="sci")
        assert fv.value == 1500.0
    
    def test_field_value_negative_number(self):
        """测试负数字符串转换"""
        fv = FieldValue(value="-50", source="test", field_name="neg")
        assert fv.value == -50
        assert isinstance(fv.value, int)


# =============================================================================
# 异常类测试
# =============================================================================

class TestCalculatorError:
    """CalculatorError 异常类测试"""
    
    def test_calculator_error_is_exception(self):
        """测试 CalculatorError 是 Exception 的子类"""
        assert issubclass(CalculatorError, Exception)
    
    def test_calculator_error_can_be_raised(self):
        """测试 CalculatorError 可以被抛出和捕获"""
        with pytest.raises(CalculatorError):
            raise CalculatorError("test error")


class TestFieldNotFoundError:
    """FieldNotFoundError 异常类测试"""
    
    def test_field_not_found_error_is_calculator_error(self):
        """测试 FieldNotFoundError 继承自 CalculatorError"""
        assert issubclass(FieldNotFoundError, CalculatorError)
    
    def test_field_not_found_error_message(self):
        """测试错误消息包含字段路径"""
        error = FieldNotFoundError("metadata.report_no")
        assert "metadata.report_no" in str(error)
        assert "Field not found" in str(error)
    
    def test_field_not_found_error_with_available_fields(self):
        """测试包含可用字段列表的错误消息"""
        available = {"metadata": ["report_no", "date"], "extracted_data": ["name"]}
        error = FieldNotFoundError("metadata.missing_field", available)
        assert "metadata.missing_field" in str(error)
        # 错误消息显示的是顶级字段名
        assert "metadata" in str(error)
        assert "extracted_data" in str(error)
    
    def test_field_not_found_error_stores_field_path(self):
        """测试错误对象存储字段路径"""
        error = FieldNotFoundError("path.to.field")
        assert error.field_path == "path.to.field"
    
    def test_field_not_found_error_stores_available_fields(self):
        """测试错误对象存储可用字段"""
        available = {"section": ["field1"]}
        error = FieldNotFoundError("path", available)
        assert error.available_fields == available


class TestFunctionNotFoundError:
    """FunctionNotFoundError 异常类测试"""
    
    def test_function_not_found_error_is_calculator_error(self):
        """测试 FunctionNotFoundError 继承自 CalculatorError"""
        assert issubclass(FunctionNotFoundError, CalculatorError)
    
    def test_function_not_found_error_message(self):
        """测试错误消息包含函数名"""
        error = FunctionNotFoundError("my_function")
        assert "my_function" in str(error)
        assert "Function not found" in str(error)
    
    def test_function_not_found_error_stores_function_name(self):
        """测试错误对象存储函数名"""
        error = FunctionNotFoundError("test_func")
        assert error.function_name == "test_func"


# =============================================================================
# DataNavigator 类测试
# =============================================================================

class TestDataNavigator:
    """DataNavigator 类的单元测试"""
    
    def test_get_value_simple_path(self):
        """测试简单路径取值"""
        data = {"name": "John", "age": 30}
        result = DataNavigator.get_value(data, "name")
        assert result == "John"
    
    def test_get_value_nested_path(self):
        """测试嵌套路径取值"""
        data = {"user": {"name": "John", "age": 30}}
        result = DataNavigator.get_value(data, "user.name")
        assert result == "John"
    
    def test_get_value_deeply_nested(self):
        """测试深层嵌套路径取值"""
        data = {"a": {"b": {"c": {"d": "value"}}}}
        result = DataNavigator.get_value(data, "a.b.c.d")
        assert result == "value"
    
    def test_get_value_nonexistent_path(self):
        """测试不存在路径返回 None"""
        data = {"name": "John"}
        result = DataNavigator.get_value(data, "nonexistent")
        assert result is None
    
    def test_get_value_nonexistent_nested_path(self):
        """测试不存在嵌套路径返回 None"""
        data = {"user": {"name": "John"}}
        result = DataNavigator.get_value(data, "user.nonexistent")
        assert result is None
    
    def test_get_value_partial_path(self):
        """测试部分路径不存在"""
        data = {"user": {"name": "John"}}
        result = DataNavigator.get_value(data, "company.department")
        assert result is None
    
    def test_get_value_empty_path(self):
        """测试空路径返回 None"""
        data = {"name": "John"}
        result = DataNavigator.get_value(data, "")
        assert result is None
    
    def test_get_value_none_path(self):
        """测试 None 路径返回 None"""
        data = {"name": "John"}
        result = DataNavigator.get_value(data, None)
        assert result is None
    
    def test_get_value_with_list_in_path(self):
        """测试路径中包含列表"""
        data = {"users": [{"name": "John"}, {"name": "Jane"}]}
        # 应该返回 None，因为不支持列表索引
        result = DataNavigator.get_value(data, "users.0.name")
        assert result is None
    
    def test_set_value_simple_path(self):
        """测试简单路径设置值"""
        data = {}
        DataNavigator.set_value(data, "name", "John")
        assert data["name"] == "John"
    
    def test_set_value_nested_path(self):
        """测试嵌套路径设置值"""
        data = {}
        DataNavigator.set_value(data, "user.name", "John")
        assert data["user"]["name"] == "John"
    
    def test_set_value_deeply_nested_creates_intermediate(self):
        """测试深层路径自动创建中间节点"""
        data = {}
        DataNavigator.set_value(data, "a.b.c.d", "value")
        assert data["a"]["b"]["c"]["d"] == "value"
    
    def test_set_value_overwrites_existing(self):
        """测试设置值覆盖已存在的值"""
        data = {"name": "Old"}
        DataNavigator.set_value(data, "name", "New")
        assert data["name"] == "New"
    
    def test_set_value_updates_nested_existing(self):
        """测试更新嵌套已存在的值"""
        data = {"user": {"name": "Old", "age": 30}}
        DataNavigator.set_value(data, "user.name", "New")
        assert data["user"]["name"] == "New"
        assert data["user"]["age"] == 30  # 其他值保持不变
    
    def test_set_value_single_part_path(self):
        """测试单部分路径"""
        data = {}
        DataNavigator.set_value(data, "key", "value")
        assert data == {"key": "value"}
    
    def test_set_and_get_roundtrip(self):
        """测试设置后获取的往返操作"""
        data = {}
        DataNavigator.set_value(data, "a.b.c", "test")
        result = DataNavigator.get_value(data, "a.b.c")
        assert result == "test"


# =============================================================================
# CalculationRegistry 类测试
# =============================================================================

class TestCalculationRegistry:
    """CalculationRegistry 类的单元测试"""
    
    def setup_method(self):
        """每个测试前清空注册表"""
        CalculationRegistry._functions.clear()
    
    def test_register_function(self):
        """测试注册函数"""
        def test_func(x):
            return x * 2
        
        CalculationRegistry.register("double", test_func)
        assert "double" in CalculationRegistry.list_functions()
        assert CalculationRegistry.get("double") == test_func
    
    def test_register_as_decorator(self):
        """测试装饰器方式注册"""
        @CalculationRegistry.register("triple")
        def triple(x):
            return x * 3
        
        assert "triple" in CalculationRegistry.list_functions()
        func = CalculationRegistry.get("triple")
        assert func(5) == 15
    
    def test_get_nonexistent_function(self):
        """测试获取不存在的函数返回 None"""
        result = CalculationRegistry.get("nonexistent")
        assert result is None
    
    def test_list_functions_empty(self):
        """测试空注册表返回空列表"""
        functions = CalculationRegistry.list_functions()
        assert functions == []
    
    def test_list_functions_returns_all(self):
        """测试列出所有函数"""
        CalculationRegistry.register("func1", lambda x: x)
        CalculationRegistry.register("func2", lambda x: x)
        
        functions = CalculationRegistry.list_functions()
        assert "func1" in functions
        assert "func2" in functions
        assert len(functions) == 2
    
    def test_register_overwrites_existing(self):
        """测试注册同名函数覆盖旧函数"""
        CalculationRegistry.register("same", lambda x: x)
        CalculationRegistry.register("same", lambda x: x * 2)
        
        func = CalculationRegistry.get("same")
        assert func(5) == 10
    
    def test_register_with_none_function(self):
        """测试注册 None 函数"""
        CalculationRegistry.register("none_func", None)
        assert CalculationRegistry.get("none_func") is None


# =============================================================================
# FieldCalculator 类测试
# =============================================================================

class TestFieldCalculator:
    """FieldCalculator 类的单元测试"""
    
    def setup_method(self):
        """每个测试前清空注册表，避免测试间状态污染"""
        CalculationRegistry._functions.clear()
    
    def test_init_with_empty_data(self):
        """测试空数据初始化"""
        calc = FieldCalculator({})
        assert calc.report_data == {}
        assert calc.calculated_values == {}
    
    def test_init_with_data(self):
        """测试带数据初始化"""
        data = {"metadata": {"name": "test"}}
        calc = FieldCalculator(data)
        assert calc.report_data == data
    
    def test_get_value_simple_path(self):
        """测试获取简单路径值"""
        data = {"metadata": {"report_no": "RPT-001"}}
        calc = FieldCalculator(data)
        result = calc.get_value("metadata.report_no")
        assert result.value == "RPT-001"
        assert result.source == "report"
    
    def test_get_value_nested_path(self):
        """测试获取嵌套路径值"""
        data = {"extracted_data": {"product": {"name": "LED"}}}
        calc = FieldCalculator(data)
        result = calc.get_value("extracted_data.product.name")
        assert result.value == "LED"
    
    def test_get_value_nonexistent_raises_error(self):
        """测试获取不存在的路径抛出错误"""
        data = {"metadata": {"name": "test"}}
        calc = FieldCalculator(data)
        
        with pytest.raises(FieldNotFoundError):
            calc.get_value("nonexistent.field")
    
    def test_get_available_fields(self):
        """测试获取可用字段列表"""
        data = {
            "metadata": {"field1": "val1"},
            "extracted_data": {"field2": "val2"},
            "calculated_data": {"field3": "val3"}
        }
        calc = FieldCalculator(data)
        available = calc._get_available_fields()
        
        assert "field1" in available["metadata"]
        assert "field2" in available["extracted_data"]
        assert "field3" in available["calculated_data"]
    
    def test_calculate_field_without_function(self):
        """测试无函数时透传第一个参数"""
        data = {"extracted_data": {"val": 10}}
        calc = FieldCalculator(data)
        
        mapping = {
            "source_field": "calculated_data.result",
            "args": ["extracted_data.val"]
        }
        
        result = calc.calculate_field(mapping)
        assert result.value == 10
    
    def test_calculate_field_with_function(self):
        """测试使用函数计算"""
        # 先注册测试函数
        CalculationRegistry.register("add_one", lambda x: x + 1)
        
        data = {"extracted_data": {"val": 5}}
        calc = FieldCalculator(data)
        
        mapping = {
            "source_field": "calculated_data.result",
            "args": ["extracted_data.val"],
            "function": "add_one"
        }
        
        result = calc.calculate_field(mapping)
        assert result.value == 6
    
    def test_calculate_field_stores_result_in_report_data(self):
        """测试结果存储在 report_data 中"""
        CalculationRegistry.register("identity", lambda x: x)
        
        data = {"extracted_data": {"val": 42}}
        calc = FieldCalculator(data)
        
        mapping = {
            "source_field": "calculated_data.result",
            "args": ["extracted_data.val"],
            "function": "identity"
        }
        
        calc.calculate_field(mapping)
        assert calc.report_data["calculated_data"]["result"] == 42
    
    def test_calculate_field_with_multiple_args(self):
        """测试多参数计算"""
        CalculationRegistry.register("add", lambda x, y: x + y)
        
        data = {"extracted_data": {"a": 10, "b": 20}}
        calc = FieldCalculator(data)
        
        mapping = {
            "source_field": "calculated_data.sum",
            "args": ["extracted_data.a", "extracted_data.b"],
            "function": "add"
        }
        
        result = calc.calculate_field(mapping)
        assert result.value == 30
    
    def test_calculate_field_function_not_found_raises(self):
        """测试函数不存在时抛出错误"""
        data = {"extracted_data": {"val": 5}}
        calc = FieldCalculator(data)
        
        mapping = {
            "source_field": "calculated_data.result",
            "args": ["extracted_data.val"],
            "function": "nonexistent_func"
        }
        
        with pytest.raises(FunctionNotFoundError):
            calc.calculate_field(mapping)
    
    def test_calculate_field_missing_arg_in_non_strict_mode(self):
        """测试非严格模式下缺失参数不会抛出异常"""
        # 使用一个能处理 None 的函数
        CalculationRegistry.register("process", lambda x: "handled_none" if x is None else x)
        
        data = {"extracted_data": {}}
        calc = FieldCalculator(data, config={"strict_mode": False})
        
        mapping = {
            "source_field": "calculated_data.result",
            "args": ["extracted_data.missing"],
            "function": "process"
        }
        
        result = calc.calculate_field(mapping)
        # 非严格模式下，缺失参数用 None 代替，函数正常执行
        assert result is not None
        assert result.value == "handled_none"
    
    def test_process_config_executes_all_calculations(self):
        """测试处理配置执行所有计算"""
        CalculationRegistry.register("double", lambda x: x * 2)
        CalculationRegistry.register("triple", lambda x: x * 3)
        
        data = {"extracted_data": {"val": 5}}
        calc = FieldCalculator(data)
        
        config = {
            "field_mappings": [
                {
                    "template_field": "doubled",
                    "source_field": "calculated_data.doubled",
                    "args": ["extracted_data.val"],
                    "function": "double"
                },
                {
                    "template_field": "tripled",
                    "source_field": "calculated_data.tripled",
                    "args": ["extracted_data.val"],
                    "function": "triple"
                }
            ]
        }
        
        results = calc.process_config(config)
        
        assert len(results) == 2
        assert results["doubled"].value == 10
        assert results["tripled"].value == 15
    
    def test_process_config_skips_mappings_without_function(self):
        """测试跳过没有函数的映射"""
        data = {"extracted_data": {"val": 5}}
        calc = FieldCalculator(data)
        
        config = {
            "field_mappings": [
                {
                    "template_field": "text_field",
                    "source_field": "extracted_data.val",
                    "type": "text"
                }
            ]
        }
        
        results = calc.process_config(config)
        assert len(results) == 0
    
    def test_process_config_handles_errors_with_raise_on_error(self):
        """测试 raise_on_error 配置"""
        data = {"extracted_data": {}}
        calc = FieldCalculator(data, config={"raise_on_error": True})
        
        config = {
            "field_mappings": [
                {
                    "template_field": "bad",
                    "source_field": "calculated_data.bad",
                    "args": ["extracted_data.missing"],
                    "function": "nonexistent"
                }
            ]
        }
        
        with pytest.raises(FunctionNotFoundError):
            calc.process_config(config)
    
    def test_get_calculated_report_returns_full_data(self):
        """测试获取完整报告数据"""
        CalculationRegistry.register("identity", lambda x: x)
        
        data = {"metadata": {"name": "test"}, "extracted_data": {"val": 10}}
        calc = FieldCalculator(data)
        
        mapping = {
            "source_field": "calculated_data.result",
            "args": ["extracted_data.val"],
            "function": "identity"
        }
        calc.calculate_field(mapping)
        
        result = calc.get_calculated_report()
        assert result["metadata"]["name"] == "test"
        assert result["calculated_data"]["result"] == 10


# =============================================================================
# 内置计算函数测试
# =============================================================================

class TestCalculateEnergyClassRating:
    """calculate_energy_class_rating 函数测试"""
    
    def test_high_efficacy_returns_a_plus_plus(self):
        """测试高能效返回 A++"""
        result = calculate_energy_class_rating(10, 2200)  # 220 lm/W
        assert result == "A++"
    
    def test_very_high_efficacy_returns_a_plus_plus(self):
        """测试极高能效返回 A++"""
        result = calculate_energy_class_rating(10, 2500)  # 250 lm/W
        assert result == "A++"
    
    def test_medium_high_efficacy_returns_a_plus(self):
        """测试中高能效返回 A+"""
        result = calculate_energy_class_rating(10, 1900)  # 190 lm/W
        assert result == "A+"
    
    def test_good_efficacy_returns_a(self):
        """测试良好能效返回 A"""
        result = calculate_energy_class_rating(10, 1650)  # 165 lm/W
        assert result == "A"
    
    def test_fair_efficacy_returns_b(self):
        """测试一般能效返回 B"""
        result = calculate_energy_class_rating(10, 1400)  # 140 lm/W
        assert result == "B"
    
    def test_below_average_efficacy_returns_c(self):
        """测试低于平均能效返回 C"""
        result = calculate_energy_class_rating(10, 1150)  # 115 lm/W
        assert result == "C"
    
    def test_low_efficacy_returns_d(self):
        """测试低能效返回 D"""
        result = calculate_energy_class_rating(10, 900)  # 90 lm/W
        assert result == "D"
    
    def test_very_low_efficacy_returns_e(self):
        """测试极低能效返回 E"""
        result = calculate_energy_class_rating(10, 700)  # 70 lm/W
        assert result == "E"
    
    def test_zero_wattage_returns_na(self):
        """测试零功率返回 N/A"""
        result = calculate_energy_class_rating(0, 1000)
        assert result == "N/A"
    
    def test_none_wattage_returns_na(self):
        """测试 None 功率返回 N/A"""
        result = calculate_energy_class_rating(None, 1000)
        assert result == "N/A"
    
    def test_zero_flux_returns_na(self):
        """测试零光通量返回 N/A"""
        result = calculate_energy_class_rating(10, 0)
        assert result == "N/A"
    
    def test_none_flux_returns_na(self):
        """测试 None 光通量返回 N/A"""
        result = calculate_energy_class_rating(10, None)
        assert result == "N/A"
    
    def test_string_numbers_work(self):
        """测试字符串数字正常工作"""
        result = calculate_energy_class_rating("10", "1650")
        assert result == "A"
    
    def test_invalid_string_returns_na(self):
        """测试无效字符串返回 N/A"""
        result = calculate_energy_class_rating("abc", "1000")
        assert result == "N/A"


class TestCalculateEnergyEfficacy:
    """calculate_energy_efficacy 函数测试"""
    
    def test_basic_calculation(self):
        """测试基本计算"""
        result = calculate_energy_efficacy(10, 1000)
        assert result == "100.00"
    
    def test_decimal_result(self):
        """测试小数结果"""
        result = calculate_energy_efficacy(3, 100)
        assert result == "33.33"
    
    def test_zero_wattage_returns_na(self):
        """测试零功率返回 N/A"""
        result = calculate_energy_efficacy(0, 1000)
        assert result == "N/A"
    
    def test_none_wattage_returns_na(self):
        """测试 None 功率返回 N/A"""
        result = calculate_energy_efficacy(None, 1000)
        assert result == "N/A"
    
    def test_zero_flux_returns_na(self):
        """测试零光通量返回 N/A"""
        result = calculate_energy_efficacy(10, 0)
        assert result == "N/A"
    
    def test_string_numbers_work(self):
        """测试字符串数字"""
        result = calculate_energy_efficacy("10", "1000")
        assert result == "100.00"
    
    def test_invalid_string_returns_na(self):
        """测试无效字符串"""
        result = calculate_energy_efficacy("abc", "1000")
        assert result == "N/A"


class TestCalculatePercentage:
    """calculate_percentage 函数测试"""
    
    def test_basic_percentage(self):
        """测试基本百分比"""
        result = calculate_percentage(25, 100)
        assert result == "25.00%"
    
    def test_decimal_percentage(self):
        """测试小数百分比"""
        result = calculate_percentage(1, 3)
        assert result == "33.33%"
    
    def test_zero_total_returns_zero(self):
        """测试零总数返回 0%"""
        result = calculate_percentage(50, 0)
        assert result == "0.00%"
    
    def test_zero_value_returns_zero(self):
        """测试零值返回 0%"""
        result = calculate_percentage(0, 100)
        assert result == "0.00%"
    
    def test_value_greater_than_total(self):
        """测试值大于总数"""
        result = calculate_percentage(150, 100)
        assert result == "150.00%"
    
    def test_string_numbers_work(self):
        """测试字符串数字"""
        result = calculate_percentage("25", "100")
        assert result == "25.00%"


class TestFormatNumber:
    """format_number 函数测试"""
    
    def test_default_two_decimals(self):
        """测试默认2位小数"""
        result = format_number(3.14159)
        assert result == "3.14"
    
    def test_custom_decimals(self):
        """测试自定义小数位"""
        result = format_number(3.14159, 4)
        assert result == "3.1416"
    
    def test_zero_decimals(self):
        """测试0位小数"""
        result = format_number(3.7, 0)
        assert result == "4"
    
    def test_integer_input(self):
        """测试整数输入"""
        result = format_number(42)
        assert result == "42.00"
    
    def test_string_number(self):
        """测试字符串数字"""
        result = format_number("3.14")
        assert result == "3.14"
    
    def test_invalid_string_returns_original(self):
        """测试无效字符串返回原值"""
        result = format_number("abc")
        assert result == "abc"


class TestConcat:
    """concat 函数测试"""
    
    def test_basic_concat(self):
        """测试基本连接"""
        result = concat("Hello", "World")
        assert result == "Hello World"
    
    def test_multiple_args(self):
        """测试多参数连接"""
        result = concat("A", "B", "C", "D")
        assert result == "A B C D"
    
    def test_custom_separator(self):
        """测试自定义分隔符"""
        result = concat("a", "b", "c", separator="-")
        assert result == "a-b-c"
    
    def test_empty_separator(self):
        """测试空分隔符"""
        result = concat("Hello", "World", separator="")
        assert result == "HelloWorld"
    
    def test_single_arg(self):
        """测试单参数"""
        result = concat("only")
        assert result == "only"
    
    def test_none_arg_skipped(self):
        """测试 None 参数被跳过"""
        result = concat("A", None, "B")
        assert result == "A B"
    
    def test_all_none_returns_empty(self):
        """测试全 None 返回空字符串"""
        result = concat(None, None)
        assert result == ""
    
    def test_number_args(self):
        """测试数字参数"""
        result = concat(1, 2, 3)
        assert result == "1 2 3"


class TestMultiply:
    """multiply 函数测试"""
    
    def test_basic_multiplication(self):
        """测试基本乘法"""
        result = multiply(5, 3)
        assert result == 15.0
    
    def test_with_floats(self):
        """测试浮点数乘法"""
        result = multiply(2.5, 4)
        assert result == 10.0
    
    def test_with_zero(self):
        """测试与零相乘"""
        result = multiply(100, 0)
        assert result == 0.0
    
    def test_negative_numbers(self):
        """测试负数乘法"""
        result = multiply(-5, 3)
        assert result == -15.0
    
    def test_string_numbers(self):
        """测试字符串数字"""
        result = multiply("5", "3")
        assert result == 15.0
    
    def test_invalid_string_returns_zero(self):
        """测试无效字符串返回 0"""
        result = multiply("abc", 5)
        assert result == 0.0


class TestDivide:
    """divide 函数测试"""
    
    def test_basic_division(self):
        """测试基本除法"""
        result = divide(10, 2)
        assert result == 5.0
    
    def test_float_result(self):
        """测试浮点结果"""
        result = divide(10, 3)
        assert abs(result - 3.333) < 0.001
    
    def test_divide_by_zero_returns_default(self):
        """测试除以零返回默认值"""
        result = divide(10, 0)
        assert result == 0.0
    
    def test_divide_by_zero_custom_default(self):
        """测试除以零返回自定义默认值"""
        result = divide(10, 0, default=-1)
        assert result == -1.0
    
    def test_negative_numbers(self):
        """测试负数除法"""
        result = divide(-10, 2)
        assert result == -5.0
    
    def test_string_numbers(self):
        """测试字符串数字"""
        result = divide("10", "2")
        assert result == 5.0
    
    def test_invalid_string_returns_default(self):
        """测试无效字符串返回默认值"""
        result = divide("abc", 2)
        assert result == 0.0


# =============================================================================
# 工具函数测试
# =============================================================================

class TestLoadJson:
    """load_json 函数测试"""
    
    def test_load_valid_json(self, tmp_path):
        """测试加载有效 JSON"""
        test_file = tmp_path / "test.json"
        test_file.write_text('{"name": "test", "value": 123}')
        
        result = load_json(str(test_file))
        assert result == {"name": "test", "value": 123}
    
    def test_load_json_with_unicode(self, tmp_path):
        """测试加载包含 Unicode 的 JSON"""
        test_file = tmp_path / "test.json"
        test_file.write_text('{"name": "测试", "value": "中文"}', encoding='utf-8')
        
        result = load_json(str(test_file))
        assert result["name"] == "测试"
        assert result["value"] == "中文"
    
    def test_load_nonexistent_file_raises(self):
        """测试加载不存在的文件抛出 FileNotFoundError"""
        with pytest.raises(FileNotFoundError):
            load_json("/nonexistent/path/file.json")
    
    def test_load_invalid_json_raises(self, tmp_path):
        """测试加载无效 JSON 抛出 JSONDecodeError"""
        test_file = tmp_path / "test.json"
        test_file.write_text('{"invalid json}')
        
        with pytest.raises(json.JSONDecodeError):
            load_json(str(test_file))


class TestSaveJson:
    """save_json 函数测试"""
    
    def test_save_dict_to_json(self, tmp_path):
        """测试保存字典到 JSON"""
        test_file = tmp_path / "test.json"
        data = {"name": "test", "value": 123}
        
        save_json(str(test_file), data)
        
        content = test_file.read_text(encoding='utf-8')
        assert '"name": "test"' in content
        assert '"value": 123' in content
    
    def test_save_nested_dict(self, tmp_path):
        """测试保存嵌套字典"""
        test_file = tmp_path / "test.json"
        data = {"user": {"name": "John", "age": 30}}
        
        save_json(str(test_file), data)
        
        result = json.loads(test_file.read_text(encoding='utf-8'))
        assert result["user"]["name"] == "John"
    
    def test_save_unicode_content(self, tmp_path):
        """测试保存 Unicode 内容"""
        test_file = tmp_path / "test.json"
        data = {"name": "中文测试"}
        
        save_json(str(test_file), data)
        
        content = test_file.read_text(encoding='utf-8')
        assert "中文测试" in content
    
    def test_save_list_to_json(self, tmp_path):
        """测试保存列表到 JSON"""
        test_file = tmp_path / "test.json"
        data = [1, 2, 3, "test"]
        
        save_json(str(test_file), data)
        
        result = json.loads(test_file.read_text(encoding='utf-8'))
        assert result == [1, 2, 3, "test"]


# =============================================================================
# main 函数测试
# =============================================================================

class TestMain:
    """main 函数单元测试"""
    
    @patch('calculator.load_json')
    @patch('calculator.FieldCalculator')
    @patch('calculator.save_json')
    def test_main_success(self, mock_save, mock_calc_class, mock_load):
        """测试 main 函数成功执行"""
        # 准备 mock 数据
        mock_load.side_effect = [
            {"field_mappings": []},     # config
            {"metadata": {}, "extracted_data": {}, "calculated_data": {}}  # report
        ]
        
        mock_calc = MagicMock()
        mock_calc.process_config.return_value = {}
        mock_calc.get_calculated_report.return_value = {"calculated_data": {}}
        mock_calc_class.return_value = mock_calc
        
        # 模拟命令行参数
        test_args = ['calculator.py', '--config', 'config.json', '--report', 'report.json', '--output', 'output.json']
        with patch.object(sys, 'argv', test_args):
            result = main()
        
        assert result == 0
        mock_load.assert_called()
        mock_save.assert_called_once()
    
    @patch('calculator.load_json')
    def test_main_file_not_found(self, mock_load):
        """测试 main 函数处理文件不存在错误"""
        mock_load.side_effect = FileNotFoundError("File not found")
        
        test_args = ['calculator.py', '--config', 'missing.json', '--report', 'report.json', '--output', 'out.json']
        with patch.object(sys, 'argv', test_args):
            result = main()
        
        assert result == 1
    
    @patch('calculator.load_json')
    def test_main_invalid_json(self, mock_load):
        """测试 main 函数处理无效 JSON 错误"""
        mock_load.side_effect = json.JSONDecodeError("Invalid JSON", "", 0)
        
        test_args = ['calculator.py', '--config', 'bad.json', '--report', 'report.json', '--output', 'out.json']
        with patch.object(sys, 'argv', test_args):
            result = main()
        
        assert result == 1
    
    @patch('calculator.load_json')
    @patch('calculator.FieldCalculator')
    def test_main_calculation_error(self, mock_calc_class, mock_load):
        """测试 main 函数处理计算错误"""
        mock_load.return_value = {"field_mappings": []}
        mock_calc_class.side_effect = Exception("Calculation failed")
        
        test_args = ['calculator.py', '--config', 'config.json', '--report', 'report.json', '--output', 'out.json']
        with patch.object(sys, 'argv', test_args):
            result = main()
        
        assert result == 1
    
    @patch('calculator.load_json')
    @patch('calculator.CalculationRegistry')
    @patch('calculator.FieldCalculator')
    @patch('calculator.save_json')
    def test_main_with_custom_functions_module(self, mock_save, mock_calc_class, mock_registry, mock_load):
        """测试 main 函数加载自定义函数模块"""
        mock_load.side_effect = [
            {"field_mappings": []},
            {"metadata": {}, "extracted_data": {}, "calculated_data": {}}
        ]
        
        mock_calc = MagicMock()
        mock_calc.process_config.return_value = {}
        mock_calc.get_calculated_report.return_value = {}
        mock_calc_class.return_value = mock_calc
        
        test_args = [
            'calculator.py', 
            '--config', 'config.json', 
            '--report', 'report.json', 
            '--output', 'out.json',
            '--functions-module', 'custom_calcs'
        ]
        with patch.object(sys, 'argv', test_args):
            result = main()
        
        assert result == 0
        mock_registry.load_from_module.assert_called_once_with('custom_calcs')
    
    @patch('calculator.load_json')
    @patch('calculator.FieldCalculator')
    @patch('calculator.save_json')
    def test_main_strict_mode(self, mock_save, mock_calc_class, mock_load):
        """测试 main 函数严格模式"""
        mock_load.side_effect = [
            {"field_mappings": []},
            {"metadata": {}, "extracted_data": {}}
        ]
        
        mock_calc = MagicMock()
        mock_calc.process_config.return_value = {}
        mock_calc.get_calculated_report.return_value = {}
        mock_calc_class.return_value = mock_calc
        
        test_args = [
            'calculator.py', 
            '--config', 'config.json', 
            '--report', 'report.json', 
            '--output', 'out.json',
            '--strict-mode'
        ]
        with patch.object(sys, 'argv', test_args):
            result = main()
        
        assert result == 0
        # 验证严格模式配置被传递
        call_args = mock_calc_class.call_args
        assert call_args[1]['config']['strict_mode'] is True
