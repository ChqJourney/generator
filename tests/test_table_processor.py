"""
单元测试：table_processor 模块
测试表格数据转换器和自定义转换器
"""

import pytest
import sys
from typing import List, Dict, Any
from unittest.mock import patch, MagicMock

sys.path.insert(0, 'src')
from table_processor.data_transformer import TableDataTransformer
from table_processor.custom_transformers import (
    CustomTransformerRegistry,
    FormatRule,
    format_number,
    parse_format_rules,
)


# =============================================================================
# TableDataTransformer 基础测试
# =============================================================================

class TestTableDataTransformerBasic:
    """TableDataTransformer 基础功能测试"""
    
    def setup_method(self):
        """每个测试前创建转换器实例"""
        self.transformer = TableDataTransformer()
    
    def test_transform_empty_data(self):
        """测试空数据转换"""
        result = self.transformer.transform([], [])
        assert result == []
    
    def test_transform_no_transformations(self):
        """测试无转换操作时数据不变"""
        data = [["a", "b"], ["c", "d"]]
        result = self.transformer.transform(data, [])
        assert result == data
    
    def test_transform_preserves_original(self):
        """测试转换不修改原始数据"""
        data = [["a", "b"], ["c", "d"]]
        original = [row[:] for row in data]
        
        transformations = [{"type": "skip_columns", "columns": [0]}]
        result = self.transformer.transform(data, transformations)
        
        assert data == original


# =============================================================================
# Skip Columns 测试
# =============================================================================

class TestSkipColumns:
    """跳过列功能测试"""
    
    def setup_method(self):
        self.transformer = TableDataTransformer()
    
    def test_skip_single_column(self):
        """测试跳过单列"""
        data = [["a", "b", "c"], ["d", "e", "f"]]
        transformations = [{"type": "skip_columns", "columns": [1]}]
        
        result = self.transformer.transform(data, transformations)
        
        assert result == [["a", "c"], ["d", "f"]]
    
    def test_skip_multiple_columns(self):
        """测试跳过多列"""
        data = [["a", "b", "c", "d"], ["e", "f", "g", "h"]]
        transformations = [{"type": "skip_columns", "columns": [0, 2]}]
        
        result = self.transformer.transform(data, transformations)
        
        assert result == [["b", "d"], ["f", "h"]]
    
    def test_skip_empty_columns_list(self):
        """测试空列列表不跳过任何列"""
        data = [["a", "b"], ["c", "d"]]
        transformations = [{"type": "skip_columns", "columns": []}]
        
        result = self.transformer.transform(data, transformations)
        
        assert result == data
    
    def test_skip_all_columns(self):
        """测试跳过所有列"""
        data = [["a", "b"], ["c", "d"]]
        transformations = [{"type": "skip_columns", "columns": [0, 1]}]
        
        result = self.transformer.transform(data, transformations)
        
        assert result == [[], []]


# =============================================================================
# Add Column 测试
# =============================================================================

class TestAddColumn:
    """添加列功能测试"""
    
    def setup_method(self):
        self.transformer = TableDataTransformer()
    
    def test_add_row_index_column(self):
        """测试添加行索引列"""
        data = [["a", "b"], ["c", "d"]]
        transformations = [{"type": "add_column", "position": 0, "source": "row_index"}]
        
        result = self.transformer.transform(data, transformations)
        
        assert result == [["1", "a", "b"], ["2", "c", "d"]]
    
    def test_add_fixed_value_column(self):
        """测试添加固定值列"""
        data = [["a", "b"], ["c", "d"]]
        transformations = [{"type": "add_column", "position": 1, "source": "value:test"}]
        
        result = self.transformer.transform(data, transformations)
        
        assert result == [["a", "test", "b"], ["c", "test", "d"]]
    
    def test_add_metadata_column(self):
        """测试从 metadata 添加列"""
        data = [["a", "b"], ["c", "d"]]
        calculated_report = {"metadata": {"model": "LED-100"}}
        transformations = [{"type": "add_column", "position": 0, "source": "metadata:model"}]
        
        result = self.transformer.transform(data, transformations, calculated_report)
        
        assert result == [["LED-100", "a", "b"], ["LED-100", "c", "d"]]
    
    def test_add_metadata_missing_key(self):
        """测试 metadata 中不存在的 key"""
        data = [["a", "b"]]
        calculated_report = {"metadata": {"model": "LED-100"}}
        transformations = [{"type": "add_column", "position": 0, "source": "metadata:missing"}]
        
        result = self.transformer.transform(data, transformations, calculated_report)
        
        assert result == [["", "a", "b"]]
    
    def test_add_extracted_data_column(self):
        """测试从 extracted_data 添加列"""
        data = [["a", "b"]]
        calculated_report = {"extracted_data": {"wattage": "100W"}}
        transformations = [{"type": "add_column", "position": 0, "source": "extracted_data:wattage"}]
        
        result = self.transformer.transform(data, transformations, calculated_report)
        
        assert result == [["100W", "a", "b"]]
    
    def test_add_calculated_data_column(self):
        """测试从 calculated_data 添加列"""
        data = [["a", "b"]]
        calculated_report = {"calculated_data": {"efficacy": "120.5"}}
        transformations = [{"type": "add_column", "position": 0, "source": "calculated_data:efficacy"}]
        
        result = self.transformer.transform(data, transformations, calculated_report)
        
        assert result == [["120.5", "a", "b"]]
    
    def test_add_column_beyond_length(self):
        """测试在超出当前行长度的位置添加列"""
        data = [["a"], ["b"]]
        transformations = [{"type": "add_column", "position": 5, "source": "value:x"}]
        
        result = self.transformer.transform(data, transformations)
        
        assert result == [["a", "x"], ["b", "x"]]


# =============================================================================
# Calculate 测试 (公式计算)
# =============================================================================

class TestCalculateFormula:
    """公式计算功能测试"""
    
    def setup_method(self):
        self.transformer = TableDataTransformer()
    
    def test_calculate_simple_formula(self):
        """测试简单公式计算"""
        data = [["10", "20"], ["30", "40"]]
        transformations = [{
            "type": "calculate",
            "column": 2,
            "operation": "formula=A{row}+B{row}"
        }]
        
        result = self.transformer.transform(data, transformations)
        
        # 公式计算结果不固定小数位
        assert result[0][2] == "30"
        assert result[1][2] == "70"
    
    def test_calculate_with_decimal(self):
        """测试带小数位的公式计算"""
        data = [["10", "3"]]
        transformations = [{
            "type": "calculate",
            "column": 2,
            "operation": "formula=A{row}/B{row}",
            "decimal": 2
        }]
        
        result = self.transformer.transform(data, transformations)
        
        assert result[0][2] == "3.33"
    
    def test_calculate_insert_new_column(self):
        """测试插入新列"""
        data = [["10", "20"]]
        transformations = [{
            "type": "calculate",
            "column": 1,
            "operation": "formula=A{row}*2",
            "insert": True
        }]
        
        result = self.transformer.transform(data, transformations)
        
        assert len(result[0]) == 3
        assert result[0][1] == "20"
    
    def test_calculate_with_non_numeric(self):
        """测试包含非数字值的公式计算"""
        data = [["10", "abc"]]
        transformations = [{
            "type": "calculate",
            "column": 2,
            "operation": "formula=A{row}+B{row}"
        }]
        
        result = self.transformer.transform(data, transformations)
        
        # B列非数字，应使用0
        assert result[0][2] == "10"


# =============================================================================
# Aggregation 测试 (聚合操作)
# =============================================================================

class TestAggregation:
    """聚合操作测试"""
    
    def setup_method(self):
        self.transformer = TableDataTransformer()
    
    def test_average_calculation(self):
        """测试平均值计算"""
        data = [["10"], ["20"], ["30"]]
        transformations = [{
            "type": "calculate",
            "column": 0,
            "operation": "average",
            "decimal": 1
        }]
        
        result = self.transformer.transform(data, transformations)
        
        assert result[-1][0] == "20.0"
    
    def test_sum_calculation(self):
        """测试求和计算"""
        data = [["10"], ["20"], ["30"]]
        transformations = [{
            "type": "calculate",
            "column": 0,
            "operation": "sum",
            "decimal": 0
        }]
        
        result = self.transformer.transform(data, transformations)
        
        assert result[-1][0] == "60"
    
    def test_max_calculation(self):
        """测试最大值计算"""
        data = [["10"], ["50"], ["30"]]
        transformations = [{
            "type": "calculate",
            "column": 0,
            "operation": "max"
        }]
        
        result = self.transformer.transform(data, transformations)
        
        # 无小数位时返回整数形式
        assert result[-1][0] == "50.0"
    
    def test_min_calculation(self):
        """测试最小值计算"""
        data = [["10"], ["50"], ["30"]]
        transformations = [{
            "type": "calculate",
            "column": 0,
            "operation": "min"
        }]
        
        result = self.transformer.transform(data, transformations)
        
        assert result[-1][0] == "10.0"
    
    def test_multiple_aggregations(self):
        """测试多个聚合操作组合"""
        data = [["10", "20"], ["30", "40"]]
        transformations = [
            {"type": "calculate", "column": 0, "operation": "sum"},
            {"type": "calculate", "column": 1, "operation": "average"}
        ]
        
        result = self.transformer.transform(data, transformations)
        
        # 无小数位时返回 .0 格式
        assert result[-1][0] == "40.0"
        assert result[-1][1] == "30.0"
    
    def test_aggregation_with_empty_data(self):
        """测试空数据聚合"""
        data = []
        transformations = [{
            "type": "calculate",
            "column": 0,
            "operation": "average"
        }]
        
        result = self.transformer.transform(data, transformations)
        
        assert result == []
    
    def test_aggregation_with_non_numeric_values(self):
        """测试包含非数字值的聚合"""
        data = [["10"], ["abc"], ["20"]]
        transformations = [{
            "type": "calculate",
            "column": 0,
            "operation": "sum"
        }]
        
        result = self.transformer.transform(data, transformations)
        
        assert result[-1][0] == "30.0"


# =============================================================================
# Format Column 测试
# =============================================================================

class TestFormatColumn:
    """列格式化测试"""
    
    def setup_method(self):
        self.transformer = TableDataTransformer()
    
    def test_format_fixed_decimal(self):
        """测试固定小数位格式化"""
        data = [["3.14159"], ["2.71828"]]
        transformations = [{
            "type": "format_column",
            "column": 0,
            "decimal": 2
        }]
        
        result = self.transformer.transform(data, transformations)
        
        assert result[0][0] == "3.14"
        assert result[1][0] == "2.72"
    
    def test_format_with_function(self):
        """测试使用 lambda 函数格式化"""
        data = [["0.5"], ["1.5"]]
        transformations = [{
            "type": "format_column",
            "column": 0,
            "function": "lambda x: f'{x:.2f}%'"
        }]
        
        result = self.transformer.transform(data, transformations)
        
        assert result[0][0] == "0.50%"
        assert result[1][0] == "1.50%"
    
    def test_format_non_numeric_unchanged(self):
        """测试非数字值保持不变"""
        data = [["abc"], ["123"]]
        transformations = [{
            "type": "format_column",
            "column": 0,
            "decimal": 2
        }]
        
        result = self.transformer.transform(data, transformations)
        
        assert result[0][0] == "abc"
        assert result[1][0] == "123.00"


# =============================================================================
# Reorder 测试
# =============================================================================

class TestReorder:
    """列重排序测试"""
    
    def setup_method(self):
        self.transformer = TableDataTransformer()
    
    def test_reorder_columns(self):
        """测试列重排序"""
        data = [["a", "b", "c", "d"]]
        transformations = [{
            "type": "reorder",
            "order": [2, 0, 1]
        }]
        
        result = self.transformer.transform(data, transformations)
        
        assert result[0] == ["c", "a", "b"]
    
    def test_reorder_with_out_of_bounds(self):
        """测试包含超出范围的索引"""
        data = [["a", "b"]]
        transformations = [{
            "type": "reorder",
            "order": [1, 0, 5, 3]  # 5和3超出范围
        }]
        
        result = self.transformer.transform(data, transformations)
        
        assert result[0] == ["b", "a"]


# =============================================================================
# Filter Rows 测试
# =============================================================================

class TestFilterRows:
    """行过滤测试"""
    
    def setup_method(self):
        self.transformer = TableDataTransformer()
    
    def test_filter_remove_empty(self):
        """测试移除空行"""
        data = [["a", "b"], ["", ""], ["c", "d"]]
        transformations = [{
            "type": "filter_rows",
            "condition": "remove_empty"
        }]
        
        result = self.transformer.transform(data, transformations)
        
        assert len(result) == 2
        assert ["a", "b"] in result
        assert ["c", "d"] in result
    
    def test_filter_remove_all_empty(self):
        """测试移除全空行（包括None）"""
        data = [["a", "b"], [None, ""], ["c", "d"]]
        transformations = [{
            "type": "filter_rows",
            "condition": "remove_all_empty"
        }]
        
        result = self.transformer.transform(data, transformations)
        
        assert len(result) == 2
    
    def test_filter_unknown_condition(self):
        """测试未知条件不过滤"""
        data = [["a", "b"], ["", ""]]
        transformations = [{
            "type": "filter_rows",
            "condition": "unknown_condition"
        }]
        
        result = self.transformer.transform(data, transformations)
        
        assert len(result) == 2


# =============================================================================
# Combined Transformations 测试
# =============================================================================

class TestCombinedTransformations:
    """组合转换测试"""
    
    def setup_method(self):
        self.transformer = TableDataTransformer()
    
    def test_skip_then_add_column(self):
        """测试先跳过列再添加列"""
        data = [["a", "b", "c"], ["d", "e", "f"]]
        transformations = [
            {"type": "skip_columns", "columns": [1]},
            {"type": "add_column", "position": 0, "source": "row_index"}
        ]
        
        result = self.transformer.transform(data, transformations)
        
        assert result == [["1", "a", "c"], ["2", "d", "f"]]
    
    def test_complex_transformation_chain(self):
        """测试复杂转换链"""
        data = [["10", "20", ""], ["30", "40", ""]]
        transformations = [
            {"type": "calculate", "column": 2, "operation": "formula=A{row}+B{row}", "insert": True},
            {"type": "format_column", "column": 2, "decimal": 1},
            {"type": "reorder", "order": [2, 0, 1]}
        ]
        
        result = self.transformer.transform(data, transformations)
        
        # 验证计算和格式化
        assert result[0][0] == "30.0"
        assert result[1][0] == "70.0"


# =============================================================================
# Helper Methods 测试
# =============================================================================

class TestHelperMethods:
    """辅助方法测试"""
    
    def setup_method(self):
        self.transformer = TableDataTransformer()
    
    def test_is_numeric_with_valid_numbers(self):
        """测试有效数字判断"""
        assert self.transformer._is_numeric("123") is True
        assert self.transformer._is_numeric("3.14") is True
        assert self.transformer._is_numeric("-10") is True
    
    def test_is_numeric_with_invalid(self):
        """测试无效数字判断"""
        assert self.transformer._is_numeric("abc") is False
        assert self.transformer._is_numeric("") is False
        assert self.transformer._is_numeric(None) is False
    
    def test_format_number_with_decimal(self):
        """测试数字格式化"""
        assert self.transformer._format_number(3.14159, 2) == "3.14"
        assert self.transformer._format_number(42, 0) == "42"
    
    def test_format_number_without_decimal(self):
        """测试无小数位格式化"""
        assert self.transformer._format_number(3.14159, None) == "3.14159"
    
    def test_letters_to_index(self):
        """测试列字母转索引"""
        assert self.transformer._letters_to_index("A") == 0
        assert self.transformer._letters_to_index("B") == 1
        assert self.transformer._letters_to_index("Z") == 25
        assert self.transformer._letters_to_index("AA") == 26
        assert self.transformer._letters_to_index("AB") == 27
    
    def test_parse_column_references(self):
        """测试列引用解析"""
        row = ["10", "20", "30"]
        formula = "A{row}+B{row}"
        
        result = self.transformer._parse_column_references(formula, 0, row)
        
        assert result == "10+20"


# =============================================================================
# CustomTransformerRegistry 测试
# =============================================================================

class TestCustomTransformerRegistry:
    """自定义转换器注册表测试"""
    
    def setup_method(self):
        """清空注册表"""
        CustomTransformerRegistry._transformers.clear()
    
    def test_register_decorator(self):
        """测试装饰器注册"""
        @CustomTransformerRegistry.register("test_transformer")
        def test_func(data, params, extracted_data):
            return [["transformed"]]
        
        assert "test_transformer" in CustomTransformerRegistry._transformers
    
    def test_transform_existing(self):
        """测试执行已注册的转换器"""
        @CustomTransformerRegistry.register("test_transformer")
        def test_func(data, params, extracted_data):
            return [[p.upper() for p in row] for row in data]
        
        data = [["a", "b"]]
        result = CustomTransformerRegistry.transform("test_transformer", data, {}, None)
        
        assert result == [["A", "B"]]
    
    def test_transform_unknown_raises(self):
        """测试未知转换器抛出异常"""
        with pytest.raises(ValueError, match="Unknown transformer"):
            CustomTransformerRegistry.transform("unknown", [], {}, None)


# =============================================================================
# FormatRule 测试
# =============================================================================

class TestFormatRule:
    """格式化规则测试"""
    
    def test_format_with_matching_condition(self):
        """测试条件匹配时的格式化"""
        rule = FormatRule(
            condition=lambda x: x >= 100,
            format_str="{:.0f}",
            default_format="{:.2f}"
        )
        
        assert rule.format(150) == "150"
    
    def test_format_with_non_matching_condition(self):
        """测试条件不匹配时的格式化"""
        rule = FormatRule(
            condition=lambda x: x >= 100,
            format_str="{:.0f}",
            default_format="{:.2f}"
        )
        
        assert rule.format(50) == "50.00"


# =============================================================================
# format_number 函数测试
# =============================================================================

class TestFormatNumber:
    """format_number 函数测试"""
    
    def test_format_with_decimal(self):
        """测试固定小数位"""
        assert format_number(3.14159, decimal=2) == "3.14"
    
    def test_format_with_rules(self):
        """测试格式化规则"""
        rules = [
            FormatRule(lambda x: x >= 100, "{:.0f}", "{:.2f}")
        ]
        
        assert format_number(150, format_rules=rules) == "150"
        assert format_number(50, format_rules=rules) == "50.00"
    
    def test_format_invalid_value(self):
        """测试无效值"""
        assert format_number("abc") == "abc"
        assert format_number(None) == ""
    
    def test_format_without_options(self):
        """测试无选项时"""
        # 默认转为浮点数字符串
        assert format_number(42) == "42.0"


# =============================================================================
# parse_format_rules 函数测试
# =============================================================================

class TestParseFormatRules:
    """parse_format_rules 函数测试"""
    
    def test_parse_greater_equal(self):
        """测试解析 >= 条件"""
        config = [{"condition": "x >= 100", "format": "{:.0f}"}]
        rules = parse_format_rules(config)
        
        assert len(rules) == 1
        assert rules[0].condition(150) is True
        assert rules[0].condition(50) is False
    
    def test_parse_less_than(self):
        """测试解析 < 条件"""
        config = [{"condition": "x < 50", "format": "{:.2f}"}]
        rules = parse_format_rules(config)
        
        assert rules[0].condition(30) is True
        assert rules[0].condition(100) is False
    
    def test_parse_invalid_condition(self):
        """测试解析无效条件"""
        config = [{"condition": "invalid", "format": "{:.2f}"}]
        rules = parse_format_rules(config)
        
        assert len(rules) == 0


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
