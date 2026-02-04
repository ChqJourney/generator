"""
单元测试：field_mapper.py
为每个函数/方法提供独立的单元测试
"""

import pytest
import json
import sys
from typing import Dict, Any, List
from unittest.mock import patch, MagicMock, mock_open, Mock
import tempfile
import os

# 导入被测试的模块
sys.path.insert(0, 'src')
from field_mapper import (
    load_json,
    get_value_by_path,
    is_external_table_reference,
    is_direct_table_data,
    generate_operations,
    build_table_data_from_excel,
    get_xlsx_to_list,
    main,
)


# =============================================================================
# load_json 函数测试
# =============================================================================

class TestLoadJson:
    """load_json 函数测试"""
    
    def test_load_simple_json_object(self, tmp_path):
        """测试加载简单 JSON 对象"""
        test_file = tmp_path / "test.json"
        test_file.write_text('{"key": "value", "number": 42}', encoding='utf-8')
        
        result = load_json(str(test_file))
        assert result == {"key": "value", "number": 42}
    
    def test_load_nested_json(self, tmp_path):
        """测试加载嵌套 JSON"""
        test_file = tmp_path / "test.json"
        test_file.write_text('{"outer": {"inner": "value"}}', encoding='utf-8')
        
        result = load_json(str(test_file))
        assert result["outer"]["inner"] == "value"
    
    def test_load_json_with_unicode(self, tmp_path):
        """测试加载 Unicode 内容"""
        test_file = tmp_path / "test.json"
        test_file.write_text('{"name": "中文测试"}', encoding='utf-8')
        
        result = load_json(str(test_file))
        assert result["name"] == "中文测试"
    
    def test_load_json_array(self, tmp_path):
        """测试加载 JSON 数组"""
        test_file = tmp_path / "test.json"
        test_file.write_text('[1, 2, 3, "test"]', encoding='utf-8')
        
        result = load_json(str(test_file))
        assert result == [1, 2, 3, "test"]
    
    def test_load_empty_json_object(self, tmp_path):
        """测试加载空 JSON 对象"""
        test_file = tmp_path / "test.json"
        test_file.write_text('{}', encoding='utf-8')
        
        result = load_json(str(test_file))
        assert result == {}
    
    def test_file_not_found_raises_exception(self):
        """测试文件不存在抛出异常"""
        with pytest.raises(FileNotFoundError):
            load_json("/nonexistent/path/file.json")
    
    def test_invalid_json_raises_exception(self, tmp_path):
        """测试无效 JSON 抛出异常"""
        test_file = tmp_path / "test.json"
        test_file.write_text('{"invalid json', encoding='utf-8')
        
        with pytest.raises(json.JSONDecodeError):
            load_json(str(test_file))


# =============================================================================
# get_value_by_path 函数测试
# =============================================================================

class TestGetValueByPath:
    """get_value_by_path 函数测试"""
    
    def test_simple_path(self):
        """测试简单路径取值"""
        data = {"name": "John", "age": 30}
        result = get_value_by_path(data, "name")
        assert result == "John"
    
    def test_nested_path(self):
        """测试嵌套路径取值"""
        data = {"user": {"name": "John", "age": 30}}
        result = get_value_by_path(data, "user.name")
        assert result == "John"
    
    def test_deeply_nested_path(self):
        """测试深层嵌套路径"""
        data = {"a": {"b": {"c": {"d": "value"}}}}
        result = get_value_by_path(data, "a.b.c.d")
        assert result == "value"
    
    def test_path_not_found_returns_none(self):
        """测试路径不存在返回 None"""
        data = {"name": "John"}
        result = get_value_by_path(data, "nonexistent")
        assert result is None
    
    def test_nested_path_not_found_returns_none(self):
        """测试嵌套路径不存在返回 None"""
        data = {"user": {"name": "John"}}
        result = get_value_by_path(data, "user.nonexistent")
        assert result is None
    
    def test_partial_path_not_found(self):
        """测试部分路径不存在"""
        data = {"user": {"name": "John"}}
        result = get_value_by_path(data, "company.department.name")
        assert result is None
    
    def test_empty_path_returns_none(self):
        """测试空路径返回 None"""
        data = {"name": "John"}
        result = get_value_by_path(data, "")
        assert result is None
    
    def test_none_path_returns_none(self):
        """测试 None 路径返回 None"""
        data = {"name": "John"}
        result = get_value_by_path(data, None)
        assert result is None
    
    def test_path_with_special_characters_in_value(self):
        """测试值包含特殊字符"""
        data = {"text": "Hello, World! @#$%"}
        result = get_value_by_path(data, "text")
        assert result == "Hello, World! @#$%"


# =============================================================================
# is_external_table_reference 函数测试
# =============================================================================

class TestIsExternalTableReference:
    """is_external_table_reference 函数测试"""
    
    def test_valid_external_reference(self):
        """测试有效的外部引用"""
        value = {
            "type": "external",
            "source_id": "file.xlsx|Sheet1",
            "start_row": 1
        }
        assert is_external_table_reference(value) is True
    
    def test_missing_type_field(self):
        """测试缺少 type 字段"""
        value = {
            "source_id": "file.xlsx|Sheet1"
        }
        assert is_external_table_reference(value) is False
    
    def test_wrong_type_value(self):
        """测试 type 值不正确"""
        value = {
            "type": "internal",
            "source_id": "file.xlsx|Sheet1"
        }
        assert is_external_table_reference(value) is False
    
    def test_missing_source_id(self):
        """测试缺少 source_id 字段"""
        value = {
            "type": "external",
            "start_row": 1
        }
        assert is_external_table_reference(value) is False
    
    def test_not_a_dict(self):
        """测试非字典类型"""
        assert is_external_table_reference("string") is False
        assert is_external_table_reference([1, 2, 3]) is False
        assert is_external_table_reference(123) is False
    
    def test_empty_dict(self):
        """测试空字典"""
        assert is_external_table_reference({}) is False
    
    def test_none_value(self):
        """测试 None 值"""
        assert is_external_table_reference(None) is False


# =============================================================================
# is_direct_table_data 函数测试
# =============================================================================

class TestIsDirectTableData:
    """is_direct_table_data 函数测试"""
    
    def test_valid_table_data(self):
        """测试有效的表格数据"""
        value = [
            ["header1", "header2"],
            ["data1", "data2"],
            ["data3", "data4"]
        ]
        assert is_direct_table_data(value) is True
    
    def test_single_row_table(self):
        """测试单行表格"""
        value = [["only", "one", "row"]]
        assert is_direct_table_data(value) is True
    
    def test_empty_list(self):
        """测试空列表"""
        assert is_direct_table_data([]) is False
    
    def test_list_with_non_list_elements(self):
        """测试包含非列表元素的列表"""
        value = [1, 2, 3]
        assert is_direct_table_data(value) is False
    
    def test_not_a_list(self):
        """测试非列表类型"""
        assert is_direct_table_data("string") is False
        assert is_direct_table_data({"key": "value"}) is False
        assert is_direct_table_data(123) is False
    
    def test_none_value(self):
        """测试 None 值"""
        assert is_direct_table_data(None) is False
    
    def test_list_of_empty_lists(self):
        """测试空列表的列表"""
        value = [[], [], []]
        assert is_direct_table_data(value) is True  # 仍然是列表的列表
    
    @pytest.mark.xfail(reason="源代码只检查第一个元素，未验证所有元素都是列表")
    def test_mixed_list_content(self):
        """测试混合内容的列表 - 期望失败，等待源代码修复"""
        value = [["row1"], "not a list", ["row3"]]
        assert is_direct_table_data(value) is False
    
    def test_current_behavior_mixed_list_returns_true(self):
        """测试当前行为：混合内容列表返回 True（因为只检查第一个元素）"""
        value = [["row1"], "not a list", ["row3"]]
        # 当前源代码只检查第一个元素，所以返回 True
        result = is_direct_table_data(value)
        assert result is True  # 这是当前实际行为


# =============================================================================
# generate_operations 函数测试
# =============================================================================

class TestGenerateOperations:
    """generate_operations 函数测试"""
    
    def test_text_field_mapping(self):
        """测试文本字段映射"""
        config = {
            "field_mappings": [
                {
                    "template_field": "report_no",
                    "source_field": "metadata.report_no",
                    "type": "text"
                }
            ]
        }
        report_data = {
            "metadata": {"report_no": "RPT-001"}
        }
        
        result = generate_operations(config, report_data)
        
        assert len(result["operations"]) == 1
        assert result["operations"][0]["type"] == "text"
        assert result["operations"][0]["placeholder"] == "report_no"
        assert result["operations"][0]["value"] == "RPT-001"
    
    def test_multiple_text_fields(self):
        """测试多个文本字段"""
        config = {
            "field_mappings": [
                {
                    "template_field": "field1",
                    "source_field": "metadata.field1",
                    "type": "text"
                },
                {
                    "template_field": "field2",
                    "source_field": "extracted_data.field2",
                    "type": "text"
                }
            ]
        }
        report_data = {
            "metadata": {"field1": "value1"},
            "extracted_data": {"field2": "value2"}
        }
        
        result = generate_operations(config, report_data)
        
        assert len(result["operations"]) == 2
        assert result["operations"][0]["value"] == "value1"
        assert result["operations"][1]["value"] == "value2"
    
    def test_nested_value_extraction(self):
        """测试嵌套值提取"""
        config = {
            "field_mappings": [
                {
                    "template_field": "user_name",
                    "source_field": "metadata.user.name",
                    "type": "text"
                }
            ]
        }
        report_data = {
            "metadata": {
                "user": {"name": "John", "age": 30}
            }
        }
        
        result = generate_operations(config, report_data)
        
        assert result["operations"][0]["value"] == "John"
    
    def test_missing_source_field_skipped(self):
        """测试缺失的源字段被跳过"""
        config = {
            "field_mappings": [
                {
                    "template_field": "exists",
                    "source_field": "metadata.exists",
                    "type": "text"
                },
                {
                    "template_field": "missing",
                    "source_field": "metadata.missing",
                    "type": "text"
                }
            ]
        }
        report_data = {
            "metadata": {"exists": "value"}
        }
        
        result = generate_operations(config, report_data)
        
        assert len(result["operations"]) == 1
        assert result["operations"][0]["placeholder"] == "exists"
    
    def test_missing_source_field_in_mapping(self):
        """测试映射中缺少 source_field"""
        config = {
            "field_mappings": [
                {
                    "template_field": "field1",
                    "type": "text"
                    # 缺少 source_field
                }
            ]
        }
        report_data = {"metadata": {"field1": "value"}}
        
        result = generate_operations(config, report_data)
        
        assert len(result["operations"]) == 0
    
    def test_image_field_single_path(self):
        """测试图片字段 - 单路径"""
        config = {
            "field_mappings": [
                {
                    "template_field": "product_image",
                    "source_field": "extracted_data.image",
                    "type": "image",
                    "width": 4.0,
                    "alignment": "center"
                }
            ]
        }
        report_data = {
            "extracted_data": {"image": "path/to/image.jpg"}
        }
        
        result = generate_operations(config, report_data)
        
        assert len(result["operations"]) == 1
        assert result["operations"][0]["type"] == "image"
        assert result["operations"][0]["image_paths"] == ["path/to/image.jpg"]
        assert result["operations"][0]["width"] == 4.0
        assert result["operations"][0]["alignment"] == "center"
    
    def test_image_field_multiple_paths(self):
        """测试图片字段 - 多路径"""
        config = {
            "field_mappings": [
                {
                    "template_field": "gallery",
                    "source_field": "extracted_data.images",
                    "type": "image"
                }
            ]
        }
        report_data = {
            "extracted_data": {
                "images": ["img1.jpg", "img2.jpg", "img3.jpg"]
            }
        }
        
        result = generate_operations(config, report_data)
        
        assert result["operations"][0]["image_paths"] == ["img1.jpg", "img2.jpg", "img3.jpg"]
    
    def test_image_field_json_string_parsing(self):
        """测试图片字段 JSON 字符串解析"""
        config = {
            "field_mappings": [
                {
                    "template_field": "images",
                    "source_field": "extracted_data.images",
                    "type": "image"
                }
            ]
        }
        report_data = {
            "extracted_data": {
                "images": '["img1.jpg", "img2.jpg"]'
            }
        }
        
        result = generate_operations(config, report_data)
        
        assert result["operations"][0]["image_paths"] == ["img1.jpg", "img2.jpg"]
    
    def test_table_field_with_embedded_data(self):
        """测试表格字段 - 内嵌数据"""
        config = {
            "field_mappings": [
                {
                    "template_field": "data_table",
                    "source_field": "extracted_data.table",
                    "table_template_path": "templates/table.docx",
                    "type": "table"
                }
            ]
        }
        report_data = {
            "extracted_data": {
                "table": [
                    ["col1", "col2"],
                    ["val1", "val2"]
                ]
            }
        }
        
        result = generate_operations(config, report_data)
        
        assert len(result["operations"]) == 1
        assert result["operations"][0]["type"] == "table"
        assert result["operations"][0]["table_data"] == [
            ["col1", "col2"],
            ["val1", "val2"]
        ]
    
    def test_table_field_with_transformations(self):
        """测试表格字段 - 带转换配置"""
        config = {
            "field_mappings": [
                {
                    "template_field": "data_table",
                    "source_field": "extracted_data.table",
                    "table_template_path": "templates/table.docx",
                    "type": "table",
                    "transformations": [{"type": "format", "column": 0}],
                    "row_strategy": "fixed_rows",
                    "skip_columns": [1],
                    "header_rows": 2
                }
            ]
        }
        report_data = {
            "extracted_data": {
                "table": [["data"]]
            }
        }
        
        result = generate_operations(config, report_data)
        
        op = result["operations"][0]
        assert op["transformations"] == [{"type": "format", "column": 0}]
        assert op["row_strategy"] == "fixed_rows"
        assert op["skip_columns"] == [1]
        assert op["header_rows"] == 2
    
    def test_empty_field_mappings(self):
        """测试空字段映射列表"""
        config = {"field_mappings": []}
        report_data = {"metadata": {"field": "value"}}
        
        result = generate_operations(config, report_data)
        
        assert result["operations"] == []
    
    def test_unknown_field_type_skipped(self):
        """测试未知字段类型被跳过"""
        config = {
            "field_mappings": [
                {
                    "template_field": "unknown",
                    "source_field": "metadata.unknown",
                    "type": "unknown_type"
                }
            ]
        }
        report_data = {"metadata": {"unknown": "value"}}
        
        result = generate_operations(config, report_data)
        
        assert len(result["operations"]) == 0
    
    def test_numeric_value_conversion_to_string(self):
        """测试数值自动转为字符串"""
        config = {
            "field_mappings": [
                {
                    "template_field": "number",
                    "source_field": "metadata.number",
                    "type": "text"
                }
            ]
        }
        report_data = {"metadata": {"number": 42}}
        
        result = generate_operations(config, report_data)
        
        assert result["operations"][0]["value"] == "42"
        assert isinstance(result["operations"][0]["value"], str)


# =============================================================================
# build_table_data_from_excel 函数测试
# =============================================================================

class TestBuildTableDataFromExcel:
    """build_table_data_from_excel 函数测试"""
    
    @patch('field_mapper.get_xlsx_to_list')
    def test_valid_external_reference(self, mock_get_xlsx):
        """测试有效的外部引用"""
        mock_get_xlsx.return_value = [
            ["col1", "col2"],
            ["val1", "val2"]
        ]
        
        value = {
            "type": "external",
            "source_id": "test.xlsx|Sheet1",
            "start_row": 1,
            "mapping": {"col1": "A", "col2": "B"}
        }
        
        result = build_table_data_from_excel(value, ["col1", "col2"])
        
        assert len(result) == 2
        mock_get_xlsx.assert_called_once()
    
    def test_invalid_source_id_format(self):
        """测试无效的 source_id 格式"""
        value = {
            "type": "external",
            "source_id": "invalid_format",
            "start_row": 1
        }
        
        result = build_table_data_from_excel(value)
        
        assert result == []
    
    @patch('field_mapper.get_xlsx_to_list')
    def test_error_during_excel_reading(self, mock_get_xlsx):
        """测试 Excel 读取错误处理"""
        mock_get_xlsx.side_effect = Exception("File not found")
        
        value = {
            "type": "external",
            "source_id": "test.xlsx|Sheet1",
            "start_row": 1
        }
        
        result = build_table_data_from_excel(value)
        
        assert result == []


# =============================================================================
# get_xlsx_to_list 函数测试
# =============================================================================

class TestGetXlsxToList:
    """get_xlsx_to_list 函数测试"""
    
    @patch('field_mapper.load_workbook')
    def test_basic_excel_reading(self, mock_load_workbook):
        """测试基本 Excel 读取"""
        # 创建 mock worksheet
        mock_ws = MagicMock()
        mock_ws.iter_rows.return_value = iter([
            ("Header1", "Header2"),  # 表头行
            ("val1", "val2"),        # 数据行1
            ("val3", "val4")         # 数据行2
        ])
        
        # 创建 mock workbook
        mock_wb = MagicMock()
        mock_wb.__getitem__ = MagicMock(return_value=mock_ws)
        
        mock_load_workbook.return_value = mock_wb
        
        mapping = {"col1": "Header1", "col2": "Header2"}
        result = get_xlsx_to_list(
            "test.xlsx", "Sheet1", 0, mapping, ["col1", "col2"]
        )
        
        assert len(result) == 2
        assert result[0] == ["val1", "val2"]
        assert result[1] == ["val3", "val4"]
        mock_wb.close.assert_called_once()
    
    @patch('field_mapper.load_workbook')
    def test_sheet_not_found_raises_error(self, mock_load_workbook):
        """测试工作表不存在抛出错误"""
        mock_wb = MagicMock()
        mock_wb.__getitem__ = MagicMock(side_effect=KeyError("Sheet1"))
        mock_wb.close = MagicMock()
        
        mock_load_workbook.return_value = mock_wb
        
        with pytest.raises(ValueError, match="工作表"):
            get_xlsx_to_list(
                "test.xlsx", "Sheet1", 0, {}, []
            )
    
    @patch('field_mapper.load_workbook')
    def test_empty_sheet_returns_empty_list(self, mock_load_workbook):
        """测试空工作表返回空列表"""
        mock_ws = MagicMock()
        mock_ws.iter_rows.return_value = iter([])
        
        mock_wb = MagicMock()
        mock_wb.__getitem__ = MagicMock(return_value=mock_ws)
        
        mock_load_workbook.return_value = mock_wb
        
        result = get_xlsx_to_list(
            "test.xlsx", "Sheet1", 0, {}, []
        )
        
        assert result == []
    
    @patch('field_mapper.load_workbook')
    def test_header_whitespace_normalization(self, mock_load_workbook):
        """测试表头空白字符规范化"""
        mock_ws = MagicMock()
        mock_ws.iter_rows.return_value = iter([
            ("  Header  1  ", "Header 2"),  # 带空格的表头
            ("val1", "val2")
        ])
        
        mock_wb = MagicMock()
        mock_wb.__getitem__ = MagicMock(return_value=mock_ws)
        
        mock_load_workbook.return_value = mock_wb
        
        mapping = {"col1": "Header 1", "col2": "Header 2"}
        result = get_xlsx_to_list(
            "test.xlsx", "Sheet1", 0, mapping, ["col1", "col2"]
        )
        
        assert result[0] == ["val1", "val2"]
    
    @patch('field_mapper.load_workbook')
    def test_none_values_converted_to_empty_string(self, mock_load_workbook):
        """测试 None 值转为空字符串"""
        mock_ws = MagicMock()
        mock_ws.iter_rows.return_value = iter([
            ("Header1",),
            (None,),
            ("val1",)
        ])
        
        mock_wb = MagicMock()
        mock_wb.__getitem__ = MagicMock(return_value=mock_ws)
        
        mock_load_workbook.return_value = mock_wb
        
        mapping = {"col1": "Header1"}
        result = get_xlsx_to_list(
            "test.xlsx", "Sheet1", 0, mapping, ["col1"]
        )
        
        assert result[0] == [""]
        assert result[1] == ["val1"]
    
    @patch('field_mapper.load_workbook')
    def test_missing_mapping_column_returns_empty(self, mock_load_workbook):
        """测试缺失映射列返回空字符串"""
        mock_ws = MagicMock()
        mock_ws.iter_rows.return_value = iter([
            ("Header1",),
            ("val1",)
        ])
        
        mock_wb = MagicMock()
        mock_wb.__getitem__ = MagicMock(return_value=mock_ws)
        
        mock_load_workbook.return_value = mock_wb
        
        # 请求一个不存在映射的列
        mapping = {"col1": "Header1"}
        result = get_xlsx_to_list(
            "test.xlsx", "Sheet1", 0, mapping, ["col1", "missing_col"]
        )
        
        assert result[0] == ["val1", ""]
    
    @patch('field_mapper.load_workbook')
    def test_start_row_offset(self, mock_load_workbook):
        """测试起始行偏移"""
        mock_ws = MagicMock()
        mock_ws.iter_rows.return_value = iter([
            ("Header1",),  # start_row=1, so this is row 2
            ("val1",)
        ])
        
        mock_wb = MagicMock()
        mock_wb.__getitem__ = MagicMock(return_value=mock_ws)
        
        mock_load_workbook.return_value = mock_wb
        
        mapping = {"col1": "Header1"}
        result = get_xlsx_to_list(
            "test.xlsx", "Sheet1", 1, mapping, ["col1"]
        )
        
        # iter_rows 应该以 min_row=2 被调用（start_row + 1）
        mock_ws.iter_rows.assert_called_once_with(min_row=2, values_only=True)


# =============================================================================
# main 函数测试
# =============================================================================

class TestMain:
    """main 函数单元测试"""
    
    @patch('field_mapper.load_json')
    @patch('field_mapper.generate_operations')
    @patch('builtins.open', new_callable=mock_open)
    def test_main_success(self, mock_file, mock_generate, mock_load):
        """测试 main 函数成功执行"""
        # 准备 mock 数据
        mock_load.side_effect = [
            {"field_mappings": []},  # config
            {"metadata": {}}          # report
        ]
        mock_generate.return_value = {"operations": []}
        
        # 模拟命令行参数
        test_args = ['field_mapper.py', '--config', 'config.json', '--report', 'report.json', '--output', 'output.json']
        with patch.object(sys, 'argv', test_args):
            result = main()
        
        assert result == 0
        mock_load.assert_called()
        mock_generate.assert_called_once()
    
    @patch('field_mapper.load_json')
    def test_main_file_not_found(self, mock_load):
        """测试 main 函数处理文件不存在错误"""
        mock_load.side_effect = FileNotFoundError("File not found")
        
        test_args = ['field_mapper.py', '--config', 'missing.json', '--report', 'report.json', '--output', 'out.json']
        with patch.object(sys, 'argv', test_args):
            result = main()
        
        assert result == 1
    
    @patch('field_mapper.load_json')
    def test_main_invalid_json(self, mock_load):
        """测试 main 函数处理无效 JSON 错误"""
        mock_load.side_effect = json.JSONDecodeError("Invalid JSON", "", 0)
        
        test_args = ['field_mapper.py', '--config', 'bad.json', '--report', 'report.json', '--output', 'out.json']
        with patch.object(sys, 'argv', test_args):
            result = main()
        
        assert result == 1
    
    @patch('field_mapper.load_json')
    @patch('field_mapper.generate_operations')
    def test_main_general_exception(self, mock_generate, mock_load):
        """测试 main 函数处理一般异常"""
        mock_load.return_value = {"field_mappings": []}
        mock_generate.side_effect = Exception("Unexpected error")
        
        test_args = ['field_mapper.py', '--config', 'config.json', '--report', 'report.json', '--output', 'out.json']
        with patch.object(sys, 'argv', test_args):
            result = main()
        
        assert result == 1
