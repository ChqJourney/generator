"""
单元测试：template_validator.py
为模板验证器提供全面的单元测试
"""

import pytest
import sys
from unittest.mock import MagicMock, patch, mock_open
from pathlib import Path

# 导入被测试的模块
sys.path.insert(0, 'src')
from template_validator import (
    TemplateValidator,
    main,
)


# =============================================================================
# Fixtures
# =============================================================================

@pytest.fixture
def mock_document():
    """创建模拟的 Document 对象"""
    doc = MagicMock()
    
    # 模拟段落 - 包含单个 run 的占位符
    para_single = MagicMock()
    para_single.text = "{{test_field}}"
    run_single = MagicMock()
    run_single.text = "{{test_field}}"
    para_single.runs = [run_single]
    
    # 模拟段落 - 包含分割的占位符
    para_split = MagicMock()
    para_split.text = "{{split_field}}"
    run_split1 = MagicMock()
    run_split1.text = "{{"
    run_split2 = MagicMock()
    run_split2.text = "split_field"
    run_split3 = MagicMock()
    run_split3.text = "}}"
    para_split.runs = [run_split1, run_split2, run_split3]
    
    # 模拟段落 - 普通文本
    para_normal = MagicMock()
    para_normal.text = "普通文本"
    run_normal = MagicMock()
    run_normal.text = "普通文本"
    para_normal.runs = [run_normal]
    
    # 模拟段落 - 多个占位符
    para_multiple = MagicMock()
    para_multiple.text = "{{field1}} and {{field2}}"
    run_multi = MagicMock()
    run_multi.text = "{{field1}} and {{field2}}"
    para_multiple.runs = [run_multi]
    
    doc.paragraphs = [para_single, para_split, para_normal, para_multiple]
    doc.tables = []
    
    return doc


@pytest.fixture
def mock_document_with_tables():
    """创建包含表格的模拟 Document 对象"""
    doc = MagicMock()
    
    # 空段落
    doc.paragraphs = []
    
    # 创建模拟单元格
    cell1 = MagicMock()
    para_cell1 = MagicMock()
    para_cell1.text = "{{table_field}}"
    para_cell1.runs = [MagicMock()]
    para_cell1.runs[0].text = "{{table_field}}"
    cell1.paragraphs = [para_cell1]
    
    cell2 = MagicMock()
    para_cell2 = MagicMock()
    para_cell2.text = "普通单元格文本"
    para_cell2.runs = [MagicMock()]
    para_cell2.runs[0].text = "普通单元格文本"
    cell2.paragraphs = [para_cell2]
    
    # 创建模拟行
    row = MagicMock()
    row.cells = [cell1, cell2]
    
    # 创建模拟表格
    table = MagicMock()
    table.rows = [row]
    
    doc.tables = [table]
    
    return doc


@pytest.fixture
def mock_config():
    """创建模拟的配置数据"""
    return {
        "field_mappings": [
            {"template_field": "test_field", "source_field": "data.field1"},
            {"template_field": "field1", "source_field": "data.field2"},
            {"template_field": "field2", "source_field": "data.field3"},
            {"template_field": "table_field", "source_field": "data.table"},
        ]
    }


# =============================================================================
# TemplateValidator 类测试
# =============================================================================

class TestTemplateValidatorInit:
    """TemplateValidator 初始化测试"""
    
    @patch('template_validator.Document')
    def test_init_without_config(self, mock_doc_class):
        """测试无配置文件初始化"""
        mock_doc = MagicMock()
        mock_doc_class.return_value = mock_doc
        
        validator = TemplateValidator("template.docx")
        
        assert validator.template_path == "template.docx"
        assert validator.config_path is None
        assert validator.config is None
        mock_doc_class.assert_called_once_with("template.docx")
    
    @patch('template_validator.Document')
    @patch('template_validator.TemplateValidator._load_config')
    def test_init_with_config(self, mock_load_config, mock_doc_class):
        """测试有配置文件初始化"""
        mock_doc = MagicMock()
        mock_doc_class.return_value = mock_doc
        mock_load_config.return_value = {"field_mappings": []}
        
        validator = TemplateValidator("template.docx", "config.json")
        
        assert validator.template_path == "template.docx"
        assert validator.config_path == "config.json"
        assert validator.config == {"field_mappings": []}


class TestLoadConfig:
    """_load_config 方法测试"""
    
    @patch('template_validator.Document')
    def test_load_valid_config(self, mock_doc_class):
        """测试加载有效配置"""
        mock_doc = MagicMock()
        mock_doc_class.return_value = mock_doc
        
        validator = TemplateValidator("template.docx", "config.json")
        
        # mock utils.jsonc_utils.load_json
        with patch('utils.jsonc_utils.load_json') as mock_load_json:
            mock_load_json.return_value = {"field_mappings": []}
            config = validator._load_config()
            
            assert config == {"field_mappings": []}
            mock_load_json.assert_called_once_with("config.json")
    
    @patch('template_validator.Document')
    @patch('template_validator.logger')
    def test_load_config_error(self, mock_logger, mock_doc_class):
        """测试加载配置出错"""
        mock_doc = MagicMock()
        mock_doc_class.return_value = mock_doc
        
        validator = TemplateValidator("template.docx", "config.json")
        
        with patch('utils.jsonc_utils.load_json') as mock_load_json:
            mock_load_json.side_effect = Exception("File not found")
            config = validator._load_config()
            
            assert config is None
            # 验证 warning 被调用（在初始化时已经调用一次，这里再调用一次）
            assert mock_logger.warning.call_count >= 1
    
    @patch('template_validator.Document')
    def test_load_config_import_path(self, mock_doc_class):
        """测试配置加载的导入路径"""
        mock_doc = MagicMock()
        mock_doc_class.return_value = mock_doc
        
        validator = TemplateValidator("template.docx", "config.json")
        
        # 验证从 utils.jsonc_utils 导入 load_json
        with patch('utils.jsonc_utils.load_json') as mock_load:
            mock_load.return_value = {"field_mappings": []}
            config = validator._load_config()
            mock_load.assert_called_once_with("config.json")


class TestGetConfiguredPlaceholders:
    """_get_configured_placeholders 方法测试"""
    
    @patch('template_validator.Document')
    def test_no_config_returns_empty(self, mock_doc_class):
        """测试无配置时返回空集合"""
        mock_doc = MagicMock()
        mock_doc_class.return_value = mock_doc
        
        validator = TemplateValidator("template.docx")
        result = validator._get_configured_placeholders()
        
        assert result == set()
    
    @patch('template_validator.Document')
    def test_extract_placeholders_from_config(self, mock_doc_class):
        """测试从配置中提取占位符"""
        mock_doc = MagicMock()
        mock_doc_class.return_value = mock_doc
        
        config = {
            "field_mappings": [
                {"template_field": "field1"},
                {"template_field": "field2"},
                {"template_field": "field3", "type": "text"},
            ]
        }
        
        validator = TemplateValidator("template.docx", "config.json")
        validator.config = config
        result = validator._get_configured_placeholders()
        
        assert result == {"field1", "field2", "field3"}


class TestExtractPlaceholderFromText:
    """_extract_placeholder_from_text 方法测试"""
    
    @patch('template_validator.Document')
    def test_single_placeholder(self, mock_doc_class):
        """测试单个占位符"""
        mock_doc = MagicMock()
        mock_doc_class.return_value = mock_doc
        
        validator = TemplateValidator("template.docx")
        result = validator._extract_placeholder_from_text("{{test_field}}")
        
        assert result == ["test_field"]
    
    @patch('template_validator.Document')
    def test_multiple_placeholders(self, mock_doc_class):
        """测试多个占位符"""
        mock_doc = MagicMock()
        mock_doc_class.return_value = mock_doc
        
        validator = TemplateValidator("template.docx")
        result = validator._extract_placeholder_from_text("{{field1}} and {{field2}}")
        
        assert result == ["field1", "field2"]
    
    @patch('template_validator.Document')
    def test_no_placeholder(self, mock_doc_class):
        """测试无占位符"""
        mock_doc = MagicMock()
        mock_doc_class.return_value = mock_doc
        
        validator = TemplateValidator("template.docx")
        result = validator._extract_placeholder_from_text("普通文本")
        
        assert result == []
    
    @patch('template_validator.Document')
    def test_placeholder_with_whitespace(self, mock_doc_class):
        """测试带空格的占位符"""
        mock_doc = MagicMock()
        mock_doc_class.return_value = mock_doc
        
        validator = TemplateValidator("template.docx")
        result = validator._extract_placeholder_from_text("{{  field_with_space  }}")
        
        assert result == ["field_with_space"]
    
    @patch('template_validator.Document')
    def test_incomplete_placeholder(self, mock_doc_class):
        """测试不完整的占位符"""
        mock_doc = MagicMock()
        mock_doc_class.return_value = mock_doc
        
        validator = TemplateValidator("template.docx")
        result = validator._extract_placeholder_from_text("{{incomplete")
        
        assert result == []


class TestCheckRunsInParagraph:
    """_check_runs_in_paragraph 方法测试"""
    
    @patch('template_validator.Document')
    def test_single_run_placeholder(self, mock_doc_class):
        """测试单个 run 的占位符"""
        mock_doc = MagicMock()
        mock_doc_class.return_value = mock_doc
        
        validator = TemplateValidator("template.docx")
        
        para = MagicMock()
        para.text = "{{test_field}}"
        run = MagicMock()
        run.text = "{{test_field}}"
        para.runs = [run]
        
        validator._check_runs_in_paragraph(para, "paragraph 0")
        
        assert "test_field" in validator.all_placeholders
        assert len(validator.single_run_placeholders) == 1
        assert len(validator.split_placeholders) == 0
    
    @patch('template_validator.Document')
    def test_split_placeholder(self, mock_doc_class):
        """测试分割的占位符"""
        mock_doc = MagicMock()
        mock_doc_class.return_value = mock_doc
        
        validator = TemplateValidator("template.docx")
        
        para = MagicMock()
        para.text = "{{split_field}}"
        run1 = MagicMock()
        run1.text = "{{"
        run2 = MagicMock()
        run2.text = "split_field"
        run3 = MagicMock()
        run3.text = "}}"
        para.runs = [run1, run2, run3]
        
        validator._check_runs_in_paragraph(para, "paragraph 1")
        
        assert "split_field" in validator.all_placeholders
        assert len(validator.single_run_placeholders) == 0
        assert len(validator.split_placeholders) == 1
        assert validator.split_placeholders[0]['placeholder'] == "split_field"
    
    @patch('template_validator.Document')
    def test_no_placeholder(self, mock_doc_class):
        """测试无占位符"""
        mock_doc = MagicMock()
        mock_doc_class.return_value = mock_doc
        
        validator = TemplateValidator("template.docx")
        
        para = MagicMock()
        para.text = "普通文本"
        run = MagicMock()
        run.text = "普通文本"
        para.runs = [run]
        
        validator._check_runs_in_paragraph(para, "paragraph 0")
        
        assert len(validator.all_placeholders) == 0
    
    @patch('template_validator.Document')
    def test_empty_runs(self, mock_doc_class):
        """测试空的 runs"""
        mock_doc = MagicMock()
        mock_doc_class.return_value = mock_doc
        
        validator = TemplateValidator("template.docx")
        
        para = MagicMock()
        para.text = ""
        para.runs = []
        
        validator._check_runs_in_paragraph(para, "paragraph 0")
        
        assert len(validator.all_placeholders) == 0


class TestValidate:
    """validate 方法测试"""
    
    @patch('template_validator.Document')
    def test_validate_without_config(self, mock_doc_class, mock_document):
        """测试无配置的验证"""
        mock_doc_class.return_value = mock_document
        
        validator = TemplateValidator("template.docx")
        report = validator.validate()
        
        assert report['total_placeholders'] == 4  # test_field, split_field, field1, field2
        assert report['split_count'] == 1  # split_field
        assert report['single_run_count'] == 3
        assert report['undefined_count'] is None  # 无配置
    
    @patch('template_validator.Document')
    def test_validate_with_config(self, mock_doc_class, mock_document, mock_config):
        """测试有配置的验证"""
        mock_doc_class.return_value = mock_document
        
        validator = TemplateValidator("template.docx", "config.json")
        validator.config = mock_config
        report = validator.validate()
        
        assert report['total_placeholders'] == 4
        assert report['split_count'] == 1
        assert report['single_run_count'] == 3
        # split_field 不在配置中，所以有一个未定义的占位符
        assert report['undefined_count'] == 1  # split_field 不在配置中
    
    @patch('template_validator.Document')
    def test_validate_with_undefined_placeholders(self, mock_doc_class):
        """测试有未定义占位符的情况"""
        doc = MagicMock()
        
        # 模拟段落 - 包含未配置的占位符
        para = MagicMock()
        para.text = "{{undefined_field}}"
        run = MagicMock()
        run.text = "{{undefined_field}}"
        para.runs = [run]
        
        doc.paragraphs = [para]
        doc.tables = []
        
        mock_doc_class.return_value = doc
        
        config = {
            "field_mappings": [
                {"template_field": "defined_field"}
            ]
        }
        
        validator = TemplateValidator("template.docx", "config.json")
        validator.config = config
        report = validator.validate()
        
        assert report['undefined_count'] == 1
        assert "undefined_field" in validator.undefined_placeholders
    
    @patch('template_validator.Document')
    def test_validate_with_table_placeholders(self, mock_doc_class, mock_document_with_tables, mock_config):
        """测试表格中的占位符"""
        mock_doc_class.return_value = mock_document_with_tables
        
        validator = TemplateValidator("template.docx", "config.json")
        validator.config = mock_config
        report = validator.validate()
        
        assert "table_field" in validator.all_placeholders
        assert report['total_placeholders'] == 1


class TestGenerateReport:
    """_generate_report 方法测试"""
    
    @patch('template_validator.Document')
    def test_generate_report_with_issues(self, mock_doc_class):
        """测试生成有问题的报告"""
        mock_doc = MagicMock()
        mock_doc_class.return_value = mock_doc
        
        validator = TemplateValidator("template.docx", "config.json")
        # 设置配置以确保 undefined_count 被正确计算
        validator.config = {"field_mappings": [{"template_field": "field1"}, {"template_field": "field2"}]}
        validator.all_placeholders = {"field1", "field2", "field3"}
        validator.split_placeholders = [
            {
                'placeholder': 'field1',
                'location': 'paragraph 0',
                'runs': ['{{', 'field1', '}}'],
                'text': '{{field1}}'
            }
        ]
        validator.single_run_placeholders = [
            {'placeholder': 'field2', 'location': 'paragraph 1', 'text': '{{field2}}'}
        ]
        validator.undefined_placeholders = ['field3']
        
        report = validator._generate_report()
        
        assert report['total_placeholders'] == 3
        assert report['split_count'] == 1
        assert report['single_run_count'] == 1
        assert report['undefined_count'] == 1
    
    @patch('template_validator.Document')
    def test_generate_report_no_issues(self, mock_doc_class):
        """测试生成无问题的报告"""
        mock_doc = MagicMock()
        mock_doc_class.return_value = mock_doc
        
        validator = TemplateValidator("template.docx", "config.json")
        validator.config = {"field_mappings": [{"template_field": "field1"}]}
        validator.all_placeholders = {"field1"}
        validator.split_placeholders = []
        validator.single_run_placeholders = [
            {'placeholder': 'field1', 'location': 'paragraph 0', 'text': '{{field1}}'}
        ]
        validator.undefined_placeholders = []
        
        report = validator._generate_report()
        
        assert report['total_placeholders'] == 1
        assert report['split_count'] == 0
        assert report['single_run_count'] == 1
        assert report['undefined_count'] == 0


# =============================================================================
# main 函数测试
# =============================================================================

class TestMain:
    """main 函数单元测试"""
    
    @patch('template_validator.TemplateValidator')
    def test_main_without_config(self, mock_validator_class):
        """测试无配置的主函数"""
        mock_validator = MagicMock()
        mock_validator.validate.return_value = {
            'total_placeholders': 5,
            'split_count': 0,
            'single_run_count': 5,
            'undefined_count': None
        }
        mock_validator_class.return_value = mock_validator
        
        test_args = ['template_validator.py', '--template', 'template.docx']
        with patch.object(sys, 'argv', test_args):
            result = main()
        
        assert result == 0
        mock_validator_class.assert_called_once_with('template.docx', None)
        mock_validator.validate.assert_called_once()
    
    @patch('template_validator.TemplateValidator')
    def test_main_with_config(self, mock_validator_class):
        """测试有配置的主函数"""
        mock_validator = MagicMock()
        mock_validator.validate.return_value = {
            'total_placeholders': 5,
            'split_count': 0,
            'single_run_count': 5,
            'undefined_count': 0
        }
        mock_validator_class.return_value = mock_validator
        
        test_args = [
            'template_validator.py',
            '--template', 'template.docx',
            '--config', 'config.json'
        ]
        with patch.object(sys, 'argv', test_args):
            result = main()
        
        assert result == 0
        mock_validator_class.assert_called_once_with('template.docx', 'config.json')
    
    @patch('template_validator.TemplateValidator')
    def test_main_with_split_placeholders(self, mock_validator_class):
        """测试有分割占位符时返回非零退出码"""
        mock_validator = MagicMock()
        mock_validator.validate.return_value = {
            'total_placeholders': 5,
            'split_count': 2,
            'single_run_count': 3,
            'undefined_count': None
        }
        mock_validator_class.return_value = mock_validator
        
        test_args = ['template_validator.py', '--template', 'template.docx']
        with patch.object(sys, 'argv', test_args):
            result = main()
        
        assert result == 1  # 有分割占位符，返回非零
    
    @patch('template_validator.TemplateValidator')
    def test_main_with_undefined_placeholders(self, mock_validator_class):
        """测试有未定义占位符时返回非零退出码"""
        mock_validator = MagicMock()
        mock_validator.validate.return_value = {
            'total_placeholders': 5,
            'split_count': 0,
            'single_run_count': 5,
            'undefined_count': 2
        }
        mock_validator_class.return_value = mock_validator
        
        test_args = [
            'template_validator.py',
            '--template', 'template.docx',
            '--config', 'config.json'
        ]
        with patch.object(sys, 'argv', test_args):
            result = main()
        
        assert result == 1  # 有未定义占位符，返回非零
