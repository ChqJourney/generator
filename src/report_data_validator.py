"""
Report Data 完整验证器
验证 report.json 数据是否符合数据结构要求和程序接口要求
"""

import json
import sys
import io
from typing import Dict, Any, List, Optional, Set, Tuple, Union
from pathlib import Path
from dataclasses import dataclass, field
from enum import Enum

# 修复 Windows 控制台编码问题
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')


class ValidationLevel(Enum):
    """验证级别"""
    ERROR = "error"      # 错误：会阻止处理
    WARNING = "warning"  # 警告：建议修复但可继续
    INFO = "info"        # 信息：仅供参考


@dataclass
class ValidationIssue:
    """验证问题详情"""
    level: ValidationLevel
    message: str
    path: str = ""                    # 字段路径
    suggestion: str = ""              # 修复建议
    current_value: Any = None         # 当前值
    expected_type: str = ""           # 期望类型
    actual_type: str = ""             # 实际类型


@dataclass
class ValidationReport:
    """完整验证报告"""
    is_valid: bool = True
    errors: List[ValidationIssue] = field(default_factory=list)
    warnings: List[ValidationIssue] = field(default_factory=list)
    infos: List[ValidationIssue] = field(default_factory=list)
    summary: Dict[str, Any] = field(default_factory=dict)
    
    def add_error(self, message: str, path: str = "", suggestion: str = "", 
                  current_value=None, expected_type: str = "", actual_type: str = ""):
        """添加错误"""
        self.errors.append(ValidationIssue(
            ValidationLevel.ERROR, message, path, suggestion, 
            current_value, expected_type, actual_type
        ))
        self.is_valid = False
    
    def add_warning(self, message: str, path: str = "", suggestion: str = "",
                    current_value=None, expected_type: str = "", actual_type: str = ""):
        """添加警告"""
        self.warnings.append(ValidationIssue(
            ValidationLevel.WARNING, message, path, suggestion,
            current_value, expected_type, actual_type
        ))
    
    def add_info(self, message: str, path: str = "", suggestion: str = ""):
        """添加信息"""
        self.infos.append(ValidationIssue(
            ValidationLevel.INFO, message, path, suggestion
        ))
    
    def print_report(self):
        """打印详细验证报告"""
        print("\n" + "=" * 70)
        print("📋 Report Data 完整验证报告")
        print("=" * 70)
        
        # 摘要信息
        if self.summary:
            print("\n📊 数据摘要:")
            for key, value in self.summary.items():
                print(f"  • {key}: {value}")
        
        # 错误
        if self.errors:
            print(f"\n❌ [错误] 发现 {len(self.errors)} 个错误（必须修复）:")
            for i, err in enumerate(self.errors, 1):
                print(f"\n  {i}. [{err.path}] {err.message}")
                if err.expected_type and err.actual_type:
                    print(f"     类型不匹配: 期望 {err.expected_type}, 实际是 {err.actual_type}")
                if err.current_value is not None:
                    print(f"     当前值: {err.current_value}")
                if err.suggestion:
                    print(f"     💡 建议: {err.suggestion}")
        
        # 警告
        if self.warnings:
            print(f"\n⚠️  [警告] 发现 {len(self.warnings)} 个警告（建议修复）:")
            for i, warn in enumerate(self.warnings, 1):
                print(f"\n  {i}. [{warn.path}] {warn.message}")
                if warn.suggestion:
                    print(f"     💡 建议: {warn.suggestion}")
        
        # 信息
        if self.infos:
            print(f"\nℹ️  [信息] {len(self.infos)} 条提示:")
            for i, info in enumerate(self.infos, 1):
                print(f"  {i}. [{info.path}] {info.message}")
        
        # 总结
        print("\n" + "=" * 70)
        if self.is_valid and not self.warnings:
            print("✅ [通过] 所有检查通过！数据格式完全符合要求。")
        elif self.is_valid:
            print("⚠️  [通过但有警告] 格式基本正确，但建议修复上述警告。")
        else:
            print("❌ [失败] 数据有错误，请先修复后再处理。")
        print("=" * 70 + "\n")


class ReportDataValidator:
    """
    Report Data 完整验证器
    
    验证内容包括：
    1. 基本结构验证 - 必需字段是否存在
    2. 数据类型验证 - 字段类型是否正确
    3. 配置一致性验证 - 数据与配置是否匹配
    4. 表格数据验证 - 表格格式是否正确
    5. 图像路径验证 - 图像文件是否存在
    6. 推荐字段检查 - 推荐字段是否齐全
    """
    
    # 必需的一级字段
    REQUIRED_TOP_LEVEL_FIELDS = ['metadata', 'extracted_data', 'calculated_data']
    
    # metadata 中必需的字段
    REQUIRED_METADATA_FIELDS = ['report_no', 'issue_date', 'applicant_name']
    
    # metadata 中推荐的字段
    RECOMMENDED_METADATA_FIELDS = [
        'report_no', 'issue_date', 'applicant_name', 'product_name',
        'manufacturer', 'test_period', 'applicant_address'
    ]
    
    # extracted_data 中推荐的字段
    RECOMMENDED_EXTRACTED_FIELDS = [
        'model_identifier', 'rated_wattage', 'useful_luminous_flux'
    ]
    
    # 表格数据类型字段的特殊结构
    TABLE_TYPE_FIELDS = ['photometric_data', 'long_term_table', 'beam_data', 
                         'eei_data', 'life_test_data', 'zone_data']
    
    def __init__(self, report_data: Dict, config_data: Optional[Dict] = None,
                 base_path: Optional[Path] = None):
        """
        初始化验证器
        
        Args:
            report_data: 报告数据
            config_data: 配置数据（可选，用于验证配置一致性）
            base_path: 基础路径（用于验证图像文件是否存在）
        """
        self.report_data = report_data
        self.config_data = config_data or {}
        self.base_path = base_path or Path.cwd()
        self.report = ValidationReport()
        self.referenced_paths: Set[str] = set()
        
    def validate(self) -> ValidationReport:
        """
        执行完整验证
        
        Returns:
            ValidationReport: 验证报告
        """
        # 1. 基本结构验证
        self._validate_basic_structure()
        
        # 2. 数据类型验证
        self._validate_data_types()
        
        # 3. 字段内容验证
        self._validate_field_contents()
        
        # 4. 表格数据验证
        self._validate_table_data()
        
        # 5. 图像路径验证
        self._validate_image_paths()
        
        # 6. 配置一致性验证（如果提供了配置）
        if self.config_data:
            self._validate_config_consistency()
        
        # 7. 生成摘要
        self._generate_summary()
        
        return self.report
    
    def _validate_basic_structure(self):
        """验证基本结构"""
        # 检查 report_data 是否为字典
        if not isinstance(self.report_data, dict):
            self.report.add_error(
                "Report 数据必须是 JSON 对象（字典）",
                "root",
                "请确保输入文件是有效的 JSON 对象"
            )
            return
        
        # 检查必需的顶级字段
        for field_name in self.REQUIRED_TOP_LEVEL_FIELDS:
            if field_name not in self.report_data:
                self.report.add_error(
                    f"缺少必需的顶级字段: {field_name}",
                    "root",
                    f"请添加 '{field_name}' 字段"
                )
        
        # 检查未知的顶级字段
        for field_name in self.report_data.keys():
            if field_name not in self.REQUIRED_TOP_LEVEL_FIELDS:
                self.report.add_warning(
                    f"未知的顶级字段: {field_name}",
                    "root",
                    f"'{field_name}' 不是标准字段，可能是拼写错误或多余字段"
                )
    
    def _validate_data_types(self):
        """验证数据类型"""
        # 验证 metadata
        if 'metadata' in self.report_data:
            metadata = self.report_data['metadata']
            if not isinstance(metadata, dict):
                self.report.add_error(
                    "metadata 必须是对象（字典）",
                    "metadata",
                    "请将 metadata 改为字典格式",
                    metadata,
                    "dict",
                    type(metadata).__name__
                )
            elif not metadata:
                self.report.add_warning(
                    "metadata 为空",
                    "metadata",
                    "建议至少添加 report_no, issue_date, applicant_name 等基本信息"
                )
            else:
                # 验证必需字段
                for field in self.REQUIRED_METADATA_FIELDS:
                    if field not in metadata:
                        self.report.add_error(
                            f"metadata 缺少必需字段: {field}",
                            f"metadata.{field}",
                            f"请添加 '{field}' 字段"
                        )
        
        # 验证 extracted_data
        if 'extracted_data' in self.report_data:
            extracted = self.report_data['extracted_data']
            if not isinstance(extracted, dict):
                self.report.add_error(
                    "extracted_data 必须是对象（字典）",
                    "extracted_data",
                    "请将 extracted_data 改为字典格式",
                    extracted,
                    "dict",
                    type(extracted).__name__
                )
            elif not extracted:
                self.report.add_warning(
                    "extracted_data 为空",
                    "extracted_data",
                    "建议添加 model_identifier, rated_wattage 等提取数据"
                )
        
        # 验证 calculated_data
        if 'calculated_data' in self.report_data:
            calculated = self.report_data['calculated_data']
            if not isinstance(calculated, dict):
                self.report.add_error(
                    "calculated_data 必须是对象（字典）",
                    "calculated_data",
                    "请将 calculated_data 改为字典格式",
                    calculated,
                    "dict",
                    type(calculated).__name__
                )
    
    def _validate_field_contents(self):
        """验证字段内容"""
        metadata = self.report_data.get('metadata', {})
        extracted = self.report_data.get('extracted_data', {})
        
        # 验证 report_no 格式
        if 'report_no' in metadata:
            report_no = metadata['report_no']
            if not isinstance(report_no, str):
                self.report.add_warning(
                    "report_no 应该是字符串类型",
                    "metadata.report_no",
                    "建议将 report_no 改为字符串格式",
                    report_no,
                    "str",
                    type(report_no).__name__
                )
        
        # 验证 rated_wattage 和 useful_luminous_flux 是否可以转换为数字
        numeric_fields = [
            ('extracted_data.rated_wattage', extracted.get('rated_wattage')),
            ('extracted_data.useful_luminous_flux', extracted.get('useful_luminous_flux')),
        ]
        
        for path, value in numeric_fields:
            if value is not None:
                try:
                    float(value)
                except (ValueError, TypeError):
                    self.report.add_warning(
                        f"{path} 应该可以转换为数值",
                        path,
                        f"当前值 '{value}' 无法转换为数字，请检查格式",
                        value,
                        "numeric",
                        type(value).__name__
                    )
    
    def _validate_table_data(self):
        """验证表格数据格式"""
        extracted = self.report_data.get('extracted_data', {})
        
        for field_name in self.TABLE_TYPE_FIELDS:
            if field_name not in extracted:
                continue
            
            table_data = extracted[field_name]
            path = f"extracted_data.{field_name}"
            
            # 检查是否有 type 和 value 结构（新的标准格式）
            if isinstance(table_data, dict):
                if 'type' not in table_data:
                    self.report.add_warning(
                        f"{path} 缺少 'type' 字段",
                        f"{path}.type",
                        "建议添加 'type': 'table' 明确数据类型"
                    )
                elif table_data.get('type') != 'table':
                    self.report.add_warning(
                        f"{path}.type 应该是 'table'",
                        f"{path}.type",
                        f"建议将 type 改为 'table'，当前是 '{table_data.get('type')}'"
                    )
                
                if 'value' not in table_data:
                    self.report.add_error(
                        f"{path} 缺少 'value' 字段",
                        f"{path}.value",
                        "请添加 'value' 字段包含表格数据"
                    )
                    continue
                
                table_rows = table_data.get('value', [])
            else:
                # 旧的直接列表格式，仍然支持但给出警告
                table_rows = table_data
                self.report.add_info(
                    f"{path} 使用旧的表格格式（直接列表）",
                    path,
                    "建议使用新的格式: {'type': 'table', 'value': [...]}"
                )
            
            # 验证表格数据结构
            if not isinstance(table_rows, list):
                self.report.add_error(
                    f"{path} 表格数据必须是列表",
                    path,
                    "请将表格数据改为列表格式",
                    table_rows,
                    "list",
                    type(table_rows).__name__
                )
                continue
            
            if len(table_rows) == 0:
                self.report.add_warning(
                    f"{path} 表格数据为空",
                    path,
                    "表格没有数据行"
                )
                continue
            
            # 验证第一行是表头
            if not isinstance(table_rows[0], list):
                self.report.add_error(
                    f"{path} 表格第一行（表头）必须是列表",
                    f"{path}[0]",
                    "请将表头改为列表格式",
                    table_rows[0],
                    "list",
                    type(table_rows[0]).__name__
                )
                continue
            
            header_cols = len(table_rows[0])
            
            # 验证数据行与表头列数一致
            for i, row in enumerate(table_rows[1:], start=1):
                if not isinstance(row, list):
                    self.report.add_error(
                        f"{path} 第 {i+1} 行必须是列表",
                        f"{path}[{i}]",
                        f"请将第 {i+1} 行改为列表格式",
                        row,
                        "list",
                        type(row).__name__
                    )
                    continue
                
                if len(row) != header_cols:
                    self.report.add_warning(
                        f"{path} 第 {i+1} 行列数与表头不一致",
                        f"{path}[{i}]",
                        f"表头有 {header_cols} 列，当前行有 {len(row)} 列",
                        f"列数: {len(row)}",
                        f"{header_cols} 列",
                        f"{len(row)} 列"
                    )
    
    def _validate_image_paths(self):
        """验证图像路径"""
        extracted = self.report_data.get('extracted_data', {})
        images = extracted.get('images', [])
        
        if not images:
            return
        
        if not isinstance(images, list):
            self.report.add_error(
                "images 必须是列表",
                "extracted_data.images",
                "请将 images 改为列表格式",
                images,
                "list",
                type(images).__name__
            )
            return
        
        for i, img_path in enumerate(images):
            if not isinstance(img_path, str):
                self.report.add_warning(
                    f"images[{i}] 必须是字符串路径",
                    f"extracted_data.images[{i}]",
                    "请将图像路径改为字符串",
                    img_path,
                    "str",
                    type(img_path).__name__
                )
                continue
            
            # 检查文件是否存在（相对路径）
            full_path = self.base_path / img_path
            if not full_path.exists():
                self.report.add_warning(
                    f"图像文件不存在: {img_path}",
                    f"extracted_data.images[{i}]",
                    f"请确认文件路径正确，或在运行时从其他位置提供图像",
                    img_path
                )
    
    def _validate_config_consistency(self):
        """验证数据与配置的一致性"""
        field_mappings = self.config_data.get('field_mappings', [])
        
        # 收集配置中引用的字段
        config_fields = {}  # {source_field: mapping_info}
        
        for mapping in field_mappings:
            source_field = mapping.get('source_field')
            if source_field:
                config_fields[source_field] = mapping
            
            # 也收集 args 中的路径
            for arg_path in mapping.get('args', []):
                if isinstance(arg_path, str):
                    config_fields[arg_path] = mapping
        
        # 检查配置引用的字段在数据中是否存在
        for field_path in config_fields:
            parts = field_path.split('.')
            
            # 跳过旧格式的路径（没有点号分隔）
            if len(parts) < 2:
                self.report.add_warning(
                    f"配置中的路径格式不标准，建议使用 'section.field' 格式: {field_path}",
                    f"config.field_mappings",
                    f"请将 '{field_path}' 改为 'extracted_data.{field_path}' 或 'metadata.{field_path}'"
                )
                continue
            
            section = parts[0]
            field = '.'.join(parts[1:])
            
            # 检查 section 是否存在
            if section not in self.report_data:
                self.report.add_error(
                    f"配置引用的 section 不存在: {section}",
                    f"config: {field_path}",
                    f"请确保 report.json 中有 '{section}' 字段"
                )
                continue
            
            section_data = self.report_data[section]
            if not isinstance(section_data, dict):
                continue
            
            # 处理嵌套路径（如 photometric_data.value）
            current = section_data
            field_parts = field.split('.')
            
            for i, part in enumerate(field_parts):
                if isinstance(current, dict) and part in current:
                    current = current[part]
                else:
                    missing_path = f"{section}.{'/'.join(field_parts[:i+1])}"
                    
                    if section == 'calculated_data':
                        # calculated_data 的字段可能由计算生成，给出信息提示
                        self.report.add_info(
                            f"calculated_data 字段可能由计算生成: {missing_path}",
                            f"config: {field_path}",
                            "如果此字段由 calculator 计算生成，可以忽略此提示"
                        )
                    else:
                        self.report.add_error(
                            f"配置引用的字段在数据中不存在: {missing_path}",
                            f"config: {field_path}",
                            f"请在 report.json 的 '{section}' 中添加 '{part}' 字段"
                        )
                    break
        
        # 检查数据中是否有未被配置引用的字段
        for section in ['metadata', 'extracted_data']:
            section_data = self.report_data.get(section, {})
            if not isinstance(section_data, dict):
                continue
            
            for field in section_data.keys():
                field_path = f"{section}.{field}"
                if field_path not in config_fields:
                    # 排除表格数据内部字段
                    if field not in ['photometric_data', 'long_term_table']:
                        self.report.add_info(
                            f"数据中的字段未被配置引用: {field_path}",
                            field_path,
                            "如果不需要此字段，可以删除；如果需要，请在 report_config.json 中添加映射"
                        )
    
    def _generate_summary(self):
        """生成数据摘要"""
        metadata = self.report_data.get('metadata', {})
        extracted = self.report_data.get('extracted_data', {})
        calculated = self.report_data.get('calculated_data', {})
        
        summary = {
            "报告编号": metadata.get('report_no', 'N/A'),
            "申请人": metadata.get('applicant_name', 'N/A'),
            "产品名称": metadata.get('product_name', 'N/A'),
            "型号标识": extracted.get('model_identifier', 'N/A'),
            "额定功率": extracted.get('rated_wattage', 'N/A'),
            "光通量": extracted.get('useful_luminous_flux', 'N/A'),
            "metadata 字段数": len(metadata) if isinstance(metadata, dict) else 0,
            "extracted_data 字段数": len(extracted) if isinstance(extracted, dict) else 0,
            "calculated_data 字段数": len(calculated) if isinstance(calculated, dict) else 0,
        }
        
        # 统计表格数据
        table_count = 0
        for field in self.TABLE_TYPE_FIELDS:
            if field in extracted:
                table_count += 1
        summary["表格数据数量"] = table_count
        
        # 统计图像
        images = extracted.get('images', [])
        summary["图像数量"] = len(images) if isinstance(images, list) else 0
        
        self.report.summary = summary
    
    def get_available_fields(self) -> Dict[str, List[str]]:
        """
        获取可用的字段列表
        
        Returns:
            Dict[str, List[str]]: 各 section 中的字段列表
        """
        available = {}
        for section in self.REQUIRED_TOP_LEVEL_FIELDS:
            section_data = self.report_data.get(section, {})
            if isinstance(section_data, dict):
                available[section] = list(section_data.keys())
        return available
    
    def get_missing_config_mappings(self) -> List[str]:
        """
        获取数据中缺少配置映射的字段
        
        Returns:
            List[str]: 缺少映射的字段路径列表
        """
        if not self.config_data:
            return []
        
        field_mappings = self.config_data.get('field_mappings', [])
        config_fields = set()
        
        for mapping in field_mappings:
            source_field = mapping.get('source_field', '')
            if '.' in source_field:
                config_fields.add(source_field)
        
        missing = []
        for section in ['metadata', 'extracted_data']:
            section_data = self.report_data.get(section, {})
            if not isinstance(section_data, dict):
                continue
            
            for field in section_data.keys():
                field_path = f"{section}.{field}"
                if field_path not in config_fields:
                    missing.append(field_path)
        
        return missing


def load_json(path: str) -> Dict:
    """加载 JSON/JSONC 文件（支持带注释的 JSON）"""
    from utils.jsonc_utils import load_json as load_jsonc
    return load_jsonc(path)


def main():
    """命令行入口"""
    import argparse
    
    parser = argparse.ArgumentParser(
        description='验证 report.json 数据格式和配置一致性'
    )
    parser.add_argument(
        '--report',
        required=True,
        help='Path to report.json'
    )
    parser.add_argument(
        '--config',
        help='Path to report_config.json (optional, 用于验证配置一致性)'
    )
    parser.add_argument(
        '--base-path',
        help='基础路径（用于验证图像文件，默认为当前目录）'
    )
    parser.add_argument(
        '--strict',
        action='store_true',
        help='严格模式：有警告时返回非零退出码'
    )
    
    args = parser.parse_args()
    
    try:
        # 加载 report.json
        try:
            report_data = load_json(args.report)
        except FileNotFoundError:
            print(f"❌ 错误: 文件不存在 - {args.report}")
            return 1
        except json.JSONDecodeError as e:
            print(f"❌ 错误: JSON 格式错误 - {e}")
            return 1
        
        # 加载 config（可选）
        config_data = None
        if args.config:
            try:
                config_data = load_json(args.config)
            except FileNotFoundError:
                print(f"⚠️ 警告: 配置文件不存在 - {args.config}")
            except json.JSONDecodeError as e:
                print(f"⚠️ 警告: 配置文件 JSON 格式错误 - {e}")
        
        # 确定基础路径
        base_path = Path(args.base_path) if args.base_path else Path(args.report).parent
        
        # 执行验证
        validator = ReportDataValidator(report_data, config_data, base_path)
        report = validator.validate()
        
        # 打印报告
        report.print_report()
        
        # 返回退出码
        if not report.is_valid:
            return 1
        if args.strict and report.warnings:
            return 2
        return 0
        
    except Exception as e:
        print(f"❌ 错误: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == '__main__':
    sys.exit(main())
