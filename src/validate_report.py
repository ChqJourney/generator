"""
Report JSON 格式验证器
用于在计算和转换前验证 report.json 格式是否符合要求
"""
import json
import sys
import re
import io
from typing import Dict, Any, List, Optional, Set, Tuple
from pathlib import Path
from dataclasses import dataclass, field
from enum import Enum
from utils.logging_config import get_logger

logger = get_logger(__name__)

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
class ValidationResult:
    """验证结果"""
    level: ValidationLevel
    message: str
    path: str = ""  # 字段路径或位置


@dataclass
class ValidationReport:
    """验证报告"""
    is_valid: bool = True
    errors: List[ValidationResult] = field(default_factory=list)
    warnings: List[ValidationResult] = field(default_factory=list)
    infos: List[ValidationResult] = field(default_factory=list)
    
    def add_error(self, message: str, path: str = ""):
        self.errors.append(ValidationResult(ValidationLevel.ERROR, message, path))
        self.is_valid = False
    
    def add_warning(self, message: str, path: str = ""):
        self.warnings.append(ValidationResult(ValidationLevel.WARNING, message, path))
    
    def add_info(self, message: str, path: str = ""):
        self.infos.append(ValidationResult(ValidationLevel.INFO, message, path))
    
    def print_report(self):
        """打印验证报告"""
        # 验证报告使用print直接输出到控制台，因为这是用户直接查看的输出
        print("\n" + "=" * 60)
        print("Report JSON 格式验证报告")
        print("=" * 60)
        
        if not self.errors and not self.warnings:
            print("[通过] 所有检查通过！格式正确。\n")
            logger.info("Validation passed: no errors or warnings")
            return
        
        if self.errors:
            print(f"\n[错误] 发现 {len(self.errors)} 个错误：")
            for err in self.errors:
                path_str = f" [{err.path}]" if err.path else ""
                print(f"  - {err.message}{path_str}")
                logger.error(f"Validation error: {err.message}{path_str}")
        
        if self.warnings:
            print(f"\n[警告] 发现 {len(self.warnings)} 个警告：")
            for warn in self.warnings:
                path_str = f" [{warn.path}]" if warn.path else ""
                print(f"  - {warn.message}{path_str}")
                logger.warning(f"Validation warning: {warn.message}{path_str}")
        
        if self.infos:
            print(f"\n[提示] 提示信息 ({len(self.infos)} 条)：")
            for info in self.infos:
                path_str = f" [{info.path}]" if info.path else ""
                print(f"  - {info.message}{path_str}")
                logger.info(f"Validation info: {info.message}{path_str}")
        
        print("\n" + "=" * 60)
        if self.is_valid:
            print("[通过] 格式基本正确，可以处理。")
        else:
            print("[失败] 格式有错误，请先修复后再处理。")
        print("=" * 60 + "\n")


class ReportValidator:
    """Report JSON 验证器"""
    
    # 必需的一级字段
    REQUIRED_TOP_LEVEL_FIELDS = ['metadata', 'extracted_data', 'calculated_data']
    
    # metadata 中推荐的字段
    RECOMMENDED_METADATA_FIELDS = [
        'report_no', 'issue_date', 'applicant_name', 'product_name',
        'manufacturer', 'test_period'
    ]
    
    # extracted_data 中推荐的字段
    RECOMMENDED_EXTRACTED_FIELDS = [
        'model_identifier', 'rated_wattage', 'useful_luminous_flux'
    ]
    
    def __init__(self, report_data: Dict, config_data: Optional[Dict] = None):
        self.report_data = report_data
        self.config_data = config_data or {}
        self.report = ValidationReport()
        self.referenced_paths: Set[str] = set()
    
    def validate(self) -> ValidationReport:
        """
        执行完整验证
        
        Returns:
            ValidationReport: 验证报告
        """
        # 1. 验证基本结构
        self._validate_structure()
        
        # 2. 验证字段类型
        self._validate_field_types()
        
        # 3. 如果提供了配置，验证字段路径
        if self.config_data:
            self._validate_config_paths()
        
        # 4. 验证 recommended 字段
        self._validate_recommended_fields()
        
        return self.report
    
    def _validate_structure(self):
        """验证基本结构"""
        # 检查 report_data 是否为字典
        if not isinstance(self.report_data, dict):
            self.report.add_error("Report 数据必须是 JSON 对象（字典）", "root")
            return
        
        # 检查必需的顶级字段
        for field_name in self.REQUIRED_TOP_LEVEL_FIELDS:
            if field_name not in self.report_data:
                self.report.add_error(f"缺少必需的顶级字段: {field_name}", "root")
        
        # 检查未知的顶级字段
        for field_name in self.report_data.keys():
            if field_name not in self.REQUIRED_TOP_LEVEL_FIELDS:
                self.report.add_warning(f"未知的顶级字段: {field_name}", "root")
    
    def _validate_field_types(self):
        """验证字段类型"""
        # 验证 metadata 是字典
        if 'metadata' in self.report_data:
            metadata = self.report_data['metadata']
            if not isinstance(metadata, dict):
                self.report.add_error("metadata 必须是对象（字典）", "metadata")
            elif not metadata:
                self.report.add_warning("metadata 为空", "metadata")
        
        # 验证 extracted_data 是字典
        if 'extracted_data' in self.report_data:
            extracted = self.report_data['extracted_data']
            if not isinstance(extracted, dict):
                self.report.add_error("extracted_data 必须是对象（字典）", "extracted_data")
            elif not extracted:
                self.report.add_warning("extracted_data 为空", "extracted_data")
        
        # 验证 calculated_data 是字典
        if 'calculated_data' in self.report_data:
            calculated = self.report_data['calculated_data']
            if not isinstance(calculated, dict):
                self.report.add_error("calculated_data 必须是对象（字典）", "calculated_data")
    
    def _validate_config_paths(self):
        """验证配置中引用的所有字段路径"""
        self._collect_referenced_paths()
        
        for path in sorted(self.referenced_paths):
            parts = path.split('.')
            
            # 检查路径格式
            if len(parts) < 2:
                self.report.add_warning(
                    f"路径格式不标准，建议格式为 'section.field': {path}",
                    path
                )
                continue
            
            section = parts[0]
            field = '.'.join(parts[1:])
            
            # 检查 section 是否存在
            if section not in self.report_data:
                self.report.add_error(
                    f"路径引用不存在的 section: {section}",
                    path
                )
                continue
            
            # 检查字段是否存在
            section_data = self.report_data.get(section, {})
            if not isinstance(section_data, dict):
                continue
            
            # 处理嵌套路径
            current = section_data
            field_parts = field.split('.')
            
            for i, part in enumerate(field_parts):
                if isinstance(current, dict) and part in current:
                    current = current[part]
                else:
                    missing_path = f"{section}.{'/'.join(field_parts[:i+1])}"
                    # 只有 calculated_data 的路径可以是可选的
                    if section == 'calculated_data':
                        self.report.add_info(
                            f"calculated_data 字段可能由计算生成: {missing_path}",
                            path
                        )
                    else:
                        self.report.add_error(
                            f"路径引用的字段不存在: {missing_path}",
                            path
                        )
                    break
    
    def _collect_referenced_paths(self):
        """收集配置中引用的所有字段路径"""
        field_mappings = self.config_data.get('field_mappings', [])
        
        for mapping in field_mappings:
            # source_field
            source_field = mapping.get('source_field')
            if source_field:
                self.referenced_paths.add(source_field)
            
            # args 中的路径
            args = mapping.get('args', [])
            for arg_path in args:
                if isinstance(arg_path, str):
                    self.referenced_paths.add(arg_path)
    
    def _validate_recommended_fields(self):
        """验证推荐的字段是否存在"""
        metadata = self.report_data.get('metadata', {})
        extracted = self.report_data.get('extracted_data', {})
        
        for field in self.RECOMMENDED_METADATA_FIELDS:
            if field not in metadata:
                self.report.add_info(
                    f"metadata 缺少推荐的字段: {field}",
                    f"metadata.{field}"
                )
        
        for field in self.RECOMMENDED_EXTRACTED_FIELDS:
            if field not in extracted:
                self.report.add_info(
                    f"extracted_data 缺少推荐的字段: {field}",
                    f"extracted_data.{field}"
                )
    
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


def load_json(path: str) -> Dict:
    """加载 JSON 文件"""
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def main():
    """命令行入口"""
    import argparse
    
    parser = argparse.ArgumentParser(
        description='验证 report.json 格式是否符合要求'
    )
    parser.add_argument(
        '--report',
        required=True,
        help='Path to report.json'
    )
    parser.add_argument(
        '--config',
        help='Path to report_config.json (optional, 用于验证字段路径)'
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
            logger.error(f"文件不存在 - {args.report}")
            return 1
        except json.JSONDecodeError as e:
            logger.error(f"JSON 格式错误 - {e}")
            return 1
        
        # 加载 config（可选）
        config_data = None
        if args.config:
            try:
                config_data = load_json(args.config)
            except FileNotFoundError:
                logger.warning(f"配置文件不存在 - {args.config}")
            except json.JSONDecodeError as e:
                logger.warning(f"配置文件 JSON 格式错误 - {e}")
        
        # 执行验证
        validator = ReportValidator(report_data, config_data)
        report = validator.validate()
        
        # 打印报告
        report.print_report()
        
        # 打印可用字段
        available = validator.get_available_fields()
        print("可用字段列表：")
        for section, fields in available.items():
            if fields:
                print(f"  {section}: {', '.join(fields)}")
            else:
                print(f"  {section}: (空)")
        print()
        
        # 返回退出码
        if not report.is_valid:
            return 1
        if args.strict and report.warnings:
            return 2
        return 0
        
    except Exception as e:
        logger.error(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == '__main__':
    sys.exit(main())
