"""
表格数据转换器
支持：跳过列、添加列、计算列、格式化列、列重排序、行过滤、自定义转换
"""

import sys
from typing import List, Dict, Optional, Any
import statistics
import math

# 导入安全评估工具
sys.path.insert(0, str(__file__).rsplit('\\', 2)[0])
from utils.safe_eval import safe_eval_formula, safe_eval_lambda, SafeEvalError
from utils.logging_config import get_logger

# 导入专用转换器
from .custom_transformers import CustomTransformerRegistry

logger = get_logger(__name__)


class TableDataTransformer:
    """表格数据转换器"""
    
    def transform(self, data, transformations, calculated_report=None):
        # print(f"calculated_report in transform: {calculated_report}")
        result = [row[:] for row in data]
    
        # 分离聚合操作和非聚合操作
        agg_configs = []
        other_transforms = []
    
        for transform in transformations:
            if transform.get('type') == 'calculate' and transform.get('operation') in ['average', 'sum', 'max', 'min']:
                agg_configs.append(transform)
            else:
                other_transforms.append(transform)
    
        # 先执行非聚合操作
        for transform in other_transforms:
            result = self._execute_transform(result, transform, calculated_report)
    
        # 统一处理所有聚合操作
        if agg_configs:
            result = self._apply_aggregations(result, agg_configs)
    
        return result
    
    def _apply_aggregations(self, data: List[List[Any]], configs: List[Dict]) -> List[List[Any]]:
        """统一处理所有聚合操作"""
        result = [row[:] for row in data]
        
        if not result:
            return result
        
        # 添加一个聚合行
        agg_row = [''] * len(result[0])
        
        for config in configs:
            column = config.get('column')
            if column is None:
                continue
            operation = config.get('operation')
            decimal = config.get('decimal')
            function = config.get('function', None)
            
            if operation == 'average':
                values = [float(row[column]) for row in result if column < len(row) and self._is_numeric(row[column])]
                if values:
                    agg_value = statistics.mean(values)
                    if function:
                        agg_row[column] = self._apply_function_value(agg_value, function)
                    else:
                        agg_row[column] = self._format_number(agg_value, decimal)
            elif operation == 'sum':
                values = [float(row[column]) for row in result if column < len(row) and self._is_numeric(row[column])]
                if values:
                    agg_value = sum(values)
                    if function:
                        agg_row[column] = self._apply_function_value(agg_value, function)
                    else:
                        agg_row[column] = self._format_number(agg_value, decimal)
            elif operation == 'max':
                values = [float(row[column]) for row in result if column < len(row) and self._is_numeric(row[column])]
                if values:
                    agg_value = max(values)
                    if function:
                        agg_row[column] = self._apply_function_value(agg_value, function)
                    else:
                        agg_row[column] = self._format_number(agg_value, decimal)
            elif operation == 'min':
                values = [float(row[column]) for row in result if column < len(row) and self._is_numeric(row[column])]
                if values:
                    agg_value = min(values)
                    if function:
                        agg_row[column] = self._apply_function_value(agg_value, function)
                    else:
                        agg_row[column] = self._format_number(agg_value, decimal)
        
        result.append(agg_row)
        return result
    
    def _execute_transform(self, data: List[List[Any]], transform: Dict, calculated_report: Optional[Dict] = None) -> List[List[Any]]:
        """执行单个转换操作"""
        transform_type = transform.get('type')
        
        if transform_type == 'skip_columns':
            return self._apply_skip_columns(data, transform)
        elif transform_type == 'add_column':
            return self._apply_add_column(data, transform, calculated_report)
        elif transform_type == 'calculate':
            return self._apply_calculate(data, transform)
        elif transform_type == 'format_column':
            return self._apply_format_column(data, transform)
        elif transform_type == 'reorder':
            return self._apply_reorder(data, transform)
        elif transform_type == 'filter_rows':
            return self._apply_filter_rows(data, transform)
        elif transform_type == 'custom_transform':
            return self._apply_custom_transform(data, transform, calculated_report)
        
        return data

    def _apply_custom_transform(self, data: List[List[Any]], config: Dict, 
                                   calculated_report: Optional[Dict] = None) -> List[List[Any]]:
        """
        应用专用自定义转换器
        
        配置示例:
        {
            "type": "custom_transform",
            "transformer": "photometric_data_transformer",
            "calculate_columns": [...],
            "average_columns": [...],
            ...
        }
        """
        transformer_name = config.get('transformer')
        if not transformer_name:
            return data
        
        try:
            # 从 calculated_report 中提取 extracted_data
            extracted_data = None
            if calculated_report and 'extracted_data' in calculated_report:
                extracted_data = calculated_report['extracted_data']
            
            return CustomTransformerRegistry.transform(
                transformer_name, data, config, extracted_data
            )
        except Exception as e:
            logger.error(f"Custom transform error ({transformer_name}): {e}")
            return data
    
    def _apply_skip_columns(self, data: List[List[Any]], config: Dict) -> List[List[Any]]:
        """
        跳过指定列
        
        配置示例:
        {
            "type": "skip_columns",
            "columns": [0, 1]  # 跳过第0列和第1列
        }
        """
        columns = config.get('columns', [])
        
        if not columns:
            return data
        
        result = []
        for row in data:
            filtered_row = [val for idx, val in enumerate(row) if idx not in columns]
            result.append(filtered_row)
        
        return result
    
    def _apply_add_column(self, data: List[List[Any]], config: Dict, calculated_report: Optional[Dict] = None) -> List[List[Any]]:
        """
        添加列
        
        配置示例:
        {
            "type": "add_column",
            "position": 0,  # 插入位置
            "source": "row_index"  # 或 "metadata:model_name" 或 "value:固定值"
        }
        """
        position = config.get('position', 0)
        source = config.get('source', '')
        logger.debug(f"calculated_report in add_column: {calculated_report}")
        result = []
        for row_idx, row in enumerate(data):
            if source == 'row_index':
                value = str(row_idx + 1)
            elif source.startswith('metadata:'):
                # 从 metadata 中获取字段值
                key = source.split(':', 1)[1]
                if calculated_report and 'metadata' in calculated_report:
                    metadata = calculated_report['metadata']
                    if isinstance(metadata, dict):
                        value = metadata.get(key, '')
                    else:
                        value = ''
                else:
                    value = ''
            elif source.startswith('extracted_data:'):
                # 从 extracted_data 中获取字段值
                key = source.split(':', 1)[1]
                if calculated_report and 'extracted_data' in calculated_report:
                    extracted_data = calculated_report['extracted_data']
                    if isinstance(extracted_data, dict):
                        value = extracted_data.get(key, '')
                    else:
                        value = ''
                else:
                    value = ''
            elif source.startswith('calculated_data:'):
                # 从 calculated_data 中获取字段值
                key = source.split(':', 1)[1]
                if calculated_report and 'calculated_data' in calculated_report:
                    calculated_data = calculated_report['calculated_data']
                    if isinstance(calculated_data, dict):
                        value = calculated_data.get(key, '')
                    else:
                        value = ''
                else:
                    value = ''
            elif source.startswith('value:'):
                value = source.split(':', 1)[1]
            else:
                value = ''
            logger.debug(f"Adding column at position {position} with value: {value}")
            new_row = row.copy()
            if position >= len(new_row):
                # 位置超出当前行长度，直接追加
                new_row.append(value)
            else:
                # 在指定位置插入值
                new_row.insert(position, value)
            result.append(new_row)
        
        return result
    

    
    def _apply_calculate(self, data: List[List[Any]], config: Dict) -> List[List[Any]]:
        """
        计算列
        
        配置示例 - 公式计算:
        {
            "type": "calculate",
            "column": 4,  # 目标列索引
            "operation": "formula=B{row}/A{row}*1000",  # {row}表示当前行号
            "decimal": 1
        }
        
        配置示例 - 插入新列:
        {
            "type": "calculate",
            "column": 4,  # 插入位置
            "operation": "formula=B{row}/A{row}*1000",
            "decimal": 1,
            "insert": true  # 插入新列，不覆盖原有数据
        }
        """
        column = config.get('column')
        operation = config.get('operation', '')
        decimal = config.get('decimal', None)
        function = config.get('function', None)
        insert = config.get('insert', False)  # 是否插入新列
        
        logger.debug(f"Calculating column {column} with operation {operation}, decimal {decimal}, insert {insert}")
        result = [row[:] for row in data]
        
        if operation.startswith('formula='):
            logger.debug(f"Applying formula calculation on column {column} with operation {operation}")
            formula = operation.split('=', 1)[1]
            
            # 如果需要插入新列，先扩展所有行
            if insert and column is not None:
                for row in result:
                    # 在指定位置插入空值，扩展现有行
                    if len(row) < column:
                        # 如果行长度不够，先扩展到 column 长度
                        row.extend([''] * (column - len(row)))
                    # 在指定位置插入空值
                    row.insert(column, '')
            
            for row_idx, row in enumerate(result):
                try:
                    # 构建变量字典
                    variables = {}
                    for col_idx, val in enumerate(row):
                        if self._is_numeric(val):
                            variables[chr(65 + col_idx)] = float(val)  # A=0, B=1, ...
                    
                    formula_exp = self._parse_column_references(formula, row_idx, row)
                    logger.debug(f"Evaluating formula for row {row_idx}: {formula_exp}")
                    
                    # 使用安全评估替代 eval
                    evaluated_value = safe_eval_formula(formula_exp, variables)
                    logger.debug(f"Evaluated value for row {row_idx}, column {column}: {evaluated_value}")
                    
                    # 确保行有足够长度
                    while len(row) <= column:
                        row.append('')
                    
                    row[column] = self._format_number(evaluated_value, decimal)
                except SafeEvalError as e:
                    logger.warning(f"Safe eval error for row {row_idx}: {e}")
                except Exception as e:
                    logger.warning(f"Error calculating row {row_idx}: {e}")
        
        return result
    
    def _apply_function_value(self, value: float, func_str: str) -> str:
        """使用函数格式化单个值"""
        try:
            return safe_eval_lambda(func_str, value)
        except SafeEvalError as e:
            logger.warning(f"Safe eval error for function '{func_str}': {e}")
            return str(value)
        except Exception as e:
            logger.warning(f"Failed to format value {value} with function: {e}")
            return str(value)
    
    def _apply_format_column(self, data: List[List[Any]], config: Dict) -> List[List[Any]]:
        """
        列格式化 - 使用函数式规则
        
        配置示例 - 函数式:
        {
            "type": "format_column",
            "column": 4,
            "function": "lambda x: f'{x:.4f}' if x < 1 else f'{x:.2f}'"
        }
        
        配置示例 - 固定小数位:
        {
            "type": "format_column",
            "column": 3,
            "decimal": 2
        }
        """
        column = config.get('column')
        
        if 'function' in config:
            func_str = config['function']
            return self._apply_function_format(data, column, func_str)
        elif 'decimal' in config:
            decimal = config['decimal']
            return self._apply_fixed_decimal(data, column, decimal)
        
        return data
    
    def _apply_function_format(self, data: List[List[Any]], column: any, func_str: str) -> List[List[Any]]:
        """函数式格式化"""
        result = [row[:] for row in data]
        
        for row in result:
            if column is not None and column < len(row) and self._is_numeric(row[column]):
                value = float(row[column])
                try:
                    row[column] = safe_eval_lambda(func_str, value)
                except SafeEvalError as e:
                    logger.warning(f"Safe eval error for function '{func_str}': {e}")
                    row[column] = str(value)
                except Exception as e:
                    logger.warning(f"Failed to format value {value}: {e}")
                    row[column] = str(value)
        
        return result
    
    def _apply_fixed_decimal(self, data: List[List[Any]], column: any, decimal: int) -> List[List[Any]]:
        """固定小数位格式化"""
        result = [row[:] for row in data]
        for row in result:
            if column < len(row) and self._is_numeric(row[column]):
                value = float(row[column])
                row[column] = f"{value:.{decimal}f}"
                logger.debug(f"Formatted value at column {column}: {row[column]} at decimal {decimal}")
        return result
    
    def _apply_reorder(self, data: List[List[Any]], config: Dict) -> List[List[Any]]:
        """列重排序"""
        order = config.get('order', [])
        return [[row[idx] for idx in order if idx < len(row)] for row in data]
    
    def _apply_filter_rows(self, data: List[List[Any]], config: Dict) -> List[List[Any]]:
        """行过滤"""
        condition = config.get('condition', '')
        if condition == 'remove_empty':
            return [row for row in data if any(str(val).strip() for val in row)]
        elif condition == 'remove_all_empty':
            return [row for row in data if any(val is not None and str(val).strip() for val in row)]
        return data
    
    def _is_numeric(self, value: Any) -> bool:
        """判断值是否为数字"""
        try:
            float(value)
            return True
        except (ValueError, TypeError):
            return False
    
    def _format_number(self, value: float, decimal: Optional[int]) -> str:
        """格式化数字"""
        if decimal is not None:
            return f"{value:.{decimal}f}"
        return str(value)
    
    def _parse_column_references(self, formula: str, row_idx: int, row: List[Any]) -> str:
        """
        解析列引用，如 A{row} -> row[0]
        支持字母列名: A=0列, B=1列, ..., AA=26列, ...
        """
        import re
        pattern = r'([A-Z]+)\{row\}'
        
        def replace_col(match):
            col_letters = match.group(1)
            col_idx = self._letters_to_index(col_letters)
            if col_idx < len(row) and self._is_numeric(row[col_idx]):
                return str(row[col_idx])
            return '0'
        
        return re.sub(pattern, replace_col, formula)
    
    def _letters_to_index(self, letters: str) -> int:
        """
        列字母转索引，如 A->0, B->1, AA->26
        """
        idx = 0
        for char in letters:
            idx = idx * 26 + (ord(char) - ord('A') + 1)
        return idx - 1
