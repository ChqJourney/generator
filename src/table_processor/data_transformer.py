"""
表格数据转换器
支持：跳过列、添加列、计算列、格式化列、列重排序、行过滤
"""

import sys
from typing import List, Dict, Optional, Any
import statistics
import math

# 导入安全评估工具
sys.path.insert(0, str(__file__).rsplit('\\', 2)[0])
from utils.safe_eval import safe_eval_formula, safe_eval_lambda, SafeEvalError


class TableDataTransformer:
    """表格数据转换器"""
    
    def transform(self, data, transformations, metadata=None, targets_data=None):
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
            result = self._execute_transform(result, transform, metadata, targets_data)
    
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
    
    def _apply_single_aggregation(self, data: List[List[Any]], config: Dict) -> List[List[Any]]:
        """执行单个聚合操作"""
        result = [row[:] for row in data]
        
        if not result:
            return result
        
        column = config.get('column')
        if column is None:
            return result
        
        operation = config.get('operation')
        decimal = config.get('decimal')
        function = config.get('function', None)
        
        values = [float(row[column]) for row in result if column < len(row) and self._is_numeric(row[column])]
        
        if not values:
            return result
        
        if operation == 'average':
            agg_value = statistics.mean(values)
        elif operation == 'sum':
            agg_value = sum(values)
        elif operation == 'max':
            agg_value = max(values)
        elif operation == 'min':
            agg_value = min(values)
        else:
            return result
        
        agg_row = [''] * len(result[0])
        agg_row[0] = 'Average'
        
        if function:
            agg_row[column] = self._apply_function_value(agg_value, function)
        else:
            agg_row[column] = self._format_number(agg_value, decimal)
        
        result.append(agg_row)
        print(f'Added aggregation row for column {column}: {agg_row}')
        return result
    
    def _execute_transform(self, data: List[List[Any]], transform: Dict, metadata: Optional[Dict] = None, targets_data: Optional[Dict] = None) -> List[List[Any]]:
        """执行单个转换操作"""
        transform_type = transform.get('type')
        
        if transform_type == 'skip_columns':
            return self._apply_skip_columns(data, transform)
        elif transform_type == 'add_column':
            return self._apply_add_column(data, transform, metadata, targets_data)
        elif transform_type == 'calculate':
            return self._apply_calculate(data, transform)
        elif transform_type == 'format_column':
            return self._apply_format_column(data, transform)
        elif transform_type == 'reorder':
            return self._apply_reorder(data, transform)
        elif transform_type == 'filter_rows':
            return self._apply_filter_rows(data, transform)
        
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
    
    def _apply_add_column(self, data: List[List[Any]], config: Dict, metadata: Optional[Dict] = None, targets_data: Optional[Dict] = None) -> List[List[Any]]:
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
        print(f"metadata in add_column: {metadata}")
        result = []
        for row_idx, row in enumerate(data):
            if source == 'row_index':
                value = str(row_idx + 1)
            elif source.startswith('metadata:'):
                key = source.split(':', 1)[1]
                if metadata is None or "fields" not in metadata:
                    value = ''
                else:
                    value=''
                    fields=metadata["fields"]
                    for field in fields:
                        if field.get("name")==key:
                            value=field.get("value","")
                            break
                    
            elif source.startswith('targets:'):
                key = source.split(':', 1)[1]
                if targets_data is None or "targets" not in targets_data:
                    value = ''
                else:
                    value = ''
                    targets=targets_data["targets"]
                    for target in targets:
                        if target.get("name")==key:
                            value=target.get("value","")
                            break
            elif source.startswith('value:'):
                value = source.split(':', 1)[1]
            else:
                value = ''
            print(f"Adding column at position {position} with value: {value}")
            new_row = row.copy()
            if position >= len(new_row) and row_idx == 0:
                new_row.append(value)
            elif row_idx == 0 and value != '':
                new_row.insert(position, value)
            else:
                new_row.insert(position, '')
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
        """
        column = config.get('column')
        operation = config.get('operation', '')
        decimal = config.get('decimal', None)
        function = config.get('function', None)
        print(f"Calculating column {column} with operation {operation} and decimal {decimal}")
        result = [row[:] for row in data]
        
        if operation.startswith('formula='):
            print(f"Applying formula calculation on column {column} with operation {operation}")
            formula = operation.split('=', 1)[1]
            for row_idx, row in enumerate(result):
                try:
                    # 构建变量字典
                    variables = {}
                    for col_idx, val in enumerate(row):
                        if self._is_numeric(val):
                            variables[chr(65 + col_idx)] = float(val)  # A=0, B=1, ...
                    
                    formula_exp = self._parse_column_references(formula, row_idx, row)
                    print(f"Evaluating formula for row {row_idx}: {formula_exp}")
                    
                    # 使用安全评估替代 eval
                    evaluated_value = safe_eval_formula(formula_exp, variables)
                    print(f"Evaluated value for row {row_idx}, column {column}: {evaluated_value}")
                    row[column] = self._format_number(evaluated_value, decimal)
                except SafeEvalError as e:
                    print(f"Safe eval error for row {row_idx}: {e}")
                except Exception as e:
                    print(f"Error calculating row {row_idx}: {e}")
        
        return result
    
    def _apply_function_value(self, value: float, func_str: str) -> str:
        """使用函数格式化单个值"""
        try:
            return safe_eval_lambda(func_str, value)
        except SafeEvalError as e:
            print(f"Safe eval error for function '{func_str}': {e}")
            return str(value)
        except Exception as e:
            print(f"Failed to format value {value} with function: {e}")
            return str(value)
    
    def _find_or_add_avg_row(self, data: List[List[Any]]) -> int:
        """查找或添加平均行"""
        for idx, row in enumerate(data):
            if len([cell for cell in row if cell is not None and 'Avg' in str(cell)]) > 0:
                return idx
        
        # 没找到，添加新行
        if data and data[0]:
            avg_row = [''] * len(data[0])
            data.append(avg_row)
            return len(data) - 1
        return 0
    
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
                    print(f"Safe eval error for function '{func_str}': {e}")
                    row[column] = str(value)
                except Exception as e:
                    print(f"Failed to format value {value}: {e}")
                    row[column] = str(value)
        
        return result
    
    def _apply_fixed_decimal(self, data: List[List[Any]], column: any, decimal: int) -> List[List[Any]]:
        """固定小数位格式化"""
        result = [row[:] for row in data]
        for row in result:
            if column < len(row) and self._is_numeric(row[column]):
                value = float(row[column])
                row[column] = f"{value:.{decimal}f}"
                print(f"Formatted value at column {column}: {row[column]} at decimal {decimal}")
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
