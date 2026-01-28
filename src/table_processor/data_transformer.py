"""
表格数据转换器
支持：跳过列、添加列、计算列、格式化列、列重排序、行过滤
"""

from typing import List, Dict, Optional, Any
import statistics
import math

class TableDataTransformer:
    """表格数据转换器"""
    
    def transform(self, 
                  data: List[List[Any]], 
                  transformations: List[Dict],
                  metadata: Optional[Dict] = None,
                  targets_data: Optional[Dict] = None) -> List[List[Any]]:
        """
        对原始数据进行转换
        
        Args:
            data: 原始二维数组数据
            transformations: 转换规则列表
            metadata: 元数据（用于metadata类型的数据源）
        
        Returns:
            转换后的数据
        """
        result = [row[:] for row in data]
        for transform in transformations:
            transform_type = transform.get('type')
            
            if transform_type == 'skip_columns':
                result = self._apply_skip_columns(result, transform)
            elif transform_type == 'add_column':
                result = self._apply_add_column(result, transform, metadata, targets_data)
            elif transform_type == 'calculate':
                result = self._apply_calculate(result, transform)
            elif transform_type == 'format_column':
                result = self._apply_format_column(result, transform)
            elif transform_type == 'reorder':
                result = self._apply_reorder(result, transform)
            elif transform_type == 'filter_rows':
                result = self._apply_filter_rows(result, transform)
        
        return result
    
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
    
    def _apply_add_column(self, data: List[List[Any]], config: Dict, metadata: Dict, targets_data: Optional[Dict] = None) -> List[List[Any]]:
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
                    # print(f"Looking for metadata key: {key}")
                    fields=metadata["fields"]
                    # print(f"Metadata fields: {fields}")
                    for field in fields:
                        # print(f"Checking field: {field} for {key}")
                        if field.get("name")==key:
                            # print(f"Found metadata field for key: {key}")
                            value=field.get("value","")
                            # print(f"Adding column at position {position} with value: {value}")
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
        
        配置示例 - 行聚合:
        {
            "type": "calculate",
            "column": 5,
            "operation": "average",  # 或 "sum", "max", "min"
            "decimal": 2
        }
        """
        column = config.get('column')
        operation = config.get('operation', '')
        decimal = config.get('decimal', None)
        print(f"Calculating column {column} with operation {operation} and decimal {decimal}")
        result = [row[:] for row in data]
        
        if operation == 'average':
            values = [float(row[column]) for row in data if column < len(row) and self._is_numeric(row[column])]
            if values and result:
                avg_value = statistics.mean(values)
                result[-1][column] = self._format_number(avg_value, decimal)
        
        elif operation == 'sum':
            values = [float(row[column]) for row in data if column < len(row) and self._is_numeric(row[column])]
            if values and result:
                sum_value = sum(values)
                result[-1][column] = self._format_number(sum_value, decimal)
        
        elif operation == 'max':
            values = [float(row[column]) for row in data if column < len(row) and self._is_numeric(row[column])]
            if values and result:
                max_value = max(values)
                result[-1][column] = self._format_number(max_value, decimal)
        
        elif operation == 'min':
            values = [float(row[column]) for row in data if column < len(row) and self._is_numeric(row[column])]
            if values and result:
                min_value = min(values)
                result[-1][column] = self._format_number(min_value, decimal)
        
        elif operation.startswith('formula='):
            print(f"Applying formula calculation on column {column} with operation {operation}")
            formula = operation.split('=', 1)[1]
            for row_idx, row in enumerate(result):
                try:
                    formula_exp = formula.replace('{row}', str(row_idx))
                    print(f"Evaluating formula for row {row_idx}: {formula_exp}")
                    formula_exp = self._parse_column_references(formula_exp, row_idx, row)
                    
                    evaluated_value = eval(formula_exp)
                    print(f"Evaluated value for row {row_idx}, column {column}: {evaluated_value}")
                    row[column] = self._format_number(evaluated_value, decimal)
                except:
                    pass
        
        return result
    
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
    
    def _apply_function_format(self, data: List[List[Any]], column: int, func_str: str) -> List[List[Any]]:
        """函数式格式化"""
        result = [row[:] for row in data]
        
        try:
            format_func = eval(func_str)
        except Exception as e:
            print(f"Failed to create format function: {e}")
            return data
        
        for row in result:
            if column < len(row) and self._is_numeric(row[column]):
                value = float(row[column])
                try:
                    row[column] = format_func(value)
                except Exception as e:
                    print(f"Failed to format value {value}: {e}")
                    row[column] = str(value)
        
        return result
    
    def _apply_fixed_decimal(self, data: List[List[Any]], column: int, decimal: int) -> List[List[Any]]:
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
