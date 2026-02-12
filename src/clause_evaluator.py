"""
通用条款判定模块
支持基于规则的条款判定，通过配置即可实现复杂的判定逻辑
"""

import re
from typing import Dict, Any, List, Optional
from enum import Enum

from src.calculator import CalculationRegistry
from src.utils.logging_config import get_logger

logger = get_logger(__name__)


class Operator(Enum):
    """支持的运算符"""
    EQUAL = "=="
    NOT_EQUAL = "!="
    GREATER = ">"
    LESS = "<"
    GREATER_EQUAL = ">="
    LESS_EQUAL = "<="
    IN = "in"
    NOT_IN = "not in"
    CONTAINS = "contains"
    IS_NULL = "is null"
    IS_NOT_NULL = "is not null"


class ClauseEvaluatorError(Exception):
    """条款判定错误"""
    pass


class ConditionEvaluator:
    """条件评估器 - 解析和评估条件表达式"""
    
    def __init__(self, params: Dict[str, Any]):
        self.params = params
    
    def evaluate(self, condition: str) -> bool:
        """评估条件表达式"""
        condition = condition.strip()
        
        # 处理括号分组
        if condition.startswith('(') and condition.endswith(')'):
            if self._is_outer_parentheses(condition):
                return self.evaluate(condition[1:-1])
        
        # 分词并解析逻辑运算符
        tokens = self._tokenize(condition)
        
        # 处理 OR (优先级较低)
        for i, token in enumerate(tokens):
            if token.upper() == 'OR' and self._is_top_level(tokens, i):
                return self.evaluate(''.join(tokens[:i])) or self.evaluate(''.join(tokens[i+1:]))
        
        # 处理 AND
        for i, token in enumerate(tokens):
            if token.upper() == 'AND' and self._is_top_level(tokens, i):
                return self.evaluate(''.join(tokens[:i])) and self.evaluate(''.join(tokens[i+1:]))
        
        # 处理 NOT
        if condition.upper().startswith('NOT '):
            return not self.evaluate(condition[4:])
        
        # 基本比较
        return self._evaluate_comparison(condition)
    
    def _is_outer_parentheses(self, expr: str) -> bool:
        """检查是否是外层括号"""
        balance = 0
        for i, char in enumerate(expr):
            if char == '(':
                balance += 1
            elif char == ')':
                balance -= 1
            if balance == 0 and i < len(expr) - 1:
                return False
        return True
    
    def _tokenize(self, expr: str) -> List[str]:
        """将表达式分词"""
        tokens = []
        current = ''
        in_string = False
        string_char = None
        
        for char in expr:
            if not in_string:
                if char in "'\"":
                    in_string = True
                    string_char = char
                    if current.strip():
                        tokens.append(current)
                    current = char
                elif char in '()':
                    if current.strip():
                        tokens.append(current)
                    tokens.append(char)
                    current = ''
                elif char.isspace():
                    if current.strip():
                        tokens.append(current)
                    current = ''
                else:
                    current += char
            else:
                current += char
                if char == string_char and not current.endswith(f'\\{string_char}'):
                    in_string = False
                    tokens.append(current)
                    current = ''
        
        if current.strip():
            tokens.append(current)
        return tokens
    
    def _is_top_level(self, tokens: List[str], index: int) -> bool:
        """检查token是否在顶级（不在括号内）"""
        depth = 0
        for i, token in enumerate(tokens):
            if i == index:
                return depth == 0
            if token == '(':
                depth += 1
            elif token == ')':
                depth -= 1
        return False
    
    def _evaluate_comparison(self, expr: str) -> bool:
        """评估比较表达式"""
        expr = expr.strip()
        
        # IS NULL / IS NOT NULL
        null_match = re.match(r'(.+?)\s+IS\s+(NOT\s+)?NULL$', expr, re.IGNORECASE)
        if null_match:
            value = self._get_value(null_match.group(1).strip())
            return (value is not None) if null_match.group(2) else (value is None)
        
        # IN / NOT IN
        in_match = re.match(r'(.+?)\s+(NOT\s+)?IN\s*\[(.*?)\]$', expr, re.IGNORECASE)
        if in_match:
            target_value = str(self._get_value(in_match.group(1).strip()))
            values = [v.strip().strip('"\'') for v in in_match.group(3).split(',')]
            return target_value not in values if in_match.group(2) else target_value in values
        
        # CONTAINS
        contains_match = re.match(r'(.+?)\s+CONTAINS\s+["\'](.+?)["\']$', expr, re.IGNORECASE)
        if contains_match:
            value = self._get_value(contains_match.group(1).strip())
            return contains_match.group(2) in str(value) if value else False
        
        # 标准比较操作符
        operators = [
            ('>=', Operator.GREATER_EQUAL), ('<=', Operator.LESS_EQUAL),
            ('!=', Operator.NOT_EQUAL), ('==', Operator.EQUAL),
            ('>', Operator.GREATER), ('<', Operator.LESS),
        ]
        
        for op_str, op_enum in operators:
            if op_str in expr:
                parts = expr.split(op_str, 1)
                if len(parts) == 2:
                    return self._compare_values(parts[0].strip(), parts[1].strip(), op_enum)
        
        logger.warning(f"无法解析表达式: {expr}")
        return False
    
    def _get_value(self, expr: str) -> Any:
        """获取表达式的值"""
        expr = expr.strip()
        
        # 字符串字面量
        if (expr.startswith('"') and expr.endswith('"')) or (expr.startswith("'") and expr.endswith("'")):
            return expr[1:-1]
        
        # 布尔值
        if expr.lower() == 'true':
            return True
        if expr.lower() == 'false':
            return False
        
        # 数值
        try:
            return float(expr) if '.' in expr else int(expr)
        except ValueError:
            pass
        
        # 从参数中获取
        if expr in self.params:
            return self.params[expr]
        
        # 尝试从report_data中获取（点号路径）
        try:
            from src.utils.path_navigator import DataNavigator
            value = DataNavigator().get_value(self.params.get('_report_data', {}), expr)
            if value is not None:
                return value
        except:
            pass
        
        return expr
    
    def _compare_values(self, left: str, right: str, operator: Operator) -> bool:
        """比较两个值"""
        left_val = self._get_value(left)
        right_val = self._get_value(right)
        
        # 处理布尔值和字符串转换
        for val in [left_val, right_val]:
            if isinstance(val, str):
                if val.lower() == 'true':
                    val = True
                elif val.lower() == 'false':
                    val = False
        
        # 执行比较
        if operator == Operator.EQUAL:
            return str(left_val) == str(right_val)
        elif operator == Operator.NOT_EQUAL:
            return str(left_val) != str(right_val)
        elif operator in [Operator.GREATER, Operator.GREATER_EQUAL, Operator.LESS, Operator.LESS_EQUAL]:
            try:
                left_num = float(left_val) if left_val is not None else 0
                right_num = float(right_val) if right_val is not None else 0
                if operator == Operator.GREATER:
                    return left_num > right_num
                elif operator == Operator.GREATER_EQUAL:
                    return left_num >= right_num
                elif operator == Operator.LESS:
                    return left_num < right_num
                elif operator == Operator.LESS_EQUAL:
                    return left_num <= right_num
            except (ValueError, TypeError):
                return False
        
        return False


@CalculationRegistry.register("evaluate_clause")
def evaluate_clause(*args, clause_config=None, **kwargs):
    """
    通用条款判定函数
    
    根据配置的条件规则评估条款，返回判定结果 (Pass/Fail/N/A)
    
    Args:
        *args: 位置参数，对应 param_names 中定义的参数
        clause_config: 条款配置字典，包含:
            - param_names: 参数名列表，与 args 一一对应
            - rules: 规则列表，每个规则包含 condition 和 result
            - default: 默认返回值
        **kwargs: 额外关键字参数
    
    Returns:
        str: 判定结果，通常为 "Pass"、"Fail" 或 "N/A"
    """
    if not clause_config:
        logger.warning("evaluate_clause: 未提供 clause_config")
        return "N/A"
    
    # 构建参数字典
    param_names = clause_config.get('param_names', [])
    params = {name: args[i] for i, name in enumerate(param_names) if i < len(args)}
    params.update(kwargs)
    
    # 按顺序评估规则
    evaluator = ConditionEvaluator(params)
    for i, rule in enumerate(clause_config.get('rules', [])):
        condition = rule.get('condition', '')
        result = rule.get('result', 'N/A')
        
        try:
            if evaluator.evaluate(condition):
                logger.info(f"规则 {i} 匹配: {condition} => {result}")
                return result
        except Exception as e:
            logger.error(f"评估规则 {i} 失败: {e}")
            continue
    
    # 返回默认值
    default = clause_config.get('default', 'N/A')
    logger.info(f"无规则匹配，返回默认值: {default}")
    return default
