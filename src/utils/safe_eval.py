"""
安全的表达式评估工具
替代 eval() 的安全实现
"""
import ast
import operator
import re
from typing import Any, Callable, Dict, Optional

# 允许的操作符
ALLOWED_OPERATORS = {
    ast.Add: operator.add,
    ast.Sub: operator.sub,
    ast.Mult: operator.mul,
    ast.Div: operator.truediv,
    ast.Pow: operator.pow,
    ast.USub: operator.neg,
    ast.UAdd: operator.pos,
    ast.FloorDiv: operator.floordiv,
    ast.Mod: operator.mod,
}

# 允许的函数
ALLOWED_FUNCTIONS = {
    'abs': abs,
    'round': round,
    'max': max,
    'min': min,
    'sum': sum,
    'len': len,
    'float': float,
    'int': int,
    'str': str,
}


class SafeEvalError(Exception):
    """安全评估错误"""
    pass


def safe_eval_formula(formula: str, variables: Optional[Dict[str, Any]] = None) -> Any:
    """
    安全地评估数学公式表达式
    
    Args:
        formula: 数学表达式字符串，如 "A / B * 1000"
        variables: 变量字典，如 {"A": 10, "B": 2}
    
    Returns:
        计算结果
    
    Raises:
        SafeEvalError: 如果表达式不安全或无法评估
    """
    if not formula or not isinstance(formula, str):
        raise SafeEvalError("Formula must be a non-empty string")
    
    variables = variables or {}
    
    try:
        tree = ast.parse(formula, mode='eval')
    except SyntaxError as e:
        raise SafeEvalError(f"Invalid syntax: {e}")
    
    def _eval_node(node):
        if isinstance(node, ast.Num):  # Python 3.7
            return node.n
        elif isinstance(node, ast.Constant):  # Python 3.8+
            if isinstance(node.value, (int, float)):
                return node.value
            raise SafeEvalError(f"Unsupported constant type: {type(node.value)}")
        elif isinstance(node, ast.Name):
            if node.id in variables:
                return variables[node.id]
            elif node.id in ALLOWED_FUNCTIONS:
                return ALLOWED_FUNCTIONS[node.id]
            raise SafeEvalError(f"Undefined variable or function: {node.id}")
        elif isinstance(node, ast.BinOp):
            op_type = type(node.op)
            if op_type not in ALLOWED_OPERATORS:
                raise SafeEvalError(f"Unsupported binary operator: {op_type.__name__}")
            left = _eval_node(node.left)
            right = _eval_node(node.right)
            return ALLOWED_OPERATORS[op_type](left, right)
        elif isinstance(node, ast.UnaryOp):
            op_type = type(node.op)
            if op_type not in ALLOWED_OPERATORS:
                raise SafeEvalError(f"Unsupported unary operator: {op_type.__name__}")
            operand = _eval_node(node.operand)
            return ALLOWED_OPERATORS[op_type](operand)
        elif isinstance(node, ast.Call):
            if not isinstance(node.func, ast.Name):
                raise SafeEvalError("Only simple function calls are allowed")
            func_name = node.func.id
            if func_name not in ALLOWED_FUNCTIONS:
                raise SafeEvalError(f"Function not allowed: {func_name}")
            args = [_eval_node(arg) for arg in node.args]
            return ALLOWED_FUNCTIONS[func_name](*args)
        elif isinstance(node, ast.Expression):
            return _eval_node(node.body)
        else:
            raise SafeEvalError(f"Unsupported node type: {type(node).__name__}")
    
    try:
        return _eval_node(tree)
    except ZeroDivisionError:
        return 0.0
    except Exception as e:
        raise SafeEvalError(f"Evaluation error: {e}")


def safe_eval_lambda(func_str: str, value: Any) -> str:
    """
    安全地评估 lambda 格式化函数
    
    限制:
    - 只允许 lambda x: ... 格式
    - 只允许 f-string 格式化
    - 不允许函数调用（除内置允许的外）
    - 不允许属性访问（如 .format）
    
    Args:
        func_str: lambda 函数字符串，如 "lambda x: f'{x:.2f}'"
        value: 要格式化的值
    
    Returns:
        格式化后的字符串
    
    Raises:
        SafeEvalError: 如果函数不安全
    """
    if not func_str or not isinstance(func_str, str):
        raise SafeEvalError("Function must be a non-empty string")
    
    func_str = func_str.strip()
    
    # 验证是 lambda 表达式
    if not func_str.startswith('lambda'):
        raise SafeEvalError("Only lambda expressions are allowed")
    
    # 提取参数名和函数体
    # lambda x: expression
    match = re.match(r'lambda\s+(\w+)\s*:\s*(.+)', func_str, re.DOTALL)
    if not match:
        raise SafeEvalError("Invalid lambda syntax")
    
    param_name = match.group(1)
    body = match.group(2).strip()
    
    # 验证函数体只包含 f-string 或简单的条件表达式
    try:
        # 尝试解析函数体
        body_tree = ast.parse(body, mode='eval')
    except SyntaxError as e:
        raise SafeEvalError(f"Invalid syntax in lambda body: {e}")
    
    def _validate_node(node, depth=0):
        """递归验证 AST 节点是否安全"""
        if depth > 10:
            raise SafeEvalError("Expression too complex")
        
        if isinstance(node, ast.JoinedStr):  # f-string
            for value_node in node.values:
                if isinstance(value_node, ast.FormattedValue):
                    _validate_node(value_node.value, depth + 1)
                    # 验证 format_spec
                    if value_node.format_spec:
                        for spec_val in value_node.format_spec.values:
                            if isinstance(spec_val, ast.Constant):
                                # 只允许简单的格式规范如 .2f, .4f 等
                                if not re.match(r'^\.?\d*[dfge%]$', str(spec_val.value)):
                                    raise SafeEvalError(f"Unsafe format spec: {spec_val.value}")
        elif isinstance(node, ast.Constant):
            pass  # 字符串常量安全
        elif isinstance(node, ast.Name):
            if node.id != param_name and node.id not in ALLOWED_FUNCTIONS:
                raise SafeEvalError(f"Undefined name: {node.id}")
        elif isinstance(node, ast.BinOp):
            if type(node.op) not in ALLOWED_OPERATORS:
                raise SafeEvalError(f"Unsupported operator: {type(node.op).__name__}")
            _validate_node(node.left, depth + 1)
            _validate_node(node.right, depth + 1)
        elif isinstance(node, ast.UnaryOp):
            if type(node.op) not in ALLOWED_OPERATORS:
                raise SafeEvalError(f"Unsupported unary operator: {type(node.op).__name__}")
            _validate_node(node.operand, depth + 1)
        elif isinstance(node, ast.Compare):
            _validate_node(node.left, depth + 1)
            for comparator in node.comparators:
                _validate_node(comparator, depth + 1)
            # 只允许简单的比较操作符
            for op in node.ops:
                if not isinstance(op, (ast.Lt, ast.LtE, ast.Gt, ast.GtE, ast.Eq, ast.NotEq)):
                    raise SafeEvalError(f"Unsupported comparison: {type(op).__name__}")
        elif isinstance(node, ast.IfExp):  # 条件表达式: a if b else c
            _validate_node(node.test, depth + 1)
            _validate_node(node.body, depth + 1)
            _validate_node(node.orelse, depth + 1)
        elif isinstance(node, ast.Call):
            # 只允许 abs() 等安全函数
            if isinstance(node.func, ast.Name):
                if node.func.id not in ALLOWED_FUNCTIONS:
                    raise SafeEvalError(f"Function call not allowed: {node.func.id}")
                for arg in node.args:
                    _validate_node(arg, depth + 1)
            else:
                raise SafeEvalError("Complex function calls not allowed")
        elif isinstance(node, ast.Expression):
            _validate_node(node.body, depth + 1)
        else:
            raise SafeEvalError(f"Unsupported node type: {type(node).__name__}")
    
    # 验证函数体
    _validate_node(body_tree)
    
    # 安全评估：使用受限的 globals 和 locals
    try:
        # 创建安全的执行环境
        safe_globals = {"__builtins__": {}}
        safe_locals = {param_name: value}
        
        # 编译并执行 lambda
        func = eval(compile(body_tree, '<string>', 'eval'), safe_globals, safe_locals)
        return str(func)
    except Exception as e:
        raise SafeEvalError(f"Function execution failed: {e}")


def validate_lambda_safety(func_str: str) -> bool:
    """
    验证 lambda 表达式是否安全（不执行）
    
    Returns:
        True if safe, False otherwise
    """
    try:
        safe_eval_lambda(func_str, 1.0)
        return True
    except SafeEvalError:
        return False
