"""Small, deliberately limited evaluator for numeric configuration formulas."""
import ast
import operator


_BINARY_OPERATORS = {
    ast.Add: operator.add,
    ast.Sub: operator.sub,
    ast.Mult: operator.mul,
    ast.Div: operator.truediv,
}


def evaluate_numeric_expression(expression, variables=None):
    """Evaluate only numeric names, constants, parentheses and + - * /.

    ``variables`` may contain scalars or pandas Series, so the same evaluator
    can be used for both template placeholders and vectorised pivot formulas.
    """
    variables = variables or {}
    tree = ast.parse(expression, mode="eval")

    def visit(node):
        if isinstance(node, ast.Expression):
            return visit(node.body)
        if isinstance(node, ast.Constant) and isinstance(node.value, (int, float)) and not isinstance(node.value, bool):
            return node.value
        if isinstance(node, ast.Name):
            if node.id not in variables:
                raise ValueError(f"不允许的变量: {node.id}")
            return variables[node.id]
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.UAdd, ast.USub)):
            value = visit(node.operand)
            return value if isinstance(node.op, ast.UAdd) else -value
        if isinstance(node, ast.BinOp) and type(node.op) in _BINARY_OPERATORS:
            return _BINARY_OPERATORS[type(node.op)](visit(node.left), visit(node.right))
        raise ValueError("表达式仅支持数字、变量、括号及 + - * /")

    return visit(tree)
