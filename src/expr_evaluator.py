import ast
import operator as op

# supported operators
operators = {ast.Add: op.add, ast.Sub: op.sub, ast.Mult: op.mul,
             ast.Div: op.truediv, ast.Pow: op.pow, ast.BitXor: op.xor,
             ast.USub: op.neg}

def evaluate_expression(expr):
    try:
        return _eval(ast.parse(expr, mode='eval').body)
    except SyntaxError:
        raise TypeError('fail')

def _eval(node):
    if isinstance(node, ast.Num): # <number>
        return node.n
    elif isinstance(node, ast.BinOp): # <left> <operator> <right>
        return operators[type(node.op)](_eval(node.left), _eval(node.right))
    elif isinstance(node, ast.UnaryOp): # <operator> <operand> e.g., -1
        return operators[type(node.op)](_eval(node.operand))
    else:
        raise TypeError(node)