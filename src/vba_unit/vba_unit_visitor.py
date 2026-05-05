from antlr4 import ParseTree
from antlr4_vba.vbaParser import vbaParser as Parser
from pyvba_interpreter.vba_visitor import VbaVisitor
from typing import TypeVar
from vba_unit.test_fail_exception import TestFailException


T = TypeVar('T', bound='VbaUnitVisitor')


class VbaUnitVisitor(VbaVisitor):

    def __init__(self: T, table: SymbolTable) -> None:
        self.visited_lines = set()
        self.source_lines = source_lines # List of strings from the VBA file
        super().__init__(table)
    
    def visit(self: T, tree: ParseTree):
        if tree is not None:
            # Get the starting line number from the context
            # ANTLR line numbers are typically 1-indexed
            line_num = tree.start.line
            self.visited_lines.add(line_num)
        
        # Call the original visit to continue traversal
        return super().visit(tree)

    def visitAssertStatement(                                      # noqa: N802
            self: T,
            ctx: Parser.AssertStatementContext) -> None:
        expr = self.visit(ctx.booleanExpression())
        if not expr:
            raise TestFailException()
