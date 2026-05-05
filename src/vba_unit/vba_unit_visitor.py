from antlr4.tree.Tree import Tree
from antlr4 import ParserRuleContext
from antlr4_vba.vbaParser import vbaParser as Parser
from pyvba_interpreter.symbol_table import SymbolTable
from pyvba_interpreter.vba_visitor import VbaVisitor
from typing import Any, TypeVar
from vba_unit.test_fail_exception import TestFailException


T = TypeVar('T', bound='VbaUnitVisitor')


class VbaUnitVisitor(VbaVisitor):

    def __init__(self: T, table: SymbolTable) -> None:
        self.visited_lines: set[int] = set()
        super().__init__(table)

    def visit(self: T, tree: Tree) -> Any:
        if isinstance(tree, ParserRuleContext):
            # Get the starting line number from the context
            # ANTLR line numbers are typically 1-indexed
            tok = tree.start.line
            if tok is not None:
                line_num = tok.line
                self.visited_lines.add(line_num)

        # Call the original visit to continue traversal
        return super().visit(tree)

    def visitAssertStatement(                                      # noqa: N802
            self: T,
            ctx: Parser.AssertStatementContext) -> None:
        expr = self.visit(ctx.booleanExpression())
        if not expr:
            raise TestFailException()
