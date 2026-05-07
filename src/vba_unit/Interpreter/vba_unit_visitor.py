from antlr4.tree.Tree import Tree
from antlr4 import ParserRuleContext
from antlr4_vba.vbaParser import vbaParser as Parser
from pyvba_interpreter.vba_visitor import VbaVisitor
from typing import Any, TypeVar
from vba_unit.test_fail_exception import TestFailException
from .coverage_table import CoverageTable


T = TypeVar('T', bound='VbaUnitVisitor')


class VbaUnitVisitor(VbaVisitor):

    def __init__(self: T, table: CoverageTable) -> None:
        self.current_line = 0
        super().__init__(table)

    def visit(self: T, tree: Tree) -> Any: 
        if isinstance(tree, ParserRuleContext):
            # Get the starting line number from the context
            # ANTLR line numbers are typically 1-indexed
            tok = tree.start
            prev_line = self.current_line
            if tok is not None:
                mods = self.table.definitions[self.context[0]]["modules"]
                if mods[self.context[1]]["cover"]:
                    line_num = tok.line
                    if line_num != self.current_line:
                        mods = self.table.definitions[self.context[0]]["modules"]
                        mods[self.context[1]]["coverage"][line_num - 1] += 1

        # Call the original visit to continue traversal
        return super().visit(tree)
        self.self.current_line = prev_line

    def visitAssertStatement(                                      # noqa: N802
            self: T,
            ctx: Parser.AssertStatementContext) -> None:
        # If there is a failed assertion within the code under test
        # an error should probably be thrown
        expr = self.visit(ctx.booleanExpression())
        if not expr:
            raise TestFailException()
