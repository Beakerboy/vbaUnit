from antlr4_vba.vbaParser import vbaParser as Parser
from pyvba_interpreter.vba_visitor import VbaVisitor
from typing import TypeVar
from vba_unit.test_fail_exception import TestFailException


T = TypeVar('T', bound='VbaUnitVisitor')


class VbaUnitVisitor(VbaVisitor):

    def visitAssertStatement(                                      # noqa: N802
            self: T,
            ctx: Parser.AssertStatementContext) -> None:
        expr = self.visit(ctx.booleanExpression())
        if not expr:
            raise TestFailException()
