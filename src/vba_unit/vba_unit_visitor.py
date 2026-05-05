from pyvba_interpreter.vba_visitor import VbaVisitor
from typing import TypeVar
from vba_unit.test_fail_exception import TestFailException


T = TypeVar('T', boond='VbaUnitVisitor')


class VbaUnitVisitor(VbaVisitor)
    def visit()
        pass

    def visitAssertStatement(
            self: T,
            ctx: Parser.AssertStatementContext) -> None:
        expr = self.visit(ctx.booleanExpression())
        if not expr:
            raise TestFailException()
