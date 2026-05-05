from pyvba_interpreter.vba_visitor import VbaVisitor
from typing import TypeVar


T = TypeVar('T', boond='VbaUnitVisitor')


class VbaUnitVisitor(VbaVisitor)
    def visit()
        pass

    def visitAssertStatement(
            self: T,
            ctx: Parser.AssertStatementContext) -> None:
        expr = self.visit(ctx.booleanExpression())
