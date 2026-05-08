from antlr4.tree.Tree import Tree
from antlr4 import ParserRuleContext
from antlr4_vba.vbaParser import vbaParser as Parser
from pyvba_interpreter.symbol_table import FunctionDefinition, SymbolTable
from pyvba_interpreter.vba_visitor import VbaVisitor
from typing import Any, TypeVar
from vba_unit.test_fail_exception import TestFailException


T = TypeVar('T', bound='VbaUnitVisitor')


class VbaUnitVisitor(VbaVisitor):

    def __init__(self: T, table: SymbolTable) -> None:
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
                        self.current_line = line_num
                        self.context_changed = False
                        name = self.context[0]
                        mods = self.table.definitions[name]["modules"]
                        extra = mods[self.context[1]]["extra"]["vba_unit"]
                        coverage = extra["coverage"]
                        coverage[line_num - 1] += 1
        # Call the original visit to continue traversal
        return super().visit(tree)
        self.current_line = prev_line

    def visitFunctionDeclaration(                                  # noqa: N802
            self: T,
            ctx: Parser.FunctionDeclarationContext) -> Any:
        # Touch the end function statement
        # If there is an Exit Function statement immediately beore the end,
        # is there a way to prohibit it from being touched...does it matter?
        line_num = ctx.stop.line
        mods = self.table.definitions[self.context[0]]["modules"]
        coverage = mods[self.context[1]]["extra"]["vba_unit"]["coverage"]
        coverage[line_num - 1] += 1

        return super().visitFunctionDeclaration(ctx)

    def visitAssertStatement(                                      # noqa: N802
            self: T,
            ctx: Parser.AssertStatementContext) -> None:
        # If there is a failed assertion within the code under test
        # an error should probably be thrown
        expr = self.visit(ctx.booleanExpression())
        if not expr:
            raise TestFailException()

    def run_function(self: T,
                     defn: FunctionDefinition,
                     args: list[Any]) -> Any:
        prev_line = self.current_line
        if (self.context[0] != defn["project"] or
                self.context[1] != defn["module"] or
                self.context[2] != defn["name"]):
            self.current_line = 0
        result = super().run_function(defn, args)
        self.current_line = prev_line
        return result
