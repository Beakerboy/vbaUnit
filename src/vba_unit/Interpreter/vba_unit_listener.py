from antlr4_vba.vbaParser import vbaParser as Parser
from pyvba_interpreter.vba_listener import VbaListener
from typing import TypeVar
from .coverage_table import CoverageTable


T = TypeVar('T', bound='VbaUnitListener')


class VbaUnitListener(VbaListener):

    def enterProceduralModuleHeader(                               # noqa: N802
            self: T,
            ctx: Parser.ProceduralModuleHeaderContext) -> None:
        super().enterProceduralModuleHeader(ctx)
        token_stream = self.parser.getInputStream()
        eof_token = token_stream.get(token_stream.size - 1)
        total_lines = eof_token.line
        mod = self.table[project]["modules"][self.module_name.lower()]
        mod["coverage"] = [0] * total_lines

    def enterFunctionDeclaration(                                  # noqa: N802
            self: T,
            ctx: Parser.FunctionDeclarationContext) -> None:
        super().enterFunctionDeclaration(ctx)
        # Add the start line number to the function definition

    def exitFunctionDeclaration(                                   # noqa: N802
            self: T,
            ctx: Parser.FunctionDeclarationContext) -> None:
        super().exitFunctionDeclaration(ctx)
        # Add the end line number to the function definition

    def enterSubroutineDeclaration(                                # noqa: N802
            self: T,
            ctx: Parser.SubroutineDeclarationContext) -> None:
        super().enterSubroutineDeclaration(ctx)
        # Add the start line number to the function definition

    def exitSubroutineDeclaration(                                 # noqa: N802
            self: T,
            ctx: Parser.SubroutineDeclarationContext) -> None:
        super().exitSubroutineDeclaration(ctx)
        # Add the end line number to the function definition

    def enterCommentBody(                                          # noqa: N802
            self: T,
            ctx: Parser.CommentBodyContext) -> None:
        super().enterCommentBody(ctx)
        # Check if the token starts at column 1 or if the token before this is
        # a wsc and it starts at column 1
        if (ctx.start.column == 0 or
                self.parser.getInputStream().get(ctx.start.tokenIndex - 1)):
            ctx.start.column
        # Add Comment Line Number to Coverage array if this line is a comment

    def enterEndOfLine(                                            # noqa: N802
            self: T,
            ctx: Parser.EndOfLineContext) -> None:
        super().enterEndOfLine(ctx)
        # Check if this token starts at column 1, or if the preceeding token
        # is a WSC and it starts at column 1.
        # Add WS Line Number to Coverage array if this line is whitespace
