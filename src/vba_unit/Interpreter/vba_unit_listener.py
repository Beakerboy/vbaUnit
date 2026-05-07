from antlr4_vba.vbaParser import vbaParser as Parser
from antlr4_vba.vbaLexer import vbaLexer as Lexer
from pyvba_interpreter.vba_listener import VbaListener
from typing import TypeVar


T = TypeVar('T', bound='VbaUnitListener')


class VbaUnitListener(VbaListener):

    def enterProceduralModuleHeader(                               # noqa: N802
            self: T,
            ctx: Parser.ProceduralModuleHeaderContext) -> None:
        super().enterProceduralModuleHeader(ctx)
        token_stream = self.parser.getInputStream()
        token_stream.fill()
        eof_token = token_stream.get(len(token_stream.tokens) - 1)
        total_lines = eof_token.line
        mods = self.table.definitions[self.project_name]["modules"]
        mods[self.module_name.lower()]["coverage"] = [0] * total_lines
        mods[self.module_name.lower()]["cover"] = True
        mods[self.module_name.lower()]["coverage"][total_lines - 1] = None
        mods[self.module_name.lower()]["coverage"][ctx.start.line] = 1

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
        in_str = self.parser.getInputStream()
        if ctx.start is not None:
            if (ctx.start.column == 0 or
                    in_str.get(ctx.start.tokenIndex - 1).column == 0):
                # Comments cannot be the first token in a file, so
                # tokenIndex - 1 cannot be less than zero
                index_num = ctx.start.line - 1
                mods = self.table.definitions[self.project_name]["modules"]
                mods[self.module_name.lower()]["coverage"][index_num] = None

    def enterEndOfLine(                                            # noqa: N802
            self: T,
            ctx: Parser.EndOfLineContext) -> None:
        super().enterEndOfLine(ctx)
        in_str = self.parser.getInputStream()
        if ctx.start is not None:
            tok_ind = ctx.start.tokenIndex
            is_wsc = in_str.get(tok_ind - 1).type == Lexer.WS
            if (
                    ctx.start.column == 0 or
                    (is_wsc and (in_str.get(tok_ind - 1).column == 0))
            ):
                index_num = ctx.start.line - 1
                mods = self.table.definitions[self.project_name]["modules"]
                mods[self.module_name.lower()]["coverage"][index_num] = None
