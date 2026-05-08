from antlr4 import CommonTokenStream
from antlr4_vba.vbaParser import vbaParser as Parser
from antlr4_vba.vbaLexer import vbaLexer as Lexer
from pyvba_interpreter.symbol_table import SymbolTable
from pyvba_interpreter.vba_listener import VbaListener
from typing import TypedDict, TypeVar


T = TypeVar('T', bound='VbaUnitListener')


class VbaUnitModuleextra(TypedDict):
    path: str                              # The file path
    cover: bool                            # Track coverage on this file?
    coverage: list[None | int]             # lines covered


class VbaUnitListener(VbaListener):
    def __init__(self: T, project: str, table: SymbolTable) -> None:
        self.parser: Parser
        super().__init__(project, table)

    def enterProceduralModuleHeader(                               # noqa: N802
            self: T,
            ctx: Parser.ProceduralModuleHeaderContext) -> None:
        super().enterProceduralModuleHeader(ctx)
        token_stream: CommonTokenStream = self.parser.getInputStream()
        token_stream.fill()
        eof_token = token_stream.get(len(token_stream.tokens) - 1)
        total_lines = eof_token.line
        name = self.module_name.lower()
        mod = self.table.definitions[self.project_name]["modules"][name]
        extra: VbaUnitModuleextra = {
            "coverage": [0] * total_lines,
            "cover": True,
            "path": ''
        }
        mod["extra"]["vba_unit"] = extra
        # EOF line is ignored.
        # Need to test the case where EOF is on the same line as code.
        mod["extra"]["vba_unit"]["coverage"][total_lines - 1] = None
        if ctx.start is not None:
            mod["extra"]["vba_unit"]["coverage"][ctx.start.line - 1] = 1

    def enterCommentBody(                                          # noqa: N802
            self: T,
            ctx: Parser.CommentBodyContext) -> None:
        super().enterCommentBody(ctx)
        # Check if the token starts at column 1 or if the token before this is
        # a wsc and it starts at column 1
        in_str = self.parser.getInputStream()
        if ctx.start is not None and ctx.start.tokenIndex is not None:
            if (ctx.start.column == 0 or
                    in_str.get(ctx.start.tokenIndex - 1).column == 0):
                # Comments cannot be the first token in a file, so
                # tokenIndex - 1 cannot be less than zero
                index_num = ctx.start.line - 1
                name = self.module_name.lower()
                mods = self.table.definitions[self.project_name]["modules"]
                mods[name]["extra"]["vba_unit"]["coverage"][index_num] = None

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
                name = self.module_name.lower()
                mods = self.table.definitions[self.project_name]["modules"]
                mods[name]["extra"]["vba_unit"]["coverage"][index_num] = None
