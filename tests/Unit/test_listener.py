from antlr4 import FileStream, CommonTokenStream, ParseTreeWalker
from antlr4_vba.vbaLexer import vbaLexer
from antlr4_vba.vbaParser import vbaParser
from pyvba_interpreter.symbol_table import SymbolTable
from vba_unit.Interpreter.vba_unit_listener import VbaUnitListener


def test_listener() -> None:
    table = SymbolTable()
    input_stream = FileStream(
        "tests/src/VbaProject/Module1.bas",
        encoding="cp1252"
    )
    lexer = vbaLexer(input_stream)
    ts = CommonTokenStream(lexer)
    parser = vbaParser(ts)
    tree = parser.module()
    listener = VbaUnitListener("vbaproject", table)
    listener.parser = parser
    walker = ParseTreeWalker()
    walker.walk(listener, tree)

    assert table.definitions["vbaproject"]["modules"]["module1"]["cover"]

    result = table.definitions["vbaproject"]["modules"]["module1"]["coverage"]
    expected = [1, 0, None, 0, 0, None]
    assert result == expected
