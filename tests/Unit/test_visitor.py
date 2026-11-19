from antlr4 import FileStream, CommonTokenStream, ParseTreeWalker
from antlr4_vba.vbaLexer import vbaLexer
from antlr4_vba.vbaParser import vbaParser
from pyvba_interpreter.symbol_table import SymbolTable
from vba_unit.Interpreter.vba_unit_listener import VbaUnitListener
from vba_unit.Interpreter.vba_unit_visitor import VbaUnitVisitor


def test_visitor() -> None:
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
    visitor = VbaUnitVisitor(table)
    mod = table.definitions["vbaproject"]["modules"]["module1"]
    func = mod["functions"]["foo"]
    assert mod["extra"]["vba_unit"]["coverage"][3] == 0
    visitor.run_function(func, [])
    assert mod["extra"]["vba_unit"]["coverage"][3] == 1
