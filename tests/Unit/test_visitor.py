from antlr4 import FileStream, CommonTokenStream, ParseTreeWalker
from antlr4_vba.vbaLexer import vbaLexer
from antlr4_vba.vbaParser import vbaParser
from vba_unit.Interpreter.coverage_table import CoverageTable
from vba_unit.Interpreter.vba_unit_listener import VbaUnitListener
from vba_unit.Interpreter.vba_unit_visitor import VbaUnitVisitor


def test_listener() -> None:
    table = CoverageTable()
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
    func = table.definitions["vbaproject"]["modules"]["module1"]["functions"]["foo"]
    visitor.run_function(func, [])
    mod = table.definitions["vbaproject"]["modules"]["vbaproject"]["module1"]
    assert mod["coverage"][3] == 1
