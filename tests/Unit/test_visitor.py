from vba_unit.Interpreter.coverage_table import CoverageTable
from vba_unit.Interpreter.vba_unit_listener import VbaUnitVisitor


def test_listener() -> None:
    table = CoverageTable()
    visitor = VbaUnitVisitor(table)
