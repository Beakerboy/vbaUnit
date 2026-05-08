from pyvba_interpreter.symbol_table import SymbolTable
from typing import TypedDict


class VbaUnitModuleExtras(TypedDict):
    path: str                              # The file path
    cover: bool                            # Track coverage on this file?
    coverage: list[None | int]             # lines covered


class CoverageTable(SymbolTable):
    pass
