from pyvba_interpreter.symbol_table import (
    FunctionDefinition, ModuleDefinition, ProjectDefinition, SymbolTable
)


class VbaUnitFuncDef(FunctionDefinition):
    visited: bool
    start_end_lines: tuple[int, int]


class VbaUnitModDef(ModuleDefinition):
    path: str
    coverage: list[None | int]
    functions: dict[str, VbaUnitFuncDef]


class VbaUnitProjDef(ProjectDefinition):
    name: str
    type: FunctionType
    modules: dict[str, VbaUnitModDef]


class CoverageTable(SymbolTable):
    pass
