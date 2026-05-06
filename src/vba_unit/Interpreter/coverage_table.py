from pyvba_interpreter.symbol_table import (
    FunctionDefinition, ModuleDefinition, ProjectDefinition, SymbolTable
)


class VbaUnitFuncDef(FunctionDefinition):
    visited: bool
    start_end_lines: tuple[int, int]


class VbaUnitModDef(ModuleDefinition):
    path: str                              # The file path
    cover: bool                            # Track coverage on this file?
    coverage: list[None | int]             # lines covered
    functions: dict[str, VbaUnitFuncDef]   # {@inheritDoc}


class VbaUnitProjDef(ProjectDefinition):
    modules: dict[str, VbaUnitModDef]


class CoverageTable(SymbolTable):
    pass
