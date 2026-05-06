from pyvba_interpreter.symbol_table import ModuleDefinition, ProjectDefinition, SymbolTable


class VbaUnitModDef(ModuleDefinition):
    path: str
    coverage: list[None | int]


class VbaUnitProjDef(ProjectDefinition):
    name: str
    type: FunctionType
    modules: dict[str, VbaUnitModDef]


class CoverageTable(SymbolTable):
    pass
