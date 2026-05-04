import glob
from antlr4 import FileStream, CommonTokenStream, ParseTreeWalker
from antlr4_vba.vbaLexer import vbaLexer
from antlr4_vba.vbaParser import vbaParser
from pyvba_interpreter.symbol_table import FunctionType, SymbolTable
from pyvba_interpreter.vba_listener import VbaListener
from pyvba_interpreter.vba_visitor import VbaVisitor
from typing import TypeVar


T = TypeVar('T', bound='TestResult')


class TestResult:
    def __init__(self: T, name: str) -> None:
        self.name = name
        self.passed = False
        self.error = None


class Debug:
    @staticmethod
    def vba_assert(expression: bool) -> None:
        if expression:
            raise Exception()


def run_tests() -> None:
    project_name = "vbaproject"
    test_project_name = "vbatests"
    table = SymbolTable()

    # 1. Parse source code
    src_files = glob.glob('src/project_name/*/*.bas')
    for file_path in src_files:
        _parse_file(file_path, project_name, table)

    # 2. Parse test code
    test_files = glob.glob('testing/*.bas')
    for file_path in test_files:
        _parse_file(file_path, test_project_name, table)

    # 3. Setup Visitor
    visitor = VbaVisitor(table)

    # Override Assert
    visitor.table.library_definitions["vbatests"]["modules"]["debug"] = {
        "name": "interaction",
        "type": FunctionType.MODULE,
        "functions": {
            "assert": {
                "name": "msgbox",
                "type": FunctionType.FUNCTION,
                "handle": getattr(Debug, "vba_assert"),
                "module": "debug"
            }
        }
    }

    # 4. Find and Execute Tests
    test_modules = table.definitions[test_project_name]["modules"]

    report = []
    for mod_name, module in test_modules.items():
        if mod_name.startswith("Test"):
            for func_name, func in module["functions"].items():
                if func_name.startswith("Test"):
                    result = TestResult(f"{mod_name}.{func_name}")
                    try:
                        visitor.current_assertion_passed = True
                        visitor.run_function(func, [])
                        result.passed = True
                    except Exception as e:
                        result.passed = False
                        result.error = str(e)
                    report.append(result)

    # 5. Generate Report
    _generate_report(report)


def _parse_file(file_path, project, table) -> None:
    input_stream = FileStream(file_path)
    lexer = vbaLexer(input_stream)
    ts = CommonTokenStream(lexer)
    parser = vbaParser(ts)
    tree = parser.module()
    listener = VbaListener(project, table)
    walker = ParseTreeWalker()
    walker.walk(listener, tree)


def _generate_report(results) -> None:
    print("\n--- VBA Test Report ---")
    passed = 0
    for r in results:
        status = "PASS" if r.passed else f"FAIL: {r.error}"
        print(f"{r.name}: {status}")
        if r.passed: passed += 1
    print(f"-----------------------\nSummary: {passed}/{len(results)} passed.")


if __name__ == "__main__":
    run_tests()
