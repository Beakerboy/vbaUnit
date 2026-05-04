import argparse
import glob
import os
from antlr4 import FileStream, CommonTokenStream, ParseTreeWalker
from antlr4_vba.vbaLexer import vbaLexer
from antlr4_vba.vbaParser import vbaParser
from pyvba_interpreter.symbol_table import FunctionType, SymbolTable
from pyvba_interpreter.vba_listener import VbaListener
from pyvba_interpreter.vba_visitor import VbaVisitor
from typing import TypeVar


T = TypeVar('T', bound='TestResult')


class TestFailException(Exception):
    pass


class TestResult:
    def __init__(self: T, name: str) -> None:
        self.name = name
        self.passed = False
        self.error = ""


class Debug:
    @staticmethod
    def vba_assert(expression: bool) -> None:
        if not expression:
            raise TestFailException()


def main() -> None:
    parser = argparse.ArgumentParser(description="VBA ANTLR Test Runner")
    parser.add_argument(
        "--src",
        type=str,
        default="./src",
        help="The path to your project."
    )
    parser.add_argument(
        "--tests",
        type=str,
        default="./tests",
        help="The path to the test files"
    )
    parser.add_argument(
        "--project",
        type=str,
        default="VbaProject",
        help="The name of the project"
    )

    args = parser.parse_args()
    run_tests(args)


def run_tests(args: "Namespace") -> None:
    project_name = args.project
    test_project_name = "vbatests"
    table = SymbolTable()

    # 1. Parse source code
    src_pattern = os.path.join(args.src, '*', '*.bas')
    src_files = glob.glob(src_pattern)
    for file_path in src_files:
        _parse_file(file_path, project_name, table)

    # 2. Parse test code
    test_pattern = os.path.join(args.tests, '*', '*.bas')
    test_files = glob.glob(test_pattern)
    for file_path in test_files:
        _parse_file(file_path, test_project_name, table)

    # 3. Setup Visitor
    visitor = VbaVisitor(table)

    # Override Assert
    visitor.table.library_definitions["special"] = {
        "name": "special",
        "type": FunctionType.PROJECT,
        "modules": {
            "debug": {
                "name": "debug",
                "type": FunctionType.MODULE,
                "functions": {
                    "assert": {
                        "name": "assert",
                        "type": FunctionType.SUB,
                        "handle": getattr(Debug, "vba_assert"),
                        "module": "debug",
                        "params": [{
                            "name": "assertion",
                            "optional": False,
                            "default": ""
                        }]
                    }
                }
            }
        }
    }

    # 4. Find and Execute Tests
    test_modules = table.definitions[test_project_name]["modules"]

    report = []
    for mod_name, module in test_modules.items():
        if mod_name.startswith("test"):
            for func_name, func in module["functions"].items():
                if func_name.startswith("test"):
                    result = TestResult(f"{mod_name}.{func_name}")
                    try:
                        visitor.run_function(func, [])
                        result.passed = True
                    except TestFailException as e:
                        result.passed = False
                        result.error = str(e)
                    report.append(result)

    # 5. Generate Report
    _generate_report(report)


def _parse_file(file_path: str, project: str, table: SymbolTable) -> None:
    input_stream = FileStream(file_path)
    lexer = vbaLexer(input_stream)
    ts = CommonTokenStream(lexer)
    parser = vbaParser(ts)
    tree = parser.module()
    listener = VbaListener(project, table)
    walker = ParseTreeWalker()
    walker.walk(listener, tree)


def _generate_report(results: list) -> None:
    print("\n--- VBA Test Report ---")
    passed = 0
    for r in results:
        status = "PASS" if r.passed else f"FAIL: {r.error}"
        print(f"{r.name}: {status}")
        if r.passed:
            passed += 1
    print(f"-----------------------\nSummary: {passed}/{len(results)} passed.")


if __name__ == "__main__":
    main()
