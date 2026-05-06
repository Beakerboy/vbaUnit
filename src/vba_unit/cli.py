import argparse
import glob
import hashlib
import json
import os
import requests
import subprocess
from antlr4 import FileStream, CommonTokenStream, ParseTreeWalker
from antlr4_vba.vbaLexer import vbaLexer
from antlr4_vba.vbaParser import vbaParser
from typing import TypeVar
from vba_unit.Interpreter.coverage_table import CoverageTable
from vba_unit.Interpreter.unit_listener import VbaUnitListener
from vba_unit.Interpreter.vba_unit_visitor import VbaUnitVisitor
from vba_unit.test_fail_exception import TestFailException


T = TypeVar('T', bound='TestResult')


class TestResult:
    def __init__(self: T, name: str) -> None:
        self.name = name
        self.passed = False
        self.error = ""


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
    parser.add_argument(
        "--coverage",
        default="no",
        const="yes",
        nargs="?",
        help="Submit Code Coverage?"
    )

    args = parser.parse_args()
    table = CoverageTable()
    run_tests(args.src, args.tests, args.project)

    # Submit Coverage
    if args.coverage == "yes":
        coverage = Coveralls(table)
        result = coverage.submit_report()
        print(result)


def run_tests(src: str, tests: str, project_name: str, table: CoverageTable) -> None:
    test_project_name = "vbatests"

    # Parse source code
    src_pattern = os.path.join(src, '*', '*.bas')
    src_files = glob.glob(src_pattern)
    for file_path in src_files:
        _parse_file(file_path, project_name, table)

    # Parse test code
    test_pattern = os.path.join(tests, '*.bas')
    test_files = glob.glob(test_pattern)
    for file_path in test_files:
        _parse_file(file_path, test_project_name, table)

    # Setup Visitor
    visitor = VbaUnitVisitor(table)

    # Find and Execute Tests
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

    # Generate Report
    _generate_report(report)


def _parse_file(file_path: str, project: str, table: CoverageTable) -> None:
    input_stream = FileStream(file_path, encoding="cp1252")
    lexer = vbaLexer(input_stream)
    ts = CommonTokenStream(lexer)
    parser = vbaParser(ts)
    tree = parser.module()
    listener = VbaUnitListener(project, table)
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
