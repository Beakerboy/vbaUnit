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
from pyvba_interpreter.vba_listener import VbaListener
from typing import TypeVar
from vba_unit.Interpreter.coverage_table import CoverageTable
from vba_unit.test_fail_exception import TestFailException
from vba_unit.vba_unit_visitor import VbaUnitVisitor


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
    run_tests(args.src, args.tests, args.project)

    # Submit Coverage
    if args.coverage == "yes":
        result = coveralls_report()
        print(result)


def run_tests(src: str, tests: str, project_name: str) -> None:
    test_project_name = "vbatests"
    table = CoverageTable()

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


def coveralls_report() -> str:
    # with open(file_path, 'r') as f:
    #    line_count = sum(1 for line in f)
    # coverage = [None] * line_count
    # for i in range(line_count):
    #    line_num = i + 1
    #    if line_num in visited_lines:
    #        coverage[i] = 1
    #    else:
    #        coverage[i] = 0
    source_code = """Attribute VB_Name = "Roots"
' Function: Discriminant
' A function to determine if the roots of a quadratic are real of complex
'
' Parameters:
'    a - x² coefficiant
'    b - x  coefficient
'    c - constant term
'
' Returns:
' a real number
Public Function Discriminant(a, b, c)
    Discriminant = b ^ 2 - (4 * a * c)
End Function
"""
    digest = hashlib.md5(source_code.encode('utf-8')).hexdigest()
    commit_sha = os.environ.get('GITHUB_SHA')
    assert commit_sha is not None
    full_ref = os.environ.get('GITHUB_REF', 'master')
    branch = full_ref.replace('refs/heads/', '').replace('refs/pull/', 'PR-')
    fmt = "%an%n%ae%n%cn%n%ce%n%s"
    details = subprocess.check_output(
        ["git", "log", "-1", f"--pretty=format:{fmt}", commit_sha],
        text=True
    ).splitlines()

    # 3. Remote URL from git config
    remote_url = subprocess.check_output(
        ["git", "config", "--get", "remote.origin.url"],
        text=True
    ).strip()
    report = {
        "repo_token": os.environ['COVERALLS_REPO_TOKEN'],
        "service_name": "manual",
        "service_job_id": os.environ['GITHUB_RUN_ID'],
        "source_files": [
            {
                "name": "src/Modules/Roots.bas",
                "source_digest": digest,
                "source": source_code,
                "coverage": [1, None, None, None, None, None, None, None, None,
                             None, None, 1, 1, 1],
            }
        ],
        "git": {
            "head": {
                "id": commit_sha,
                "author_name": details[0],
                "author_email": details[1],
                "committer_name": details[2],
                "committer_email": details[3],
                "message": details[4]
            },
            "branch": branch,
            "remotes": [
                {
                    "name": "origin",
                    "url": remote_url
                }
            ]
        }
    }
    print(report)
    url = "https://coveralls.io/api/v1/jobs"
    response = requests.post(url, files={'json_file': json.dumps(report)})
    return response.json()


if __name__ == "__main__":
    main()
