import json
import requests
from typing import TypeVar
from .git_repo import GitRepo
from vba_unit.Interpreter.coverage_table import CoverageTable


T = TypeVar('T', bound='Coverage')


class Coverage():
    def __init__(self: T) -> None:
        self.endpoint = ''
        self.table: CoverageTable
        self.git: GitRepo

    def submit_report(self: T) -> str:
        report = self.generate_report()
        response = requests.post(
            self.endpoint, files={'json_file': json.dumps(report)})
        return response.json()
