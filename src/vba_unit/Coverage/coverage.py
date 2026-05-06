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

    def generate_report() -> str:
        raise Exception("Must be implemented by an extending class")

    def submit_report(self: T) -> str:
        report = self.generate_report()
        response = requests.post(
            self.endpoint, files={'json_file': json.dumps(report)})
        return response.json()
        # should either use response.raise_for_status() to raise
        # an exception if 5xx or 4xx errors
        # or inspect response.status_code. Pass response instead
        # of response.json?
