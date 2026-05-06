import json
import requests
from typing import TypeVar
from vba_unit.Coverage.coverage_table import CoverageTable


T = TypeVar('T', bound='Coverage')


class Coverage():
    def __init__(self: T) -> None:
        self.table: CoverageTable

    def submit_report(self: T) -> str:
        report = self.generate_report()
        response = requests.post(
            self.endpoint, files={'json_file': json.dumps(report)})
        return response.json()
