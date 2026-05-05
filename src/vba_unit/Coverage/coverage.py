import json
import requests
from typing import TypeVar


T = TypeVar('T', bound='Coverage')


class Coverage():
    def submit_report(self: T) -> str:
        report = self.generate_report()
        response = requests.post(self.endpoint, files={'json_file': json.dumps(report)})
        return response.json()
