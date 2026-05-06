import hashlib
import os
from typing import TypeVar
from vba_unit.Coverage.coverage import Coverage
from vba_unit.Coverage.git_repo import GitRepo
from vba_unit.Interpreter.coverage_table import VbaUnitModDef


T = TypeVar('T', bound='Coveralls')


class Coveralls(Coverage):
    def __init__(self: T) -> None:
        self.endpoint = "https://coveralls.io/api/v1/jobs"
        self.git: GitRepo

    def generate_report(self: T) -> dict:
        source_files = []
        for lib in self.table.values():
            for module in lib["modules"].values():
                if module["cover"]:
                    file_cov = self.file_coverage(module)
                    source_files.append([file_cov])

        commit_sha = os.environ.get('GITHUB_SHA')
        assert commit_sha is not None

        report = {
            "repo_token": os.environ['COVERALLS_REPO_TOKEN'],
            "service_name": "manual",
            "service_job_id": self.git.job_id,
            "source_files": source_files,
            "git": self.git.repo()
        }
        return report

    def file_coverage(self: T, module: VbaUnitModDef) -> dict:
        file_path = module["path"]
        with open(file_path, 'r') as f:
            source_code = f.read()
        digest = hashlib.md5(source_code.encode('utf-8')).hexdigest()
        return {
            "name": file_path,
            "source_digest": digest,
            "coverage": module["coverage"],
        }
