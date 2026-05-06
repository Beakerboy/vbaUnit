import hashlib
import os
from typing import TypeVar
from vba_unit.Coverage.coverage import Coverage
from vba_unit.Coverage.git_repo import GitRepo


T = TypeVar('T', bound='Coveralls')


class Coveralls(Coverage):
    def __init__(self: T, git: GitRepo) -> None:
        self.endpoint = "https://coveralls.io/api/v1/jobs"
        self.git = git

    def coveralls_report(self: T) -> str:
        file_paths = ["src/Modules/Roots.bas"]
        source_files = []
        for file in file_paths:
            file_cov = self.file_coverage(file)
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
        print(report)

    def file_coverage(self: T, file_path: str) -> dict:
        with open(file_path, 'r') as f:
            line_count = sum(1 for line in f)

        coverage = [None] * line_count
        for i in range(line_count):
            line_num = i + 1
            if line_num in visited_lines:
                coverage[i] = 1
            else:
                coverage[i] = 0
        with open(file_path, 'r') as f:
            source_code = f.read()
        digest = hashlib.md5(source_code.encode('utf-8')).hexdigest()
        return {
            "name": file_path,
            "source_digest": digest,
            "coverage": [1, None, None, None, None, None, None, None,
                         None, None, None, 1, 1, 1],
        }
