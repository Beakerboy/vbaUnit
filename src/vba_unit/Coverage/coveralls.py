import os
import subprocess
from typing import TypeVar
from vba_unit.Coverage.coverage import Coverage
from vba_unit.Coverage.git_repo import GitRepo


T = TypeVar('T', bound='Coveralls')


class Coveralls(Coverage):
    def __init__(self: T, git: GitRepo) -> None:
        self.endpoint = "https://coveralls.io/api/v1/jobs"
        self.git = git

    def coveralls_report(self: T) -> str:
        file_path = "src/Modules/Roots.bas"
        with open(file_path, 'r') as f:
            line_count = sum(1 for line in f)
        with open(file_path, 'r') as f:
            source_code = f.read()
        coverage = [None] * line_count
        for i in range(line_count):
            line_num = i + 1
            if line_num in visited_lines:
                coverage[i] = 1
            else:
                coverage[i] = 0
        digest = hashlib.md5(source_code.encode('utf-8')).hexdigest()
        commit_sha = os.environ.get('GITHUB_SHA')
        assert commit_sha is not None

        # 3. Remote URL from git config
        remote_url = subprocess.check_output(
            ["git", "config", "--get", "remote.origin.url"],
            text=True
        ).strip()
        report = {
            "repo_token": os.environ['COVERALLS_REPO_TOKEN'],
            "service_name": "manual",
            "service_job_id": self.git.job_id,
            "source_files": [
                {
                    "name": file_path,
                    "source_digest": digest,
                    "source": source_code,
                    "coverage": [1, None, None, None, None, None, None, None,
                                 None, None, None, 1, 1, 1],
                }
            ],
            "git": self.git.repo()
        }
        print(report)
