import os
from typing import TypeVar
from vba_unit.Coverage.git_repo import GitRepo


T = Typevar('T', bound='Github')


class Coveralls(GitRepo):
    def __init__(self: T) -> None:
        self.job_id = os.environ['GITHUB_RUN_ID']
        self.commit_sha = commit_sha = os.environ.get('GITHUB_SHA')
