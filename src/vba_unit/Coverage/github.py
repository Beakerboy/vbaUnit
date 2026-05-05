import os
from typing import TypeVar
from vba_unit.Coverage.git_repo import GitRepo


T = Typevar('T', bound='Github')


class Coveralls(GitRepo):
    def __init__(self: T) -> None:
        self.job_id = os.environ['GITHUB_RUN_ID']
        self.commit_sha = commit_sha = os.environ.get('GITHUB_SHA')
        full_ref = os.environ.get('GITHUB_REF', 'master')
    branch = full_ref.replace('refs/heads/', '').replace('refs/pull/', 'PR-')
    fmt = "%an%n%ae%n%cn%n%ce%n%s"
    details = subprocess.check_output(
        ["git", "log", "-1", f"--pretty=format:{fmt}", self.commit_sha],
        text=True
    ).splitlines()
