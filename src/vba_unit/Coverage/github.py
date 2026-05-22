import os
import subprocess
from typing import TypeVar
from vba_unit.Coverage.git_repo import GitRepo


T = TypeVar('T', bound='GitRepo')


class Github(GitRepo):
    def __init__(self: T) -> None:
        self.job_id = os.environ['GITHUB_RUN_ID']
        self.commit_sha = os.environ.get('GITHUB_SHA', '')
        full_ref = os.environ.get('GITHUB_REF', 'master')
        self.branch = full_ref.replace(
            'refs/heads/', ''
        ).replace('refs/pull/', 'PR-')
        fmt = "%an%n%ae%n%cn%n%ce%n%s"
        (self.author_name,
         self.author_email,
         self.committer_name,
         self.committer_email,
         self.message) = subprocess.check_output(
            ["git", "log", "-1", f"--pretty=format:{fmt}", self.commit_sha],
            text=True
        ).splitlines()
        self.remote_url = subprocess.check_output(
            ["git", "config", "--get", "remote.origin.url"],
            text=True
        ).strip()
