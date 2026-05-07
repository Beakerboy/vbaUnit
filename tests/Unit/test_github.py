import os
import subprocess
from unittest import mock
from vba_unit.Coverage.github import Github


@mock.patch.dict(os.environ, {
    "GITHUB_RUN_ID": "25396149145",
    "GITHUB_SHA": "123",
    "GITHUB_REF": "master"
})
def test_construct() -> None:
    github = Github()
    assert github.job_id == "25396149145"
    
def test_subprocess() -> None:
    commit_sha = os.environ.get('GITHUB_SHA', '')
    fmt = "%an%n%ae%n%cn%n%ce%n%s"
    message = subprocess.check_output(
            ["git", "log", "-1", f"--pretty=format:{fmt}", commit_sha],
            text=True)
    assert message == ""
