import os
from unittest import mock
from vba_unit.Coverage.github import Github


@mock.patch.dict(os.environ, {"GITHUB_RUN_ID": "123"})
def test_construct() -> None:
    github = Github()
    assert github.job_id == "123"
    
