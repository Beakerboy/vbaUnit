import os
import subprocess
from unittest import mock
from vba_unit.Coverage.github import Github


@mock.patch.dict(os.environ, {
    "GITHUB_RUN_ID": "25396149145",
    "GITHUB_SHA": "123",
    "GITHUB_REF": "master"
})
@patch("subprocess.check_output")
def test_construct(mock_check_output: str) -> None:
    mock_check_output.return_value = "John Doe\nme@me.com\nGitHub\nnoreply@github.com\nmessage"
    github = Github()
    assert github.job_id == "25396149145"
    assert github.author_name == John Doe
