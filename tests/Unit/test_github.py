import os
import subprocess
from unittest import mock
from vba_unit.Coverage.github import Github


@mock.patch.dict(os.environ, {
    "GITHUB_RUN_ID": "25396149145",
    "GITHUB_SHA": "036c36dfac1d00cb37b6510fc423641cda7b1f08",
    "GITHUB_REF": "refs/pull/6/merge"
})
@mock.patch("subprocess.check_output")
def test_construct(mock_check_output: str) -> None:
    mock_check_output.side_effect = [
        "John Doe\nme@me.com\nGitHub\nnoreply@github.com\ncommit message",
        'https://github.com/Beakerboy/FooProject'
    ]
    github = Github()
    assert github.job_id == "25396149145"
    assert github.author_name == "John Doe"
    assert github.author_email == "me@me.com"
    assert github.committer_name == "GitHub"
    assert github.committer_email == "noreply@github.com"
    assert github.message == "commit message"
    assert github.remote_url == 'https://github.com/Beakerboy/FooProject'
    assert github.branch == "PR-6/merge"
    expected_calls = [
            mock.call(
                ["git", "log", "-1", "--pretty=format:%an%n%ae%n%cn%n%ce%n%s", "036c36dfac1d00cb37b6510fc423641cda7b1f08"],
                text=True
            ),
            mock.call(
                ["git", "config", "--get", "remote.origin.url"],
                text=True
            ),
        ]
    mock_check_output.assert_has_calls(expected_calls)
    expected = {
        'head': {
            'id': '036c36dfac1d00cb37b6510fc423641cda7b1f08',
            'author_name': 'John Doe',
            'author_email': 'me@me.com',
            'committer_name': 'GitHub',
            'committer_email': 'noreply@github.com',
            'message': 'commit message'
        },
        'branch': 'PR-6/merge',
        'remotes': [
            {
                'name': 'origin',
                'url': 'https://github.com/Beakerboy/VBA-Projects'
            }
        ]
    }
    assert github.git() == expected
